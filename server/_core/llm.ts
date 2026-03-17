/**
 * llm.ts — Multi-provider LLM client
 *
 * Routes requests to the correct provider based on the model ID prefix:
 *   claude-*   → Anthropic Messages API
 *   gemini-*   → Google AI (Gemini) via OpenAI-compatible endpoint
 *   gpt-* / o* → OpenAI Chat Completions API
 *
 * Required env vars (provide the key(s) for the model(s) you use):
 *   OPENAI_API_KEY
 *   ANTHROPIC_API_KEY
 *   GOOGLE_AI_API_KEY
 *   DEFAULT_LLM_MODEL  (default: "gemini-2.5-flash")
 */

import { ENV } from "./env";

// ─── Shared types (kept compatible with existing callsites) ──────────────────

export type Role = "system" | "user" | "assistant" | "tool" | "function";

export type TextContent = { type: "text"; text: string };

export type ImageContent = {
  type: "image_url";
  image_url: { url: string; detail?: "auto" | "low" | "high" };
};

export type FileContent = {
  type: "file_url";
  file_url: {
    url: string;
    mime_type?: "audio/mpeg" | "audio/wav" | "application/pdf" | "audio/mp4" | "video/mp4";
  };
};

export type MessageContent = string | TextContent | ImageContent | FileContent;

export type Message = {
  role: Role;
  content: MessageContent | MessageContent[];
  name?: string;
  tool_call_id?: string;
};

export type Tool = {
  type: "function";
  function: {
    name: string;
    description?: string;
    parameters?: Record<string, unknown>;
  };
};

export type ToolChoicePrimitive = "none" | "auto" | "required";
export type ToolChoiceByName = { name: string };
export type ToolChoiceExplicit = { type: "function"; function: { name: string } };
export type ToolChoice = ToolChoicePrimitive | ToolChoiceByName | ToolChoiceExplicit;

export type JsonSchema = {
  name: string;
  schema: Record<string, unknown>;
  strict?: boolean;
};
export type OutputSchema = JsonSchema;
export type ResponseFormat =
  | { type: "text" }
  | { type: "json_object" }
  | { type: "json_schema"; json_schema: JsonSchema };

export type InvokeParams = {
  model?: string;
  messages: Message[];
  tools?: Tool[];
  toolChoice?: ToolChoice;
  tool_choice?: ToolChoice;
  maxTokens?: number;
  max_tokens?: number;
  outputSchema?: OutputSchema;
  output_schema?: OutputSchema;
  responseFormat?: ResponseFormat;
  response_format?: ResponseFormat;
};

export type ToolCall = {
  id: string;
  type: "function";
  function: { name: string; arguments: string };
};

export type InvokeResult = {
  id: string;
  created: number;
  model: string;
  choices: Array<{
    index: number;
    message: {
      role: Role;
      content: string | Array<TextContent | ImageContent | FileContent>;
      tool_calls?: ToolCall[];
    };
    finish_reason: string | null;
  }>;
  usage?: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
};

// ─── Helpers ─────────────────────────────────────────────────────────────────

const ensureArray = (v: MessageContent | MessageContent[]): MessageContent[] =>
  Array.isArray(v) ? v : [v];

/**
 * Fetch the raw bytes of a file_url and return as a base64 data-URL string.
 * Works with GCS signed URLs, public URLs, etc.
 */
async function fileUrlToBase64DataUrl(
  url: string,
  mimeType: string
): Promise<string> {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Failed to fetch file_url ${url}: ${resp.status}`);
  const buf = Buffer.from(await resp.arrayBuffer());
  return `data:${mimeType};base64,${buf.toString("base64")}`;
}

// ─── OpenAI ──────────────────────────────────────────────────────────────────

/**
 * Normalise a MessageContent array to OpenAI's content format.
 * file_url → converted to an image_url with a base64 data URL so that
 * vision-capable OpenAI models (gpt-4o etc.) can read PDFs / images.
 */
async function normaliseContentOpenAI(
  content: MessageContent | MessageContent[]
): Promise<unknown> {
  const parts = ensureArray(content);

  // Single plain string → keep as string for backward compat
  if (parts.length === 1 && typeof parts[0] === "string") return parts[0];
  if (parts.length === 1 && parts[0].type === "text") return (parts[0] as TextContent).text;

  const result: unknown[] = [];
  for (const part of parts) {
    if (typeof part === "string") {
      result.push({ type: "text", text: part });
    } else if (part.type === "text") {
      result.push(part);
    } else if (part.type === "image_url") {
      result.push(part);
    } else if (part.type === "file_url") {
      const mime = part.file_url.mime_type ?? "application/pdf";
      const dataUrl = await fileUrlToBase64DataUrl(part.file_url.url, mime);
      result.push({ type: "image_url", image_url: { url: dataUrl, detail: "high" } });
    }
  }
  return result;
}

async function invokeOpenAI(model: string, params: InvokeParams): Promise<InvokeResult> {
  const apiKey = ENV.openaiApiKey;
  if (!apiKey) throw new Error("OPENAI_API_KEY is not configured.");

  const messages: unknown[] = [];
  for (const msg of params.messages) {
    messages.push({
      role: msg.role,
      content: await normaliseContentOpenAI(msg.content),
      ...(msg.name ? { name: msg.name } : {}),
      ...(msg.tool_call_id ? { tool_call_id: msg.tool_call_id } : {}),
    });
  }

  const maxTokens = params.maxTokens ?? params.max_tokens ?? 8192;

  const payload: Record<string, unknown> = { model, messages, max_tokens: maxTokens };

  if (params.tools && params.tools.length > 0) {
    payload.tools = params.tools;
    const tc = params.toolChoice ?? params.tool_choice;
    if (tc) payload.tool_choice = tc;
  }

  const rf = params.responseFormat ?? params.response_format;
  const os = params.outputSchema ?? params.output_schema;
  if (rf) {
    payload.response_format = rf;
  } else if (os) {
    payload.response_format = { type: "json_schema", json_schema: os };
  }

  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify(payload),
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`OpenAI API error ${resp.status}: ${err}`);
  }

  return (await resp.json()) as InvokeResult;
}

// ─── Gemini (via OpenAI-compatible endpoint) ──────────────────────────────────

async function invokeGemini(model: string, params: InvokeParams): Promise<InvokeResult> {
  const apiKey = ENV.googleAiApiKey;
  if (!apiKey) throw new Error("GOOGLE_AI_API_KEY is not configured.");

  // Gemini's OpenAI-compatible endpoint accepts the same payload shape as OpenAI.
  // PDF / image content is sent as a data URL inside an image_url part.
  const messages: unknown[] = [];
  for (const msg of params.messages) {
    messages.push({
      role: msg.role === "assistant" ? "assistant" : msg.role === "system" ? "system" : "user",
      content: await normaliseContentOpenAI(msg.content),
      ...(msg.name ? { name: msg.name } : {}),
    });
  }

  const maxTokens = params.maxTokens ?? params.max_tokens ?? 8192;

  const payload: Record<string, unknown> = { model, messages, max_tokens: maxTokens };

  const rf = params.responseFormat ?? params.response_format;
  const os = params.outputSchema ?? params.output_schema;
  if (rf) {
    payload.response_format = rf;
  } else if (os) {
    payload.response_format = { type: "json_schema", json_schema: os };
  }

  const resp = await fetch(
    "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions",
    {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify(payload),
    }
  );

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Gemini API error ${resp.status}: ${err}`);
  }

  return (await resp.json()) as InvokeResult;
}

// ─── Anthropic ────────────────────────────────────────────────────────────────

/**
 * Convert a single MessageContent to Anthropic's content block format.
 */
async function toAnthropicContentBlock(part: MessageContent): Promise<unknown> {
  if (typeof part === "string") return { type: "text", text: part };
  if (part.type === "text") return { type: "text", text: part.text };
  if (part.type === "image_url") {
    const url = part.image_url.url;
    if (url.startsWith("data:")) {
      const [header, data] = url.split(",");
      const mimeType = header.replace("data:", "").replace(";base64", "");
      return { type: "image", source: { type: "base64", media_type: mimeType, data } };
    }
    // Remote URL — fetch and embed
    const resp = await fetch(url);
    const buf = Buffer.from(await resp.arrayBuffer());
    const ct = resp.headers.get("content-type") ?? "image/jpeg";
    return { type: "image", source: { type: "base64", media_type: ct, data: buf.toString("base64") } };
  }
  if (part.type === "file_url") {
    // Anthropic supports PDF documents via the document content type
    const mime = part.file_url.mime_type ?? "application/pdf";
    const resp = await fetch(part.file_url.url);
    const buf = Buffer.from(await resp.arrayBuffer());
    return {
      type: "document",
      source: { type: "base64", media_type: mime, data: buf.toString("base64") },
    };
  }
  throw new Error("Unsupported content part type");
}

async function invokeAnthropic(model: string, params: InvokeParams): Promise<InvokeResult> {
  const apiKey = ENV.anthropicApiKey;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY is not configured.");

  // Extract system messages (Anthropic puts system at the top level)
  let systemText = "";
  const conversationMessages: unknown[] = [];

  for (const msg of params.messages) {
    if (msg.role === "system") {
      const parts = ensureArray(msg.content);
      systemText += parts.map(p => (typeof p === "string" ? p : (p as TextContent).text ?? "")).join("\n");
    } else {
      const parts = ensureArray(msg.content);
      const blocks = await Promise.all(parts.map(toAnthropicContentBlock));
      // Anthropic collapses single text blocks to a string
      const content = blocks.length === 1 && (blocks[0] as any).type === "text"
        ? (blocks[0] as any).text
        : blocks;
      conversationMessages.push({ role: msg.role, content });
    }
  }

  const maxTokens = params.maxTokens ?? params.max_tokens ?? 8192;

  const payload: Record<string, unknown> = {
    model,
    messages: conversationMessages,
    max_tokens: maxTokens,
  };
  if (systemText) payload.system = systemText;

  if (params.tools && params.tools.length > 0) {
    payload.tools = params.tools.map(t => ({
      name: t.function.name,
      description: t.function.description,
      input_schema: t.function.parameters ?? { type: "object", properties: {} },
    }));
  }

  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "pdfs-2024-09-25",
    },
    body: JSON.stringify(payload),
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Anthropic API error ${resp.status}: ${err}`);
  }

  // Normalise to InvokeResult shape
  const raw = (await resp.json()) as {
    id: string;
    model: string;
    content: Array<{ type: string; text?: string }>;
    usage?: { input_tokens: number; output_tokens: number };
    stop_reason?: string;
  };

  const textContent = raw.content
    .filter(b => b.type === "text")
    .map(b => b.text ?? "")
    .join("");

  return {
    id: raw.id,
    created: Math.floor(Date.now() / 1000),
    model: raw.model,
    choices: [
      {
        index: 0,
        message: { role: "assistant", content: textContent },
        finish_reason: raw.stop_reason ?? null,
      },
    ],
    usage: raw.usage
      ? {
          prompt_tokens: raw.usage.input_tokens,
          completion_tokens: raw.usage.output_tokens,
          total_tokens: raw.usage.input_tokens + raw.usage.output_tokens,
        }
      : undefined,
  };
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Route an LLM invocation to the correct provider based on the model ID.
 * Falls back to ENV.defaultLlmModel if no model is specified in params.
 */
export async function invokeLLM(params: InvokeParams): Promise<InvokeResult> {
  const model = params.model ?? ENV.defaultLlmModel ?? "gemini-2.5-flash";

  if (model.startsWith("claude-")) {
    return invokeAnthropic(model, params);
  }

  if (model.startsWith("gemini-")) {
    return invokeGemini(model, params);
  }

  // Default: OpenAI (gpt-*, o1-*, etc.)
  return invokeOpenAI(model, params);
}
