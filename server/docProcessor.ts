import { exec } from "child_process";
import { promisify } from "util";
import fs from "fs";
import path from "path";
import os from "os";

const execAsync = promisify(exec);

// ─── Cost constants ──────────────────────────────────────────────────────────
// GPT-4o-mini: $0.15/1M input tokens, $0.60/1M output tokens
export const COST_PER_PAGE_EXTRACTION = 0.00038;
export const COST_PER_PAGE_TRANSLATION = 0.00090;
export const COST_PER_PAGE_TOTAL = COST_PER_PAGE_EXTRACTION + COST_PER_PAGE_TRANSLATION;

// ─── Processing safeguards ────────────────────────────────────────────────────
/** Maximum pages to process with Vision AI (LLM per-page calls). Pages beyond this are skipped. */
export const MAX_PAGES_VISION = 30;
/** Maximum characters to send to translation LLM (prevents runaway token usage). */
export const MAX_CHARS_TRANSLATION = 100_000;
/** Maximum file size accepted by the server (50 MB). */
export const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024;
/** Job processing timeout in milliseconds (10 minutes). */
export const JOB_TIMEOUT_MS = 10 * 60 * 1000;
/** Cost threshold (USD) above which the UI should warn the user before processing. */
export const COST_WARNING_THRESHOLD = 1.00;

// ─── Available LLM models ─────────────────────────────────────────────────────
export const AVAILABLE_MODELS = [
  { id: "claude-sonnet-4-6",          label: "Claude Sonnet 4.6",   provider: "Anthropic", tier: "fast"    },
  { id: "claude-sonnet-4-5",          label: "Claude Sonnet 4.5",   provider: "Anthropic", tier: "fast"    },
  { id: "claude-opus-4-5",            label: "Claude Opus 4.5",     provider: "Anthropic", tier: "best"    },
  { id: "claude-3-7-sonnet-20250219", label: "Claude Sonnet 3.7",   provider: "Anthropic", tier: "fast"    },
  { id: "gpt-4o",                     label: "GPT-4o",              provider: "OpenAI",    tier: "fast"    },
  { id: "gpt-4.1",                    label: "GPT-4.1",             provider: "OpenAI",    tier: "best"    },
  { id: "gpt-4o-mini",                label: "GPT-4o mini",         provider: "OpenAI",    tier: "economy" },
  { id: "gemini-2.5-flash",           label: "Gemini 2.5 Flash",    provider: "Google",    tier: "fast"    },
  { id: "gemini-2.0-flash",           label: "Gemini 2.0 Flash",    provider: "Google",    tier: "economy" },
] as const;

export type ModelId = (typeof AVAILABLE_MODELS)[number]["id"];
export const DEFAULT_MODEL: ModelId = "claude-sonnet-4-6";

// ─── Supported formats ───────────────────────────────────────────────────────
export const SUPPORTED_INPUT_FORMATS = ["docx", "pptx", "xlsx", "txt"] as const;
export const SUPPORTED_OUTPUT_FORMATS = ["pdf", "docx", "pptx", "xlsx", "txt", "csv"] as const;
export type SupportedFormat = (typeof SUPPORTED_INPUT_FORMATS)[number] | "pdf" | "csv";

export function getMimeType(format: string): string {
  const map: Record<string, string> = {
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    txt: "text/plain",
    csv: "text/csv",
    html: "text/html",
  };
  return map[format] ?? "application/octet-stream";
}

export function getFormatFromMime(mime: string): SupportedFormat {
  if (mime.includes("pdf")) return "pdf";
  if (mime.includes("wordprocessingml") || mime.includes("msword")) return "docx";
  if (mime.includes("presentationml") || mime.includes("powerpoint")) return "pptx";
  if (mime.includes("spreadsheetml") || mime.includes("excel")) return "xlsx";
  return "txt";
}

export const IMAGE_INPUT_FORMATS = ["png", "jpg", "jpeg", "webp", "gif"] as const;
export type ImageInputFormat = (typeof IMAGE_INPUT_FORMATS)[number];

export function isImageFormat(ext: string): ext is ImageInputFormat {
  return IMAGE_INPUT_FORMATS.includes(ext as ImageInputFormat);
}

export function getFormatFromFilename(filename: string): SupportedFormat {
  const ext = path.extname(filename).toLowerCase().replace(".", "");
  if (SUPPORTED_INPUT_FORMATS.includes(ext as any)) return ext as SupportedFormat;
  // Image formats are converted to PDF before processing
  if (isImageFormat(ext)) return "pdf";
  return "txt";
}

// ─── Model-aware LLM invoker ─────────────────────────────────────────────────
/**
 * Creates a model-aware wrapper around the base invokeLLM function.
 * If modelId is provided, it overrides the default model.
 */
export function makeModelInvoker(baseInvokeLLM: Function, modelId?: string) {
  return async (params: any) => {
    if (modelId) {
      return baseInvokeLLM({ ...params, model: modelId });
    }
    return baseInvokeLLM(params);
  };
}

// ─── Text extraction ─────────────────────────────────────────────────────────
export async function extractTextFromBuffer(
  buffer: Buffer,
  format: SupportedFormat,
  filename: string
): Promise<{ text: string; pageCount: number; charCount: number }> {
  let text = "";
  let pageCount = 1;

  if (format === "pdf") {
    // Use pdfjs-dist directly — fully ESM-compatible, no interop issues
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const uint8 = new Uint8Array(buffer);
    const doc = await (pdfjs as any).getDocument({ data: uint8 }).promise;
    pageCount = doc.numPages || 1;
    const pageTexts: string[] = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const content = await page.getTextContent();
      const pageText = (content.items as any[])
        .map((item: any) => ("str" in item ? item.str : ""))
        .join(" ");
      pageTexts.push(pageText);
    }
    text = pageTexts.join("\n\n");
  } else if (format === "docx") {
    const mammoth = await import("mammoth");
    const result = await mammoth.extractRawText({ buffer });
    text = result.value;
    pageCount = Math.max(1, Math.ceil(text.split(/\s+/).length / 500));
  } else if (format === "xlsx") {
    const XLSX = await import("xlsx");
    const wb = XLSX.read(buffer, { type: "buffer" });
    const sheets: string[] = [];
    wb.SheetNames.forEach((name) => {
      const ws = wb.Sheets[name];
      sheets.push(`=== Sheet: ${name} ===\n${XLSX.utils.sheet_to_csv(ws)}`);
    });
    text = sheets.join("\n\n");
    pageCount = wb.SheetNames.length;
  } else if (format === "txt") {
    text = buffer.toString("utf-8");
    pageCount = Math.max(1, Math.ceil(text.split(/\s+/).length / 500));
  } else if (format === "pptx") {
    try {
      const AdmZip = (await import("adm-zip")).default;
      const zip = new AdmZip(buffer);
      const slideEntries = zip.getEntries().filter((e: any) => e.entryName.match(/ppt\/slides\/slide\d+\.xml$/));
      const texts: string[] = [];
      for (const entry of slideEntries) {
        const xml = entry.getData().toString("utf-8");
        const matches = xml.match(/<a:t[^>]*>([^<]*)<\/a:t>/g) || [];
        const slideText = matches.map(m => m.replace(/<[^>]+>/g, "")).join(" ").trim();
        if (slideText) texts.push(slideText);
      }
      text = texts.join("\n\n");
      pageCount = Math.max(1, slideEntries.length);
    } catch {
      text = "[PPTX content could not be extracted]";
      pageCount = 1;
    }
  }

  // Strip invalid XML 1.0 characters (null bytes, control chars) from extracted text
  const sanitized = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "").trim();
  return { text: sanitized, pageCount, charCount: sanitized.length };
}

// ─── Translation via LLM ─────────────────────────────────────────────────────
export async function translateWithLLM(
  text: string,
  targetLanguageName: string,
  invokeLLM: Function
): Promise<string> {
  if (!text.trim()) return text;

  // Cap text length to prevent runaway token usage
  let cappedText = text;
  if (text.length > MAX_CHARS_TRANSLATION) {
    cappedText = text.slice(0, MAX_CHARS_TRANSLATION) + `\n\n[Content truncated at ${MAX_CHARS_TRANSLATION.toLocaleString()} characters to stay within processing limits]`;
  }

  const chunks = splitIntoChunks(cappedText, 6000);
  const translated: string[] = [];

  for (const chunk of chunks) {
    const response = await invokeLLM({
      messages: [
        {
          role: "system",
          content: `You are a professional document translator. Translate the following document content to ${targetLanguageName}.

Rules:
- Preserve ALL markdown formatting exactly: headings (#, ##, ###), bold (**text**), italic (*text*), tables (|col|col|), lists (-, 1.), code blocks
- Preserve ALL structural markers like "--- PAGE BREAK ---"  
- Only translate the text content, not the markdown syntax
- Maintain the same paragraph structure and line breaks
- Do not add any commentary or explanations
- Return ONLY the translated content`,
        },
        { role: "user", content: chunk },
      ],
    });
    translated.push(response.choices[0].message.content as string);
  }

  return translated.join("\n");
}

/**
 * Translate an array of text strings in a single LLM call.
 * Returns an array of the same length with translated strings.
 * Empty strings are passed through unchanged.
 *
 * Uses a line-by-line parser so bracketed numbers inside translated text
 * (e.g. Chinese citations like "[2]") never corrupt the output array.
 */
export async function translateBatchWithLLM(
  texts: string[],
  targetLanguageName: string,
  invokeLLM: Function
): Promise<string[]> {
  const nonEmpty: { idx: number; text: string }[] = [];
  texts.forEach((t, i) => { if (t.trim()) nonEmpty.push({ idx: i, text: t }); });

  if (nonEmpty.length === 0) return texts;

  const result = [...texts];

  // Split into batches of up to 4 000 chars of combined input text
  const batches: typeof nonEmpty[] = [];
  let current: typeof nonEmpty = [];
  let currentLen = 0;
  for (const item of nonEmpty) {
    if (currentLen + item.text.length > 4000 && current.length > 0) {
      batches.push(current);
      current = [];
      currentLen = 0;
    }
    current.push(item);
    currentLen += item.text.length;
  }
  if (current.length > 0) batches.push(current);

  for (const batch of batches) {
    // Use §N§ delimiters instead of [N] so translated text containing
    // bracket-numbers (e.g. "[2]") never confuses the parser.
    const numbered = batch.map((item, i) => `§${i + 1}§ ${item.text}`).join("\n");

    const response = await invokeLLM({
      messages: [
        {
          role: "system",
          content: `You are a professional document translator. Translate each numbered item below to ${targetLanguageName}.

Rules:
- Output EXACTLY one translated line per input item.
- Keep the delimiter at the start of each line: §1§ translation, §2§ translation, etc.
- Do NOT merge, split, or reorder items.
- Do NOT add commentary or explanations.
- Preserve numbers, punctuation, and formatting symbols exactly.`,
        },
        { role: "user", content: numbered },
      ],
    });

    const responseText = response.choices[0].message.content as string;

    // Line-by-line parser: only the START of a line can be a §N§ delimiter
    const lineMap = new Map<number, string>();
    for (const line of responseText.split("\n")) {
      const m = line.match(/^§(\d+)§\s*(.*)/);
      if (m) {
        const n = parseInt(m[1], 10);
        // Accumulate multi-line translations (rare but safe)
        const existing = lineMap.get(n);
        lineMap.set(n, existing ? `${existing} ${m[2]}` : m[2]);
      }
    }

    batch.forEach((item, i) => {
      const translation = lineMap.get(i + 1);
      if (translation && translation.trim()) {
        result[item.idx] = translation.trim();
      }
      // If nothing parsed, keep original text (result was initialised with texts)
    });
  }

  return result;
}

function splitIntoChunks(text: string, maxLen: number): string[] {
  const chunks: string[] = [];
  let start = 0;
  while (start < text.length) {
    let end = start + maxLen;
    if (end < text.length) {
      const breakAt = text.lastIndexOf("\n\n", end);
      if (breakAt > start) end = breakAt;
    }
    chunks.push(text.slice(start, end));
    start = end;
  }
  return chunks;
}

// ─── LLM-based extraction with layout metadata ───────────────────────────────
export async function extractWithLLM(
  fileUrl: string,
  format: SupportedFormat,
  invokeLLM: Function
): Promise<{ structuredContent: string; pageCount: number }> {
  const systemPrompt = `You are a document content extractor. Extract ALL text content from the document, preserving:
- Headings and their levels (use # ## ### for heading levels)
- Paragraph structure (preserve paragraph breaks)
- Table content (use markdown table format)
- List items (use - or 1. for lists)
- Bold/italic emphasis where visible
- Page breaks (mark with --- PAGE BREAK ---)
- Any metadata like title, author, date

Return ONLY the extracted content in structured markdown format. Do not add commentary.`;

  const userMessage = format === "pdf"
    ? { role: "user" as const, content: [{ type: "file_url" as const, file_url: { url: fileUrl, mime_type: "application/pdf" as const } }, { type: "text" as const, text: "Extract all text content from this document, preserving structure and formatting as described." }] }
    : { role: "user" as const, content: `Extract all text content from this ${format.toUpperCase()} document at URL: ${fileUrl}. Preserve structure and formatting.` };

  const response = await invokeLLM({
    messages: [{ role: "system", content: systemPrompt }, userMessage],
  });

  const content = response.choices[0].message.content as string;
  const pageBreaks = (content.match(/--- PAGE BREAK ---/g) || []).length;
  const pageCount = Math.max(1, pageBreaks + 1);

  return { structuredContent: content, pageCount };
}

// ─────────────────────────────────────────────────────────────────────────────
// PDF → PPTX VISION PIPELINE
// Renders each PDF page as a PNG image, uses LLM Vision to extract structured
// slide content (text positions, colors, font sizes), then builds a PPTX where:
//   - The original page image is used as the slide background
//   - Translated text boxes are overlaid at the detected positions
// This preserves ALL visual design: images, colors, logos, backgrounds.
// ─────────────────────────────────────────────────────────────────────────────

interface SlideElement {
  type: "text" | "image" | "logo" | "shape";
  content?: string;
  position: { x: number; y: number };
  size: { w: number; h: number };
  style?: {
    fontSize?: number;
    bold?: boolean;
    color?: string;
    align?: "left" | "center" | "right";
  };
}

interface SlideLayout {
  background: string;
  elements: SlideElement[];
}

/**
 * Convert a PDF to PPTX using LLM Vision for layout-preserving reconstruction.
 *
 * NEW APPROACH (no system binaries required):
 * 1. Upload the PDF to S3 once
 * 2. For each page, send the PDF as a file_url to the LLM with a per-page prompt
 *    asking it to extract text elements with positions as percentages
 * 3. Build a PPTX where each slide uses the page image (rendered via pdfjs-dist
 *    canvas-free pixel data) as a data-URI background + translated text overlays
 *
 * Falls back to text-only slides if image rendering is unavailable.
 */
export async function convertPdfToPptxWithVision(
  buffer: Buffer,
  targetLanguageName: string | null,
  invokeLLM: Function,
  storagePutFn: (key: string, data: Buffer, mime: string) => Promise<{ url: string }>,
  onProgress?: (msg: string, level?: "info" | "progress" | "success" | "warning" | "error") => void,
  isCancelled?: () => Promise<boolean>
): Promise<Buffer> {
  const baseName = `pdf_vision_${Date.now()}`;

  // Get page count
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const uint8 = new Uint8Array(buffer);
  const pdfDoc = await (pdfjs as any).getDocument({ data: uint8 }).promise;
  const totalPages = pdfDoc.numPages as number;

  // Apply page cap — only process up to MAX_PAGES_VISION pages via Vision AI
  const pagesToProcess = Math.min(totalPages, MAX_PAGES_VISION);
  if (totalPages > MAX_PAGES_VISION) {
    onProgress?.(`Document has ${totalPages} pages. Processing first ${MAX_PAGES_VISION} pages via Vision AI (limit).`, "warning");
  }

  onProgress?.(`Uploading PDF for Vision AI analysis (${pagesToProcess} of ${totalPages} pages)...`, "info");

  // Check cancellation before expensive upload
  if (await isCancelled?.()) throw new Error("Job cancelled by user");

  // Upload the PDF to S3 once — LLM reads it directly via file_url
  const pdfKey = `tmp/vision/${baseName}.pdf`;
  const { url: pdfUrl } = await storagePutFn(pdfKey, buffer, "application/pdf");

  onProgress?.(`Analyzing ${pagesToProcess} pages with Vision AI...`, "info");

  // Extract layout for every page by sending the PDF + page number to the LLM
  const slideLayouts: SlideLayout[] = new Array(pagesToProcess).fill(null);

  // Process pages in batches of 5 to balance speed vs rate limits
  const BATCH_SIZE = 5;
  for (let batchStart = 0; batchStart < pagesToProcess; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE, pagesToProcess);
    const batchPromises = Array.from({ length: batchEnd - batchStart }, (_, idx) => {
      const pageNum = batchStart + idx + 1;
      return (async () => {
        try {
          const response = await invokeLLM({
            messages: [{
              role: "user",
              content: [
                {
                  type: "file_url",
                  file_url: { url: pdfUrl, mime_type: "application/pdf" },
                },
                {
                  type: "text",
                  text: `Look at page ${pageNum} of this PDF only. Extract ALL visible text elements with their positions on that page.

Return ONLY valid JSON (no markdown, no code blocks):
{
  "background": "brief description of background color/image",
  "elements": [
    {
      "type": "text",
      "content": "exact text content",
      "x": 5,
      "y": 10,
      "w": 40,
      "h": 8,
      "fontSize": 24,
      "bold": true,
      "color": "#FFFFFF",
      "align": "left"
    }
  ]
}

Rules:
- x, y, w, h are PERCENTAGES of page width/height (0-100)
- Include ALL visible text: titles, subtitles, body text, captions, footnotes, labels, bullet points
- fontSize in approximate points; color as hex code
- Only return the JSON object, nothing else`,
                },
              ],
            }],
          });

          const raw = (response.choices[0].message.content as string).trim();
          const jsonStr = raw.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/i, "").trim();
          const parsed = JSON.parse(jsonStr);
          // Normalize element shape
          const elements: SlideElement[] = (parsed.elements || []).map((el: any) => ({
            type: el.type || "text",
            content: el.content,
            position: { x: el.x ?? 0, y: el.y ?? 0 },
            size: { w: el.w ?? 80, h: el.h ?? 5 },
            style: {
              fontSize: el.fontSize,
              bold: el.bold,
              color: el.color,
              align: el.align,
            },
          }));
          return { pageNum, layout: { background: parsed.background || "", elements } as SlideLayout };
        } catch {
          return { pageNum, layout: { background: "", elements: [] } as SlideLayout };
        }
      })();
    });

    const batchResults = await Promise.all(batchPromises);
    for (const { pageNum, layout } of batchResults) {
      slideLayouts[pageNum - 1] = layout;
    }
    onProgress?.(`Analyzed pages ${batchStart + 1}–${batchEnd} of ${pagesToProcess}${totalPages > pagesToProcess ? ` (${totalPages - pagesToProcess} pages skipped — limit reached)` : ''}`, "progress");

    // Check cancellation between batches
    if (await isCancelled?.()) throw new Error("Job cancelled by user");
  }

  // Translate all text elements if a target language is specified
  if (targetLanguageName) {
    onProgress?.(`Translating text to ${targetLanguageName}...`, "info");
    const allTexts: { slideIdx: number; elemIdx: number; text: string }[] = [];
    slideLayouts.forEach((layout, si) => {
      layout.elements.forEach((elem, ei) => {
        if (elem.type === "text" && elem.content?.trim()) {
          allTexts.push({ slideIdx: si, elemIdx: ei, text: elem.content });
        }
      });
    });
    if (allTexts.length > 0) {
      const textStrings = allTexts.map(t => t.text);
      const translated = await translateBatchWithLLM(textStrings, targetLanguageName, invokeLLM);
      allTexts.forEach((item, i) => {
        slideLayouts[item.slideIdx].elements[item.elemIdx].content = translated[i] ?? item.text;
      });
    }
  }

  // Check cancellation before building output
  if (await isCancelled?.()) throw new Error("Job cancelled by user");
  onProgress?.(`Building PPTX with ${slideLayouts.length} slides...`, "info");

  // Build PPTX using pure Node.js:
  //   1. Render each PDF page to PNG via pdfjs-dist + @napi-rs/canvas (no Python/system binaries)
  //   2. Embed the PNG as a full-slide background image via pptxgenjs
  //   3. Overlay translated text boxes with correct positions/styles
  //   4. Produces a valid PPTX (single slide master, no corruption)
  const { createCanvas: napiCreateCanvas } = await import("@napi-rs/canvas");

  // Re-open the PDF document for rendering
  const renderPdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const renderUint8 = new Uint8Array(buffer);
  const renderPdfDoc = await (renderPdfjs as any).getDocument({ data: renderUint8 }).promise;

  onProgress?.(`Rendering ${slideLayouts.length} PDF pages to images...`, "progress");

  // Render all pages to PNG buffers
  const pageImages: Buffer[] = [];
  for (let i = 0; i < slideLayouts.length; i++) {
    const pageNum = i + 1;
    try {
      const page = await renderPdfDoc.getPage(pageNum);
      // Use 1.5x scale for good quality without huge file sizes
      const viewport = page.getViewport({ scale: 1.5 });
      const canvasWidth = Math.floor(viewport.width);
      const canvasHeight = Math.floor(viewport.height);
      const renderCanvas = napiCreateCanvas(canvasWidth, canvasHeight);
      const renderCtx = renderCanvas.getContext("2d");
      // Fill white background first
      renderCtx.fillStyle = "#ffffff";
      renderCtx.fillRect(0, 0, canvasWidth, canvasHeight);
      await page.render({ canvasContext: renderCtx as any, viewport }).promise;
      const pngBuf = renderCanvas.toBuffer("image/png");
      pageImages.push(pngBuf);
    } catch {
      // If rendering fails for a page, push an empty white image
      const fallback = napiCreateCanvas(960, 540);
      const fCtx = fallback.getContext("2d");
      fCtx.fillStyle = "#ffffff";
      fCtx.fillRect(0, 0, 960, 540);
      pageImages.push(fallback.toBuffer("image/png"));
    }
    if (await isCancelled?.()) throw new Error("Job cancelled by user");
  }

  onProgress?.(`Assembling PPTX slides...`, "progress");

  // Build PPTX with pptxgenjs — one presentation, one master
  const PptxGenJS = (await import("pptxgenjs")).default;
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.33" x 7.5"

  const SLIDE_W = 13.33;
  const SLIDE_H = 7.5;

  for (let i = 0; i < slideLayouts.length; i++) {
    const layout = slideLayouts[i];
    const slide = pptx.addSlide();

    // Add the rendered PDF page as full-slide background image
    const imgBase64 = pageImages[i].toString("base64");
    slide.addImage({
      data: `data:image/png;base64,${imgBase64}`,
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    });

    // Overlay text elements on top
    for (const el of layout.elements) {
      if (el.type !== "text" || !el.content?.trim()) continue;
      const x = (el.position.x / 100) * SLIDE_W;
      const y = (el.position.y / 100) * SLIDE_H;
      const w = Math.max((el.size.w / 100) * SLIDE_W, 0.5);
      const h = Math.max((el.size.h / 100) * SLIDE_H, 0.2);
      const fontSize = el.style?.fontSize ?? 14;
      const bold = el.style?.bold ?? false;
      const color = (el.style?.color ?? "#000000").replace("#", "");
      const align = (el.style?.align ?? "left") as "left" | "center" | "right";
      slide.addText(el.content, {
        x, y, w, h,
        fontSize,
        bold,
        color,
        align,
        fontFace: "Calibri",
        wrap: true,
        valign: "top",
        // Transparent background so the image shows through
        fill: { type: "none" } as any,
      });
    }
  }

  const pptxBuffer = await pptx.write({ outputType: "nodebuffer" }) as unknown as Buffer;

  // ── Fix pptxgenjs Content_Types corruption ──────────────────────────────────
  // pptxgenjs adds one <Override PartName="/ppt/slideMasters/slideMasterN.xml"> entry
  // per slide in [Content_Types].xml, but only writes slideMaster1.xml.
  // PowerPoint refuses to open the file because slideMaster2..N don't exist.
  // Fix: deduplicate slideMaster Override entries, keeping only the first one.
  const AdmZip = (await import("adm-zip")).default;
  const fixZip = new AdmZip(pptxBuffer);
  const ctEntry = fixZip.getEntry("[Content_Types].xml");
  if (ctEntry) {
    let ctXml = ctEntry.getData().toString("utf-8");
    // Remove all slideMaster Override entries after the first one
    let firstSlideMasterSeen = false;
    ctXml = ctXml.replace(
      /<Override[^>]*\/ppt\/slideMasters\/slideMaster\d+\.xml[^>]*\/>/g,
      (match) => {
        if (!firstSlideMasterSeen) { firstSlideMasterSeen = true; return match; }
        return ""; // remove duplicates
      }
    );
    fixZip.updateFile("[Content_Types].xml", Buffer.from(ctXml, "utf-8"));
  }
  return fixZip.toBuffer();
}

// ─────────────────────────────────────────────────────────────────────────────
// IN-PLACE XML TRANSLATION
// For PPTX, DOCX, and XLSX: unzip → find text nodes → translate → rezip
// This preserves ALL images, shapes, backgrounds, fonts, and formatting.
// ─────────────────────────────────────────────────────────────────────────────

// ─── PPTX paragraph-level skip logic (mirrors pptx-translator skill) ─────────
const PPTX_SKIP_PATTERNS = [
  /^\s*$/,                                   // empty / whitespace
  /^\d[\d\s\.\,\%\-\/]*$/,                  // pure numbers / percentages
  /^https?:\/\//i,                           // URLs
  /^www\./i,                                 // URLs
  /^[<>]?number[<>]?$/i,                     // slide number placeholders
];

/** Returns true if the target language uses a CJK / Arabic / RTL script (i.e. is itself a non-Latin language). */
function isTargetCJKOrRTL(targetLanguageName: string): boolean {
  // Match common CJK / Arabic / Hebrew / Thai language names
  return /chinese|japanese|korean|arabic|hebrew|thai|hindi|urdu|persian|farsi/i.test(targetLanguageName);
}

/**
 * Returns true if this paragraph should be skipped for translation.
 * @param text            The paragraph text to evaluate.
 * @param targetLanguageName  The human-readable target language (e.g. "English", "Chinese").
 *                        Used to decide whether CJK source text should be translated.
 */
function pptxShouldSkip(text: string, targetLanguageName: string): boolean {
  const t = text.trim();
  if (!t || t.length <= 1) return true;

  const hasCJKOrRTL = /[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af\u0600-\u06ff]/.test(t);

  if (hasCJKOrRTL) {
    // If the target is also a CJK/RTL language (e.g. Chinese→Japanese), skip —
    // the text is likely already in or very close to the target script.
    // If the target is a Latin-script language (e.g. Chinese→English), DO NOT skip —
    // this is the source text that needs translating.
    if (isTargetCJKOrRTL(targetLanguageName)) return true;
    // CJK source → Latin target: translate it
    return false;
  }

  // No Latin characters at all — symbols, numbers, icons
  if (!/[a-zA-Z]/.test(t)) return true;

  for (const pattern of PPTX_SKIP_PATTERNS) {
    if (pattern.test(t)) return true;
  }
  return false;
}

/**
 * Build a numbered-list translation prompt that instructs the LLM to:
 * - Keep brand names, technical abbreviations, URLs unchanged within sentences
 * - Return natural, fluent output (not word-by-word)
 * - Preserve punctuation and formatting
 */
function buildPptxTranslationPrompt(strings: string[], targetLanguageName: string): string {
  const numbered = strings.map((s, i) => `${i + 1}. ${s}`).join("\n");
  return (
    `Translate the following numbered list to ${targetLanguageName}.\n` +
    `Rules:\n` +
    `- Return ONLY the translated numbered list in the same format.\n` +
    `- Keep brand names, proper nouns, technical abbreviations ` +
    `(e.g. CAPEX, OPEX, CO2, BSG, GMO, LCA, DACH, MVP, TVP, KPI, ROI, EBITDA), ` +
    `URLs, and email addresses unchanged within the translation.\n` +
    `- Produce natural, fluent ${targetLanguageName} — do NOT translate word-by-word.\n` +
    `- Preserve all punctuation and formatting within each item.\n` +
    `- Do NOT merge or split items.\n\n` +
    numbered
  );
}

/**
 * Strip any XML/HTML tags that the LLM may have hallucinated into a translation.
 * Also unescape any HTML entities the LLM may have introduced.
 */
function sanitizeTranslation(text: string): string {
  // Remove any XML/HTML tags (e.g. <a:rPr .../>, <a:tab/>, <br/>, etc.)
  let cleaned = text.replace(/<[^>]+>/g, "");
  // Unescape common HTML entities that the LLM might have introduced
  cleaned = cleaned
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
  return cleaned.trim();
}

function parsePptxTranslationResponse(responseText: string, originals: string[]): Map<string, string> {
  const map = new Map<string, string>();
  const lines = responseText.trim().split("\n").map(l => l.trim()).filter(Boolean);
  for (const line of lines) {
    const m = line.match(/^(\d+)\.\s*([\s\S]*)/);
    if (m) {
      const idx = parseInt(m[1]) - 1;
      if (idx >= 0 && idx < originals.length) {
        map.set(originals[idx], sanitizeTranslation(m[2]));
      }
    }
  }
  return map;
}

/**
 * Translate a PPTX buffer in-place using the pptx-translator skill approach:
 * 1. Extract unique paragraph texts (joining all <a:r> runs in each <a:p>)
 * 2. Skip pure numbers / URLs / already-translated text
 * 3. Batch-translate with brand-name preservation prompt
 * 4. Apply back: consolidate all runs in a paragraph into one run,
 *    preserving the first run's formatting (font, size, bold, color)
 */
export async function translatePptxInPlace(
  buffer: Buffer,
  targetLanguageName: string,
  invokeLLM: Function,
  onProgress?: (msg: string) => void
): Promise<Buffer> {
  // Use JSZip instead of AdmZip — AdmZip 0.5.x mishandles PPTX binary entries
  // (images, media) that use streaming ZIP format (bit-3 data descriptors),
  // producing a re-written ZIP that PowerPoint rejects. JSZip handles all ZIP
  // variants and re-writes them as standard DEFLATE-compressed archives.
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);

  const slideNames = Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)?.[1] ?? "0");
      const nb = parseInt(b.match(/slide(\d+)/)?.[1] ?? "0");
      return na - nb;
    });

  const totalSlides = slideNames.length;

  // ── Phase 1: Extract all unique paragraph texts across all slides ──────────
  const uniqueParas = new Set<string>();
  const slideXmls: { name: string; xml: string }[] = [];

  for (const name of slideNames) {
    const xml = await zip.file(name)!.async("text");
    slideXmls.push({ name, xml });

    // Extract paragraphs: each <a:p>...</a:p> block
    const paraRegex = /<a:p[\s>][\s\S]*?<\/a:p>/g;
    let pm;
    while ((pm = paraRegex.exec(xml)) !== null) {
      const paraXml = pm[0];
      // Collect all run texts within this paragraph
      const runTexts: string[] = [];
      const runRegex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
      let rm;
      while ((rm = runRegex.exec(paraXml)) !== null) {
        runTexts.push(rm[1]);
      }
      // Unescape HTML entities from the raw XML text content before comparing/translating
      const rawText = runTexts.join("").trim();
      const fullText = rawText
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      if (fullText && !pptxShouldSkip(fullText, targetLanguageName)) {
        uniqueParas.add(fullText);
      }
    }
  }

  const toTranslate = Array.from(uniqueParas);
  if (toTranslate.length === 0) return zip.toBuffer();

  onProgress?.(`Translating ${toTranslate.length} unique paragraphs across ${totalSlides} slides...`);

  // ── Phase 2: Batch-translate with brand-name preservation ─────────────────
  const BATCH_SIZE = 60;
  const translationMap = new Map<string, string>();
  const totalBatches = Math.ceil(toTranslate.length / BATCH_SIZE);

  for (let b = 0; b < totalBatches; b++) {
    const batch = toTranslate.slice(b * BATCH_SIZE, (b + 1) * BATCH_SIZE);
    onProgress?.(`Translating batch ${b + 1} of ${totalBatches} (${batch.length} paragraphs)...`);

    const prompt = buildPptxTranslationPrompt(batch, targetLanguageName);
    let attempts = 0;
    while (attempts < 3) {
      try {
        const response = await invokeLLM({
          messages: [{ role: "user", content: prompt }],
        });
        const responseText = response.choices[0].message.content as string;
        const batchMap = parsePptxTranslationResponse(responseText, batch);
        // Fall back to original for any unparsed strings
        for (const s of batch) {
          translationMap.set(s, batchMap.get(s) ?? s);
        }
        break;
      } catch (e) {
        attempts++;
        if (attempts >= 3) {
          for (const s of batch) translationMap.set(s, s);
        } else {
          await new Promise(r => setTimeout(r, 3000));
        }
      }
    }
  }

  const translatedCount = Array.from(translationMap.entries()).filter(([k, v]) => k !== v).length;
  onProgress?.(`Applied ${translatedCount} translations — rebuilding slides...`);

  // ── Phase 3: Apply translations with run consolidation ────────────────────
  // For each paragraph: if its full text matches a translation,
  // replace all <a:r> runs with a single run containing the translated text,
  // preserving the first run's formatting.
  for (const { name, xml } of slideXmls) {
    let updatedXml = xml;

    // Replace each <a:p>...</a:p> block
    updatedXml = updatedXml.replace(/<a:p[\s>][\s\S]*?<\/a:p>/g, (paraXml) => {
      // Collect run texts
      const runTexts: string[] = [];
      const runRegex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
      let rm;
      while ((rm = runRegex.exec(paraXml)) !== null) {
        runTexts.push(rm[1]);
      }
      // Unescape HTML entities to match the keys stored in translationMap (same as Phase 1)
      const rawText = runTexts.join("").trim();
      const fullText = rawText
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      if (!fullText) return paraXml;

      const translated = translationMap.get(fullText);
      if (!translated || translated === fullText) return paraXml;

      // Find the first <a:r> run to preserve its formatting
      const firstRunMatch = paraXml.match(/<a:r>[\s\S]*?<\/a:r>/);
      if (!firstRunMatch) return paraXml;

      // Extract the <a:rPr> (run properties) from the first run.
      // Use greedy match: <a:rPr .../> (self-closing) OR <a:rPr ...>...</a:rPr> (with children).
      // The non-greedy .*? would stop at the first /> inside a child element (e.g. <a:srgbClr val="FF8C00"/>)
      // causing the rPr to be cut off mid-element and producing invalid XML.
      const rPrSelfClosing = firstRunMatch[0].match(/<a:rPr[^>]*\/>/);
      const rPrWithChildren = firstRunMatch[0].match(/<a:rPr[\s\S]*?<\/a:rPr>/);
      // Prefer the full element with children; fall back to self-closing
      const rPr = rPrWithChildren ? rPrWithChildren[0] : (rPrSelfClosing ? rPrSelfClosing[0] : "");

      // Build a single consolidated run with the translated text
      const consolidatedRun = `<a:r>${rPr}<a:t>${escapeXml(translated)}</a:t></a:r>`;

      // Replace first <a:r>...</a:r> with consolidated run, remove all subsequent runs
      let firstReplaced = false;
      const updatedPara = paraXml.replace(/<a:r>[\s\S]*?<\/a:r>/g, () => {
        if (!firstReplaced) { firstReplaced = true; return consolidatedRun; }
        return "";
      });

      return updatedPara;
    });

    zip.file(name, updatedXml);
  }

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

/**
 * Apply two-pass regex translation to a single DOCX XML string.
 * Returns the translated XML string.
 * Throws if no text was found.
 */
async function translateDocxXml(
  xml: string,
  targetLanguageName: string,
  invokeLLM: Function,
  partLabel = "word/document.xml"
): Promise<string> {
  // Strip invalid XML 1.0 control characters
  const cleanXml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");

  // ── Pass 1: collect text strings ─────────────────────────────────────────
  //
  // We track ALL non-empty w:t nodes in order.
  // Nodes whose raw content contains entity-encoded XML markup (&lt; / &gt;)
  // are flagged as "markup storage" nodes and must NOT be translated — sending
  // them to the LLM causes it to return the markup mixed with the translation,
  // which then gets double-encoded and corrupts the ZIP.
  //
  // skipSet maps each non-empty node's sequential index → true if it should
  // be skipped.  Only non-skip nodes are sent to the LLM.

  const isMarkupNode = (raw: string) =>
    (raw.includes("&lt;") || raw.includes("&gt;"));

  const nodeTexts: string[] = [];   // unescaped text for every non-empty node
  const skipSet   = new Set<number>(); // indices of markup-storage nodes

  // NOTE: Pattern is <w:t(?:\s[^>]*)?>  — the char after "t" must be ">" or
  // whitespace so we don't accidentally match <w:tbl>, <w:tr>, <w:tblPr>, etc.
  const wTextRegex = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g;
  let m: RegExpExecArray | null;
  while ((m = wTextRegex.exec(cleanXml)) !== null) {
    const raw = m[2];
    if (!raw.trim()) continue;
    if (isMarkupNode(raw)) {
      skipSet.add(nodeTexts.length);
      nodeTexts.push(raw); // keep raw for mapping; won't be sent to LLM
    } else {
      nodeTexts.push(unescapeXml(raw));
    }
  }

  const textsToTranslate = nodeTexts.filter((_, i) => !skipSet.has(i));

  console.log(`[DOCX:${partLabel}] Pass 1: ${nodeTexts.length} non-empty nodes, ${skipSet.size} markup-storage nodes skipped, ${textsToTranslate.length} sent to LLM`);
  if (textsToTranslate.length === 0) return cleanXml;

  console.log(`[DOCX:${partLabel}] Sample texts:`, textsToTranslate.slice(0, 5).map(t => t.slice(0, 60)));

  const translatedBatch = await translateBatchWithLLM(textsToTranslate, targetLanguageName, invokeLLM);

  // Map translations back onto nodeTexts (skipped nodes keep their raw value)
  let batchIdx = 0;
  const allTranslations: string[] = nodeTexts.map((orig, i) => {
    if (skipSet.has(i)) return orig; // keep original (will not be used for replacement)
    return translatedBatch[batchIdx++] ?? orig;
  });

  console.log(`[DOCX:${partLabel}] Sample translations:`, allTranslations.filter((_, i) => !skipSet.has(i)).slice(0, 5).map(t => String(t).slice(0, 60)));

  // ── Pass 2: build replacement list with exact character positions ─────────
  const allMatches: { start: number; end: number; replacement: string }[] = [];
  const wTextRegex2 = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g;
  let m2: RegExpExecArray | null;
  let nodeIdx = 0;
  while ((m2 = wTextRegex2.exec(cleanXml)) !== null) {
    const raw = m2[2];
    if (!raw.trim()) continue;

    // Skip markup-storage nodes — leave their content completely untouched
    if (!skipSet.has(nodeIdx)) {
      const translated = allTranslations[nodeIdx];
      const contentStart = m2.index + m2[1].length;
      const contentEnd   = contentStart + m2[2].length;
      allMatches.push({
        start: contentStart,
        end:   contentEnd,
        replacement: escapeXml(translated),
      });
    }
    nodeIdx++;
  }

  console.log(`[DOCX:${partLabel}] Pass 2: ${allMatches.length} replacements will be applied`);

  // Apply in reverse so earlier positions remain valid
  let translatedXml = cleanXml;
  for (let i = allMatches.length - 1; i >= 0; i--) {
    const { start, end, replacement } = allMatches[i];
    translatedXml = translatedXml.slice(0, start) + replacement + translatedXml.slice(end);
  }

  // ── Structural validation ────────────────────────────────────────────────
  // Compare tag counts between original and translated XML. The translation
  // only replaces text inside <w:t> tags, so structural tag counts should
  // stay the same. If they differ, the translation introduced corruption.
  // (We compare against the original rather than checking opens==closes
  // because the original document itself may have self-closing tags like
  // <w:p/> that legitimately cause an imbalance.)
  const countTag = (s: string, tag: string) => {
    const open = (s.match(new RegExp(`<${tag}[\\s>/]`, "g")) || []).length;
    const close = (s.match(new RegExp(`</${tag}>`, "g")) || []).length;
    return { open, close };
  };

  const criticalTags = ["w:body", "w:document", "w:p", "w:r", "w:tbl", "w:tr", "w:tc"];
  let structureOk = true;
  for (const tag of criticalTags) {
    const orig = countTag(cleanXml, tag);
    const translated = countTag(translatedXml, tag);
    if (orig.open !== translated.open || orig.close !== translated.close) {
      console.error(`[DOCX:${partLabel}] TAG COUNT CHANGED: <${tag}> original=${orig.open}/${orig.close} translated=${translated.open}/${translated.close}`);
      structureOk = false;
    }
  }

  if (!structureOk) {
    console.error(`[DOCX:${partLabel}] XML structure corrupted after translation — returning original XML`);
    return cleanXml;
  }

  if (!translatedXml.includes("<w:document") && !translatedXml.includes("<w:hdr") && !translatedXml.includes("<w:ftr")) {
    console.error(`[DOCX:${partLabel}] XML root tag lost — returning original XML`);
    return cleanXml;
  }

  console.log(`[DOCX:${partLabel}] Done. Output length: ${translatedXml.length} (input: ${cleanXml.length})`);

  return translatedXml;
}

/**
 * Translate a DOCX buffer in-place.
 *
 * Translates the main document body (word/document.xml) AND any headers /
 * footers (word/header*.xml, word/footer*.xml) so the full document is
 * correctly translated without XML corruption.
 */
export async function translateDocxInPlace(
  buffer: Buffer,
  targetLanguageName: string,
  invokeLLM: Function
): Promise<Buffer> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);

  // Collect all XML parts that contain translatable <w:t> nodes
  const partNames: string[] = [];
  for (const name of Object.keys(zip.files)) {
    if (
      name === "word/document.xml" ||
      /^word\/header\d*\.xml$/.test(name) ||
      /^word\/footer\d*\.xml$/.test(name)
    ) {
      partNames.push(name);
    }
  }

  if (partNames.length === 0) return buffer;

  let modified = false;
  for (const partName of partNames) {
    const file = zip.file(partName);
    if (!file) continue;

    const originalXml = await file.async("text");
    try {
      const translatedXml = await translateDocxXml(originalXml, targetLanguageName, invokeLLM, partName);
      if (translatedXml !== originalXml) {
        zip.file(partName, translatedXml);
        modified = true;
        console.log(`[DOCX] Updated ${partName} in ZIP`);
      }
    } catch (err) {
      console.error(`[DOCX] Skipping ${partName} due to error:`, err);
    }
  }

  if (!modified) return buffer;

  try {
    return Buffer.from(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }));
  } catch (err) {
    console.error("[DOCX] Error rebuilding ZIP:", err);
    return buffer;
  }
}

/**
 * Translate an XLSX buffer in-place.
 */
export async function translateXlsxInPlace(
  buffer: Buffer,
  targetLanguageName: string,
  invokeLLM: Function
): Promise<Buffer> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(buffer);

  const ssFile = zip.file("xl/sharedStrings.xml");
  if (!ssFile) return buffer;

  let xml = await ssFile.async("text");

  const textMatches: string[] = [];
  const regex = /(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    const text = unescapeXml(match[2]);
    if (text.trim()) textMatches.push(text);
  }

  if (textMatches.length === 0) return buffer;

  const translatedTexts = await translateBatchWithLLM(textMatches, targetLanguageName, invokeLLM);

  const allMatches: { start: number; end: number; replacement: string }[] = [];
  const regex2 = /(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/g;
  let match2;
  let nonEmptyIdx = 0;
  while ((match2 = regex2.exec(xml)) !== null) {
    const rawText = match2[2];
    if (unescapeXml(rawText).trim()) {
      const translated = translatedTexts[nonEmptyIdx] ?? unescapeXml(rawText);
      allMatches.push({
        start: match2.index + match2[1].length,
        end: match2.index + match2[1].length + rawText.length,
        replacement: escapeXml(translated),
      });
      nonEmptyIdx++;
    }
  }

  let translatedXml = xml;
  for (let i = allMatches.length - 1; i >= 0; i--) {
    const m = allMatches[i];
    translatedXml = translatedXml.slice(0, m.start) + m.replacement + translatedXml.slice(m.end);
  }

  zip.file("xl/sharedStrings.xml", translatedXml);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

function escapeXml(str: string): string {
  // Strip invalid XML 1.0 characters: null bytes and control chars except tab, LF, CR
  const sanitized = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
  return sanitized
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeXml(str: string): string {
  return str
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

// ─── Build output document from translated markdown text ─────────────────────
export async function buildTranslatedDocument(
  translatedText: string,
  outputFormat: SupportedFormat,
  originalFilename: string
): Promise<Buffer> {
  if (outputFormat === "txt") {
    return Buffer.from(translatedText, "utf-8");
  }

  if (outputFormat === "docx") {
    return await buildDocx(translatedText);
  }

  if (outputFormat === "pptx") {
    return await buildPptx(translatedText);
  }

  if (outputFormat === "xlsx") {
    return await buildXlsx(translatedText);
  }

  if (outputFormat === "pdf") {
    // Use CJK-aware font stack so Chinese/Japanese/Korean characters render correctly
    const htmlContent = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>@import url('file:///usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc');body{font-family:'Noto Serif CJK SC','Noto Sans CJK SC','WenQuanYi Zen Hei','WenQuanYi Micro Hei',sans-serif;max-width:900px;margin:40px auto;padding:0 20px;line-height:1.8;color:#333;word-break:break-word;}h1,h2,h3{font-family:inherit;}p{margin:0.8em 0;}</style></head><body>${markdownToHtml(translatedText)}</body></html>`;
    const htmlBuf = Buffer.from(htmlContent, "utf-8");
    let libreOfficePdfPath: string | null = null;
    try {
      const tmpDir2 = fs.mkdtempSync(path.join(os.tmpdir(), "pdf-build-"));
      const baseName = `pdf_${Date.now()}`;
      const htmlFile = path.join(tmpDir2, `${baseName}.html`);
      const pdfFile = path.join(tmpDir2, `${baseName}.pdf`);
      fs.writeFileSync(htmlFile, htmlBuf);
      await execAsync(`libreoffice --headless --convert-to pdf --outdir ${tmpDir2} ${htmlFile}`, { timeout: 60000 });
      if (fs.existsSync(pdfFile) && fs.statSync(pdfFile).size > 100) {
        libreOfficePdfPath = pdfFile;
        const result = fs.readFileSync(pdfFile);
        try { fs.unlinkSync(pdfFile); } catch {}
        try { fs.unlinkSync(htmlFile); } catch {}
        try { fs.rmdirSync(tmpDir2); } catch {}
        return result;
      }
      try { fs.rmSync(tmpDir2, { recursive: true, force: true }); } catch {}
    } catch {
      // LibreOffice failed — fall back to pdf-lib plain text PDF
    }
    // Fallback: build a plain PDF with pdf-lib (CJK chars become boxes but at least it's a valid PDF)
    const { PDFDocument: PDFDoc2, StandardFonts: SF2, rgb: rgb2 } = await import("pdf-lib");
    const pdfDoc2 = await PDFDoc2.create();
    const font2 = await pdfDoc2.embedFont(SF2.Helvetica);
    const pageW = 595, pageH = 842, margin = 50, lh = 14, fs2 = 11;
    const maxW = pageW - margin * 2;
    const maxLpp = Math.floor((pageH - margin * 2) / lh);
    const rawLines = translatedText.split(/\r?\n/);
    const wrapped: string[] = [];
    for (const raw of rawLines) {
      if (!raw) { wrapped.push(""); continue; }
      let rem = raw;
      while (rem.length > 0) {
        let end = rem.length;
        while (end > 0 && font2.widthOfTextAtSize(rem.slice(0, end), fs2) > maxW) end--;
        if (end === 0) end = 1;
        wrapped.push(rem.slice(0, end));
        rem = rem.slice(end);
      }
    }
    let pg2: any = null; let ly = 0; let lc = 0;
    for (const line of wrapped) {
      if (!pg2 || lc >= maxLpp) {
        pg2 = pdfDoc2.addPage([pageW, pageH]);
        ly = pageH - margin - lh; lc = 0;
      }
      if (line.length > 0) {
        // Filter out non-latin chars for Helvetica (they'd render as boxes)
        const safe = line.replace(/[^\x00-\xFF]/g, "?");
        pg2.drawText(safe, { x: margin, y: ly, size: fs2, font: font2, color: rgb2(0.1, 0.1, 0.1) });
      }
      ly -= lh; lc++;
    }
    if (pdfDoc2.getPageCount() === 0) pdfDoc2.addPage([pageW, pageH]);
    return Buffer.from(await pdfDoc2.save());
  }

  return Buffer.from(translatedText, "utf-8");
}

// ─── Build DOCX from scratch ──────────────────────────────────────────────────
async function buildDocx(markdownText: string): Promise<Buffer> {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = await import("docx");
  // Strip invalid XML 1.0 characters (null bytes, control chars) before building DOCX
  const sanitizedText = markdownText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");
  const lines = sanitizedText.split("\n");
  const children: any[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.trim() === "--- PAGE BREAK ---") {
      children.push(new Paragraph({ pageBreakBefore: true, children: [] }));
    } else if (line.startsWith("### ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: line.slice(4), bold: true })] }));
    } else if (line.startsWith("## ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: line.slice(3), bold: true })] }));
    } else if (line.startsWith("# ")) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: line.slice(2), bold: true })] }));
    } else if (line.startsWith("|") && line.endsWith("|")) {
      const tableLines: string[] = [];
      let j = i;
      while (j < lines.length && lines[j].startsWith("|")) {
        if (!lines[j].match(/^\|[-| :]+\|$/)) tableLines.push(lines[j]);
        j++;
      }
      i = j - 1;
      if (tableLines.length > 0) {
        const rows = tableLines.map((tl, rowIdx) => {
          const cells = tl.split("|").filter((_, ci) => ci > 0 && ci < tl.split("|").length - 1).map(c => c.trim());
          return new TableRow({
            children: cells.map(cell => new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: cell.trim(), bold: rowIdx === 0 })] })],
              width: { size: Math.floor(9000 / cells.length), type: WidthType.DXA },
            })),
          });
        });
        children.push(new Table({ rows, width: { size: 9000, type: WidthType.DXA } }));
      }
    } else if (line.match(/^[-*] /)) {
      children.push(new Paragraph({ bullet: { level: 0 }, children: [new TextRun(line.slice(2))] }));
    } else if (line.trim() === "") {
      children.push(new Paragraph({ children: [] }));
    } else {
      children.push(new Paragraph({ children: [new TextRun(line)] }));
    }
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  return await Packer.toBuffer(doc);
}

// ─── Build PPTX from scratch (fallback for non-PDF sources) ──────────────────
async function buildPptx(markdownText: string): Promise<Buffer> {
  const PptxGenJS = (await import("pptxgenjs")).default;
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  const sections = markdownText.split(/--- PAGE BREAK ---|\n(?=# )/);

  for (const section of sections) {
    if (!section.trim()) continue;
    const slide = pptx.addSlide();
    const lines = section.trim().split("\n").filter(l => l.trim());
    let titleLine = "";
    const bodyLines: string[] = [];

    for (const line of lines) {
      if (line.startsWith("# ") && !titleLine) titleLine = line.slice(2);
      else if (line.startsWith("## ") && !titleLine) titleLine = line.slice(3);
      else bodyLines.push(line.replace(/^#{1,3} /, "").replace(/\*\*/g, "").replace(/\*/g, ""));
    }

    slide.background = { color: "FFFFFF" };

    if (titleLine) {
      slide.addText(titleLine, { x: 0.5, y: 0.3, w: "90%", h: 1.0, fontSize: 24, bold: true, color: "1a1a2e", fontFace: "Calibri" });
    }

    if (bodyLines.length > 0) {
      const bodyText = bodyLines
        .filter(l => l.trim() && !l.match(/^\|[-| :]+\|$/))
        .map(l => {
          const isBullet = l.match(/^[-*] /);
          const text = l.replace(/^[-*\d.] /, "").trim();
          return { text: (isBullet ? "• " : "") + text, options: { fontSize: 14, color: "333333", breakLine: true } };
        });

      if (bodyText.length > 0) {
        slide.addText(bodyText, { x: 0.5, y: titleLine ? 1.5 : 0.5, w: "90%", h: titleLine ? 4.5 : 5.5, fontFace: "Calibri", valign: "top", wrap: true });
      }
    }
  }

  const rawBuffer = await pptx.write({ outputType: "nodebuffer" }) as unknown as Buffer;

  // Fix pptxgenjs Content_Types corruption (duplicate slideMaster entries)
  const AdmZip2 = (await import("adm-zip")).default;
  const fixZip2 = new AdmZip2(rawBuffer);
  const ctEntry2 = fixZip2.getEntry("[Content_Types].xml");
  if (ctEntry2) {
    let ctXml2 = ctEntry2.getData().toString("utf-8");
    let firstSeen2 = false;
    ctXml2 = ctXml2.replace(
      /<Override[^>]*\/ppt\/slideMasters\/slideMaster\d+\.xml[^>]*\/>/g,
      (match) => { if (!firstSeen2) { firstSeen2 = true; return match; } return ""; }
    );
    fixZip2.updateFile("[Content_Types].xml", Buffer.from(ctXml2, "utf-8"));
  }
  return fixZip2.toBuffer();
}

// ─── Build XLSX from scratch ──────────────────────────────────────────────────
async function buildXlsx(markdownText: string): Promise<Buffer> {
  const XLSX = await import("xlsx");
  const wb = XLSX.utils.book_new();
  const sections = markdownText.split(/--- PAGE BREAK ---/);

  sections.forEach((section, idx) => {
    const lines = section.trim().split("\n").filter(l => l.trim());
    const rows: string[][] = [];
    for (const line of lines) {
      if (line.startsWith("|") && line.endsWith("|")) {
        if (line.match(/^\|[-| :]+\|$/)) continue;
        const cells = line.split("|").filter((_, ci, arr) => ci > 0 && ci < arr.length - 1).map(c => c.trim());
        rows.push(cells);
      } else if (line.startsWith("# ") || line.startsWith("## ") || line.startsWith("### ")) {
        rows.push([line.replace(/^#{1,3} /, "")]);
      } else if (line.trim()) {
        rows.push([line.replace(/^[-*\d.] /, "").replace(/\*\*/g, "")]);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, `Sheet${idx + 1}`);
  });

  if (wb.SheetNames.length === 0) {
    const ws = XLSX.utils.aoa_to_sheet([[markdownText]]);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  }

  return Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
}

// ─── Simple markdown to HTML converter ───────────────────────────────────────
function markdownToHtml(md: string): string {
  return md
    .replace(/--- PAGE BREAK ---/g, "<hr>")
    .replace(/^### (.+)$/gm, "<h3>$1</h3>")
    .replace(/^## (.+)$/gm, "<h2>$1</h2>")
    .replace(/^# (.+)$/gm, "<h1>$1</h1>")
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/^[-*] (.+)$/gm, "<li>$1</li>")
    .replace(/(<li>.*<\/li>\n?)+/g, (m) => `<ul>${m}</ul>`)
    .replace(/^\d+\. (.+)$/gm, "<li>$1</li>")
    .replace(/\n\n/g, "</p><p>")
    .replace(/^(?!<[hul]|<li|<hr)(.+)$/gm, "<p>$1</p>");
}

// ─── High-fidelity PDF → DOCX (structural, editable) ────────────────────────
/**
 * Converts a PDF buffer to a fully editable DOCX using MuPDF.js for structural
 * text/image extraction and the docx package for reconstruction.
 *
 * Strategy per page:
 *  1. Extract all text lines (with x/y position, font name, size, bold/italic)
 *     and image blocks (bounding boxes) via MuPDF structured text.
 *  2. Detect multi-column layout by clustering x positions.
 *  3. For single-column pages: emit paragraphs sorted top-to-bottom.
 *  4. For two-column pages: build a borderless 2-column table, placing each
 *     text block in the appropriate column cell.
 *  5. Image regions are cropped from a 2× rendered pixmap and embedded as
 *     inline ImageRun elements at the correct position.
 *
 * All text is real, selectable, and translatable. No screenshots.
 */
export async function convertPdfToDocxWithPdf2Docx(
  pdfBuffer: Buffer,
  onProgress?: (msg: string, level?: string) => void
): Promise<Buffer> {
  onProgress?.("Extracting PDF structure for editable DOCX...", "info");

  const mupdf = await import("mupdf");
  const { createCanvas } = await import("@napi-rs/canvas");
  const {
    Document, Packer, Paragraph, TextRun, ImageRun,
    Table, TableRow, TableCell, WidthType, BorderStyle,
    AlignmentType, HeadingLevel,
  } = await import("docx");

  const doc = (mupdf as any).Document.openDocument(pdfBuffer, "application/pdf");
  const numPages = doc.countPages();

  // ── helpers ──────────────────────────────────────────────────────────────────

  /** Crop a region from a rendered pixmap and return a PNG Buffer */
  function cropRegionToPng(
    pixels: Uint8ClampedArray,
    stride: number,
    n: number,
    scale: number,
    bbox: { x: number; y: number; w: number; h: number }
  ): Buffer {
    const x0 = Math.round(bbox.x * scale);
    const y0 = Math.round(bbox.y * scale);
    const x1 = Math.round((bbox.x + bbox.w) * scale);
    const y1 = Math.round((bbox.y + bbox.h) * scale);
    const cropW = Math.max(1, x1 - x0);
    const cropH = Math.max(1, y1 - y0);
    const canvas = createCanvas(cropW, cropH);
    const ctx = canvas.getContext("2d");
    const imageData = ctx.createImageData(cropW, cropH);
    for (let y = 0; y < cropH; y++) {
      for (let x = 0; x < cropW; x++) {
        const srcIdx = (y0 + y) * stride + (x0 + x) * n;
        const dstIdx = (y * cropW + x) * 4;
        imageData.data[dstIdx]     = pixels[srcIdx];
        imageData.data[dstIdx + 1] = pixels[srcIdx + 1];
        imageData.data[dstIdx + 2] = pixels[srcIdx + 2];
        imageData.data[dstIdx + 3] = 255;
      }
    }
    ctx.putImageData(imageData, 0, 0);
    return canvas.toBuffer("image/png");
  }

  /** Convert a pt font size to half-points (docx uses half-points) */
  const ptToHalfPt = (pt: number) => Math.round(pt * 2);

  /** Map a font name to bold/italic flags */
  function parseFontFlags(fontName: string, weight: string, style: string) {
    const name = fontName.toLowerCase();
    const bold = weight === "bold" || name.includes("bold") || name.includes("-b") || name.includes("heavy") || name.includes("black");
    const italic = style === "italic" || name.includes("italic") || name.includes("oblique");
    return { bold, italic };
  }

  /** Build a docx Paragraph from a MuPDF text line */
  function lineToDocxParagraph(line: any): any {
    const { bold, italic } = parseFontFlags(
      line.font?.name ?? "",
      line.font?.weight ?? "normal",
      line.font?.style ?? "normal"
    );
    const fontSize = line.font?.size ?? 10;
    const run = new TextRun({
      text: line.text ?? "",
      bold,
      italics: italic,
      size: ptToHalfPt(fontSize),
      font: "Calibri", // closest web-safe fallback
    });
    return new Paragraph({
      children: [run],
      spacing: { before: 0, after: 40 },
    });
  }

  /** Build a no-border TableCell containing the given children */
  function noBorderCell(children: any[], widthPct: number): any {
    const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    return new TableCell({
      children,
      width: { size: widthPct, type: WidthType.PERCENTAGE },
      borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
    });
  }

  const sections: any[] = [];

  for (let pageIdx = 0; pageIdx < numPages; pageIdx++) {
    onProgress?.(`Processing page ${pageIdx + 1} of ${numPages}...`, "progress");

    const page = doc.loadPage(pageIdx);
    const bounds = page.getBounds(); // [x0, y0, x1, y1]
    const pageW = bounds[2] - bounds[0];

    // Render page at 2× for image cropping
    const scale = 2;
    const matrix = (mupdf as any).Matrix.scale(scale, scale);
    const pixmap = page.toPixmap(matrix, (mupdf as any).ColorSpace.DeviceRGB, false, true);
    const pixels = pixmap.getPixels() as Uint8ClampedArray;
    const stride = pixmap.getStride() as number;
    const n = pixmap.getNumberOfComponents() as number;

    // Extract structured text
    const stext = page.toStructuredText("preserve-spans,preserve-images");
    const stextJson = JSON.parse(stext.asJSON());
    const blocks: any[] = stextJson.blocks ?? [];

    // Detect column layout: find distinct x-clusters
    const xPositions = blocks
      .filter((b: any) => b.type === "text")
      .flatMap((b: any) => (b.lines ?? []).map((l: any) => l.x as number));
    const midX = pageW / 2;
    const leftCount  = xPositions.filter((x: number) => x < midX).length;
    const rightCount = xPositions.filter((x: number) => x >= midX).length;
    const isTwoColumn = leftCount > 2 && rightCount > 2 && Math.min(leftCount, rightCount) / Math.max(leftCount, rightCount) > 0.15;

    // Separate blocks into left/right columns (or single column)
    const leftBlocks: any[]  = [];
    const rightBlocks: any[] = [];
    const allBlocksSorted = [...blocks].sort((a: any, b: any) => (a.bbox?.y ?? 0) - (b.bbox?.y ?? 0));

    for (const block of allBlocksSorted) {
      const bx = block.bbox?.x ?? 0;
      if (!isTwoColumn || bx < midX) {
        leftBlocks.push(block);
      } else {
        rightBlocks.push(block);
      }
    }

    const blocksToChildren = (blist: any[]): any[] => {
      const children: any[] = [];
      for (const block of blist) {
        if (block.type === "image") {
          // Crop and embed the image
          try {
            const pngBuf = cropRegionToPng(pixels, stride, n, scale, block.bbox);
            const imgW = Math.round((block.bbox.w / pageW) * 650 * 0.9); // scale to ~90% of column width
            const imgH = Math.round(imgW * (block.bbox.h / block.bbox.w));
            children.push(
              new Paragraph({
                children: [
                  new ImageRun({ data: pngBuf, transformation: { width: imgW, height: imgH }, type: "png" }),
                ],
                spacing: { before: 0, after: 40 },
              })
            );
          } catch {
            // skip unrenderable images
          }
        } else if (block.type === "text") {
          for (const line of block.lines ?? []) {
            if ((line.text ?? "").trim()) {
              children.push(lineToDocxParagraph(line));
            }
          }
        }
      }
      return children;
    };

    let sectionChildren: any[];

    if (isTwoColumn) {
      // Build a borderless 2-column table
      const leftChildren  = blocksToChildren(leftBlocks);
      const rightChildren = blocksToChildren(rightBlocks);
      // Ensure cells are non-empty
      if (!leftChildren.length)  leftChildren.push(new Paragraph({ children: [] }));
      if (!rightChildren.length) rightChildren.push(new Paragraph({ children: [] }));

      const table = new Table({
        rows: [
          new TableRow({
            children: [
              noBorderCell(leftChildren,  60),
              noBorderCell(rightChildren, 40),
            ],
          }),
        ],
        width: { size: 100, type: WidthType.PERCENTAGE },
      });
      sectionChildren = [table];
    } else {
      sectionChildren = blocksToChildren(leftBlocks);
    }

    if (!sectionChildren.length) sectionChildren.push(new Paragraph({ children: [] }));

    sections.push({
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4 in twips
          margin: { top: 720, bottom: 720, left: 1080, right: 1080 },
        },
      },
      children: sectionChildren,
    });
  }

  const docxDoc = new Document({ sections });
  const docxBuffer = await Packer.toBuffer(docxDoc);
  onProgress?.("Editable DOCX created — text, fonts, and images preserved", "info");
  return docxBuffer;
}

// ─── Format conversion (for non-translated conversions) ──────────────────────
export async function convertDocument(
  inputBuffer: Buffer,
  inputFormat: SupportedFormat,
  outputFormat: SupportedFormat,
  filename: string
): Promise<Buffer> {
  if (inputFormat === outputFormat) return inputBuffer;

  // ── PDF → DOCX: use pdf2docx for layout-preserving conversion ──────────────
  if (inputFormat === "pdf" && outputFormat === "docx") {
    return await convertPdfToDocxWithPdf2Docx(inputBuffer);
  }

  // ── XLSX → CSV: use xlsx library for accurate conversion ──────────────────
  if (inputFormat === "xlsx" && outputFormat === "csv") {
    const XLSX = await import("xlsx");
    const wb = XLSX.read(inputBuffer, { type: "buffer" });
    // Convert all sheets to CSV, separated by sheet name headers
    const csvParts: string[] = [];
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(ws);
      if (wb.SheetNames.length > 1) {
        csvParts.push(`# Sheet: ${sheetName}\n${csv}`);
      } else {
        csvParts.push(csv);
      }
    }
    return Buffer.from(csvParts.join("\n\n"), "utf-8");
  }

  // ── DOCX/PPTX/XLSX → PDF: use LibreOffice for layout-preserving conversion ──
  if ((inputFormat === "docx" || inputFormat === "pptx" || inputFormat === "xlsx") && outputFormat === "pdf") {
    const { convertToPdfBuffer } = await import("./watermark");
    return await convertToPdfBuffer(inputBuffer, inputFormat);
  }

  const extracted = await extractTextFromBuffer(inputBuffer, inputFormat, filename);
  return await buildTranslatedDocument(extracted.text, outputFormat, filename);
}

// ─── Language list ────────────────────────────────────────────────────────────
export const LANGUAGES = [
  { code: "af", name: "Afrikaans" }, { code: "sq", name: "Albanian" }, { code: "am", name: "Amharic" },
  { code: "ar", name: "Arabic" }, { code: "hy", name: "Armenian" }, { code: "az", name: "Azerbaijani" },
  { code: "eu", name: "Basque" }, { code: "be", name: "Belarusian" }, { code: "bn", name: "Bengali" },
  { code: "bs", name: "Bosnian" }, { code: "bg", name: "Bulgarian" }, { code: "ca", name: "Catalan" },
  { code: "ceb", name: "Cebuano" }, { code: "zh-CN", name: "Chinese (Simplified)" }, { code: "zh-TW", name: "Chinese (Traditional)" },
  { code: "co", name: "Corsican" }, { code: "hr", name: "Croatian" }, { code: "cs", name: "Czech" },
  { code: "da", name: "Danish" }, { code: "nl", name: "Dutch" }, { code: "en", name: "English" },
  { code: "eo", name: "Esperanto" }, { code: "et", name: "Estonian" }, { code: "fi", name: "Finnish" },
  { code: "fr", name: "French" }, { code: "fy", name: "Frisian" }, { code: "gl", name: "Galician" },
  { code: "ka", name: "Georgian" }, { code: "de", name: "German" }, { code: "el", name: "Greek" },
  { code: "gu", name: "Gujarati" }, { code: "ht", name: "Haitian Creole" }, { code: "ha", name: "Hausa" },
  { code: "haw", name: "Hawaiian" }, { code: "he", name: "Hebrew" }, { code: "hi", name: "Hindi" },
  { code: "hmn", name: "Hmong" }, { code: "hu", name: "Hungarian" }, { code: "is", name: "Icelandic" },
  { code: "ig", name: "Igbo" }, { code: "id", name: "Indonesian" }, { code: "ga", name: "Irish" },
  { code: "it", name: "Italian" }, { code: "ja", name: "Japanese" }, { code: "jv", name: "Javanese" },
  { code: "kn", name: "Kannada" }, { code: "kk", name: "Kazakh" }, { code: "km", name: "Khmer" },
  { code: "rw", name: "Kinyarwanda" }, { code: "ko", name: "Korean" }, { code: "ku", name: "Kurdish" },
  { code: "ky", name: "Kyrgyz" }, { code: "lo", name: "Lao" }, { code: "la", name: "Latin" },
  { code: "lv", name: "Latvian" }, { code: "lt", name: "Lithuanian" }, { code: "lb", name: "Luxembourgish" },
  { code: "mk", name: "Macedonian" }, { code: "mg", name: "Malagasy" }, { code: "ms", name: "Malay" },
  { code: "ml", name: "Malayalam" }, { code: "mt", name: "Maltese" }, { code: "mi", name: "Maori" },
  { code: "mr", name: "Marathi" }, { code: "mn", name: "Mongolian" }, { code: "my", name: "Myanmar (Burmese)" },
  { code: "ne", name: "Nepali" }, { code: "no", name: "Norwegian" }, { code: "ny", name: "Nyanja (Chichewa)" },
  { code: "or", name: "Odia (Oriya)" }, { code: "ps", name: "Pashto" }, { code: "fa", name: "Persian" },
  { code: "pl", name: "Polish" }, { code: "pt", name: "Portuguese" }, { code: "pa", name: "Punjabi" },
  { code: "ro", name: "Romanian" }, { code: "ru", name: "Russian" }, { code: "sm", name: "Samoan" },
  { code: "gd", name: "Scots Gaelic" }, { code: "sr", name: "Serbian" }, { code: "st", name: "Sesotho" },
  { code: "sn", name: "Shona" }, { code: "sd", name: "Sindhi" }, { code: "si", name: "Sinhala (Sinhalese)" },
  { code: "sk", name: "Slovak" }, { code: "sl", name: "Slovenian" }, { code: "so", name: "Somali" },
  { code: "es", name: "Spanish" }, { code: "su", name: "Sundanese" }, { code: "sw", name: "Swahili" },
  { code: "sv", name: "Swedish" }, { code: "tl", name: "Tagalog (Filipino)" }, { code: "tg", name: "Tajik" },
  { code: "ta", name: "Tamil" }, { code: "tt", name: "Tatar" }, { code: "te", name: "Telugu" },
  { code: "th", name: "Thai" }, { code: "tr", name: "Turkish" }, { code: "tk", name: "Turkmen" },
  { code: "uk", name: "Ukrainian" }, { code: "ur", name: "Urdu" }, { code: "ug", name: "Uyghur" },
  { code: "uz", name: "Uzbek" }, { code: "vi", name: "Vietnamese" }, { code: "cy", name: "Welsh" },
  { code: "xh", name: "Xhosa" }, { code: "yi", name: "Yiddish" }, { code: "yo", name: "Yoruba" },
  { code: "zu", name: "Zulu" },
];
