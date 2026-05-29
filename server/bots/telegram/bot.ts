import { Bot, InlineKeyboard, InputFile, type Context } from "grammy";
import { getFormatFromFilename, getMimeType } from "../../docProcessor";
import { startDocumentJob, waitForJobDone } from "../../documentJob";
import { TRIGGER_KEYWORD, TELEGRAM_MAX_FILE_BYTES, POPULAR_LANGUAGES, OUTPUT_FORMAT_LABELS, BOT_OUTPUT_FORMATS } from "./constants";
import { getSession, resetSession, setSession, type TelegramSession } from "./session";

function isTrigger(text: string): boolean {
  return text.trim().toLowerCase() === TRIGGER_KEYWORD;
}

function mainMenuKeyboard() {
  return new InlineKeyboard()
    .text("1️⃣ Convert file format", "mode:convert")
    .row()
    .text("2️⃣ Translate document", "mode:translate");
}

function languageKeyboard() {
  const kb = new InlineKeyboard();
  POPULAR_LANGUAGES.forEach((lang, i) => {
    if (i > 0 && i % 2 === 0) kb.row();
    kb.text(lang.name, `lang:${lang.code}`);
  });
  kb.row().text("« Back", "nav:menu");
  return kb;
}

function formatKeyboard(prefix: "fmt" | "tfmt", includeSame = false) {
  const kb = new InlineKeyboard();
  if (includeSame) {
    kb.text("Keep same format as upload", `${prefix}:same`).row();
  }
  BOT_OUTPUT_FORMATS.forEach((fmt, i) => {
    if (i > 0 && i % 2 === 0) kb.row();
    kb.text(OUTPUT_FORMAT_LABELS[fmt] ?? fmt.toUpperCase(), `${prefix}:${fmt}`);
  });
  kb.row().text("« Back", "nav:menu");
  return kb;
}

function confirmKeyboard() {
  return new InlineKeyboard().text("✅ Confirm & upload file", "confirm:yes").text("✖️ Start over", "confirm:no");
}

function summaryText(session: TelegramSession): string {
  const lines: string[] = ["<b>Please confirm your request:</b>", ""];
  if (session.mode === "translate") {
    lines.push(`• Service: <b>Translate document</b>`);
    lines.push(`• Target language: <b>${session.targetLanguageName ?? "—"}</b>`);
  } else {
    lines.push(`• Service: <b>Convert file format</b>`);
  }
  if (session.keepSameFormat) {
    lines.push(`• Output format: <b>Same as your file</b>`);
  } else {
    lines.push(`• Output format: <b>${(session.outputFormat ?? "auto").toUpperCase()}</b>`);
  }
  lines.push("", "Tap <b>Confirm & upload file</b>, then send your document as a file attachment.");
  return lines.join("\n");
}

async function showMainMenu(ctx: Context, chatId: number) {
  setSession(chatId, { state: "menu", mode: undefined, targetLanguage: undefined, targetLanguageName: undefined, outputFormat: undefined, keepSameFormat: false });
  await ctx.reply(
    "Welcome to <b>PDFGodWork</b> 👋\n\nWhat would you like to do?",
    { parse_mode: "HTML", reply_markup: mainMenuKeyboard() },
  );
}

async function downloadTelegramFile(bot: Bot, fileId: string): Promise<Buffer> {
  const file = await bot.api.getFile(fileId);
  if (!file.file_path) throw new Error("Could not read file from Telegram.");
  const token = bot.token;
  const url = `https://api.telegram.org/file/bot${token}/${file.file_path}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("Failed to download file from Telegram.");
  const ab = await res.arrayBuffer();
  return Buffer.from(ab);
}

async function processUploadedFile(
  ctx: Context,
  chatId: number,
  buffer: Buffer,
  filename: string,
  mimeType: string,
) {
  const session = getSession(chatId);
  if (session.state !== "awaiting_file") {
    await ctx.reply(`Send <code>${TRIGGER_KEYWORD}</code> to start a new request.`, { parse_mode: "HTML" });
    return;
  }

  if (buffer.length > TELEGRAM_MAX_FILE_BYTES) {
    await ctx.reply(`File is too large for Telegram (max ${TELEGRAM_MAX_FILE_BYTES / 1024 / 1024} MB). Try a smaller file or use the website.`);
    resetSession(chatId);
    return;
  }

  setSession(chatId, { state: "processing" });
  const statusMsg = await ctx.reply("⏳ Processing your document… This may take a few minutes.");

  try {
    const inputFormat = getFormatFromFilename(filename);
    let outputFormat = session.keepSameFormat ? inputFormat : session.outputFormat;

    if (session.mode === "translate" && !session.targetLanguage) {
      throw new Error("Missing target language. Send pdfgodwork to start again.");
    }

    if (session.mode === "convert" && !outputFormat && !session.keepSameFormat) {
      throw new Error("Missing output format. Send pdfgodwork to start again.");
    }

    if (session.mode === "translate" && session.keepSameFormat) {
      outputFormat = inputFormat;
    }

    const { jobId } = await startDocumentJob({
      filename,
      mimeType,
      buffer,
      outputFormat: outputFormat ?? undefined,
      targetLanguage: session.mode === "translate" ? session.targetLanguage : undefined,
      targetLanguageName: session.mode === "translate" ? session.targetLanguageName : undefined,
      skipPayment: true,
    });

    const job = await waitForJobDone(jobId);

    if (job?.status === "error") {
      throw new Error(job.errorMessage ?? "Processing failed.");
    }
    if (job?.status === "paused") {
      throw new Error("Processing took too long and was paused. Please try a smaller file or use the website.");
    }
    if (job?.status === "cancelled") {
      throw new Error("Processing was cancelled.");
    }
    if (!job?.outputFileUrl) {
      throw new Error("No output file was produced.");
    }

    const outRes = await fetch(job.outputFileUrl);
    if (!outRes.ok) throw new Error("Could not fetch processed file.");
    const outBuf = Buffer.from(await outRes.arrayBuffer());

    const ext = job.outputFormat ?? getFormatFromFilename(filename);
    const base = filename.replace(/\.[^.]+$/, "") || "document";
    const outName = `${base}_pdfgodwork.${ext}`;

    await ctx.api.editMessageText(chatId, statusMsg.message_id, "✅ Done! Sending your file…");
    await ctx.replyWithDocument(new InputFile(outBuf, outName), {
      caption: session.mode === "translate"
        ? `Translated to ${session.targetLanguageName} (${ext.toUpperCase()})`
        : `Converted to ${ext.toUpperCase()}`,
    });
  } catch (err: any) {
    const msg = err?.message ?? "Unknown error";
    await ctx.api.editMessageText(chatId, statusMsg.message_id, `❌ ${msg}`).catch(() => {});
    await ctx.reply(`Something went wrong: ${msg}\n\nSend <code>${TRIGGER_KEYWORD}</code> to try again.`, { parse_mode: "HTML" });
  } finally {
    resetSession(chatId);
  }
}

export function createTelegramBot(token: string): Bot {
  const bot = new Bot(token);

  bot.command("start", async ctx => {
    const chatId = ctx.chat?.id;
    if (!chatId) return;
    resetSession(chatId);
    await ctx.reply(
      "Hi! I'm the <b>PDFGodWork</b> bot — translate documents or convert file formats.\n\n" +
        `To begin, send the keyword <code>${TRIGGER_KEYWORD}</code> (any time you're ready).`,
      { parse_mode: "HTML" },
    );
  });

  bot.on("message:text", async ctx => {
    const chatId = ctx.chat?.id;
    if (!chatId || !ctx.message.text) return;

    if (isTrigger(ctx.message.text)) {
      await showMainMenu(ctx, chatId);
      return;
    }

    const session = getSession(chatId);
    if (session.state === "idle") {
      await ctx.reply(
        `Send <code>${TRIGGER_KEYWORD}</code> when you want to translate or convert a document.`,
        { parse_mode: "HTML" },
      );
    }
  });

  bot.on("callback_query:data", async ctx => {
    const chatId = ctx.chat?.id;
    if (!chatId) {
      await ctx.answerCallbackQuery();
      return;
    }

    const data = ctx.callbackQuery.data;
    await ctx.answerCallbackQuery();

    if (data === "nav:menu") {
      await showMainMenu(ctx, chatId);
      return;
    }

    if (data === "mode:convert") {
      setSession(chatId, { state: "convert_format", mode: "convert" });
      await ctx.editMessageText("Choose the <b>output file format</b>:", {
        parse_mode: "HTML",
        reply_markup: formatKeyboard("fmt", false),
      });
      return;
    }

    if (data === "mode:translate") {
      setSession(chatId, { state: "translate_language", mode: "translate" });
      await ctx.editMessageText("Which language should we translate <b>to</b>?", {
        parse_mode: "HTML",
        reply_markup: languageKeyboard(),
      });
      return;
    }

    if (data.startsWith("lang:")) {
      const code = data.slice(5);
      const lang = POPULAR_LANGUAGES.find(l => l.code === code);
      if (!lang) return;
      setSession(chatId, {
        state: "translate_format",
        targetLanguage: lang.code,
        targetLanguageName: lang.name,
      });
      await ctx.editMessageText(
        `Target language: <b>${lang.name}</b>\n\nChoose output format (or keep the same as your upload):`,
        { parse_mode: "HTML", reply_markup: formatKeyboard("tfmt", true) },
      );
      return;
    }

    if (data.startsWith("fmt:") || data.startsWith("tfmt:")) {
      const fmt = data.split(":")[1];
      if (fmt === "same") {
        setSession(chatId, { keepSameFormat: true, outputFormat: undefined, state: "confirm" });
      } else {
        setSession(chatId, { keepSameFormat: false, outputFormat: fmt, state: "confirm" });
      }
      const session = getSession(chatId);
      await ctx.editMessageText(summaryText(session), {
        parse_mode: "HTML",
        reply_markup: confirmKeyboard(),
      });
      return;
    }

    if (data === "confirm:yes") {
      setSession(chatId, { state: "awaiting_file" });
      await ctx.editMessageText(
        "📎 <b>Upload your file now</b>\n\nAttach your document here (PDF, DOCX, PPTX, XLSX, TXT, or an image). " +
          "Use “send as file” rather than a photo when possible.",
        { parse_mode: "HTML" },
      );
      return;
    }

    if (data === "confirm:no") {
      await showMainMenu(ctx, chatId);
    }
  });

  bot.on("message:document", async ctx => {
    const chatId = ctx.chat?.id;
    if (!chatId) return;
    const doc = ctx.message.document;
    if (!doc) return;

    const filename = doc.file_name ?? "document.bin";
    const mimeType = doc.mime_type ?? getMimeType(getFormatFromFilename(filename));
    const buffer = await downloadTelegramFile(bot, doc.file_id);
    await processUploadedFile(ctx, chatId, buffer, filename, mimeType);
  });

  bot.on("message:photo", async ctx => {
    const chatId = ctx.chat?.id;
    if (!chatId) return;
    const session = getSession(chatId);
    if (session.state !== "awaiting_file") return;

    const photos = ctx.message.photo;
    const largest = photos[photos.length - 1];
    const buffer = await downloadTelegramFile(bot, largest.file_id);
    await processUploadedFile(ctx, chatId, buffer, "photo.jpg", "image/jpeg");
  });

  return bot;
}
