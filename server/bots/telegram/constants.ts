import { LANGUAGES, SUPPORTED_OUTPUT_FORMATS } from "../../docProcessor";

export const TRIGGER_KEYWORD = "pdfgodwork";

/** Telegram Bot API document upload limit */
export const TELEGRAM_MAX_FILE_BYTES = 20 * 1024 * 1024;

export const POPULAR_LANGUAGES = LANGUAGES.filter(l =>
  [
    "en",
    "es",
    "fr",
    "de",
    "it",
    "pt",
    "zh-CN",
    "zh-TW",
    "ja",
    "ko",
    "ar",
    "hi",
    "ru",
    "nl",
    "pl",
    "vi",
    "th",
    "id",
    "tr",
    "sv",
  ].includes(l.code),
);

export const OUTPUT_FORMAT_LABELS: Record<string, string> = {
  pdf: "PDF",
  docx: "Word (DOCX)",
  pptx: "PowerPoint (PPTX)",
  xlsx: "Excel (XLSX)",
  txt: "Plain text",
  csv: "CSV",
};

export const BOT_OUTPUT_FORMATS = [...SUPPORTED_OUTPUT_FORMATS];
