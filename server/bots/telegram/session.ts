/** In-memory conversation state per Telegram chat (resets on server restart). */

export type BotMode = "convert" | "translate";

export type SessionState =
  | "idle"
  | "menu"
  | "translate_language"
  | "translate_format"
  | "convert_format"
  | "confirm"
  | "awaiting_file"
  | "processing";

export interface TelegramSession {
  state: SessionState;
  mode?: BotMode;
  targetLanguage?: string;
  targetLanguageName?: string;
  outputFormat?: string;
  /** When true, keep output format same as uploaded file */
  keepSameFormat?: boolean;
}

const sessions = new Map<number, TelegramSession>();

export function getSession(chatId: number): TelegramSession {
  let s = sessions.get(chatId);
  if (!s) {
    s = { state: "idle" };
    sessions.set(chatId, s);
  }
  return s;
}

export function resetSession(chatId: number): void {
  sessions.set(chatId, { state: "idle" });
}

export function setSession(chatId: number, patch: Partial<TelegramSession>): TelegramSession {
  const s = { ...getSession(chatId), ...patch };
  sessions.set(chatId, s);
  return s;
}
