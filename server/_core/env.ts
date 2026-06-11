export const ENV = {
  // ── Google Cloud ──────────────────────────────────────────────────────────
  gcsBucketName: process.env.GCS_BUCKET_NAME ?? "",
  gcsProjectId: process.env.GCS_PROJECT_ID ?? "",
  // Path to GCP service-account JSON key file (local dev).
  // In production (Cloud Run) leave empty and use Workload Identity / ADC instead.
  gcsKeyFilePath: process.env.GCS_KEY_FILE_PATH ?? "",

  // ── LLM API keys ─────────────────────────────────────────────────────────
  // Provide the key(s) for the model(s) you want to use.
  openaiApiKey: process.env.OPENAI_API_KEY ?? "",
  anthropicApiKey: process.env.ANTHROPIC_API_KEY ?? "",
  // Cursor API key (OpenAI-compatible chat completions).
  cursorApiKey: process.env.CURSOR_API_KEY ?? "",
  // Base URL for Cursor chat API (include /v1 if your proxy expects it).
  cursorApiBaseUrl: process.env.CURSOR_API_BASE_URL ?? "https://api.cursor.com/v1",
  // Which model is used when no modelId is specified by the caller.
  defaultLlmModel: process.env.DEFAULT_LLM_MODEL ?? "composer-2.5",

  // ── Telegram bot ──────────────────────────────────────────────────────────
  telegramBotToken: process.env.TELEGRAM_BOT_TOKEN ?? "",

  // ── App / Auth ────────────────────────────────────────────────────────────
  cookieSecret: process.env.JWT_SECRET ?? "",
  databaseUrl: process.env.DATABASE_URL ?? "",
  isProduction: process.env.NODE_ENV === "production",

  // ── Legacy Manus fields (kept so existing imports don't break) ────────────
  // These will be empty strings in the self-hosted version.
  appId: process.env.VITE_APP_ID ?? "",
  oAuthServerUrl: process.env.OAUTH_SERVER_URL ?? "",
  ownerOpenId: process.env.OWNER_OPEN_ID ?? "",
  forgeApiUrl: process.env.BUILT_IN_FORGE_API_URL ?? "",
  forgeApiKey: process.env.BUILT_IN_FORGE_API_KEY ?? "",
};
