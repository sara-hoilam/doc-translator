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
  // Google AI Studio key (for Gemini models). Also accepts a GCP API key.
  googleAiApiKey: process.env.GOOGLE_AI_API_KEY ?? "",
  // Which model is used when no modelId is specified by the caller.
  defaultLlmModel: process.env.DEFAULT_LLM_MODEL ?? "gemini-2.5-flash",

  // ── Stripe ────────────────────────────────────────────────────────────────
  stripeSecretKey: process.env.STRIPE_SECRET_KEY ?? "",
  stripeWebhookSecret: process.env.STRIPE_WEBHOOK_SECRET ?? "",

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
