import type { Express } from "express";
import { webhookCallback } from "grammy";
import { createTelegramBot } from "./bot";

let botStarted = false;

/**
 * Register Telegram bot routes and start webhook or long-polling.
 * Requires TELEGRAM_BOT_TOKEN. Optional TELEGRAM_WEBHOOK_URL for production webhook mode.
 */
export async function initTelegramBot(app: Express): Promise<void> {
  const token = process.env.TELEGRAM_BOT_TOKEN?.trim();
  if (!token) {
    console.log("[Telegram] TELEGRAM_BOT_TOKEN not set — bot disabled.");
    return;
  }

  const bot = createTelegramBot(token);
  const secret = process.env.TELEGRAM_WEBHOOK_SECRET?.trim();

  app.post(
    "/api/bots/telegram/webhook",
    webhookCallback(bot, "express", {
      secretToken: secret || undefined,
    }),
  );

  const webhookUrl = process.env.TELEGRAM_WEBHOOK_URL?.trim();

  if (webhookUrl) {
    await bot.api.setWebhook(webhookUrl, {
      secret_token: secret || undefined,
      allowed_updates: ["message", "callback_query"],
    });
    console.log(`[Telegram] Webhook registered: ${webhookUrl}`);
  } else if (!botStarted) {
    botStarted = true;
    void bot.start({
      onStart: () => console.log("[Telegram] Long polling started (set TELEGRAM_WEBHOOK_URL for webhook mode)"),
    });
  }
}
