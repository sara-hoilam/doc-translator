import express, { type Express } from "express";
import { stripe } from "./stripe";
import { ENV } from "./_core/env";
import { getJobById, updateJob } from "./db";

export function registerStripeWebhook(app: Express) {
  // MUST use express.raw() before express.json() for signature verification
  app.post(
    "/api/stripe/webhook",
    express.raw({ type: "application/json" }),
    async (req, res) => {
      const sig = req.headers["stripe-signature"];

      // ── Test event shortcut (Stripe dashboard "Send test webhook") ─────────
      let event: ReturnType<typeof stripe.webhooks.constructEvent>;
      try {
        event = stripe.webhooks.constructEvent(
          req.body,
          sig as string,
          ENV.stripeWebhookSecret
        );
      } catch (err: any) {
        console.error("[Webhook] Signature verification failed:", err.message);
        return res.status(400).send(`Webhook Error: ${err.message}`);
      }

      // Test events — return immediately so the dashboard shows success
      if (event.id.startsWith("evt_test_")) {
        console.log("[Webhook] Test event detected, returning verification response");
        return res.json({ verified: true });
      }

      console.log(`[Webhook] Received event: ${event.type} (${event.id})`);

      if (event.type === "checkout.session.completed") {
        const session = event.data.object as any;
        const jobId = parseInt(session.metadata?.job_id ?? "0");
        if (jobId) {
          const job = await getJobById(jobId);
          if (job) {
            await updateJob(jobId, {
              paid: true,
              stripeSessionId: session.id,
            });
            console.log(`[Webhook] Job ${jobId} marked as paid (session: ${session.id})`);
          }
        }
      }

      res.json({ received: true });
    }
  );
}
