import Stripe from "stripe";
import { ENV } from "./_core/env";

// ─── Stripe client ────────────────────────────────────────────────────────────
export const stripe = new Stripe(ENV.stripeSecretKey, {
  apiVersion: "2026-02-25.clover",
});

// ─── Pricing logic ────────────────────────────────────────────────────────────
// Download price = $2–$5 based on file size (bytes) and token cost (USD).
// Formula:
//   base = $2.00
//   size tier:  <1MB → +$0, 1–5MB → +$1, 5–20MB → +$2, >20MB → +$3
//   token tier: <$0.01 → +$0, $0.01–$0.05 → +$0.50, >$0.05 → +$1
//   capped at $5.00, floored at $2.00
export function calculateDownloadPrice(fileSizeBytes: number, conversionCostUsd: number): number {
  const MB = 1024 * 1024;
  let price = 2.0;

  // Size tier
  if (fileSizeBytes >= 20 * MB) price += 3.0;
  else if (fileSizeBytes >= 5 * MB) price += 2.0;
  else if (fileSizeBytes >= 1 * MB) price += 1.0;

  // Token cost tier
  if (conversionCostUsd >= 0.05) price += 1.0;
  else if (conversionCostUsd >= 0.01) price += 0.5;

  return Math.min(5.0, Math.max(2.0, price));
}

// ─── Create a Stripe Checkout Session for a download ─────────────────────────
export async function createDownloadCheckoutSession({
  jobId,
  userId,
  userEmail,
  userName,
  originalFileName,
  downloadPriceUsd,
  origin,
}: {
  jobId: number;
  userId: number;
  userEmail?: string | null;
  userName?: string | null;
  originalFileName: string;
  downloadPriceUsd: number;
  origin: string;
}): Promise<string> {
  // Sanitize: Stripe rejects empty strings for customer fields — use undefined instead
  // Handle both null and empty string cases
  const safeEmail = userEmail && userEmail.trim() ? userEmail.trim() : undefined;
  const safeName = userName && userName.trim() ? userName.trim() : undefined;

  // Create a Stripe Customer with a guaranteed non-empty name so that if
  // Stripe Link activates and pulls from this customer profile, the name
  // field won't be empty (which causes the customer_data[name] error).
  const customer = await stripe.customers.create({
    name: safeName || "Customer",
    ...(safeEmail ? { email: safeEmail } : {}),
    metadata: { user_id: userId.toString(), job_id: jobId.toString() },
  });

  const sessionParams: any = {
    mode: "payment",
    customer: customer.id,
    payment_method_types: ["card"],
    saved_payment_method_options: { payment_method_save: "disabled" },
    allow_promotion_codes: true,
    line_items: [
      {
        quantity: 1,
        price_data: {
          currency: "usd",
          unit_amount: Math.round(downloadPriceUsd * 100), // cents
          product_data: {
            name: "Document Download",
            description: `Converted document: ${originalFileName}`,
          },
        },
      },
    ],
    metadata: {
      job_id: jobId.toString(),
      user_id: userId.toString(),
      // Store for our own reference only — NOT passed as top-level Stripe fields
      // to avoid triggering Stripe Link auto-population of customer_data[name].
      ...(safeEmail ? { email: safeEmail } : {}),
      ...(safeName ? { name: safeName } : {}),
    },
    success_url: `${origin}/job/${jobId}?payment=success&session_id={CHECKOUT_SESSION_ID}`,
    cancel_url: `${origin}/job/${jobId}?payment=cancelled`,
  };

  // The checkout session is tied to a customer with a non-empty name.
  // This prevents the "customer_data[name] cannot be unset" error from
  // Stripe Link auto-fill. If the error still persists, disable Link in
  // the Stripe Dashboard: Settings → Payment methods → Link → Turn off.

  try {
    const session = await stripe.checkout.sessions.create(sessionParams);
    if (!session.url) throw new Error("Stripe did not return a checkout URL");
    return session.url;
  } catch (error: any) {
    console.error(`[Stripe] Checkout session creation failed:`, error.message);
    throw error;
  }
}
