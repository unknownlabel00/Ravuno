import crypto from "node:crypto";

function verifyStripeSignature(payload: string, sigHeader: string, secret: string): boolean {
  const parts = Object.fromEntries(sigHeader.split(",").map((p) => p.split("=")));
  if (!parts.t || !parts.v1) return false;
  const signedPayload = `${parts.t}.${payload}`;
  const expected = crypto.createHmac("sha256", secret).update(signedPayload, "utf8").digest("hex");
  try {
    return crypto.timingSafeEqual(Buffer.from(expected), Buffer.from(parts.v1));
  } catch {
    return false;
  }
}

const SUPABASE_URL = "https://qixrcxwinvcwqddnvowq.supabase.co";

async function upsertProfileByUserId(serviceKey: string, userId: string, fields: Record<string, unknown>) {
  await fetch(`${SUPABASE_URL}/rest/v1/profiles?on_conflict=user_id`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: serviceKey,
      Authorization: `Bearer ${serviceKey}`,
      Prefer: "resolution=merge-duplicates,return=minimal",
    },
    body: JSON.stringify({ user_id: userId, ...fields, updated_at: new Date().toISOString() }),
  });
}

async function findUserIdByCustomerId(serviceKey: string, customerId: string): Promise<string | null> {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/profiles?stripe_customer_id=eq.${customerId}&select=user_id`, {
    headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}` },
  });
  const rows = await res.json();
  return rows?.[0]?.user_id || null;
}

export default async (req: Request) => {
  const sig = req.headers.get("stripe-signature") || "";
  const payload = await req.text();
  const webhookSecret = Netlify.env.get("STRIPE_WEBHOOK_SECRET") || "";
  const serviceKey = Netlify.env.get("SUPABASE_SERVICE_ROLE_KEY") || "";

  if (!verifyStripeSignature(payload, sig, webhookSecret)) {
    return new Response("Invalid signature", { status: 400 });
  }
  if (!serviceKey) {
    return new Response("Server not configured", { status: 500 });
  }

  const event = JSON.parse(payload);

  try {
    if (event.type === "checkout.session.completed") {
      const session = event.data.object;
      const userId = session.client_reference_id;
      if (userId) {
        await upsertProfileByUserId(serviceKey, userId, {
          stripe_customer_id: session.customer,
          subscription_status: "active",
        });
      }
    } else if (event.type === "customer.subscription.updated") {
      const sub = event.data.object;
      const userId = await findUserIdByCustomerId(serviceKey, sub.customer);
      if (userId) await upsertProfileByUserId(serviceKey, userId, { subscription_status: sub.status, price_id: sub.items?.data?.[0]?.price?.id || null });
    } else if (event.type === "customer.subscription.deleted") {
      const sub = event.data.object;
      const userId = await findUserIdByCustomerId(serviceKey, sub.customer);
      if (userId) await upsertProfileByUserId(serviceKey, userId, { subscription_status: "cancelled" });
    } else if (event.type === "invoice.payment_failed") {
      const invoice = event.data.object;
      const userId = await findUserIdByCustomerId(serviceKey, invoice.customer);
      if (userId) await upsertProfileByUserId(serviceKey, userId, { subscription_status: "past_due" });
    }
  } catch (e) {
    // still acknowledge receipt so Stripe doesn't endlessly retry a transient error
    console.error("webhook handling error", e);
  }

  return new Response(JSON.stringify({ received: true }), { headers: { "Content-Type": "application/json" } });
};

export const config = { path: "/.netlify/functions/stripe-webhook" };
