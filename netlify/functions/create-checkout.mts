export default async (req: Request) => {
  if (req.method !== "POST") return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405, headers: { "Content-Type": "application/json" } });

  let body: { amount?: number; userId?: string; email?: string };
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Invalid JSON" }), { status: 400, headers: { "Content-Type": "application/json" } });
  }
  const { amount, userId, email } = body;
  if (!amount || amount <= 0 || !userId || !email) {
    return new Response(JSON.stringify({ error: "Missing fields", received: { amount, userId: !!userId, email: !!email } }), { status: 400, headers: { "Content-Type": "application/json" } });
  }

  const siteUrl = Netlify.env.get("URL") || "https://ravuno.netlify.app";
  const params = new URLSearchParams();
  params.set("mode", "payment"); // one-time support payment, not a subscription
  params.set("line_items[0][price_data][currency]", "myr");
  params.set("line_items[0][price_data][product_data][name]", "Support Ravuno");
  params.set("line_items[0][price_data][unit_amount]", String(Math.round(amount * 100)));
  params.set("line_items[0][quantity]", "1");
  params.set("customer_email", email);
  params.set("client_reference_id", userId);
  params.set("success_url", `${siteUrl}/?support=thanks`);
  params.set("cancel_url", `${siteUrl}/?support=cancelled`);

  const secretKey = Netlify.env.get("STRIPE_SECRET_KEY");
  if (!secretKey) return new Response(JSON.stringify({ error: "Stripe not configured (missing STRIPE_SECRET_KEY)" }), { status: 500, headers: { "Content-Type": "application/json" } });

  const res = await fetch("https://api.stripe.com/v1/checkout/sessions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${secretKey}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params.toString(),
  });

  const data = await res.json();
  if (!res.ok) return new Response(JSON.stringify({ error: data?.error?.message || "Stripe error", stripe: data }), { status: 400, headers: { "Content-Type": "application/json" } });
  return new Response(JSON.stringify({ url: data.url }), { headers: { "Content-Type": "application/json" } });
};

export const config = { path: "/api/create-checkout" };
