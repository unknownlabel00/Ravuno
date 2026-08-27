export default async (req: Request) => {
  if (req.method !== "POST") return new Response("Method not allowed", { status: 405 });

  let body: { priceId?: string; userId?: string; email?: string };
  try {
    body = await req.json();
  } catch {
    return new Response("Invalid JSON", { status: 400 });
  }
  const { priceId, userId, email } = body;
  if (!priceId || !userId || !email) return new Response("Missing fields", { status: 400 });

  const siteUrl = Netlify.env.get("URL") || "https://ravuno.netlify.app";
  const params = new URLSearchParams();
  params.set("mode", "subscription");
  params.set("line_items[0][price]", priceId);
  params.set("line_items[0][quantity]", "1");
  params.set("customer_email", email);
  params.set("client_reference_id", userId);
  params.set("success_url", `${siteUrl}/?checkout=success`);
  params.set("cancel_url", `${siteUrl}/?checkout=cancelled`);

  const secretKey = Netlify.env.get("STRIPE_SECRET_KEY");
  if (!secretKey) return new Response("Stripe not configured", { status: 500 });

  const res = await fetch("https://api.stripe.com/v1/checkout/sessions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${secretKey}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params.toString(),
  });

  const data = await res.json();
  if (!res.ok) return new Response(JSON.stringify(data), { status: 400 });
  return new Response(JSON.stringify({ url: data.url }), { headers: { "Content-Type": "application/json" } });
};

export const config = { path: "/api/create-checkout" };
