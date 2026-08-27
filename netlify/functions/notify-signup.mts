export default async (req: Request) => {
  if (req.method !== "POST") return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405, headers: { "Content-Type": "application/json" } });

  const url = new URL(req.url);
  const secret = url.searchParams.get("secret");
  const expected = Netlify.env.get("SIGNUP_WEBHOOK_SECRET");
  if (!expected || secret !== expected) {
    return new Response(JSON.stringify({ error: "Unauthorized" }), { status: 401, headers: { "Content-Type": "application/json" } });
  }

  let payload: any;
  try {
    payload = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Invalid JSON" }), { status: 400, headers: { "Content-Type": "application/json" } });
  }

  const email = payload?.record?.email || "unknown";

  const token = Netlify.env.get("TELEGRAM_BOT_TOKEN");
  const chatId = Netlify.env.get("TELEGRAM_CHAT_ID");
  if (!token || !chatId) {
    return new Response(JSON.stringify({ error: "Telegram not configured" }), { status: 500, headers: { "Content-Type": "application/json" } });
  }

  const text = `🎉 New Ravuno signup\n\n${email}`;

  await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text }),
  });

  return new Response(JSON.stringify({ received: true }), { headers: { "Content-Type": "application/json" } });
};

export const config = { path: "/api/notify-signup" };
