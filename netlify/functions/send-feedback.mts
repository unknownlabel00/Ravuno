export default async (req: Request) => {
  if (req.method !== "POST") return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405, headers: { "Content-Type": "application/json" } });

  let body: { message?: string; email?: string };
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Invalid JSON" }), { status: 400, headers: { "Content-Type": "application/json" } });
  }
  const { message, email } = body;
  if (!message || !message.trim()) {
    return new Response(JSON.stringify({ error: "Message is empty" }), { status: 400, headers: { "Content-Type": "application/json" } });
  }

  const token = Netlify.env.get("TELEGRAM_BOT_TOKEN");
  const chatId = Netlify.env.get("TELEGRAM_CHAT_ID");
  if (!token || !chatId) {
    return new Response(JSON.stringify({ error: "Telegram not configured" }), { status: 500, headers: { "Content-Type": "application/json" } });
  }

  const text = `📩 New Ravuno feedback\n\nFrom: ${email || "unknown"}\n\n${message.trim()}`;

  const res = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text }),
  });

  const data = await res.json();
  if (!data.ok) return new Response(JSON.stringify({ error: data.description || "Telegram error" }), { status: 400, headers: { "Content-Type": "application/json" } });
  return new Response(JSON.stringify({ success: true }), { headers: { "Content-Type": "application/json" } });
};

export const config = { path: "/api/send-feedback" };
