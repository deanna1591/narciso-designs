// Vercel serverless function — proxies the app's AI calls so your
// ANTHROPIC_API_KEY never touches the browser.
//
// Lives at  /api/anthropic  (repo-root /api folder is auto-detected
// by Vercel). The React app calls fetch("/api/anthropic", { body }).
//
// Set the key in Vercel → Project → Settings → Environment Variables:
//   ANTHROPIC_API_KEY = sk-ant-...

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY is not set" });
  }
  try {
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(req.body),
    });
    const data = await upstream.json();
    return res.status(upstream.status).json(data);
  } catch (e) {
    return res.status(500).json({ error: String(e) });
  }
}
