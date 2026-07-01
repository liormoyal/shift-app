// api/monday.js — server-side Monday.com proxy.
// The API key never reaches the browser: it lives in MONDAY_API_KEY
// (NO VITE_ prefix) as a Vercel environment variable.
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "method_not_allowed" });
  }
  var apiKey = process.env.MONDAY_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "monday_key_not_configured" });
  }
  try {
    var query = (req.body && req.body.query) || "";
    if (!query) return res.status(400).json({ error: "missing_query" });

    var r = await fetch("https://api.monday.com/v2", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": apiKey,
      },
      body: JSON.stringify({ query: query }),
    });
    var data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(502).json({ error: "monday_upstream_error", detail: String(err) });
  }
}
