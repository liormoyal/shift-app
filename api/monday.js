// api/monday.js — structured Monday.com proxy.
//
// The browser never sends raw GraphQL. It sends a small typed intent and the
// server builds the exact query. Only these operations, this one board, and
// this fixed set of columns are possible, so the endpoint can't be used as an
// open read/write gateway to the board even though its URL is public.
//
// Accepted intents:
//   { action: "list", cursor? }                 -> paginated { id, name } read
//   { action: "update", itemId, columnValues }  -> column update on one item
//
// The API key lives in MONDAY_API_KEY (no VITE_ prefix). The board id lives in
// MONDAY_BOARD_ID (falls back to the known board).

const BOARD_ID = String(process.env.MONDAY_BOARD_ID || "18419606261");

// Only columns this app is allowed to write. Anything else is dropped.
const ALLOWED_COLUMNS = {
  color_mm4qvjcs: true, // status
  text_mm4qxbn0: true,  // day
  text_mm4qsdfw: true,  // hours
};

// Best-effort in-memory rate limit. Serverless instances aren't shared, so this
// isn't distributed — it's a cheap speed bump against a single client hammering
// a warm instance. The structured API above is the real blast-radius control.
const RATE = { windowMs: 60000, max: 600 };
let hits = [];
function rateLimited() {
  const now = Date.now();
  hits = hits.filter((t) => now - t < RATE.windowMs);
  if (hits.length >= RATE.max) return true;
  hits.push(now);
  return false;
}

async function monday(query, apiKey) {
  const r = await fetch("https://api.monday.com/v2", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: apiKey },
    body: JSON.stringify({ query }),
  });
  return r.json();
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "method_not_allowed" });
  }
  const apiKey = process.env.MONDAY_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "monday_key_not_configured" });
  }
  if (rateLimited()) {
    return res.status(429).json({ error: "rate_limited" });
  }

  const body = req.body || {};
  const action = body.action;

  try {
    if (action === "list") {
      let query;
      if (body.cursor != null && body.cursor !== "") {
        if (typeof body.cursor !== "string") {
          return res.status(400).json({ error: "bad_cursor" });
        }
        // JSON.stringify safely quotes/escapes the cursor into the query.
        query =
          "query { next_items_page(limit: 500, cursor: " +
          JSON.stringify(body.cursor) +
          ") { cursor items { id name } } }";
      } else {
        query =
          "{ boards(ids: " +
          Number(BOARD_ID) +
          ") { items_page(limit: 500) { cursor items { id name } } } }";
      }
      const data = await monday(query, apiKey);
      return res.status(200).json(data);
    }

    if (action === "update") {
      const itemId = String(body.itemId == null ? "" : body.itemId);
      if (!/^\d+$/.test(itemId)) {
        return res.status(400).json({ error: "bad_item_id" });
      }
      const cv =
        body.columnValues && typeof body.columnValues === "object"
          ? body.columnValues
          : null;
      if (!cv) {
        return res.status(400).json({ error: "missing_column_values" });
      }
      // Keep only allow-listed columns; ignore anything else the caller sent.
      const safe = {};
      Object.keys(cv).forEach(function (k) {
        if (ALLOWED_COLUMNS[k]) safe[k] = cv[k];
      });
      if (Object.keys(safe).length === 0) {
        return res.status(400).json({ error: "no_allowed_columns" });
      }
      // Double-encode the column values into a GraphQL string literal (safe).
      const colValStr = JSON.stringify(JSON.stringify(safe));
      const query =
        "mutation { change_multiple_column_values(board_id: " +
        Number(BOARD_ID) +
        ", item_id: " +
        itemId +
        ", column_values: " +
        colValStr +
        ") { id } }";
      const data = await monday(query, apiKey);
      return res.status(200).json(data);
    }

    return res.status(400).json({ error: "unknown_action" });
  } catch (err) {
    return res.status(502).json({ error: "monday_upstream_error", detail: String(err) });
  }
}
