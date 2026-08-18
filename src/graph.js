import { GRAPH_BASE, FETCH_CONCURRENCY } from "./config.js";

const LIST_SELECT = [
  "id",
  "subject",
  "from",
  "toRecipients",
  "ccRecipients",
  "receivedDateTime",
  "hasAttachments",
].join(",");

async function graphGet(url, token) {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      // Ask Graph for plain text instead of HTML so no parsing is needed.
      Prefer: 'outlook.body-content-type="text"',
    },
  });
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status} ${res.statusText}${detail ? " - " + detail.slice(0, 300) : ""}`);
  }
  return res.json();
}

// Graph rejects $filter + $orderby together on conversationId, so sorting is
// done client-side in format.js instead.
export async function listConversationMessages(conversationId, token) {
  const filter = encodeURIComponent(`conversationId eq '${conversationId.replace(/'/g, "''")}'`);
  let url = `${GRAPH_BASE}/me/messages?$filter=${filter}&$select=${LIST_SELECT}&$top=50`;

  const out = [];
  while (url) {
    const page = await graphGet(url, token);
    out.push(...(page.value || []));
    url = page["@odata.nextLink"] || null;
  }
  return out;
}

// uniqueBody is the message minus the quoted history. It is not returned on
// collection queries, so each message is fetched individually.
export async function fetchUniqueBody(messageId, token) {
  const url = `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}?$select=uniqueBody`;
  try {
    const r = await graphGet(url, token);
    return r.uniqueBody?.content ?? "";
  } catch {
    // Fall back to the full body if uniqueBody is unavailable for this item.
    try {
      const r = await graphGet(
        `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}?$select=body`,
        token
      );
      return r.body?.content ?? "";
    } catch {
      return "";
    }
  }
}

export async function fetchAttachmentNames(messageId, token) {
  const url = `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/attachments?$select=name`;
  try {
    const r = await graphGet(url, token);
    return (r.value || []).map((a) => a.name).filter(Boolean);
  } catch {
    return [];
  }
}

// Simple bounded pool so a long thread does not fire 60 requests at once.
export async function mapLimit(items, fn, limit = FETCH_CONCURRENCY) {
  const results = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (true) {
      const i = next++;
      if (i >= items.length) return;
      results[i] = await fn(items[i], i);
    }
  });
  await Promise.all(workers);
  return results;
}
