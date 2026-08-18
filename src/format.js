import { NEWEST_FIRST } from "./config.js";

function addr(recipient) {
  const name = recipient?.emailAddress?.name?.trim();
  const mail = recipient?.emailAddress?.address?.trim();
  if (name && mail && name.toLowerCase() !== mail.toLowerCase()) return `${name} <${mail}>`;
  return mail || name || "(unknown)";
}

function addrList(list) {
  return (list || []).map(addr).join(", ");
}

function fmtDate(iso) {
  if (!iso) return "(no date)";
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function cleanSubject(s) {
  let out = (s || "").trim();
  // Strip repeated RE:/FW:/FWD: prefixes, including localised spacings.
  while (/^(re|fw|fwd)\s*:/i.test(out)) out = out.replace(/^(re|fw|fwd)\s*:/i, "").trim();
  return out || "(no subject)";
}

function tidyBody(text) {
  return (text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+$/gm, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function sortMessages(messages) {
  const sorted = [...messages].sort(
    (a, b) => new Date(a.receivedDateTime || 0) - new Date(b.receivedDateTime || 0)
  );
  return NEWEST_FIRST ? sorted.reverse() : sorted;
}

export function buildMarkdown(messages, threadSubject) {
  const lines = [];
  lines.push(`# Mail Thread: ${cleanSubject(threadSubject)}`);
  lines.push("");
  lines.push(`Exported: ${fmtDate(new Date().toISOString())}`);
  lines.push(`Messages: ${messages.length}`);
  lines.push("");
  lines.push("---");

  messages.forEach((m, i) => {
    lines.push("");
    lines.push(`## [${i + 1}] ${cleanSubject(m.subject)}`);
    lines.push("");
    lines.push(`**From:** ${addr(m.from)}`);
    const to = addrList(m.toRecipients);
    if (to) lines.push(`**To:** ${to}`);
    const cc = addrList(m.ccRecipients);
    if (cc) lines.push(`**Cc:** ${cc}`);
    lines.push(`**Date:** ${fmtDate(m.receivedDateTime)}`);
    if (m.attachmentNames?.length) {
      lines.push(`**Attachments:** ${m.attachmentNames.join(", ")}`);
    }
    lines.push("");
    lines.push(tidyBody(m.bodyText) || "_(empty message)_");
    lines.push("");
    lines.push("---");
  });

  return lines.join("\n") + "\n";
}

export function buildFileName(subject) {
  const safe = cleanSubject(subject)
    .replace(/[\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 60);
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
  return `${safe || "thread"}_${stamp}.md`;
}
