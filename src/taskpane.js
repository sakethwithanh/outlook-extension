import { getGraphToken } from "./auth.js";
import { listConversationMessages, fetchUniqueBody, fetchAttachmentNames, mapLimit } from "./graph.js";
import { buildMarkdown, buildFileName, sortMessages, cleanSubject } from "./format.js";

const els = {};
let markdown = "";
let fileName = "thread.md";

function setStatus(text, kind = "") {
  els.status.textContent = text;
  els.status.className = "status" + (kind ? " " + kind : "");
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  els.subject = document.getElementById("subject");
  els.status = document.getElementById("status");
  els.export = document.getElementById("btn-export");
  els.copy = document.getElementById("btn-copy");
  els.previewWrap = document.getElementById("preview-wrap");
  els.preview = document.getElementById("preview");

  els.export.addEventListener("click", onExport);
  els.copy.addEventListener("click", onCopy);

  const item = Office.context.mailbox.item;
  els.subject.textContent = item?.subject ? cleanSubject(item.subject) : "(no subject)";

  if (!item?.conversationId) {
    setStatus("No conversation found for this item.", "error");
    return;
  }

  setStatus("Ready. Click Export to build the thread.");
  els.export.disabled = false;
});

async function buildThread() {
  const item = Office.context.mailbox.item;

  setStatus("Signing in…");
  const token = await getGraphToken();

  setStatus("Finding messages in this conversation…");
  const listed = await listConversationMessages(item.conversationId, token);

  if (listed.length === 0) {
    throw new Error("No messages returned. The conversation may have been moved or deleted.");
  }

  const ordered = sortMessages(listed);

  setStatus(`Fetching ${ordered.length} message${ordered.length === 1 ? "" : "s"}…`);
  await mapLimit(ordered, async (m) => {
    m.bodyText = await fetchUniqueBody(m.id, token);
    m.attachmentNames = m.hasAttachments ? await fetchAttachmentNames(m.id, token) : [];
  });

  markdown = buildMarkdown(ordered, item.subject);
  fileName = buildFileName(item.subject);

  els.preview.value = markdown;
  els.previewWrap.hidden = false;
  els.copy.disabled = false;

  return ordered.length;
}

async function onExport() {
  els.export.disabled = true;
  els.copy.disabled = true;
  try {
    const count = await buildThread();
    const saved = downloadFile(markdown, fileName);
    if (saved) {
      setStatus(`Exported ${count} message${count === 1 ? "" : "s"} to ${fileName}`, "ok");
    } else {
      setStatus(
        `Built ${count} message${count === 1 ? "" : "s"}, but this Outlook build blocked the download. ` +
          `Use "Copy to clipboard" instead.`,
        "error"
      );
    }
  } catch (e) {
    setStatus(describeError(e), "error");
  } finally {
    els.export.disabled = false;
  }
}

async function onCopy() {
  try {
    if (!markdown) await buildThread();
    await copyText(markdown);
    setStatus("Copied to clipboard.", "ok");
  } catch (e) {
    setStatus(describeError(e), "error");
  }
}

// Classic Outlook hosts the pane in WebView2, where a blob download can be
// silently blocked. Report failure so the clipboard path can be offered.
function downloadFile(text, name) {
  try {
    const blob = new Blob([text], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    return true;
  } catch {
    return false;
  }
}

async function copyText(text) {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch {
      // fall through to the textarea method
    }
  }
  els.previewWrap.hidden = false;
  els.previewWrap.open = true;
  els.preview.focus();
  els.preview.select();
  if (!document.execCommand("copy")) {
    throw new Error("Clipboard blocked. Select the preview text and copy manually.");
  }
}

function describeError(e) {
  const msg = String(e?.message || e);
  if (/consent|AADSTS65001/i.test(msg)) {
    return "Permission not granted yet. Approve the Mail.Read consent prompt and retry.";
  }
  if (/popup|user_cancelled/i.test(msg)) {
    return "Sign-in was cancelled. Click Export to try again.";
  }
  if (/Graph 401/.test(msg)) return "Sign-in expired. Click Export to sign in again.";
  if (/Graph 403/.test(msg)) return "Access denied. The Mail.Read permission is missing or blocked by policy.";
  if (/Graph 429/.test(msg)) return "Microsoft throttled the request. Wait a moment and retry.";
  return msg;
}
