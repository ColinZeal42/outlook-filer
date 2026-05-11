"use strict";

const AUTH_CLIENT_ID = "75dc31c8-0515-4c64-849c-3958218e2c5f";
const AUTH_TOKEN_URL = "https://login.microsoftonline.com/hmflaw.com/oauth2/v2.0/token";
const USER_DOMAIN = "hmflaw.com";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

Office.actions.associate("onMessageSend", onMessageSend);

async function onMessageSend(event) {
  try {
    const item = Office.context.mailbox.item;
    const [subject, recipients] = await Promise.all([
      getAsync(item.subject),
      getAsync(item.to),
    ]);

    if (isCalendarMessage(subject)) {
      return event.completed({ allowEvent: true });
    }

    const emails = recipients.map(r => r.emailAddress);
    if (!hasExternalRecipient(emails, USER_DOMAIN)) {
      return event.completed({ allowEvent: true });
    }

    const token = getAccessToken();
    if (!token) {
      return event.completed({ allowEvent: true });
    }

    // Second pass: user clicked "Send Anyway" — file it
    const pending = getPending();
    if (pending) {
      const itemId = await getItemIdAsync(item).catch(() => null);
      if (itemId) moveAfterSend(token, itemId, pending.folderId);
      clearPending();
      return event.completed({ allowEvent: true });
    }

    // First pass: match folder (synchronous read from roamingSettings — no Graph call)
    let folders;
    try {
      folders = getCaseFolders();
    } catch (e) {
      return event.completed({ allowEvent: false, errorMessage: "DEBUG: " + e.message });
    }

    const participantText = recipients.map(r => `${r.displayName} ${r.emailAddress}`).join(" ");
    const match = matchFolder({ subject, participantText, bodyText: "" }, folders);

    if (!match) {
      return event.completed({ allowEvent: false, errorMessage: `DEBUG: No match. Subject="${subject}", folders: ${folders.map(f => f.displayName).join(", ")}` });
    }

    setPending({ folderId: match.id, folderName: match.displayName });
    event.completed({
      allowEvent: false,
      errorMessage: `File to "${match.displayName}"? Click Send Anyway to confirm.`,
    });
  } catch (err) {
    event.completed({ allowEvent: false, errorMessage: "DEBUG error: " + err.message });
  }
}

// --- Auth ---

function getAccessToken() {
  return Office.context.roamingSettings.get("access_token") || null;
}

// --- Folders (cached in roamingSettings by setup pane) ---

function getCaseFolders() {
  const stored = Office.context.roamingSettings.get("case_folders");
  if (!stored) throw new Error("No folders cached. Open HMF Setup and connect.");
  return JSON.parse(stored).map(f => ({
    displayName: f.displayName,
    id: f.id,
    keywords: f.displayName.split("/").map(k => k.trim().toLowerCase()),
  }));
}

// --- Pending confirmation (in-memory; requires lifetime="long" runtime) ---

let _pending = null;

function getPending() {
  if (!_pending) return null;
  if (Date.now() - _pending.ts > 90000) { _pending = null; return null; }
  return _pending;
}

function setPending(folder) {
  _pending = { folderId: folder.folderId, folderName: folder.folderName, ts: Date.now() };
}

function clearPending() {
  _pending = null;
}

// --- Graph (used only for fire-and-forget move after send) ---

async function graphGet(token, url) {
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!resp.ok) throw new Error(`Graph GET ${url} → ${resp.status}`);
  return resp.json();
}

async function graphPost(token, url, body) {
  const resp = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!resp.ok) throw new Error(`Graph POST ${url} → ${resp.status}`);
  return resp.json();
}

async function moveMessage(token, internetMessageId, folderId) {
  const findMsg = async () => {
    const enc = encodeURIComponent(`internetMessageId eq '${internetMessageId}'`);
    const res = await graphGet(token, `${GRAPH_BASE}/me/mailFolders/SentItems/messages?$filter=${enc}&$select=id&$top=1`);
    return res.value?.[0]?.id ?? null;
  };
  for (let i = 0; i < 5; i++) {
    if (i > 0) await delay(2000);
    const msgId = await findMsg();
    if (msgId) {
      await graphPost(token, `${GRAPH_BASE}/me/messages/${msgId}/move`, { destinationId: folderId });
      return;
    }
  }
}

function moveAfterSend(token, internetMessageId, folderId) {
  moveMessage(token, internetMessageId, folderId).catch(e => console.error("moveAfterSend:", e));
}

// --- Matching ---

const CALENDAR_PREFIXES = ["accepted:", "declined:", "tentative:", "cancelled:", "meeting request:"];

function isCalendarMessage(subject) {
  return CALENDAR_PREFIXES.some(p => subject.toLowerCase().startsWith(p));
}

function hasExternalRecipient(emails, domain) {
  return emails.some(a => a && !a.toLowerCase().endsWith(`@${domain}`));
}

function matchFolder(email, folders) {
  for (const text of [email.subject, email.participantText]) {
    const lower = text.toLowerCase();
    for (const f of folders) {
      if (f.keywords.some(kw => lower.includes(kw))) return f;
    }
  }
  return null;
}

// --- Office.js helpers ---

function getAsync(prop) {
  return new Promise((res, rej) =>
    prop.getAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error))
  );
}

function getItemIdAsync(item) {
  return new Promise((res, rej) => {
    if (!item.getItemIdAsync) return rej(new Error("getItemIdAsync not available"));
    item.getItemIdAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error));
  });
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
