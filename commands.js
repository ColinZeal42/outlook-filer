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

    const token = await getAccessToken();
    if (!token) {
      return event.completed({ allowEvent: false, errorMessage: "DEBUG: No token found in roamingSettings." });
    }

    // Second pass: user clicked "Send Anyway" — get item ID now and file it
    const pending = getPending();
    if (pending) {
      const itemId = await getItemIdAsync(item).catch(() => null);
      if (itemId) moveAfterSend(token, itemId, pending.folderId);
      clearPending();
      return event.completed({ allowEvent: true });
    }

    // First pass: find matching folder (uses cached folders after first call)
    let folders;
    try {
      folders = await getCaseFolders(token);
    } catch (e) {
      return event.completed({ allowEvent: false, errorMessage: "DEBUG: getCaseFolders failed: " + e.message });
    }

    const participantText = recipients.map(r => `${r.displayName} ${r.emailAddress}`).join(" ");
    const match = matchFolder({ subject, participantText, bodyText: "" }, folders);

    if (!match) {
      return event.completed({ allowEvent: false, errorMessage: `DEBUG: No match. Subject="${subject}", ${folders.length} folders: ${folders.map(f=>f.displayName).join(",")}` });
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

// --- Auth (OfficeRuntime.storage) ---

async function getAccessToken() {
  const refreshToken = Office.context.roamingSettings.get("refresh_token");
  if (!refreshToken) return null;

  const accessToken = Office.context.roamingSettings.get("access_token");
  const expiry = parseInt(Office.context.roamingSettings.get("token_expiry") || "0");

  if (accessToken && Date.now() < expiry - 300000) return accessToken;

  const resp = await fetch(AUTH_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      client_id: AUTH_CLIENT_ID,
      refresh_token: refreshToken,
      scope: "https://graph.microsoft.com/Mail.ReadWrite offline_access",
    }),
  });

  if (!resp.ok) return null;
  const data = await resp.json();
  return data.access_token || null;
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

// --- Graph ---

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

let _folderCache = null;
let _folderCacheTs = 0;

async function getCaseFolders(token) {
  if (_folderCache && Date.now() - _folderCacheTs < 5 * 60 * 1000) return _folderCache;
  const res = await graphGet(token, `${GRAPH_BASE}/me/mailFolders?$top=100&$expand=childFolders($top=100)`);
  const casesFolder = res.value.find(f => f.displayName === "__Cases");
  if (!casesFolder) throw new Error("__Cases folder not found");
  _folderCache = (casesFolder.childFolders || []).map(f => ({
    displayName: f.displayName,
    id: f.id,
    keywords: f.displayName.split("/").map(k => k.trim().toLowerCase()),
  }));
  _folderCacheTs = Date.now();
  return _folderCache;
}

// Pre-warm folder cache on runtime startup so first send doesn't pay the Graph cost
(async () => {
  try {
    const token = await getAccessToken();
    if (token) await getCaseFolders(token);
  } catch (_) {}
})();

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
  for (const text of [email.subject, email.participantText, email.bodyText]) {
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

function getBodyAsync(item) {
  return new Promise((res, rej) =>
    item.body.getAsync(Office.CoercionType.Text, r =>
      r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error)
    )
  );
}

function getItemIdAsync(item) {
  return new Promise((res, rej) => {
    if (!item.getItemIdAsync) return rej(new Error("getItemIdAsync not available"));
    item.getItemIdAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error));
  });
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
