"use strict";

const USER_DOMAIN = "hmflaw.com";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

Office.onReady(() => {
  Office.actions.associate("onMessageSend", onMessageSend);
});

async function onMessageSend(event) {
  try {
    const item = Office.context.mailbox.item;
    const [subject, recipients, body, itemId] = await Promise.all([
      getAsync(item.subject),
      getAsync(item.to),
      getBodyAsync(item),
      getItemIdAsync(item),
    ]);

    if (isCalendarMessage(subject)) {
      return event.completed({ allowEvent: true });
    }

    const emails = recipients.map(r => r.emailAddress);
    if (!hasExternalRecipient(emails, USER_DOMAIN)) {
      return event.completed({ allowEvent: true });
    }

    // Second pass: user clicked "Send Anyway" to confirm filing
    const pending = await getPending();
    if (pending) {
      const token = await getAccessToken();
      moveAfterSend(token, itemId, pending.folderId);
      await clearPending();
      return event.completed({ allowEvent: true });
    }

    // First pass: find matching folder
    const token = await getAccessToken();
    const folders = await getCaseFolders(token);
    const match = matchFolder(
      {
        subject,
        participantText: recipients.map(r => `${r.displayName} ${r.emailAddress}`).join(" "),
        bodyText: body,
      },
      folders
    );

    if (!match) {
      return event.completed({ allowEvent: false, errorMessage: `HMF: No match. Checked ${folders.length} folders. Subject: "${subject}"` });
    }

    await setPending({ folderId: match.id, folderName: match.displayName });
    event.completed({
      allowEvent: false,
      errorMessage: `File to "${match.displayName}"? Click Send Anyway to confirm.`,
    });
  } catch (err) {
    console.error("onMessageSend error:", err);
    event.completed({ allowEvent: false, errorMessage: "HMF Filer: " + (err.message || String(err)) });
  }
}

// --- Pending confirmation (OfficeRuntime.storage) ---

async function getPending() {
  const s = await OfficeRuntime.storage.getItems(["pendingFolderId", "pendingFolderName", "pendingTs"]);
  if (!s.pendingFolderId || !s.pendingTs) return null;
  if (Date.now() - parseInt(s.pendingTs) > 90000) return null; // expire after 90s
  return { folderId: s.pendingFolderId, folderName: s.pendingFolderName };
}

async function setPending(folder) {
  await OfficeRuntime.storage.setItems({
    pendingFolderId: folder.folderId,
    pendingFolderName: folder.folderName,
    pendingTs: String(Date.now()),
  });
}

async function clearPending() {
  await OfficeRuntime.storage.removeItems(["pendingFolderId", "pendingFolderName", "pendingTs"]);
}

// --- Auth ---

async function getAccessToken() {
  return Office.auth.getAccessToken({
    allowSignInPrompt: false,
    allowConsentPrompt: false,
    forMSGraphAccess: true,
  });
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

async function getCaseFolders(token) {
  const top = await graphGet(token, `${GRAPH_BASE}/me/mailFolders?$top=100`);
  const casesFolder = top.value.find(f => f.displayName === "__Cases");
  if (!casesFolder) throw new Error("__Cases folder not found");
  const children = await graphGet(token, `${GRAPH_BASE}/me/mailFolders/${casesFolder.id}/childFolders?$top=100`);
  return children.value.map(f => ({
    displayName: f.displayName,
    id: f.id,
    keywords: f.displayName.split("/").map(k => k.trim().toLowerCase()),
  }));
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
  throw new Error("Sent message not found after retries");
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
