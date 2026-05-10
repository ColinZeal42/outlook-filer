"use strict";

const DIALOG_URL = "https://ColinZeal42.github.io/outlook-filer/dialog.html";
const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
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

    const token = await getAccessToken();

    const folders = await getCaseFolders(token);
    const emailData = {
      subject,
      participantText: recipients.map(r => `${r.displayName} ${r.emailAddress}`).join(" "),
      bodyText: body,
    };

    const match = matchFolder(emailData, folders);
    if (!match) {
      return event.completed({ allowEvent: true });
    }

    const confirmed = await showConfirmDialog(match);
    if (confirmed) moveAfterSend(token, itemId, match.id);
    event.completed({ allowEvent: true });
  } catch (err) {
    console.error("onMessageSend error:", err);
    event.completed({ allowEvent: true });
  }
}

// --- Auth ---

async function getAccessToken() {
  try {
    return await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      forMSGraphAccess: true,
    });
  } catch (e) {
    throw new Error("SSO failed (code=" + (e.code ?? e.message) + "); add-in must be deployed via Centralized Deployment");
  }
}

function getTokenViaDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      AUTH_URL,
      { height: 60, width: 30, displayInIframe: true },
      result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(new Error("Cannot open auth dialog: code=" + result.error.code + " " + result.error.message));
        }
        const dlg = result.value;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, args => {
          dlg.close();
          try {
            const msg = JSON.parse(args.message);
            if (msg.token) resolve(msg.token);
            else reject(new Error(msg.error || "Auth failed"));
          } catch (e) { reject(e); }
        });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
          reject(new Error("Auth dialog closed"));
        });
      }
    );
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

// --- Confirm dialog ---

function showConfirmDialog(folder) {
  return new Promise(resolve => {
    const url = `${DIALOG_URL}#folderName=${encodeURIComponent(folder.displayName)}&folderId=${encodeURIComponent(folder.id)}`;
    Office.context.ui.displayDialogAsync(url, { height: 30, width: 35, displayInIframe: true }, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return resolve(false);
      const dlg = result.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, args => {
        dlg.close();
        try { resolve(JSON.parse(args.message).action === "move"); }
        catch (e) { resolve(false); }
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => resolve(false));
    });
  });
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
