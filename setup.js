"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

let _pendingMode = null;
let _emailQueue = []; // [{msg, match, opts, done, body}]
let _emailIndex = 0;

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("processBtn").addEventListener("click", processUnfiled);
  document.getElementById("fileSelectedBtn").addEventListener("click", fileSelected);
  document.getElementById("fileInboxBtn").addEventListener("click", fileInbox);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";
  wireCardButtons();

  const params = new URLSearchParams(window.location.search);
  const mode = params.get("mode");
  checkStatus();

  const token = Office.context.roamingSettings.get("access_token");
  if (mode) {
    if (!token) {
      _pendingMode = mode;
      document.getElementById("queue-status").textContent = "Connecting to Microsoft...";
      signIn();
    } else if (mode === "unsent") {
      processUnfiled();
    } else if (mode === "selected") {
      fileSelected();
    } else if (mode === "inbox") {
      fileInbox();
    }
  }
});

// --- Status ---

function checkStatus() {
  const refreshToken = Office.context.roamingSettings.get("refresh_token");
  const expiry = parseInt(Office.context.roamingSettings.get("token_expiry") || "0");
  const foldersJson = Office.context.roamingSettings.get("case_folders");
  const folderCount = foldersJson ? JSON.parse(foldersJson).length : 0;
  const statusEl = document.getElementById("status");
  const refreshBtn = document.getElementById("refreshBtn");
  const processBtn = document.getElementById("processBtn");
  const fileSelectedBtn = document.getElementById("fileSelectedBtn");
  const fileInboxBtn = document.getElementById("fileInboxBtn");

  if (refreshToken) {
    if (Date.now() < expiry) {
      statusEl.textContent = `Connected. ${folderCount} case folder${folderCount !== 1 ? "s" : ""} cached.`;
      statusEl.style.color = "green";
    } else {
      statusEl.textContent = "Token expired. Click Connect to refresh.";
      statusEl.style.color = "darkorange";
    }
    refreshBtn.style.display = "inline-block";
    processBtn.style.display = "inline-block";
    fileSelectedBtn.style.display = "inline-block";
    fileInboxBtn.style.display = "inline-block";
  } else {
    statusEl.textContent = "Not connected. Click Connect to sign in.";
    statusEl.style.color = "#555";
    refreshBtn.style.display = "none";
    processBtn.style.display = "none";
    fileSelectedBtn.style.display = "none";
    fileInboxBtn.style.display = "none";
  }
}

// --- Graph helpers ---

async function moveMessage(token, msgId, destinationId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${msgId}/move`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ destinationId })
  });
  if (!res.ok) throw new Error("Move failed: " + res.status);
}

async function fetchEmailBody(token, msgId) {
  try {
    const res = await fetch(`${GRAPH_BASE}/me/messages/${msgId}?$select=body`, {
      headers: {
        Authorization: "Bearer " + token,
        "Prefer": 'outlook.body-content-type="text"'
      }
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.body && data.body.content ? data.body.content : null;
  } catch(e) {
    return null;
  }
}

// --- Body / ID helpers ---

function extractPreviewLines(text, maxLines) {
  if (!text) return "";
  const lines = text.split(/\r?\n/);
  const result = [];
  let blankRun = 0;
  for (const line of lines) {
    const t = line.trim();
    if (!t) {
      if (blankRun === 0 && result.length > 0) result.push("");
      blankRun++;
    } else {
      blankRun = 0;
      result.push(t);
    }
    if (result.length >= maxLines) break;
  }
  return result.join("\n").trim();
}

function convertToRestIdAsync(ewsId) {
  return new Promise(resolve => {
    Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0, resolve);
  });
}

function getItemBodyAsync(item) {
  return new Promise(resolve => {
    item.body.getAsync(Office.CoercionType.Text, {}, result => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null);
    });
  });
}

// --- Email card ---

function wireCardButtons() {
  document.getElementById("prevEmailBtn").onclick = () => {
    if (_emailIndex > 0) showEmailAt(_emailIndex - 1);
  };
  document.getElementById("nextEmailBtn").onclick = () => {
    if (_emailIndex < _emailQueue.length - 1) showEmailAt(_emailIndex + 1);
  };
  document.getElementById("fileItBtn").onclick = () => cardFileIt();
  document.getElementById("deleteItBtn").onclick = () => cardDelete();
  document.getElementById("ignoreItBtn").onclick = () => cardIgnore();
}

function initEmailCard(entries) {
  _emailQueue = entries.map(e => ({
    msg: e.msg,
    match: e.match,
    opts: e.opts,
    done: false,
    body: e.body !== undefined ? e.body : null
  }));
  _emailIndex = 0;
  document.getElementById("email-card").style.display = "block";
  showEmailAt(0);
}

async function showEmailAt(index) {
  _emailIndex = index;
  const total = _emailQueue.length;
  const entry = _emailQueue[index];

  document.getElementById("email-counter").textContent = `${index + 1} / ${total}`;
  document.getElementById("prevEmailBtn").disabled = index === 0;
  document.getElementById("nextEmailBtn").disabled = index === total - 1;

  document.getElementById("email-subject-text").textContent = entry.msg.subject || "(no subject)";

  const metaParts = [entry.opts.senderLabel, entry.opts.dateStr].filter(Boolean);
  document.getElementById("email-meta").textContent = metaParts.join("  •  ");

  const matchEl = document.getElementById("email-match-line");
  if (entry.match) {
    matchEl.textContent = "→ " + entry.match.displayName;
    matchEl.className = "";
  } else if (entry.opts.isInternal) {
    matchEl.textContent = "Internal email — no filing";
    matchEl.className = "no-match";
  } else {
    matchEl.textContent = "No case folder match";
    matchEl.className = "no-match";
  }

  const fileBtn = document.getElementById("fileItBtn");
  fileBtn.style.display = entry.match ? "inline-block" : "none";
  fileBtn.textContent = "File It";
  document.getElementById("deleteItBtn").textContent = "Delete";
  document.getElementById("ignoreItBtn").textContent = "Ignore";

  if (entry.done) {
    disableCardBtns();
  } else {
    if (entry.match) fileBtn.disabled = false;
    document.getElementById("deleteItBtn").disabled = false;
    document.getElementById("ignoreItBtn").disabled = false;
  }

  const bodyEl = document.getElementById("email-body");
  if (entry.body !== null) {
    bodyEl.textContent = entry.body || "(no preview)";
  } else {
    bodyEl.textContent = "Loading preview…";
    const token = Office.context.roamingSettings.get("access_token");
    const raw = await fetchEmailBody(token, entry.msg.id);
    entry.body = extractPreviewLines(raw, 10) || "(no preview)";
    if (_emailIndex === index) bodyEl.textContent = entry.body;
  }
}

function disableCardBtns() {
  ["fileItBtn", "deleteItBtn", "ignoreItBtn"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.disabled = true;
  });
}

function enableCardBtns() {
  const entry = _emailQueue[_emailIndex];
  if (!entry || entry.done) return;
  if (entry.match) document.getElementById("fileItBtn").disabled = false;
  document.getElementById("deleteItBtn").disabled = false;
  document.getElementById("ignoreItBtn").disabled = false;
}

function autoAdvance() {
  setTimeout(() => {
    let next = -1;
    for (let i = _emailIndex + 1; i < _emailQueue.length; i++) {
      if (!_emailQueue[i].done) { next = i; break; }
    }
    if (next !== -1) {
      showEmailAt(next);
    } else if (_emailQueue.filter(e => !e.done).length === 0) {
      document.getElementById("queue-status").textContent = "All done ✓";
    }
  }, 700);
}

async function cardFileIt() {
  const entry = _emailQueue[_emailIndex];
  if (!entry || !entry.match || entry.done) return;
  const token = Office.context.roamingSettings.get("access_token");
  disableCardBtns();
  document.getElementById("fileItBtn").textContent = "Filing…";
  try {
    await moveMessage(token, entry.msg.id, entry.match.id);
    entry.done = true;
    document.getElementById("fileItBtn").textContent = "Filed ✓";
    autoAdvance();
  } catch(e) {
    document.getElementById("fileItBtn").textContent = "Error";
    enableCardBtns();
  }
}

async function cardDelete() {
  const entry = _emailQueue[_emailIndex];
  if (!entry || entry.done) return;
  const token = Office.context.roamingSettings.get("access_token");
  disableCardBtns();
  document.getElementById("deleteItBtn").textContent = "Deleting…";
  try {
    await moveMessage(token, entry.msg.id, "deleteditems");
    entry.done = true;
    document.getElementById("deleteItBtn").textContent = "Deleted ✓";
    autoAdvance();
  } catch(e) {
    document.getElementById("deleteItBtn").textContent = "Error";
    enableCardBtns();
  }
}

async function cardIgnore() {
  const entry = _emailQueue[_emailIndex];
  if (!entry || entry.done) return;
  const token = Office.context.roamingSettings.get("access_token");
  disableCardBtns();
  if (entry.opts.moveOnIgnore) {
    document.getElementById("ignoreItBtn").textContent = "Moving…";
    try {
      await moveMessage(token, entry.msg.id, "SentItems");
      entry.done = true;
      document.getElementById("ignoreItBtn").textContent = "Moved ✓";
      autoAdvance();
    } catch(e) {
      document.getElementById("ignoreItBtn").textContent = "Error";
      enableCardBtns();
    }
  } else {
    entry.done = true;
    autoAdvance();
  }
}

// --- Process Unfiled ---

async function ensureSentUnfiledFolder(token) {
  const res = await fetch(`${GRAPH_BASE}/me/mailFolders?$top=100`, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!res.ok) throw new Error("Graph " + res.status);
  const data = await res.json();
  const existing = (data.value || []).find(f => f.displayName === "Sent-Unfiled");
  if (existing) return existing.id;

  const createRes = await fetch(`${GRAPH_BASE}/me/mailFolders`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: "Sent-Unfiled" })
  });
  if (!createRes.ok) throw new Error("Could not create Sent-Unfiled folder");
  const created = await createRes.json();
  return created.id;
}

async function processUnfiled() {
  const btn = document.getElementById("processBtn");
  const statusEl = document.getElementById("queue-status");

  btn.disabled = true;
  document.getElementById("email-card").style.display = "none";
  statusEl.textContent = "Checking Sent-Unfiled…";

  const token = Office.context.roamingSettings.get("access_token");
  if (!token) { statusEl.textContent = "Not connected."; btn.disabled = false; return; }

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Click Refresh Folders first.";
    btn.disabled = false;
    return;
  }

  try {
    const sentUnfiledId = await ensureSentUnfiledFolder(token);

    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/${sentUnfiledId}/messages` +
      `?$select=id,subject,toRecipients,ccRecipients,sentDateTime,from&$top=50&$orderby=sentDateTime asc`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];

    if (messages.length === 0) {
      statusEl.textContent = "No emails to process.";
      btn.disabled = false;
      return;
    }

    messages.sort((a, b) => (a.subject || "").localeCompare(b.subject || ""));

    const folders = parseFolders(foldersJson);
    const entries = [];
    const calendarMoves = [];

    for (const msg of messages) {
      if (isCalendarMessage(msg.subject || "")) {
        calendarMoves.push(moveMessage(token, msg.id, "SentItems").catch(() => {}));
        continue;
      }

      const allRecipients = [...(msg.toRecipients || []), ...(msg.ccRecipients || [])];
      const emails = allRecipients.map(r => r.emailAddress.address);
      const participantText = allRecipients.map(r =>
        r.emailAddress.name + " " + r.emailAddress.address
      ).join(" ");

      const isInternal = !hasExternalRecipient(emails, USER_DOMAIN);
      const match = isInternal ? null : matchFolder({ subject: msg.subject || "", participantText }, folders);

      const toNames = (msg.toRecipients || []).map(r => r.emailAddress.name || r.emailAddress.address).join(", ");
      entries.push({
        msg,
        match,
        opts: {
          isInternal,
          senderLabel: toNames ? "To: " + toNames : "",
          dateStr: formatDate(msg.sentDateTime),
          moveOnIgnore: true
        }
      });
    }

    await Promise.allSettled(calendarMoves);

    if (entries.length === 0) {
      statusEl.textContent = "No emails to process.";
      btn.disabled = false;
      return;
    }

    statusEl.textContent = `${entries.length} email${entries.length !== 1 ? "s" : ""} to review:`;
    initEmailCard(entries);
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
  }

  btn.disabled = false;
}

// --- File Selected ---

async function fileSelected() {
  const btn = document.getElementById("fileSelectedBtn");
  const statusEl = document.getElementById("queue-status");

  btn.disabled = true;
  document.getElementById("email-card").style.display = "none";

  const token = Office.context.roamingSettings.get("access_token");
  if (!token) { statusEl.textContent = "Not connected."; btn.disabled = false; return; }

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Click Refresh Folders first.";
    btn.disabled = false;
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) {
    statusEl.textContent = "No email is currently open.";
    btn.disabled = false;
    return;
  }

  statusEl.textContent = "Loading…";

  const subject = item.subject || "";
  const allRecipients = [...(item.to || []), ...(item.cc || [])];
  const emails = allRecipients.map(r => r.emailAddress);
  const participantText = allRecipients.map(r => r.displayName + " " + r.emailAddress).join(" ");

  const folders = parseFolders(foldersJson);
  const match = matchFolder({ subject, participantText }, folders);

  const toNames = (item.to || []).map(r => r.displayName || r.emailAddress).join(", ");
  const dateStr = item.dateTimeCreated ? formatDate(item.dateTimeCreated.toISOString()) : "";

  try {
    const [restId, bodyText] = await Promise.all([
      convertToRestIdAsync(item.itemId),
      getItemBodyAsync(item)
    ]);

    const preview = extractPreviewLines(bodyText, 10) || "(no preview)";

    initEmailCard([{
      msg: { id: restId, subject },
      match,
      opts: {
        isInternal: false,
        senderLabel: toNames ? "To: " + toNames : "",
        dateStr,
        moveOnIgnore: true
      },
      body: preview
    }]);

    statusEl.textContent = "1 email to review:";
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
  }

  btn.disabled = false;
}

// --- File Inbox ---

async function fileInbox() {
  const btn = document.getElementById("fileInboxBtn");
  const statusEl = document.getElementById("queue-status");

  btn.disabled = true;
  document.getElementById("email-card").style.display = "none";
  statusEl.textContent = "Scanning Inbox…";

  const token = Office.context.roamingSettings.get("access_token");
  if (!token) { statusEl.textContent = "Not connected."; btn.disabled = false; return; }

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Click Refresh Folders first.";
    btn.disabled = false;
    return;
  }

  try {
    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/Inbox/messages` +
      `?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime&$top=50&$orderby=receivedDateTime desc`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];

    if (messages.length === 0) {
      statusEl.textContent = "Inbox is empty.";
      btn.disabled = false;
      return;
    }

    messages.sort((a, b) => (a.subject || "").localeCompare(b.subject || ""));

    const folders = parseFolders(foldersJson);
    const entries = [];

    for (const msg of messages) {
      if (isCalendarMessage(msg.subject || "")) continue;

      const fromAddr = msg.from && msg.from.emailAddress ? msg.from.emailAddress.address || "" : "";
      const fromName = msg.from && msg.from.emailAddress ? msg.from.emailAddress.name || "" : "";
      const allRecipients = [...(msg.toRecipients || []), ...(msg.ccRecipients || [])];
      const recipientEmails = allRecipients.map(r => r.emailAddress ? r.emailAddress.address || "" : "");
      const participantText = [fromName, fromAddr,
        ...allRecipients.map(r => r.emailAddress ? (r.emailAddress.name || "") + " " + (r.emailAddress.address || "") : "")
      ].join(" ");

      const isInternal = fromAddr.toLowerCase().endsWith("@" + USER_DOMAIN) && !hasExternalRecipient(recipientEmails, USER_DOMAIN);
      const match = isInternal ? null : matchFolder({ subject: msg.subject || "", participantText }, folders);

      entries.push({
        msg,
        match,
        opts: {
          isInternal,
          senderLabel: fromName ? "From: " + fromName : (fromAddr ? "From: " + fromAddr : ""),
          dateStr: formatDate(msg.receivedDateTime),
          moveOnIgnore: false
        }
      });
    }

    if (entries.length === 0) {
      statusEl.textContent = "No emails in Inbox.";
      btn.disabled = false;
      return;
    }

    statusEl.textContent = `${entries.length} email${entries.length !== 1 ? "s" : ""} to review:`;
    initEmailCard(entries);
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
  }

  btn.disabled = false;
}

// --- Folder refresh ---

async function fetchCaseFolders(token) {
  const res = await fetch(`${GRAPH_BASE}/me/mailFolders?$top=100&$expand=childFolders($top=100)`, {
    headers: { Authorization: "Bearer " + token },
  });
  if (!res.ok) throw new Error("Graph " + res.status);
  const data = await res.json();
  const casesFolder = data.value.find(f => f.displayName === "__Cases");
  if (!casesFolder) throw new Error("__Cases folder not found");
  return (casesFolder.childFolders || []).map(f => ({ displayName: f.displayName, id: f.id }));
}

async function refreshFolders() {
  const btn = document.getElementById("refreshBtn");
  const statusEl = document.getElementById("status");
  const token = Office.context.roamingSettings.get("access_token");
  if (!token) { statusEl.textContent = "Not connected."; statusEl.style.color = "red"; return; }
  btn.disabled = true;
  statusEl.textContent = "Refreshing folder list...";
  try {
    const folders = await fetchCaseFolders(token);
    Office.context.roamingSettings.set("case_folders", JSON.stringify(folders));
    Office.context.roamingSettings.saveAsync(() => { btn.disabled = false; checkStatus(); });
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
    statusEl.style.color = "red";
    btn.disabled = false;
  }
}

// --- Sign in ---

function signIn() {
  const btn = document.getElementById("connectBtn");
  const statusEl = document.getElementById("status");
  btn.disabled = true;
  statusEl.textContent = "Opening sign-in window...";

  Office.context.ui.displayDialogAsync(AUTH_URL, { height: 60, width: 40, displayInIframe: false },
    result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        statusEl.textContent = "Could not open sign-in window: " + result.error.message;
        statusEl.style.color = "red";
        btn.disabled = false;
        return;
      }
      const dlg = result.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, async args => {
        dlg.close();
        try {
          const msg = JSON.parse(args.message);
          if (!msg.token) {
            statusEl.textContent = "Sign-in failed: " + (msg.error || "Unknown error");
            statusEl.style.color = "red";
            btn.disabled = false;
            return;
          }
          statusEl.textContent = "Fetching case folders...";
          let folders = [];
          try { folders = await fetchCaseFolders(msg.token); } catch(e) {}
          Office.context.roamingSettings.set("access_token", msg.token);
          Office.context.roamingSettings.set("token_expiry", String(msg.expiry || (Date.now() + 3600000)));
          if (msg.refreshToken) Office.context.roamingSettings.set("refresh_token", msg.refreshToken);
          Office.context.roamingSettings.set("case_folders", JSON.stringify(folders));
          Office.context.roamingSettings.saveAsync(r => {
            btn.disabled = false;
            if (r.status === Office.AsyncResultStatus.Succeeded) {
              checkStatus();
              if (_pendingMode) {
                const pendingMode = _pendingMode;
                _pendingMode = null;
                if (pendingMode === "unsent") processUnfiled();
                else if (pendingMode === "selected") fileSelected();
                else if (pendingMode === "inbox") fileInbox();
              }
            } else {
              statusEl.textContent = "Error saving: " + r.error.message;
              statusEl.style.color = "red";
            }
          });
        } catch(e) {
          statusEl.textContent = "Error: " + e.message;
          statusEl.style.color = "red";
          btn.disabled = false;
        }
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
        statusEl.textContent = "Sign-in cancelled.";
        btn.disabled = false;
      });
    }
  );
}

// --- Email matching helpers ---

function parseFolders(foldersJson) {
  return JSON.parse(foldersJson).map(f => ({
    displayName: f.displayName,
    id: f.id,
    keywords: f.displayName.split("/").map(k => k.trim().toLowerCase())
  }));
}

const CALENDAR_PREFIXES = ["accepted:", "declined:", "tentative:", "cancelled:", "meeting request:"];

function isCalendarMessage(subject) {
  return CALENDAR_PREFIXES.some(p => subject.toLowerCase().indexOf(p) === 0);
}

function hasExternalRecipient(emails, domain) {
  return emails.some(a => a && a.toLowerCase().slice(-(domain.length + 1)) !== "@" + domain);
}

function matchFolder(email, folders) {
  const texts = [email.subject, email.participantText];
  for (let t = 0; t < texts.length; t++) {
    const lower = texts[t].toLowerCase();
    for (let f = 0; f < folders.length; f++) {
      const kws = folders[f].keywords;
      for (let k = 0; k < kws.length; k++) {
        if (lower.indexOf(kws[k]) !== -1) return folders[f];
      }
    }
  }
  return null;
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}
