"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

let _pendingMode = null;
let _emailQueue = []; // [{msg, match, opts, done, body, isReplied, isForwarded}]
let _emailIndex = 0;

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("processBtn").addEventListener("click", processUnfiled);
  document.getElementById("setBaselineBtn").addEventListener("click", setBaseline);
  document.getElementById("fileInboxBtn").addEventListener("click", fileInbox);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";
  wireCardButtons();

  // SETUP_MODE is set as a global by the mode-specific HTML files (setup-unsent.html etc.)
  // Fall back to URL param for direct loads of setup.html
  const mode = (typeof window.SETUP_MODE !== "undefined" ? window.SETUP_MODE : null)
    || new URLSearchParams(window.location.search).get("mode");

  checkStatus();

  const token = Office.context.roamingSettings.get("access_token");
  if (mode) {
    if (!token) {
      _pendingMode = mode;
      document.getElementById("queue-status").textContent = "Connecting to Microsoft...";
      signIn();
    } else if (mode === "unsent") {
      processUnfiled();
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
  const connectBtn = document.getElementById("connectBtn");
  const refreshBtn = document.getElementById("refreshBtn");
  const processBtn = document.getElementById("processBtn");
  const fileInboxBtn = document.getElementById("fileInboxBtn");

  if (refreshToken && Date.now() < expiry) {
    statusEl.textContent = `Connected. ${folderCount} case folder${folderCount !== 1 ? "s" : ""} cached.`;
    statusEl.style.color = "green";
    connectBtn.style.display = "none";
    refreshBtn.style.display = "inline-block";
    processBtn.style.display = "inline-block";
    fileInboxBtn.style.display = "inline-block";
  } else {
    statusEl.textContent = refreshToken
      ? "Token expired. Click Connect to refresh."
      : "Not connected. Click Connect to sign in.";
    statusEl.style.color = refreshToken ? "darkorange" : "#555";
    connectBtn.style.display = "inline-block";
    refreshBtn.style.display = "none";
    processBtn.style.display = "none";
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

// Fetches body (plain text) + replied/forwarded status in one Graph call.
async function fetchEmailDetails(token, msgId) {
  try {
    const res = await fetch(
      `${GRAPH_BASE}/me/messages/${msgId}?$select=body` +
      `&$expand=singleValueExtendedProperties($filter=id eq 'Integer 0x1081')`,
      {
        headers: {
          Authorization: "Bearer " + token,
          "Prefer": 'outlook.body-content-type="text"'
        }
      }
    );
    if (!res.ok) return { body: null, isReplied: false, isForwarded: false };
    const data = await res.json();
    const verb = parseInt(((data.singleValueExtendedProperties || [])[0] || {}).value || "0");
    return {
      body: data.body && data.body.content ? data.body.content : null,
      isReplied: verb === 102 || verb === 103,
      isForwarded: verb === 104
    };
  } catch(e) {
    return { body: null, isReplied: false, isForwarded: false };
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
    body: e.body !== undefined ? e.body : null,
    isReplied: e.isReplied,
    isForwarded: e.isForwarded
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

  // Reply/forward badge — shown if fetched, hidden while pending
  const badgeEl = document.getElementById("email-reply-badge");
  if (entry.isReplied !== undefined) {
    const show = entry.isReplied || entry.isForwarded;
    badgeEl.style.display = show ? "block" : "none";
    badgeEl.textContent = entry.isForwarded ? "↪ Forwarded" : "↩ Replied";
  } else {
    badgeEl.style.display = "none";
  }

  const line1 = [entry.opts.senderLabel, entry.opts.dateStr].filter(Boolean).join("  •  ");
  const line2 = entry.opts.toLabel || "";
  const line3 = entry.opts.ccLabel || "";
  document.getElementById("email-meta").textContent = [line1, line2, line3].filter(Boolean).join("\n");

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

  // Body + reply status — lazy-load from Graph if not yet fetched
  const bodyEl = document.getElementById("email-body");
  if (entry.body !== null) {
    bodyEl.textContent = entry.body || "(no preview)";
  } else {
    bodyEl.textContent = "Loading preview…";
    const token = Office.context.roamingSettings.get("access_token");
    const details = await fetchEmailDetails(token, entry.msg.id);
    entry.body = extractPreviewLines(details.body, 10) || "(no preview)";
    entry.isReplied = details.isReplied;
    entry.isForwarded = details.isForwarded;
    if (_emailIndex === index) {
      bodyEl.textContent = entry.body;
      const show = details.isReplied || details.isForwarded;
      badgeEl.style.display = show ? "block" : "none";
      badgeEl.textContent = details.isForwarded ? "↪ Forwarded" : "↩ Replied";
    }
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

async function processUnfiled() {
  const btn = document.getElementById("processBtn");
  const baselineBtn = document.getElementById("setBaselineBtn");
  const statusEl = document.getElementById("queue-status");

  btn.disabled = true;
  baselineBtn.style.display = "none";
  document.getElementById("email-card").style.display = "none";

  const token = Office.context.roamingSettings.get("access_token");
  if (!token) { statusEl.textContent = "Not connected."; btn.disabled = false; return; }

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Click Refresh Folders first.";
    btn.disabled = false;
    return;
  }

  const lastRun = Office.context.roamingSettings.get("sent_last_run");
  if (!lastRun) {
    statusEl.textContent = "No baseline set. Set a baseline to begin tracking sent emails.";
    baselineBtn.style.display = "inline-block";
    btn.disabled = false;
    return;
  }

  statusEl.textContent = "Checking Sent Items…";
  const newTimestamp = new Date().toISOString();

  try {
    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
      `?$filter=sentDateTime gt ${lastRun}` +
      `&$top=100&$orderby=sentDateTime asc` +
      `&$select=id,subject,toRecipients,ccRecipients,sentDateTime,from`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];

    Office.context.roamingSettings.set("sent_last_run", newTimestamp);
    Office.context.roamingSettings.saveAsync(() => {});

    const nonCalendar = messages.filter(m => !isCalendarMessage(m.subject || ""));
    nonCalendar.sort((a, b) => (a.subject || "").localeCompare(b.subject || ""));

    if (nonCalendar.length === 0) {
      statusEl.textContent = "No new sent emails to process.";
      btn.disabled = false;
      return;
    }

    const folders = parseFolders(foldersJson);
    const entries = [];

    for (const msg of nonCalendar) {
      const allRecipients = [...(msg.toRecipients || []), ...(msg.ccRecipients || [])];
      const emails = allRecipients.map(r => r.emailAddress.address);
      const participantText = allRecipients.map(r =>
        r.emailAddress.name + " " + r.emailAddress.address
      ).join(" ");

      const isInternal = !hasExternalRecipient(emails, USER_DOMAIN);
      const match = isInternal ? null : matchFolder({ subject: msg.subject || "", participantText }, folders);

      const toNames = (msg.toRecipients || []).map(r => r.emailAddress.name || r.emailAddress.address).join(", ");
      const ccNames = (msg.ccRecipients || []).map(r => r.emailAddress.name || r.emailAddress.address).join(", ");
      entries.push({
        msg,
        match,
        opts: {
          isInternal,
          senderLabel: toNames ? "To: " + toNames : "",
          ccLabel: ccNames ? "CC: " + ccNames : "",
          dateStr: formatDate(msg.sentDateTime),
          moveOnIgnore: false
        }
      });
    }

    statusEl.textContent = `${entries.length} email${entries.length !== 1 ? "s" : ""} to review:`;
    initEmailCard(entries);
  } catch(e) {
    statusEl.textContent = "Error: " + e.message;
  }

  btn.disabled = false;
}

function setBaseline() {
  const baselineBtn = document.getElementById("setBaselineBtn");
  const statusEl = document.getElementById("queue-status");
  baselineBtn.disabled = true;
  Office.context.roamingSettings.set("sent_last_run", new Date().toISOString());
  Office.context.roamingSettings.saveAsync(() => {
    baselineBtn.style.display = "none";
    baselineBtn.disabled = false;
    statusEl.textContent = "Baseline set. Next run will file emails sent from now on.";
  });
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

      const toNames = (msg.toRecipients || []).map(r => r.emailAddress ? (r.emailAddress.name || r.emailAddress.address || "") : "").filter(Boolean).join(", ");
      const ccNames = (msg.ccRecipients || []).map(r => r.emailAddress ? (r.emailAddress.name || r.emailAddress.address || "") : "").filter(Boolean).join(", ");
      entries.push({
        msg,
        match,
        opts: {
          isInternal,
          senderLabel: fromName ? "From: " + fromName : (fromAddr ? "From: " + fromAddr : ""),
          toLabel: toNames ? "To: " + toNames : "",
          ccLabel: ccNames ? "CC: " + ccNames : "",
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
