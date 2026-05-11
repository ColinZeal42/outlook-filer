"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("processBtn").addEventListener("click", processUnfiled);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";
  checkStatus();
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

  if (refreshToken) {
    if (Date.now() < expiry) {
      statusEl.textContent = `Connected. ${folderCount} case folder${folderCount !== 1 ? "s" : ""} cached.`;
      statusEl.style.color = "green";
    } else {
      statusEl.textContent = "Token expired. Click Connect to refresh.";
      statusEl.style.color = "darkorange";
    }
    refreshBtn.style.display = "inline-block";
    processBtn.style.display = "block";
  } else {
    statusEl.textContent = "Not connected. Click Connect to sign in.";
    statusEl.style.color = "#555";
    refreshBtn.style.display = "none";
    processBtn.style.display = "none";
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

async function moveMessage(token, msgId, destinationId) {
  const res = await fetch(`${GRAPH_BASE}/me/messages/${msgId}/move`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ destinationId })
  });
  if (!res.ok) throw new Error("Move failed: " + res.status);
}

async function fileEmail(token, msgId, folderId, rowEl) {
  const fileBtn = rowEl.querySelector(".file-btn");
  const skipBtn = rowEl.querySelector(".skip-btn");
  fileBtn.disabled = true;
  skipBtn.disabled = true;
  fileBtn.textContent = "Filing...";
  try {
    await moveMessage(token, msgId, folderId);
    fileBtn.textContent = "Filed ✓";
    rowEl.classList.add("done");
  } catch(e) {
    fileBtn.textContent = "Error";
    fileBtn.disabled = false;
    skipBtn.disabled = false;
  }
}

async function skipEmail(token, msgId, rowEl) {
  const fileBtn = rowEl.querySelector(".file-btn");
  const skipBtn = rowEl.querySelector(".skip-btn");
  fileBtn.disabled = true;
  skipBtn.disabled = true;
  skipBtn.textContent = "Moving...";
  try {
    await moveMessage(token, msgId, "SentItems");
    skipBtn.textContent = "Moved to Sent";
    rowEl.classList.add("done");
  } catch(e) {
    skipBtn.textContent = "Error";
    fileBtn.disabled = false;
    skipBtn.disabled = false;
  }
}

async function processUnfiled() {
  const btn = document.getElementById("processBtn");
  const statusEl = document.getElementById("queue-status");
  const queueEl = document.getElementById("queue");

  btn.disabled = true;
  queueEl.innerHTML = "";
  statusEl.textContent = "Checking Sent-Unfiled...";

  const token = Office.context.roamingSettings.get("access_token");
  if (!token) {
    statusEl.textContent = "Not connected.";
    btn.disabled = false;
    return;
  }

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
      `?$select=id,subject,toRecipients,sentDateTime&$top=50&$orderby=sentDateTime asc`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];

    if (messages.length === 0) {
      statusEl.textContent = "No emails to process.";
      btn.disabled = false;
      return;
    }

    const folders = JSON.parse(foldersJson).map(f => ({
      displayName: f.displayName,
      id: f.id,
      keywords: f.displayName.split("/").map(k => k.trim().toLowerCase())
    }));

    let matchCount = 0;
    let autoMovedCount = 0;
    const autoMoves = [];

    for (const msg of messages) {
      const emails = (msg.toRecipients || []).map(r => r.emailAddress.address);
      const participantText = (msg.toRecipients || []).map(r =>
        r.emailAddress.name + " " + r.emailAddress.address
      ).join(" ");

      if (isCalendarMessage(msg.subject || "") || !hasExternalRecipient(emails, USER_DOMAIN)) {
        autoMoves.push(moveMessage(token, msg.id, "SentItems").catch(() => {}));
        autoMovedCount++;
        continue;
      }

      const match = matchFolder({ subject: msg.subject || "", participantText }, folders);
      if (!match) {
        autoMoves.push(moveMessage(token, msg.id, "SentItems").catch(() => {}));
        autoMovedCount++;
        continue;
      }

      matchCount++;
      const subject = (msg.subject || "(no subject)").slice(0, 50);
      const row = document.createElement("div");
      row.className = "queue-row";
      row.innerHTML =
        `<span class="queue-subject" title="${escapeHtml(msg.subject || "")}">${escapeHtml(subject)}</span>` +
        `<span class="queue-arrow">→</span>` +
        `<span class="queue-folder" title="${escapeHtml(match.displayName)}">${escapeHtml(match.displayName)}</span>` +
        `<button class="file-btn">File It</button>` +
        `<button class="skip-btn">Skip</button>`;
      queueEl.appendChild(row);

      const capturedMsgId = msg.id;
      const capturedFolderId = match.id;
      row.querySelector(".file-btn").addEventListener("click", () => fileEmail(token, capturedMsgId, capturedFolderId, row));
      row.querySelector(".skip-btn").addEventListener("click", () => skipEmail(token, capturedMsgId, row));
    }

    await Promise.allSettled(autoMoves);

    if (matchCount === 0) {
      statusEl.textContent = `All ${autoMovedCount} email${autoMovedCount !== 1 ? "s" : ""} moved to Sent Items (no case matches found).`;
    } else {
      const autoNote = autoMovedCount > 0 ? `; ${autoMovedCount} moved to Sent automatically` : "";
      statusEl.textContent = `${matchCount} email${matchCount !== 1 ? "s" : ""} to review${autoNote}:`;
    }
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
            if (r.status === Office.AsyncResultStatus.Succeeded) checkStatus();
            else { statusEl.textContent = "Error saving: " + r.error.message; statusEl.style.color = "red"; }
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

function escapeHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
