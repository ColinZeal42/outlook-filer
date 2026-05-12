"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

let _pendingMode = null;
let _threadGroups = [];
let _threadFolders = [];

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("processBtn").addEventListener("click", processUnfiled);
  document.getElementById("setBaselineBtn").addEventListener("click", setBaseline);
  document.getElementById("fileInboxBtn").addEventListener("click", fileInbox);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";

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

// --- Token refresh ---

async function refreshAccessToken() {
  const storedRefresh = Office.context.roamingSettings.get("refresh_token");
  if (!storedRefresh) return false;
  try {
    const res = await fetch("https://login.microsoftonline.com/hmflaw.com/oauth2/v2.0/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: "75dc31c8-0515-4c64-849c-3958218e2c5f",
        grant_type: "refresh_token",
        refresh_token: storedRefresh,
        scope: "https://graph.microsoft.com/Mail.ReadWrite offline_access"
      }).toString()
    });
    if (!res.ok) return false;
    const data = await res.json();
    if (!data.access_token) return false;
    Office.context.roamingSettings.set("access_token", data.access_token);
    Office.context.roamingSettings.set("token_expiry", String(Date.now() + data.expires_in * 1000));
    if (data.refresh_token) Office.context.roamingSettings.set("refresh_token", data.refresh_token);
    await new Promise(resolve => Office.context.roamingSettings.saveAsync(resolve));
    return true;
  } catch(e) {
    return false;
  }
}

async function ensureFreshToken() {
  const expiry = parseInt(Office.context.roamingSettings.get("token_expiry") || "0");
  if (Date.now() >= expiry) {
    const ok = await refreshAccessToken();
    if (!ok) throw new Error("Session expired. Please reconnect.");
  }
  return Office.context.roamingSettings.get("access_token");
}

// --- Status ---

async function checkStatus() {
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
  } else if (refreshToken) {
    statusEl.textContent = "Refreshing session…";
    statusEl.style.color = "#555";
    connectBtn.style.display = "none";
    refreshBtn.style.display = "none";
    processBtn.style.display = "none";
    fileInboxBtn.style.display = "none";
    const ok = await refreshAccessToken();
    if (ok) {
      checkStatus();
    } else {
      statusEl.textContent = "Session expired. Click Connect to sign in.";
      statusEl.style.color = "darkorange";
      connectBtn.style.display = "inline-block";
    }
  } else {
    statusEl.textContent = "Not connected. Click Connect to sign in.";
    statusEl.style.color = "#555";
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

// --- Helpers ---

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

function esc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// --- Thread grouping ---

function groupByThread(messages, folders) {
  const map = {};
  const order = [];
  for (const msg of messages) {
    const cid = msg.conversationId || msg.id;
    if (!map[cid]) {
      map[cid] = { conversationId: cid, subject: msg.subject || "", emails: [], latestDate: 0 };
      order.push(cid);
    }
    const d = new Date(msg.sentDateTime || msg.receivedDateTime || 0).getTime();
    if (d > map[cid].latestDate) {
      map[cid].latestDate = d;
      map[cid].subject = msg.subject || map[cid].subject;
    }
    map[cid].emails.push({ msg, checked: true, body: null, isReplied: undefined, isForwarded: undefined });
  }

  return order.map(cid => {
    const group = map[cid];
    const counts = {};
    for (const e of group.emails) {
      const allRecip = [...(e.msg.toRecipients || []), ...(e.msg.ccRecipients || [])];
      const ea = (e.msg.from && e.msg.from.emailAddress) || {};
      const pt = [ea.name || "", ea.address || "",
        ...allRecip.map(r => { const a = r.emailAddress || {}; return (a.name || "") + " " + (a.address || ""); })
      ].join(" ");
      const m = matchFolder({ subject: e.msg.subject || "", participantText: pt }, folders);
      if (m) {
        if (!counts[m.id]) counts[m.id] = { folder: m, n: 0 };
        counts[m.id].n++;
      }
    }
    const hits = Object.values(counts);
    const best = hits.length ? hits.reduce((a, b) => b.n > a.n ? b : a).folder : null;

    return {
      conversationId: cid,
      subject: group.subject,
      emails: group.emails,
      match: best,
      manualMatch: null,
      expanded: false,
      done: false,
      latestDate: group.latestDate
    };
  }).sort((a, b) => b.latestDate - a.latestDate);
}

// --- Thread list UI ---

function initThreadList(groups, folders) {
  _threadGroups = groups;
  _threadFolders = folders;
  const el = document.getElementById("thread-list");
  el.style.display = "block";
  renderThreadList();
}

function renderThreadList() {
  const el = document.getElementById("thread-list");
  if (!el) return;
  let html = "";

  _threadGroups.forEach((group, idx) => {
    const doneClass = group.done ? " tl-done" : "";
    const subject = esc(group.subject || "(no subject)");
    const matchHtml = group.match
      ? '<span class="tl-match">→ ' + esc(group.match.displayName) + '</span>'
      : '<span class="tl-match tl-no-match">(no match)</span>';
    const chevron = group.expanded ? "▲" : "▼";
    const headerAttrs = group.done ? "" : ' onclick="toggleThread(' + idx + ')" style="cursor:pointer"';

    html += '<div class="tl-group' + doneClass + '" id="tg-' + idx + '">';
    html += '<div class="tl-header"' + headerAttrs + '>';
    html += '<span class="tl-pill">' + group.emails.length + '</span>';
    html += '<span class="tl-subject">' + subject + '</span>';
    html += matchHtml;
    html += '<span class="tl-chevron">' + chevron + '</span>';
    html += '</div>';

    if (group.expanded && !group.done) {
      html += '<div class="tl-body">';
      group.emails.forEach(e => {
        const ea = (e.msg.from && e.msg.from.emailAddress) || {};
        const sender = esc(ea.name || ea.address || "Unknown");
        const dateStr = esc(formatDate(e.msg.sentDateTime || e.msg.receivedDateTime));
        const badge = e.isForwarded ? '<span class="tl-badge tl-fwd">↪</span> '
                    : e.isReplied   ? '<span class="tl-badge">↩</span> '
                    : '';
        const bodyHtml = e.body === null
          ? '<em>Loading…</em>'
          : esc(e.body || "(no preview)").replace(/\n/g, "<br>");
        const checked = e.checked ? " checked" : "";

        html += '<div class="tl-email">';
        html += '<label class="tl-email-label">';
        html += '<input type="checkbox" id="chk-' + esc(e.msg.id) + '"' + checked + ' onchange="onCheckChange(' + idx + ')">';
        html += '<div class="tl-email-content">';
        html += '<div class="tl-email-meta">' + badge + sender + ' <span class="tl-email-date">· ' + dateStr + '</span></div>';
        html += '<div class="tl-email-body">' + bodyHtml + '</div>';
        html += '</div></label></div>';
      });

      if (!group.match) {
        html += '<select class="tl-folder-select" onchange="onFolderPick(' + idx + ', this.value)">';
        html += '<option value="">Choose folder…</option>';
        _threadFolders.forEach(f => {
          const sel = (group.manualMatch && group.manualMatch.id === f.id) ? " selected" : "";
          html += '<option value="' + esc(f.id) + '"' + sel + '>' + esc(f.displayName) + '</option>';
        });
        html += '</select>';
      }

      html += '<div class="tl-actions" id="tl-actions-' + idx + '">' + buildActionButtons(idx) + '</div>';
      html += '</div>';
    }

    html += '</div>';
  });

  el.innerHTML = html;
}

function buildActionButtons(idx) {
  const group = _threadGroups[idx];
  const checkedCount = group.emails.filter(e => e.checked).length;
  const folder = group.match || group.manualMatch;
  const fileOff = (!folder || checkedCount === 0) ? " disabled" : "";
  const delOff  = checkedCount === 0 ? " disabled" : "";
  const flagOff = checkedCount === 0 ? " disabled" : "";
  const n = checkedCount > 0 ? " (" + checkedCount + ")" : "";
  return '<button class="tl-btn tl-file"'   + fileOff + ' onclick="fileThread('   + idx + ')">File'   + n + '</button>' +
         '<button class="tl-btn tl-delete"' + delOff  + ' onclick="deleteThread(' + idx + ')">Delete' + n + '</button>' +
         '<button class="tl-btn tl-flag"'   + flagOff + ' onclick="flagThread('   + idx + ')">Flag'   + n + '</button>' +
         '<button class="tl-btn tl-skip" onclick="skipThread(' + idx + ')">Skip</button>';
}

function toggleThread(idx) {
  const group = _threadGroups[idx];
  if (!group || group.done) return;
  const willExpand = !group.expanded;
  const needsLoad = willExpand && group.emails.some(e => e.body === null);
  group.expanded = willExpand;
  renderThreadList();
  if (needsLoad) loadThreadBodies(group);
}

async function loadThreadBodies(group) {
  const token = await ensureFreshToken().catch(() => null);
  if (!token) return;
  const rawBodies = {};
  await Promise.all(group.emails.map(async e => {
    if (e.body !== null) return;
    const details = await fetchEmailDetails(token, e.msg.id)
      .catch(() => ({ body: null, isReplied: false, isForwarded: false }));
    rawBodies[e.msg.id] = details.body || "";
    e.body = extractPreviewLines(details.body, 5) || "(no preview)";
    e.isReplied = details.isReplied;
    e.isForwarded = details.isForwarded;
  }));
  if (!group.match) {
    const counts = {};
    for (const e of group.emails) {
      const allRecip = [...(e.msg.toRecipients||[]), ...(e.msg.ccRecipients||[])];
      const fromAddr = e.msg.from?.emailAddress?.address || "";
      const fromName = e.msg.from?.emailAddress?.name || "";
      const pt = [fromName, fromAddr, ...allRecip.map(r => (r.emailAddress?.name||"") + " " + (r.emailAddress?.address||""))].join(" ");
      const m = matchFolder({ subject: e.msg.subject || "", participantText: pt, bodyText: rawBodies[e.msg.id] || "" }, _threadFolders);
      if (m) { if (!counts[m.id]) counts[m.id] = { folder: m, n: 0 }; counts[m.id].n++; }
    }
    const entries = Object.values(counts);
    if (entries.length) group.match = entries.reduce((a, b) => b.n > a.n ? b : a).folder;
  }
  if (group.expanded && !group.done) renderThreadList();
}

function onCheckChange(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  group.emails.forEach(e => {
    const chk = document.getElementById("chk-" + e.msg.id);
    if (chk) e.checked = chk.checked;
  });
  const actionsEl = document.getElementById("tl-actions-" + idx);
  if (actionsEl) actionsEl.innerHTML = buildActionButtons(idx);
}

function onFolderPick(idx, folderId) {
  const group = _threadGroups[idx];
  if (!group) return;
  group.manualMatch = folderId ? (_threadFolders.find(f => f.id === folderId) || null) : null;
  const actionsEl = document.getElementById("tl-actions-" + idx);
  if (actionsEl) actionsEl.innerHTML = buildActionButtons(idx);
}

function setThreadWorking(idx, msg) {
  const actionsEl = document.getElementById("tl-actions-" + idx);
  if (actionsEl) actionsEl.innerHTML = '<span class="tl-working">' + esc(msg) + '</span>';
}

async function fileThread(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const folder = group.match || group.manualMatch;
  if (!folder) return;
  const checked = group.emails.filter(e => e.checked);
  if (!checked.length) return;
  setThreadWorking(idx, "Filing…");
  try {
    const token = await ensureFreshToken();
    for (const e of checked) await moveMessage(token, e.msg.id, folder.id);
    markThreadDone(idx);
  } catch(err) {
    setThreadWorking(idx, "Error — try again");
  }
}

async function deleteThread(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const checked = group.emails.filter(e => e.checked);
  if (!checked.length) return;
  setThreadWorking(idx, "Deleting…");
  try {
    const token = await ensureFreshToken();
    for (const e of checked) await moveMessage(token, e.msg.id, "deleteditems");
    markThreadDone(idx);
  } catch(err) {
    setThreadWorking(idx, "Error — try again");
  }
}

function skipThread(idx) {
  markThreadDone(idx);
}

async function flagThread(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  const checked = group.emails.filter(e => e.checked);
  if (!checked.length) return;
  setThreadWorking(idx, "Flagging…");
  try {
    const token = await ensureFreshToken();
    for (const e of checked) {
      await fetch(`${GRAPH_BASE}/me/messages/${e.msg.id}`, {
        method: "PATCH",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ flag: { flagStatus: "flagged" } })
      });
    }
    markThreadDone(idx);
  } catch(err) {
    setThreadWorking(idx, "Error — try again");
  }
}

function markThreadDone(idx) {
  const group = _threadGroups[idx];
  if (!group) return;
  group.done = true;
  group.expanded = false;

  const nextIdx = _threadGroups.findIndex((g, i) => i > idx && !g.done);
  if (nextIdx !== -1) _threadGroups[nextIdx].expanded = true;

  renderThreadList();

  if (nextIdx !== -1) {
    const nextEl = document.getElementById("tg-" + nextIdx);
    if (nextEl) nextEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
    if (_threadGroups[nextIdx].emails.some(e => e.body === null)) {
      loadThreadBodies(_threadGroups[nextIdx]);
    }
  } else if (_threadGroups.every(g => g.done)) {
    document.getElementById("queue-status").textContent = "All done ✓";
  }
}

// --- Process Unfiled ---

async function processUnfiled() {
  const btn = document.getElementById("processBtn");
  const baselineBtn = document.getElementById("setBaselineBtn");
  const statusEl = document.getElementById("queue-status");

  btn.disabled = true;
  baselineBtn.style.display = "none";
  document.getElementById("thread-list").style.display = "none";

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
    const token = await ensureFreshToken();
    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
      `?$filter=sentDateTime gt ${lastRun}` +
      `&$top=100&$orderby=sentDateTime asc` +
      `&$select=id,subject,toRecipients,ccRecipients,sentDateTime,from,conversationId,flag`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];

    Office.context.roamingSettings.set("sent_last_run", newTimestamp);
    Office.context.roamingSettings.saveAsync(() => {});

    const nonCalendar = messages.filter(m => !isCalendarMessage(m.subject || "") && (m.flag?.flagStatus || "notFlagged") === "notFlagged");
    if (nonCalendar.length === 0) {
      statusEl.textContent = "No new sent emails to process.";
      btn.disabled = false;
      return;
    }

    const folders = parseFolders(foldersJson);
    const groups = groupByThread(nonCalendar, folders);
    statusEl.textContent = `${groups.length} thread${groups.length !== 1 ? "s" : ""} to review:`;
    initThreadList(groups, folders);
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
  document.getElementById("thread-list").style.display = "none";
  statusEl.textContent = "Scanning Inbox…";

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) {
    statusEl.textContent = "No case folders cached. Click Refresh Folders first.";
    btn.disabled = false;
    return;
  }

  try {
    const token = await ensureFreshToken();
    const msgsRes = await fetch(
      `${GRAPH_BASE}/me/mailFolders/Inbox/messages` +
      `?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId,flag` +
      `&$top=100&$orderby=receivedDateTime desc`,
      { headers: { Authorization: "Bearer " + token } }
    );
    if (!msgsRes.ok) throw new Error("Graph " + msgsRes.status);
    const messages = (await msgsRes.json()).value || [];

    const nonCalendar = messages.filter(m => !isCalendarMessage(m.subject || "") && (m.flag?.flagStatus || "notFlagged") === "notFlagged");
    if (nonCalendar.length === 0) {
      statusEl.textContent = "Inbox is empty.";
      btn.disabled = false;
      return;
    }

    const folders = parseFolders(foldersJson);
    const groups = groupByThread(nonCalendar, folders);
    statusEl.textContent = `${groups.length} thread${groups.length !== 1 ? "s" : ""} to review:`;
    initThreadList(groups, folders);
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
  btn.disabled = true;
  statusEl.textContent = "Refreshing folder list...";
  try {
    const token = await ensureFreshToken();
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
  const texts = [email.subject, email.participantText, email.bodyText || ""].filter(Boolean);
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
