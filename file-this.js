"use strict";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

let _currentItem = null;
let _folders = [];
let _match = null;
let _manualMatch = null;
let _bodyFetched = false;
let _lastSubject = null;
let _lastParticipantText = null;
let _isInternal = false;
let _pendingRestId = null;
let _pendingTimestamp = null;
let _pendingSubject = null;

Office.onReady(async () => {
  document.getElementById("ver").textContent =
    typeof FILE_THIS_VERSION !== "undefined" ? FILE_THIS_VERSION : "?";

  // Handle auto-file pending item (set by OnMessageSent event handler)
  const raw = localStorage.getItem("hmf_auto_file_pending");
  if (raw) {
    localStorage.removeItem("hmf_auto_file_pending");
    loadPendingItem(JSON.parse(raw));
    // Still register storage listener for subsequent sends while pane stays open
    registerStorageListener();
    return;
  }

  await loadCurrentItem();
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
  registerStorageListener();
});

function registerStorageListener() {
  window.addEventListener("storage", (e) => {
    if (e.key === "hmf_auto_file_pending" && e.newValue) {
      localStorage.removeItem("hmf_auto_file_pending");
      loadPendingItem(JSON.parse(e.newValue));
    }
  });
}

function onItemChanged() {
  if (_pendingRestId) return; // don't clobber a pending auto-file
  loadCurrentItem();
}

// --- Normal mode: read from live mailbox item ---

async function loadCurrentItem() {
  _match = null;
  _manualMatch = null;
  _bodyFetched = false;
  _pendingRestId = null;

  const item = Office.context.mailbox.item;
  _currentItem = item;

  if (!item) {
    renderIdle("No email selected.");
    return;
  }

  const foldersJson = Office.context.roamingSettings.get("case_folders");
  if (!foldersJson) {
    renderIdle("No case folders cached — open the HMF Email Filer settings pane first.");
    return;
  }
  _folders = parseFolders(foldersJson);

  const subject = item.subject || "";
  const fromAddr = (item.from && item.from.emailAddress) || "";
  const fromName = (item.from && item.from.displayName) || "";
  const toText = (item.to || []).map(r => (r.displayName || "") + " " + (r.emailAddress || "")).join(" ");
  const ccText = (item.cc || []).map(r => (r.displayName || "") + " " + (r.emailAddress || "")).join(" ");
  const participantText = [fromName, fromAddr, toText, ccText].join(" ");

  _lastSubject = subject;
  _lastParticipantText = participantText;

  const allAddresses = [fromAddr,
    ...(item.to || []).map(r => r.emailAddress),
    ...(item.cc || []).map(r => r.emailAddress)
  ].filter(Boolean);
  _isInternal = !hasExternalRecipient(allAddresses, USER_DOMAIN);

  _match = _isInternal ? null : matchFolder({ subject, participantText }, _folders);

  renderUI({ subject, fromName, fromAddr });

  if (!_bodyFetched && !_isInternal && !_match) {
    fetchBodyAndRefine();
  }
}

async function fetchBodyAndRefine() {
  _bodyFetched = true;
  try {
    const token = await ensureFreshToken();
    const restId = Office.context.mailbox.convertToRestId(
      _currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0
    );
    const res = await fetch(`${GRAPH_BASE}/me/messages/${restId}?$select=body`, {
      headers: { Authorization: "Bearer " + token, Prefer: 'outlook.body-content-type="text"' }
    });
    if (!res.ok) return;
    const data = await res.json();
    const bodyText = (data.body && data.body.content) || "";
    const refined = matchFolder({ subject: _lastSubject, participantText: _lastParticipantText, bodyText }, _folders);
    if (refined && !_manualMatch) {
      _match = refined;
      updateMatchDisplay();
    }
  } catch(e) {}
}

// --- Pending mode: data from OnMessageSent event handler via localStorage ---

function loadPendingItem(pending) {
  _currentItem = null;
  _pendingRestId = pending.restId;
  _pendingTimestamp = pending.timestamp || null;
  _pendingSubject = pending.subject || null;
  _manualMatch = null;
  _bodyFetched = false;
  _isInternal = false;

  const foldersJson = Office.context.roamingSettings.get("case_folders") || "[]";
  _folders = parseFolders(foldersJson);
  _match = pending.match ? (_folders.find(f => f.id === pending.match.id) || null) : null;

  _lastSubject = pending.subject;
  _lastParticipantText = pending.fromName + " " + pending.fromAddr;

  renderUI({ subject: pending.subject, fromName: pending.fromName, fromAddr: pending.fromAddr });

  if (!_match) {
    fetchBodyAndRefineById(pending.restId);
  }
}

async function fetchBodyAndRefineById(restId) {
  _bodyFetched = true;
  try {
    const token = await ensureFreshToken();
    const res = await fetch(`${GRAPH_BASE}/me/messages/${restId}?$select=body`, {
      headers: { Authorization: "Bearer " + token, Prefer: 'outlook.body-content-type="text"' }
    });
    if (!res.ok) return;
    const data = await res.json();
    const bodyText = (data.body && data.body.content) || "";
    const refined = matchFolder({ subject: _lastSubject, participantText: _lastParticipantText, bodyText }, _folders);
    if (refined && !_manualMatch) {
      _match = refined;
      updateMatchDisplay();
    }
  } catch(e) {}
}

function resumeNormalMode() {
  _pendingRestId = null;
  _pendingTimestamp = null;
  _pendingSubject = null;
  loadCurrentItem().then(() => {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
  });
}

// --- Shared UI ---

function renderUI({ subject, fromName, fromAddr }) {
  document.getElementById("email-subject").textContent = subject || "(no subject)";
  document.getElementById("email-from").textContent =
    fromName ? fromName + " <" + fromAddr + ">" : fromAddr || "";
  document.getElementById("match-head").style.display = "block";

  const effectiveFolder = _manualMatch || _match;
  const matchValueEl = document.getElementById("match-value");
  const folderRow = document.getElementById("folder-row");
  const actionsEl = document.getElementById("actions");
  document.getElementById("status").textContent = "";
  document.getElementById("status").className = "status";

  if (_isInternal) {
    matchValueEl.textContent = "Internal";
    matchValueEl.className = "match-value internal";
    folderRow.style.display = "none";
    actionsEl.innerHTML =
      '<button class="btn btn-delete" onclick="deleteIt()">Delete</button>' +
      '<button class="btn btn-ignore" onclick="ignoreIt()">Ignore</button>';
  } else if (effectiveFolder) {
    matchValueEl.textContent = "→ " + effectiveFolder.displayName;
    matchValueEl.className = "match-value";
    renderFolderPicker(effectiveFolder.id);
    folderRow.style.display = "block";
    actionsEl.innerHTML =
      '<button class="btn btn-file" onclick="fileIt()">File</button>' +
      '<button class="btn btn-delete" onclick="deleteIt()">Delete</button>' +
      '<button class="btn btn-ignore" onclick="ignoreIt()">Ignore</button>';
  } else {
    matchValueEl.textContent = "(no match)";
    matchValueEl.className = "match-value nomatch";
    renderFolderPicker(null);
    folderRow.style.display = "block";
    actionsEl.innerHTML =
      '<button class="btn btn-file" id="fileBtn" onclick="fileIt()" disabled>File</button>' +
      '<button class="btn btn-delete" onclick="deleteIt()">Delete</button>' +
      '<button class="btn btn-ignore" onclick="ignoreIt()">Ignore</button>';
  }
}

function renderFolderPicker(selectedId) {
  const picker = document.getElementById("folderPicker");
  picker.innerHTML = '<option value="">Choose folder…</option>';
  for (const f of _folders) {
    const opt = document.createElement("option");
    opt.value = f.id;
    opt.textContent = f.displayName;
    if (f.id === selectedId) opt.selected = true;
    picker.appendChild(opt);
  }
}

function updateMatchDisplay() {
  const effectiveFolder = _manualMatch || _match;
  if (!effectiveFolder) return;
  document.getElementById("match-value").textContent = "→ " + effectiveFolder.displayName;
  document.getElementById("match-value").className = "match-value";
  renderFolderPicker(effectiveFolder.id);
  const actionsEl = document.getElementById("actions");
  if (actionsEl.querySelector("#fileBtn")) {
    actionsEl.innerHTML =
      '<button class="btn btn-file" onclick="fileIt()">File</button>' +
      '<button class="btn btn-delete" onclick="deleteIt()">Delete</button>' +
      '<button class="btn btn-ignore" onclick="ignoreIt()">Ignore</button>';
  }
}

function onFolderChange() {
  const picker = document.getElementById("folderPicker");
  const folderId = picker.value;
  _manualMatch = folderId ? (_folders.find(f => f.id === folderId) || null) : null;
  const fileBtn = document.getElementById("fileBtn");
  if (fileBtn) fileBtn.disabled = !folderId;
  if (_manualMatch) {
    document.getElementById("match-value").textContent = "→ " + _manualMatch.displayName;
    document.getElementById("match-value").className = "match-value";
  } else {
    document.getElementById("match-value").textContent = "(no match)";
    document.getElementById("match-value").className = "match-value nomatch";
  }
}

async function resolveRestId(token) {
  if (!_pendingRestId) {
    return Office.context.mailbox.convertToRestId(
      _currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0
    );
  }
  // Try stored restId first; if the send changed the message ID, fall back to
  // searching recent Sent Items by subject and approximate send time.
  const probe = await fetch(`${GRAPH_BASE}/me/messages/${_pendingRestId}?$select=id`, {
    headers: { Authorization: "Bearer " + token }
  });
  if (probe.ok) return _pendingRestId;

  // Fallback: find the message in Sent Items
  const since = _pendingTimestamp
    ? new Date(_pendingTimestamp - 5000).toISOString()
    : new Date(Date.now() - 120000).toISOString();
  const search = await fetch(
    `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
    `?$filter=sentDateTime ge ${since}&$select=id,subject&$top=20&$orderby=sentDateTime desc`,
    { headers: { Authorization: "Bearer " + token } }
  );
  if (!search.ok) throw new Error("Could not locate sent message");
  const msgs = (await search.json()).value || [];
  const match = msgs.find(m => m.subject === _pendingSubject);
  if (!match) throw new Error("Sent message not found — try again in a moment");
  return match.id;
}

async function fileIt() {
  const folder = _manualMatch || _match;
  if (!folder) return;
  const wasPending = !!_pendingRestId;
  setWorking("Filing…");
  try {
    const token = await ensureFreshToken();
    const restId = await resolveRestId(token);
    const res = await fetch(`${GRAPH_BASE}/me/messages/${restId}/move`, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: folder.id })
    });
    if (!res.ok) throw new Error("Graph " + res.status);
    setDone("Filed to " + folder.displayName);
    if (wasPending) setTimeout(resumeNormalMode, 2000);
  } catch(e) {
    setError("Error — " + e.message);
  }
}

async function deleteIt() {
  const wasPending = !!_pendingRestId;
  setWorking("Deleting…");
  try {
    const token = await ensureFreshToken();
    const restId = await resolveRestId(token);
    const res = await fetch(`${GRAPH_BASE}/me/messages/${restId}/move`, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: "deleteditems" })
    });
    if (!res.ok) throw new Error("Graph " + res.status);
    setDone("Moved to Deleted Items");
    if (wasPending) setTimeout(resumeNormalMode, 2000);
  } catch(e) {
    setError("Error — " + e.message);
  }
}

function ignoreIt() {
  const wasPending = !!_pendingRestId;
  if (wasPending) {
    resumeNormalMode();
  } else {
    renderIdle();
  }
}

function setWorking(msg) {
  document.getElementById("actions").innerHTML =
    '<span class="working">' + esc(msg) + '</span>';
  document.getElementById("status").textContent = "";
  document.getElementById("status").className = "status";
}

function setDone(msg) {
  document.getElementById("actions").innerHTML = "";
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status success";
}

function setError(msg) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status error";
}

function renderIdle(msg) {
  document.getElementById("email-subject").textContent = "";
  document.getElementById("email-from").textContent = "";
  document.getElementById("match-head").style.display = "none";
  document.getElementById("match-value").textContent = "";
  document.getElementById("match-value").className = "match-value";
  document.getElementById("folder-row").style.display = "none";
  document.getElementById("actions").innerHTML = "";
  const statusEl = document.getElementById("status");
  statusEl.textContent = msg || "";
  statusEl.className = "status";
}

// --- Token management ---

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
    if (!ok) throw new Error("Session expired. Please reconnect via the HMF Email Filer settings pane.");
  }
  return Office.context.roamingSettings.get("access_token");
}

// --- Matching helpers ---

function parseFolders(foldersJson) {
  return JSON.parse(foldersJson).map(f => ({
    displayName: f.displayName,
    id: f.id,
    keywords: f.displayName.split("/").map(k => k.trim().toLowerCase())
  }));
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

function hasExternalRecipient(emails, domain) {
  return emails.some(a => a && a.toLowerCase().slice(-(domain.length + 1)) !== "@" + domain);
}

function esc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
