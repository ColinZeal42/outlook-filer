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
let _candidates = [];
let _ambiguous = false;
let _learnedMatch = false;
let _externalAddresses = [];

Office.onReady(async () => {
  document.getElementById("ver").textContent =
    typeof FILE_THIS_VERSION !== "undefined" ? FILE_THIS_VERSION : "?";

  await loadCurrentItem();
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
});

function onItemChanged() {
  loadCurrentItem();
}

// --- Normal mode: read from live mailbox item ---

async function loadCurrentItem() {
  _match = null;
  _manualMatch = null;
  _bodyFetched = false;
  _candidates = [];
  _ambiguous = false;
  _learnedMatch = false;
  _externalAddresses = [];

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

  _externalAddresses = allAddresses.filter(
    a => a && !a.toLowerCase().endsWith("@" + USER_DOMAIN)
  );

  if (!_isInternal) {
    const allCandidates = matchAllFolders({ subject, participantText }, _folders);
    _candidates = allCandidates;

    if (allCandidates.length === 0) {
      _match = null;
    } else if (allCandidates.length === 1) {
      _match = allCandidates[0];
    } else {
      const learnedContacts = JSON.parse(
        Office.context.roamingSettings.get("learned_contacts") || "{}"
      );
      const resolved = resolveAmbiguity(_externalAddresses, allCandidates, learnedContacts);
      if (resolved) {
        _match = resolved;
        _learnedMatch = true;
      } else {
        _match = null;
        _ambiguous = true;
      }
    }
  }

  renderUI({ subject, fromName, fromAddr });

  if (!_bodyFetched && !_isInternal && !_match && !_ambiguous) {
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
  } else if (effectiveFolder && _learnedMatch && !_manualMatch) {
    matchValueEl.textContent = "→ " + effectiveFolder.displayName + " ✓";
    matchValueEl.className = "match-value match-learned";
    renderFolderPicker(effectiveFolder.id);
    folderRow.style.display = "block";
    actionsEl.innerHTML =
      '<button class="btn btn-file" onclick="fileIt()">File</button>' +
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
  } else if (_ambiguous) {
    matchValueEl.textContent = "(pick folder)";
    matchValueEl.className = "match-value match-ambiguous";
    renderDisambigPicker();
    folderRow.style.display = "block";
    actionsEl.innerHTML =
      '<button class="btn btn-file" id="fileBtn" onclick="fileIt()" disabled>File</button>' +
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
  const pinned = JSON.parse(Office.context.roamingSettings.get("pinned_folders") || "[]");
  picker.innerHTML = '<option value="">Choose folder…</option>';
  if (pinned.length > 0) {
    for (const f of pinned) {
      const opt = document.createElement("option");
      opt.value = f.id;
      opt.textContent = "★ " + f.displayName;
      if (f.id === selectedId) opt.selected = true;
      picker.appendChild(opt);
    }
    const sep = document.createElement("option");
    sep.disabled = true;
    sep.textContent = "──────────";
    picker.appendChild(sep);
  }
  for (const f of _folders) {
    const opt = document.createElement("option");
    opt.value = f.id;
    opt.textContent = f.displayName;
    if (f.id === selectedId) opt.selected = true;
    picker.appendChild(opt);
  }
}

function renderDisambigPicker() {
  const picker = document.getElementById("folderPicker");
  picker.innerHTML = '<option value="">' + _candidates.length + ' matches — choose one…</option>';
  for (const c of _candidates) {
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = c.displayName;
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
  if (_ambiguous) {
    _manualMatch = folderId ? (_candidates.find(f => f.id === folderId) || null) : null;
  } else {
    const pinned = JSON.parse(Office.context.roamingSettings.get("pinned_folders") || "[]");
    _manualMatch = folderId ? ([...pinned, ..._folders].find(f => f.id === folderId) || null) : null;
  }
  const fileBtn = document.getElementById("fileBtn");
  if (fileBtn) fileBtn.disabled = !folderId;
  if (_manualMatch) {
    document.getElementById("match-value").textContent = "→ " + _manualMatch.displayName;
    document.getElementById("match-value").className = "match-value";
  } else if (_ambiguous) {
    document.getElementById("match-value").textContent = "(pick folder)";
    document.getElementById("match-value").className = "match-value match-ambiguous";
  } else {
    document.getElementById("match-value").textContent = "(no match)";
    document.getElementById("match-value").className = "match-value nomatch";
  }
}

function learnFromDisambiguation(folder) {
  const learned = JSON.parse(Office.context.roamingSettings.get("learned_contacts") || "{}");
  for (const addr of _externalAddresses) {
    learned[addr] = { folderId: folder.id, folderName: folder.displayName };
  }
  Office.context.roamingSettings.set("learned_contacts", JSON.stringify(learned));
  Office.context.roamingSettings.saveAsync(() => {});
}

async function fileIt() {
  const folder = _manualMatch || _match;
  if (!folder) return;
  if (_ambiguous) {
    learnFromDisambiguation(folder);
    _ambiguous = false;
    _learnedMatch = true;
  }
  setWorking("Filing…");
  try {
    const token = await ensureFreshToken();
    const restId = Office.context.mailbox.convertToRestId(
      _currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0
    );
    const res = await fetch(`${GRAPH_BASE}/me/messages/${restId}/move`, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: folder.id })
    });
    if (!res.ok) throw new Error("Graph " + res.status);
    setDone("Filed to " + folder.displayName);
  } catch(e) {
    setError("Error — " + e.message);
  }
}

async function deleteIt() {
  setWorking("Deleting…");
  try {
    const token = await ensureFreshToken();
    const restId = Office.context.mailbox.convertToRestId(
      _currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0
    );
    const res = await fetch(`${GRAPH_BASE}/me/messages/${restId}/move`, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: "deleteditems" })
    });
    if (!res.ok) throw new Error("Graph " + res.status);
    setDone("Moved to Deleted Items");
  } catch(e) {
    setError("Error — " + e.message);
  }
}

function ignoreIt() {
  renderIdle();
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

function matchAllFolders(email, folders) {
  const texts = [email.subject, email.participantText, email.bodyText || ""].filter(Boolean);
  const seen = new Set();
  const matches = [];
  for (let t = 0; t < texts.length; t++) {
    const lower = texts[t].toLowerCase();
    for (let f = 0; f < folders.length; f++) {
      if (seen.has(folders[f].id)) continue;
      const kws = folders[f].keywords;
      for (let k = 0; k < kws.length; k++) {
        if (lower.indexOf(kws[k]) !== -1) {
          seen.add(folders[f].id);
          matches.push(folders[f]);
          break;
        }
      }
    }
  }
  return matches;
}

function resolveAmbiguity(externalAddresses, candidates, learnedContacts) {
  for (const addr of externalAddresses) {
    const entry = learnedContacts[addr.toLowerCase()];
    if (entry) {
      const found = candidates.find(c => c.id === entry.folderId);
      if (found) return found;
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
