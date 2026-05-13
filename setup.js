"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_DOMAIN = "hmflaw.com";

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("setBaselineBtn").addEventListener("click", setBaseline);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";

  await checkStatus();

  const token = Office.context.roamingSettings.get("access_token");
  if (token) populateFolderPicker();

  const mode = new URLSearchParams(window.location.search).get("mode");
  if (mode === "sent" || mode === "inbox") {
    setTimeout(() => openDialog(mode), 300);
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
  const accountEl = document.getElementById("account-status");
  const connectBtn = document.getElementById("connectBtn");

  if (refreshToken && Date.now() < expiry) {
    accountEl.textContent = "Connected";
    accountEl.style.color = "green";
    connectBtn.style.display = "none";
    renderFolderSection();
    renderPinnedSection();
    renderBaselineSection();
    renderBehaviorSection();
  } else if (refreshToken) {
    accountEl.textContent = "Refreshing session…";
    accountEl.style.color = "#555";
    connectBtn.style.display = "none";
    const ok = await refreshAccessToken();
    if (ok) {
      checkStatus();
    } else {
      accountEl.textContent = "Session expired — reconnect to continue";
      accountEl.style.color = "darkorange";
      connectBtn.style.display = "inline-block";
    }
  } else {
    accountEl.textContent = "Not connected";
    accountEl.style.color = "#555";
    connectBtn.style.display = "inline-block";
    renderBaselineSection();
    renderBehaviorSection();
  }
}

function renderFolderSection() {
  const foldersJson = Office.context.roamingSettings.get("case_folders");
  const count = foldersJson ? JSON.parse(foldersJson).length : 0;
  document.getElementById("folder-count").textContent =
    count ? `${count} case folder${count !== 1 ? "s" : ""} cached` : "No folders cached — select a root folder and refresh";
}

function renderBaselineSection() {
  const lastRun = Office.context.roamingSettings.get("sent_last_run");
  document.getElementById("baseline-status").textContent = lastRun
    ? "Sent emails filed through: " + new Date(lastRun).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })
    : "No baseline set";
}

function renderBehaviorSection() {
  populateSortPicker();
}

// --- Folder picker ---

async function fetchRootFolders(token) {
  const res = await fetch(
    `${GRAPH_BASE}/me/mailFolders?$top=100&$expand=childFolders($top=100)`,
    { headers: { Authorization: "Bearer " + token } }
  );
  if (!res.ok) throw new Error("Graph " + res.status);
  const data = await res.json();
  const folders = [];
  for (const f of data.value) {
    folders.push({ id: f.id, displayName: f.displayName, depth: 0 });
    for (const c of (f.childFolders || [])) {
      folders.push({ id: c.id, displayName: f.displayName + "/" + c.displayName, depth: 1 });
    }
  }
  return folders;
}

async function populateFolderPicker() {
  const picker = document.getElementById("rootFolderPicker");
  const refreshBtn = document.getElementById("refreshBtn");
  picker.innerHTML = '<option value="">Choose root folder…</option>';
  picker.disabled = true;
  refreshBtn.disabled = true;
  try {
    const token = await ensureFreshToken();
    const folders = await fetchRootFolders(token);
    const storedId = Office.context.roamingSettings.get("root_folder_id");
    for (const f of folders) {
      const opt = document.createElement("option");
      opt.value = f.id;
      opt.textContent = f.displayName;
      if (f.id === storedId) opt.selected = true;
      picker.appendChild(opt);
    }
    picker.disabled = false;
    if (!storedId) {
      const casesOpt = Array.from(picker.options).find(o => o.textContent === "__Cases" || o.textContent.endsWith("/__Cases"));
      if (casesOpt) casesOpt.selected = true;
    }
    if (picker.value) refreshBtn.disabled = false;
  } catch(e) {
    picker.innerHTML = '<option value="">Could not load folders</option>';
  }
}

function onRootFolderChange() {
  const picker = document.getElementById("rootFolderPicker");
  const folderId = picker.value;
  const folderName = picker.options[picker.selectedIndex]?.textContent || "";
  document.getElementById("refreshBtn").disabled = !folderId;
  if (!folderId) return;
  Office.context.roamingSettings.set("root_folder_id", folderId);
  Office.context.roamingSettings.set("root_folder_name", folderName);
  Office.context.roamingSettings.saveAsync(() => refreshFolders());
}

// --- Folder refresh ---

async function fetchCaseFolders(token) {
  let rootId = Office.context.roamingSettings.get("root_folder_id");

  if (!rootId) {
    const res = await fetch(`${GRAPH_BASE}/me/mailFolders?$top=100`, {
      headers: { Authorization: "Bearer " + token }
    });
    if (!res.ok) throw new Error("Graph " + res.status);
    const data = await res.json();
    const casesFolder = data.value.find(f => f.displayName === "__Cases");
    if (!casesFolder) throw new Error("No root folder selected and __Cases not found");
    rootId = casesFolder.id;
    Office.context.roamingSettings.set("root_folder_id", rootId);
    Office.context.roamingSettings.set("root_folder_name", "__Cases");
    Office.context.roamingSettings.saveAsync(() => {});
    const picker = document.getElementById("rootFolderPicker");
    if (picker) {
      const opt = Array.from(picker.options).find(o => o.value === rootId);
      if (opt) opt.selected = true;
    }
  }

  const result = [];
  await fetchFolderChildren(token, rootId, "", result);
  result.sort((a, b) => a.displayName.localeCompare(b.displayName));
  return result;
}

async function fetchFolderChildren(token, folderId, prefix, result) {
  const children = [];
  let url = `${GRAPH_BASE}/me/mailFolders/${folderId}/childFolders` +
            `?$top=100&$select=id,displayName,childFolderCount`;
  while (url) {
    const res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!res.ok) throw new Error("Graph " + res.status);
    const data = await res.json();
    children.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }
  await Promise.all(children.map(async child => {
    const fullName = prefix ? prefix + " / " + child.displayName : child.displayName;
    result.push({ id: child.id, displayName: fullName });
    if (child.childFolderCount > 0) {
      await fetchFolderChildren(token, child.id, fullName, result);
    }
  }));
}

// --- Pinned folders ---

function renderPinnedSection() {
  const pinned = JSON.parse(Office.context.roamingSettings.get("pinned_folders") || "[]");
  const listEl = document.getElementById("pinned-list");
  if (!listEl) return;
  listEl.innerHTML = pinned.length === 0
    ? '<span style="color:#aaa">None</span>'
    : pinned.map(f =>
        '<div class="ss-pinned-row">' + esc(f.displayName) +
        ' <button class="ss-pin-remove" onclick="removePinnedFolder(\'' + esc(f.id) + '\')">✕</button>' +
        '</div>'
      ).join("");
  const pinBtn = document.getElementById("pinBtn");
  if (pinBtn) pinBtn.disabled = pinned.length >= 8;
  populatePinnedPicker(pinned);
}

async function populatePinnedPicker(pinned) {
  const picker = document.getElementById("pinnedFolderPicker");
  if (!picker) return;
  picker.disabled = true;
  try {
    const token = await ensureFreshToken();
    const folders = await fetchRootFolders(token);
    const pinnedIds = new Set(pinned.map(f => f.id));
    picker.innerHTML = '<option value="">Choose folder to pin…</option>';
    for (const f of folders) {
      if (!pinnedIds.has(f.id)) {
        const opt = document.createElement("option");
        opt.value = f.id;
        opt.textContent = f.displayName;
        picker.appendChild(opt);
      }
    }
    picker.disabled = pinned.length >= 8;
  } catch(e) {
    picker.innerHTML = '<option value="">Could not load folders</option>';
  }
}

function addPinnedFolder() {
  const picker = document.getElementById("pinnedFolderPicker");
  const folderId = picker.value;
  if (!folderId) return;
  const folderName = picker.options[picker.selectedIndex].textContent;
  const pinned = JSON.parse(Office.context.roamingSettings.get("pinned_folders") || "[]");
  if (pinned.length >= 8 || pinned.find(f => f.id === folderId)) return;
  pinned.push({ id: folderId, displayName: folderName });
  Office.context.roamingSettings.set("pinned_folders", JSON.stringify(pinned));
  Office.context.roamingSettings.saveAsync(() => renderPinnedSection());
}

function removePinnedFolder(id) {
  const pinned = JSON.parse(Office.context.roamingSettings.get("pinned_folders") || "[]");
  const updated = pinned.filter(f => f.id !== id);
  Office.context.roamingSettings.set("pinned_folders", JSON.stringify(updated));
  Office.context.roamingSettings.saveAsync(() => renderPinnedSection());
}

async function refreshFolders() {
  const refreshBtn = document.getElementById("refreshBtn");
  const folderCountEl = document.getElementById("folder-count");
  refreshBtn.disabled = true;
  folderCountEl.textContent = "Refreshing…";
  try {
    const token = await ensureFreshToken();
    const folders = await fetchCaseFolders(token);
    Office.context.roamingSettings.set("case_folders", JSON.stringify(folders));
    Office.context.roamingSettings.saveAsync(() => {
      refreshBtn.disabled = false;
      renderFolderSection();
    });
  } catch(e) {
    folderCountEl.textContent = "Error: " + e.message;
    refreshBtn.disabled = false;
  }
}

// --- Sort picker ---

function populateSortPicker() {
  const val = Office.context.roamingSettings.get("sort_order") || "date-desc";
  const picker = document.getElementById("sortOrderPicker");
  if (picker) picker.value = val;
}

function onSortOrderChange() {
  const val = document.getElementById("sortOrderPicker").value;
  Office.context.roamingSettings.set("sort_order", val);
  Office.context.roamingSettings.saveAsync(() => {});
}

// --- Baseline ---

function setBaseline() {
  const btn = document.getElementById("setBaselineBtn");
  btn.disabled = true;
  Office.context.roamingSettings.set("sent_last_run", new Date().toISOString());
  Office.context.roamingSettings.saveAsync(() => {
    btn.disabled = false;
    renderBaselineSection();
  });
}

// --- Dialog launcher ---

function openDialog(mode) {
  localStorage.setItem("hmf_access_token", Office.context.roamingSettings.get("access_token") || "");
  localStorage.setItem("hmf_token_expiry", Office.context.roamingSettings.get("token_expiry") || "0");
  localStorage.setItem("hmf_refresh_token", Office.context.roamingSettings.get("refresh_token") || "");
  localStorage.setItem("hmf_case_folders", Office.context.roamingSettings.get("case_folders") || "[]");
  localStorage.setItem("hmf_mode", mode);
  localStorage.setItem("hmf_sort_order", Office.context.roamingSettings.get("sort_order") || "date-desc");
  localStorage.setItem("hmf_pinned_folders", Office.context.roamingSettings.get("pinned_folders") || "[]");
  const sentLastRun = Office.context.roamingSettings.get("sent_last_run");
  if (sentLastRun) localStorage.setItem("hmf_sent_last_run", sentLastRun);

  const dialogUrl = "https://ColinZeal42.github.io/outlook-filer/dialog.html";
  Office.context.ui.displayDialogAsync(dialogUrl, { width: 70, height: 80, displayInIframe: false },
    result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return;
      const dlg = result.value;

      dlg.addEventHandler(Office.EventType.DialogMessageReceived, args => {
        try {
          const msg = JSON.parse(args.message);
          if (msg.action === "open-item" && msg.restId) {
            const ewsId = Office.context.mailbox.convertToEwsId(msg.restId, Office.MailboxEnums.RestVersion.v2_0);
            Office.context.mailbox.displayMessageForm(ewsId);
          }
        } catch(e) {}
      });

      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
        const tok = localStorage.getItem("hmf_access_token");
        const exp = localStorage.getItem("hmf_token_expiry");
        const ref = localStorage.getItem("hmf_refresh_token");
        if (tok) {
          Office.context.roamingSettings.set("access_token", tok);
          Office.context.roamingSettings.set("token_expiry", exp || "0");
          if (ref) Office.context.roamingSettings.set("refresh_token", ref);
          Office.context.roamingSettings.saveAsync(() => {});
        }
      });
    }
  );
}

// --- Sign in ---

function signIn() {
  const btn = document.getElementById("connectBtn");
  const accountEl = document.getElementById("account-status");
  btn.disabled = true;
  accountEl.textContent = "Opening sign-in window…";
  accountEl.style.color = "#555";

  Office.context.ui.displayDialogAsync(AUTH_URL, { height: 60, width: 40, displayInIframe: false },
    result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        accountEl.textContent = "Could not open sign-in window: " + result.error.message;
        accountEl.style.color = "red";
        btn.disabled = false;
        return;
      }
      const dlg = result.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, async args => {
        dlg.close();
        try {
          const msg = JSON.parse(args.message);
          if (!msg.token) {
            accountEl.textContent = "Sign-in failed: " + (msg.error || "Unknown error");
            accountEl.style.color = "red";
            btn.disabled = false;
            return;
          }
          accountEl.textContent = "Fetching case folders…";
          accountEl.style.color = "#555";
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
              populateFolderPicker();
            } else {
              accountEl.textContent = "Error saving: " + r.error.message;
              accountEl.style.color = "red";
            }
          });
        } catch(e) {
          accountEl.textContent = "Error: " + e.message;
          accountEl.style.color = "red";
          btn.disabled = false;
        }
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
        accountEl.textContent = "Sign-in cancelled.";
        btn.disabled = false;
      });
    }
  );
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

function stripClutter(text) {
  return text
    .replace(/\[.*?\]/g, "")
    .replace(/<https?:\/\/[^>]*>/g, "")
    .replace(/https?:\/\/\S+/g, "")
    .replace(/\s{2,}/g, " ");
}

function extractPreviewLines(text, maxLines) {
  if (!text) return "";
  const lines = text.split(/\r?\n/);
  const result = [];
  let blankRun = 0;
  for (const line of lines) {
    const t = stripClutter(line).trim();
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

function recipientAddresses(msg) {
  return [...(msg.toRecipients || []), ...(msg.ccRecipients || [])]
    .map(r => r.emailAddress && r.emailAddress.address)
    .filter(Boolean);
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

function openEmail(restId) {
  const ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
  Office.context.mailbox.displayMessageForm(ewsId);
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}
