"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const CONFIRM_URL = "https://ColinZeal42.github.io/outlook-filer/confirm.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";
  checkStatus();
  startFilingPoller();
});

// --- Polling for pending filings ---

var _dialogOpen = false;

function startFilingPoller() {
  setInterval(checkForPendingFiling, 5000);
}

function checkForPendingFiling() {
  if (_dialogOpen) return;
  Office.context.roamingSettings.loadAsync(function() {
    var pendingJson = Office.context.roamingSettings.get("pending_filing");
    if (!pendingJson) return;
    try {
      var pending = JSON.parse(pendingJson);
      if (Date.now() - pending.ts > 300000) {
        Office.context.roamingSettings.remove("pending_filing");
        Office.context.roamingSettings.saveAsync(function() {});
        return;
      }
      // Clear it immediately so we don't show it twice
      Office.context.roamingSettings.remove("pending_filing");
      Office.context.roamingSettings.saveAsync(function() {});
      showConfirmDialog(pending);
    } catch(e) {}
  });
}

function showConfirmDialog(pending) {
  var subject = (pending.subject || "").slice(0, 50);
  var url = CONFIRM_URL +
    "?folder=" + encodeURIComponent(pending.folderName) +
    "&subject=" + encodeURIComponent(subject);

  _dialogOpen = true;
  Office.context.ui.displayDialogAsync(url, { height: 25, width: 35, displayInIframe: false },
    function(result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        _dialogOpen = false;
        return;
      }
      var dlg = result.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, function(args) {
        dlg.close();
        _dialogOpen = false;
        if (args.message === "yes") {
          moveEmail(pending);
        }
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, function() {
        _dialogOpen = false;
      });
    }
  );
}

function moveEmail(pending) {
  var token = Office.context.roamingSettings.get("access_token");
  if (!token) return;

  var sentAfter = new Date(pending.ts - 15000).toISOString();
  var safeSubject = pending.subject.replace(/'/g, "''");
  var attempt = 0;

  function tryMove() {
    var filter = encodeURIComponent(
      "subject eq '" + safeSubject + "' and sentDateTime ge " + sentAfter
    );
    var url = GRAPH_BASE + "/me/mailFolders/SentItems/messages?$filter=" + filter +
              "&$orderby=sentDateTime desc&$select=id&$top=1";

    graphFetch("GET", token, url, null, function(data) {
      var msgId = data && data.value && data.value[0] && data.value[0].id;
      if (msgId) {
        graphFetch("POST", token, GRAPH_BASE + "/me/messages/" + msgId + "/move",
          { destinationId: pending.folderId }, function() {});
      } else if (attempt < 5) {
        attempt++;
        setTimeout(tryMove, 2000);
      }
    }, function() {
      if (attempt < 5) { attempt++; setTimeout(tryMove, 2000); }
    });
  }
  tryMove();
}

function graphFetch(method, token, url, body, onSuccess, onError) {
  var xhr = new XMLHttpRequest();
  xhr.open(method, url);
  xhr.setRequestHeader("Authorization", "Bearer " + token);
  if (body) xhr.setRequestHeader("Content-Type", "application/json");
  xhr.onload = function() {
    if (xhr.status >= 200 && xhr.status < 300) {
      try { onSuccess(xhr.responseText ? JSON.parse(xhr.responseText) : null); } catch(e) {}
    } else { if (onError) onError(); }
  };
  xhr.onerror = function() { if (onError) onError(); };
  xhr.send(body ? JSON.stringify(body) : null);
}

// --- Status ---

function checkStatus() {
  const refreshToken = Office.context.roamingSettings.get("refresh_token");
  const expiry = parseInt(Office.context.roamingSettings.get("token_expiry") || "0");
  const foldersJson = Office.context.roamingSettings.get("case_folders");
  const folderCount = foldersJson ? JSON.parse(foldersJson).length : 0;
  const statusEl = document.getElementById("status");
  const refreshBtn = document.getElementById("refreshBtn");

  if (refreshToken) {
    if (Date.now() < expiry) {
      statusEl.textContent = `Connected. ${folderCount} case folder${folderCount !== 1 ? "s" : ""} cached.`;
      statusEl.style.color = "green";
    } else {
      statusEl.textContent = "Token expired. Click Connect to refresh.";
      statusEl.style.color = "darkorange";
    }
    refreshBtn.style.display = "inline-block";
  } else {
    statusEl.textContent = "Not connected. Click Connect to sign in.";
    statusEl.style.color = "#555";
    refreshBtn.style.display = "none";
  }
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
