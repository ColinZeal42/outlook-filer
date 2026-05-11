"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  document.getElementById("refreshBtn").addEventListener("click", refreshFolders);
  document.getElementById("confirmBtn").addEventListener("click", confirmFiling);
  document.getElementById("cancelBtn").addEventListener("click", cancelFiling);
  document.getElementById("ver").textContent = typeof SETUP_VERSION !== "undefined" ? SETUP_VERSION : "?";
  await loadAndCheck();
});

async function loadAndCheck() {
  await new Promise(resolve =>
    Office.context.roamingSettings.loadAsync(resolve)
  );
  checkStatus();
  checkPending();
}

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

function checkPending() {
  const pendingJson = Office.context.roamingSettings.get("pending_filing");
  const pendingEl = document.getElementById("pending");
  const pendingMsg = document.getElementById("pendingMsg");

  if (!pendingJson) {
    pendingEl.style.display = "none";
    return;
  }

  try {
    const pending = JSON.parse(pendingJson);
    if (Date.now() - pending.ts > 300000) {
      Office.context.roamingSettings.remove("pending_filing");
      Office.context.roamingSettings.saveAsync(() => {});
      pendingEl.style.display = "none";
      return;
    }
    pendingMsg.textContent = `File "${pending.subject}" to "${pending.folderName}"?`;
    pendingEl.dataset.folderId = pending.folderId;
    pendingEl.style.display = "block";
  } catch(e) {
    pendingEl.style.display = "none";
  }
}

function confirmFiling() {
  const pendingJson = Office.context.roamingSettings.get("pending_filing");
  if (!pendingJson) return;

  try {
    const pending = JSON.parse(pendingJson);
    const approval = { folderId: pending.folderId, ts: Date.now() };
    Office.context.roamingSettings.set("filing_approved", JSON.stringify(approval));
    Office.context.roamingSettings.remove("pending_filing");
    Office.context.roamingSettings.saveAsync(() => {
      document.getElementById("pending").style.display = "none";
      document.getElementById("status").textContent = "Approved. Click Send again in your email.";
      document.getElementById("status").style.color = "green";
    });
  } catch(e) {}
}

function cancelFiling() {
  Office.context.roamingSettings.remove("pending_filing");
  Office.context.roamingSettings.saveAsync(() => {
    document.getElementById("pending").style.display = "none";
  });
}

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
  statusEl.style.color = "#555";
  try {
    const folders = await fetchCaseFolders(token);
    Office.context.roamingSettings.set("case_folders", JSON.stringify(folders));
    Office.context.roamingSettings.saveAsync(() => { btn.disabled = false; checkStatus(); });
  } catch(e) {
    statusEl.textContent = "Error refreshing folders: " + e.message;
    statusEl.style.color = "red";
    btn.disabled = false;
  }
}

function signIn() {
  const btn = document.getElementById("connectBtn");
  const statusEl = document.getElementById("status");
  btn.disabled = true;
  statusEl.textContent = "Opening sign-in window...";
  statusEl.style.color = "#555";

  Office.context.ui.displayDialogAsync(
    AUTH_URL,
    { height: 60, width: 40, displayInIframe: false },
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

          statusEl.textContent = "Signed in. Fetching case folders...";
          statusEl.style.color = "#555";

          let folders = [];
          try { folders = await fetchCaseFolders(msg.token); } catch(e) {}

          Office.context.roamingSettings.set("access_token", msg.token);
          Office.context.roamingSettings.set("token_expiry", String(msg.expiry || (Date.now() + 3600000)));
          if (msg.refreshToken) Office.context.roamingSettings.set("refresh_token", msg.refreshToken);
          Office.context.roamingSettings.set("case_folders", JSON.stringify(folders));

          Office.context.roamingSettings.saveAsync(saveResult => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
              checkStatus();
            } else {
              statusEl.textContent = "Error saving credentials: " + saveResult.error.message;
              statusEl.style.color = "red";
            }
            btn.disabled = false;
          });
        } catch(e) {
          statusEl.textContent = "Error: " + e.message;
          statusEl.style.color = "red";
          btn.disabled = false;
        }
      });

      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
        statusEl.textContent = "Sign-in cancelled.";
        statusEl.style.color = "#555";
        btn.disabled = false;
      });
    }
  );
}
