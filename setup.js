"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  checkStatus();
});

function checkStatus() {
  const refreshToken = Office.context.roamingSettings.get("refresh_token");
  const expiry = parseInt(Office.context.roamingSettings.get("token_expiry") || "0");
  const statusEl = document.getElementById("status");

  if (refreshToken) {
    if (Date.now() < expiry) {
      statusEl.textContent = "Connected. Token valid until " + new Date(expiry).toLocaleTimeString();
      statusEl.style.color = "green";
    } else {
      statusEl.textContent = "Token will auto-refresh on next send. Reconnect if filing stops working.";
      statusEl.style.color = "darkorange";
    }
  } else {
    statusEl.textContent = "Not connected. Click Connect to sign in.";
    statusEl.style.color = "#555";
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

      dlg.addEventHandler(Office.EventType.DialogMessageReceived, args => {
        dlg.close();
        try {
          const msg = JSON.parse(args.message);
          if (msg.token) {
            Office.context.roamingSettings.set("access_token", msg.token);
            Office.context.roamingSettings.set("token_expiry", String(msg.expiry || (Date.now() + 3600000)));
            if (msg.refreshToken) {
              Office.context.roamingSettings.set("refresh_token", msg.refreshToken);
            }
            Office.context.roamingSettings.saveAsync(saveResult => {
              if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                checkStatus();
              } else {
                statusEl.textContent = "Error saving credentials: " + saveResult.error.message;
                statusEl.style.color = "red";
              }
              btn.disabled = false;
            });
          } else {
            statusEl.textContent = "Sign-in failed: " + (msg.error || "Unknown error");
            statusEl.style.color = "red";
            btn.disabled = false;
          }
        } catch (e) {
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
