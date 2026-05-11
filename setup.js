"use strict";

const AUTH_URL = "https://ColinZeal42.github.io/outlook-filer/auth.html";

Office.onReady(async () => {
  document.getElementById("connectBtn").addEventListener("click", signIn);
  await checkStatus();
});

async function checkStatus() {
  const stored = await OfficeRuntime.storage.getItems(["access_token", "token_expiry", "refresh_token"]);
  const statusEl = document.getElementById("status");

  if (stored.refresh_token) {
    const expiry = parseInt(stored.token_expiry || "0");
    if (Date.now() < expiry) {
      statusEl.textContent = "Connected. Token valid until " + new Date(expiry).toLocaleTimeString();
      statusEl.style.color = "green";
    } else {
      statusEl.textContent = "Token expired — will refresh automatically on next send. Reconnect if filing stops working.";
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

      dlg.addEventHandler(Office.EventType.DialogMessageReceived, async args => {
        dlg.close();
        try {
          const msg = JSON.parse(args.message);
          if (msg.token) {
            await OfficeRuntime.storage.setItems({
              access_token: msg.token,
              token_expiry: String(msg.expiry || (Date.now() + 3600000)),
              ...(msg.refreshToken ? { refresh_token: msg.refreshToken } : {}),
            });
            await checkStatus();
          } else {
            statusEl.textContent = "Sign-in failed: " + (msg.error || "Unknown error");
            statusEl.style.color = "red";
          }
        } catch (e) {
          statusEl.textContent = "Error: " + e.message;
          statusEl.style.color = "red";
        }
        btn.disabled = false;
      });

      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => {
        statusEl.textContent = "Sign-in cancelled.";
        statusEl.style.color = "#555";
        btn.disabled = false;
      });
    }
  );
}
