"use strict";

const AUTH_CLIENT_ID = "75dc31c8-0515-4c64-849c-3958218e2c5f";
const AUTH_AUTHORITY = "https://login.microsoftonline.com/hmflaw.com";
const AUTH_REDIRECT_URI = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const AUTH_SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite", "offline_access"];

const msalInstance = new msal.PublicClientApplication({
  auth: { clientId: AUTH_CLIENT_ID, authority: AUTH_AUTHORITY, redirectUri: AUTH_REDIRECT_URI },
  cache: { cacheLocation: "sessionStorage" },
});

const redirectPromise = msalInstance.handleRedirectPromise();

Office.onReady(async () => {
  try {
    const redirectResult = await redirectPromise;
    if (redirectResult) {
      Office.context.ui.messageParent(JSON.stringify({
        token: redirectResult.accessToken,
        refreshToken: extractRefreshToken(),
        expiry: redirectResult.expiresOn.getTime(),
      }));
      return;
    }

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      try {
        const silent = await msalInstance.acquireTokenSilent({ scopes: AUTH_SCOPES, account: accounts[0] });
        Office.context.ui.messageParent(JSON.stringify({
          token: silent.accessToken,
          refreshToken: extractRefreshToken(),
          expiry: silent.expiresOn.getTime(),
        }));
        return;
      } catch (e) {
        // Fall through to interactive
      }
    }

    await msalInstance.loginRedirect({ scopes: AUTH_SCOPES });
  } catch (e) {
    try {
      Office.context.ui.messageParent(JSON.stringify({ error: e.message || String(e) }));
    } catch (_) {}
  }
});

function extractRefreshToken() {
  for (let i = 0; i < sessionStorage.length; i++) {
    const key = sessionStorage.key(i);
    try {
      const val = JSON.parse(sessionStorage.getItem(key));
      if (val && val.credentialType === "RefreshToken" && val.secret) {
        return val.secret;
      }
    } catch (e) {}
  }
  return null;
}
