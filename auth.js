"use strict";

const AUTH_CLIENT_ID = "75dc31c8-0515-4c64-849c-3958218e2c5f";
const AUTH_AUTHORITY = "https://login.microsoftonline.com/hmflaw.com";
const AUTH_REDIRECT_URI = "https://ColinZeal42.github.io/outlook-filer/auth.html";
const AUTH_SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite"];

const msalInstance = new msal.PublicClientApplication({
  auth: { clientId: AUTH_CLIENT_ID, authority: AUTH_AUTHORITY, redirectUri: AUTH_REDIRECT_URI },
  cache: { cacheLocation: "sessionStorage" },
});

// Call immediately per MSAL recommendation
const redirectPromise = msalInstance.handleRedirectPromise();

Office.onReady(async () => {
  try {
    const redirectResult = await redirectPromise;
    if (redirectResult) {
      Office.context.ui.messageParent(JSON.stringify({ token: redirectResult.accessToken }));
      return;
    }

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      try {
        const silent = await msalInstance.acquireTokenSilent({ scopes: AUTH_SCOPES, account: accounts[0] });
        Office.context.ui.messageParent(JSON.stringify({ token: silent.accessToken }));
        return;
      } catch (e) {
        // Fall through to interactive
      }
    }

    // Redirect to Microsoft login — page navigates away, comes back after auth
    await msalInstance.loginRedirect({ scopes: AUTH_SCOPES });
  } catch (e) {
    try {
      Office.context.ui.messageParent(JSON.stringify({ error: e.message || String(e) }));
    } catch (_) {}
  }
});
