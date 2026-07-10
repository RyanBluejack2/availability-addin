// ============================================================
//  auth.js — MSAL browser auth (delegated, user reads own calendar)
// ============================================================

var AvailabilityAuth = (function () {
  var cfg = window.AVAILABILITY_CONFIG;
  var msalInstance = null;

  function init() {
    if (msalInstance) return msalInstance;
    if (typeof msal === "undefined" || !msal.PublicClientApplication) {
      throw new Error("Sign-in library (MSAL) has not loaded yet. " +
        "Check your internet connection and that the page can reach alcdn.msauth.net.");
    }
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + cfg.TENANT_ID,
        redirectUri: cfg.REDIRECT_URI
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true // helps inside the Outlook webview
      }
    });
    return msalInstance;
  }

  // Safe check used at startup: returns null instead of throwing if MSAL
  // isn't loaded, so the UI can still render.
  function getAccount() {
    if (typeof msal === "undefined" || !msal.PublicClientApplication) return null;
    try {
      var accts = init().getAllAccounts();
      return accts && accts.length ? accts[0] : null;
    } catch (e) {
      return null;
    }
  }

  // Try silent token first; fall back to popup.
  async function getToken() {
    var instance = init();
    var account = getAccount();
    var request = { scopes: cfg.SCOPES };

    if (account) {
      try {
        request.account = account;
        var silent = await instance.acquireTokenSilent(request);
        return silent.accessToken;
      } catch (e) {
        // fall through to interactive
      }
    }

    var result = await instance.acquireTokenPopup(request);
    return result.accessToken;
  }

  async function signIn() {
    var instance = init();
    var result = await instance.loginPopup({ scopes: cfg.SCOPES });
    return result.account;
  }

  function signOut() {
    var account = getAccount();
    if (account) init().logoutPopup({ account: account });
  }

  return { init: init, getAccount: getAccount, getToken: getToken, signIn: signIn, signOut: signOut };
})();
