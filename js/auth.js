/* ==========================================================================
   EAT COE — Microsoft SSO sign-in (MSAL.js)
   Sign-in ONLY. No Graph document/permission calls happen here — this just
   proves who the visitor is (their real Microsoft work identity) and feeds
   their real display name into the rest of the site (activity log, tags,
   etc). The Viewer/Contributor toggle in site.js is unchanged by this file.

   TO ENABLE:
   1. Ask your Entra ID admin to register this site as a Single-page
      application (see README.md for the exact steps).
   2. Paste the Application (client) ID and Directory (tenant) ID below.
   3. That's it — no other code changes needed. Until you do this, the two
      placeholders below keep SSO OFF and the site behaves exactly as before.
   ========================================================================== */

// This MUST exactly match a Redirect URI registered on the Entra ID app
// registration (Authentication blade) — protocol, host, path, and trailing
// slash all have to match character-for-character, or sign-in fails with
// "AADSTS50011: redirect URI mismatch". It does NOT need to be login.html —
// MSAL's popup flow just briefly loads this URL to capture the auth
// response, then closes the popup automatically; it's independent of which
// page in the app the person actually clicked "Log In" from.
const REGISTERED_REDIRECT_URI = "https://malkiattestingxperts.github.io/eat-coe-site";

const MSAL_CONFIG = {
  auth: {
    clientId: "9a817d03-ec3b-4e4f-8fa6-b7278cab47fe",
    authority: "https://login.microsoftonline.com/d7e861c9-d924-4413-86db-05780e928657",
    redirectUri: REGISTERED_REDIRECT_URI,
    postLogoutRedirectUri: REGISTERED_REDIRECT_URI
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

// Sign-in is considered "configured" only once both placeholders are replaced.
const SSO_ENABLED =
  MSAL_CONFIG.auth.clientId.indexOf("PASTE-YOUR") === -1 &&
  MSAL_CONFIG.auth.authority.indexOf("PASTE-YOUR") === -1 &&
  typeof msal !== "undefined";

// Set to false if you'd rather make sign-in optional (a button, not a gate).
const REQUIRE_SIGNIN = true;

let msalInstance = null;
if (SSO_ENABLED) {
  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
}

function getActiveAccount() {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  return accounts.length ? accounts[0] : null;
}

async function signIn() {
  if (!SSO_ENABLED) {
    alert(
      "Log in isn't wired up yet.\n\n" +
      "Your Entra ID admin needs to register this site (Single-page application) " +
      "and give you a Client ID + Tenant ID. Paste those into js/auth.js — see the " +
      "\"SSO login\" section in README.md for the exact steps.\n\n" +
      "Until then, everyone browses as a read-only Viewer."
    );
    return;
  }
  if (!msalInstance) return;
  try {
    const result = await msalInstance.loginPopup({ scopes: ["User.Read"] });
    msalInstance.setActiveAccount(result.account);
    onSignedIn(result.account);
    // On the dedicated login page, a successful sign-in sends you straight
    // to the Home page — Home itself has no login button, it assumes
    // whoever reaches it is already authenticated.
    if (isLoginPage()) {
      window.location.href = "index.html";
    }
  } catch (e) {
    console.error("Microsoft sign-in failed:", e);
    const errorCode = e && e.errorCode ? e.errorCode : (e && e.name) || "unknown_error";
    const errorMessage = e && e.errorMessage ? e.errorMessage : (e && e.message) || String(e);

    let hint = "";
    if (errorCode.indexOf("50011") !== -1 || /redirect/i.test(errorMessage)) {
      hint =
        "\n\nThis usually means the redirect URI this page is sending doesn't " +
        "exactly match what's registered in Entra ID.\n\n" +
        "This page is sending:\n  " + MSAL_CONFIG.auth.redirectUri + "\n\n" +
        "Ask your Entra ID admin to check App registrations → this app → " +
        "Authentication → Redirect URIs, and add that exact URL (protocol, " +
        "host, port, and path all have to match — e.g. \"localhost\" and " +
        "\"127.0.0.1\" are treated as different origins).";
    } else if (errorCode === "popup_window_error" || errorCode === "user_cancelled") {
      hint =
        "\n\nThis usually means the sign-in popup was blocked or closed. " +
        "Check for a popup-blocked icon in the address bar and allow popups " +
        "for this site, then try again.";
    } else if (errorCode.indexOf("65001") !== -1 || /consent/i.test(errorMessage)) {
      hint =
        "\n\nThis usually means the app needs admin consent for the requested " +
        "permissions. Ask your Entra ID admin to grant admin consent in " +
        "Entra ID → App registrations → this app → API permissions.";
    }

    alert(
      "Sign-in failed.\n\n" +
      "Error code: " + errorCode + "\n" +
      "Details: " + errorMessage +
      hint
    );
  }
}

async function signOut() {
  if (!msalInstance) return;
  const account = getActiveAccount();
  try {
    await msalInstance.logoutPopup({ account });
  } catch (e) {
    console.error("Sign-out failed:", e);
  }
  onSignedOut();
}

function onSignedIn(account) {
  const displayName = account.name || account.username || "Signed-in user";
  if (typeof setUserName === "function") setUserName(displayName);
  // No manual role switcher anymore: a real, signed-in identity is treated as
  // Contributor. Once real SharePoint-group enforcement is added later, this
  // is the spot to check actual group membership instead of granting it outright.
  if (typeof setRole === "function") setRole("contributor");
  document.body.classList.remove("signin-required");
  renderAuthUI();

  // Load the real SharePoint document list (Microsoft Graph) in the
  // background. This is a separate token request from the basic sign-in
  // above (User.Read), so if Sites/Files permissions need a moment of
  // incremental consent or aren't granted yet, it fails quietly and the
  // site just keeps showing the built-in document list instead.
  if (typeof initGraphCatalog === "function") {
    initGraphCatalog();
  }
}

function onSignedOut() {
  if (typeof setUserName === "function") setUserName("Guest");
  if (typeof setRole === "function") setRole("viewer");
  if (REQUIRE_SIGNIN && !isLoginPage()) document.body.classList.add("signin-required");
  renderAuthUI();
}

function escapeHtmlAuth(str) {
  return String(str).replace(/[&<>"']/g, s => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[s]));
}

function renderAuthUI() {
  const box = document.getElementById("authBox");
  if (!box) return;

  const account = getActiveAccount();
  if (account) {
    box.innerHTML = `
      <span class="signed-in-name">👤 ${escapeHtmlAuth(account.name || account.username)}</span>
      <button class="signout-btn" onclick="signOut()">Log Out</button>`;
  } else {
    box.innerHTML = `
      <button class="signin-btn" onclick="signIn()">
        Log In
      </button>${SSO_ENABLED ? "" : '<span class="sso-note">(setup pending)</span>'}`;
  }
}

function isLoginPage() {
  return /(^|\/)login\.html$/.test(window.location.pathname);
}

document.addEventListener("DOMContentLoaded", async () => {
  if (!SSO_ENABLED) {
    renderAuthUI();
    return;
  }
  try {
    await msalInstance.handleRedirectPromise();
  } catch (e) {
    console.error("MSAL redirect handling failed:", e);
  }
  const account = getActiveAccount();
  if (account) {
    onSignedIn(account);
    // Already signed in but landed on the login page anyway (e.g. via a
    // bookmark) — no need to show it, go straight to Home.
    if (isLoginPage()) {
      window.location.href = "index.html";
    }
  } else {
    onSignedOut();
    // Not signed in and this isn't the login page itself — send them there
    // instead of showing a blurred/gated version of the real page.
    if (REQUIRE_SIGNIN && !isLoginPage()) {
      window.location.href = "login.html";
    }
  }
});
