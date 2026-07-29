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

const MSAL_CONFIG = {
  auth: {
    clientId: "9a817d03-ec3b-4e4f-8fa6-b7278cab47fe",
    authority: "https://login.microsoftonline.com/d7e861c9-d924-4413-86db-05780e928657",
    redirectUri: window.location.origin + window.location.pathname
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
  } catch (e) {
    console.error("Microsoft sign-in failed:", e);
    alert("Sign-in failed or was cancelled. Check the browser console for details.");
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
}

function onSignedOut() {
  if (typeof setUserName === "function") setUserName("Guest");
  if (typeof setRole === "function") setRole("viewer");
  if (REQUIRE_SIGNIN) document.body.classList.add("signin-required");
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

document.addEventListener("DOMContentLoaded", async () => {
  if (!SSO_ENABLED) {
    renderAuthUI();
    return;
  }

  try {
    await msalInstance.initialize();      // <-- Required in MSAL v3
    await msalInstance.handleRedirectPromise();
  } catch (e) {
    console.error("MSAL initialization failed:", e);
  }

  const account = getActiveAccount();
  if (account) {
    onSignedIn(account);
  } else {
    onSignedOut();
  }
});
