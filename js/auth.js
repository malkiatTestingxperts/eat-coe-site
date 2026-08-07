/* ==========================================================================
   TREAT COE — Microsoft SSO sign-in (MSAL.js)
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
    cacheLocation: "sessionStorage",
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
// MSAL Browser v3+ requires calling and awaiting initialize() before any
// other API (loginPopup, acquireTokenSilent, handleRedirectPromise, etc.)
// — skipping this throws "uninitialized_public_client_application". Every
// function below that touches msalInstance awaits this first.
let msalReady = Promise.resolve();
if (SSO_ENABLED) {
  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  msalReady = msalInstance.initialize().catch(e => {
    console.error("MSAL initialize() failed:", e);
  });
}

function getActiveAccount() {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  return accounts.length ? accounts[0] : null;
}

async function signIn() {
  if (!SSO_ENABLED) {
    if (typeof showToast === "function") {
      showToast(
        "Log in isn't wired up yet. Your Entra ID admin needs to register this site and provide a Client ID + Tenant ID — see the \"SSO login\" section in README.md. Until then, everyone browses as a read-only Viewer.",
        "error", 0
      );
    }
    return;
  }
  if (!msalInstance) return;
  try {
    await msalReady;
    // Request the Graph scope (Sites.Read.All) together with the basic
    // sign-in scope in this one popup, instead of asking for it separately
    // the first time a document is opened. This used to be two sequential
    // popups — safe to merge now that Sites.Read.All has real tenant-wide
    // admin consent already, so including it upfront doesn't trigger any
    // extra approval screen, just a single normal sign-in.
    const loginScopes = (typeof GRAPH_SCOPES !== "undefined" && GRAPH_SCOPES && GRAPH_SCOPES.length)
      ? GRAPH_SCOPES
      : ["User.Read"];
    const result = await msalInstance.loginPopup({ scopes: loginScopes });
    msalInstance.setActiveAccount(result.account);
    onSignedIn(result.account);
    // On the dedicated login page, a successful sign-in sends you straight
    // to the Home page — Home itself has no login button, it assumes
    // whoever reaches it is already authenticated.
    if (isLoginPage()) {
      window.location.href = "index.html";
    }
  } catch (e) {
    const errorCode = e && e.errorCode ? e.errorCode : (e && e.name) || "unknown_error";
    const errorMessage = e && e.errorMessage ? e.errorMessage : (e && e.message) || String(e);

    // Someone closing or blocking the sign-in popup themselves is a normal,
    // deliberate action, not a real error — show nothing at all, the same
    // way the page looked before they clicked Log In.
    if (errorCode === "popup_window_error" || errorCode === "user_cancelled") {
      console.warn("Sign-in popup was closed or blocked before completing:", e);
      return;
    }

    console.error("Microsoft sign-in failed:", e);
    let hint = "";
    if (errorCode.indexOf("50011") !== -1 || /redirect/i.test(errorMessage)) {
      hint =
        " This usually means the redirect URI this page is sending doesn't exactly match what's registered in Entra ID. " +
        "This page is sending: " + MSAL_CONFIG.auth.redirectUri + ". Ask your Entra ID admin to check App registrations → " +
        "this app → Authentication → Redirect URIs.";
    } else if (errorCode.indexOf("65001") !== -1 || /consent/i.test(errorMessage)) {
      hint = " This usually means the app needs admin consent for the requested permissions — ask your Entra ID admin to grant it under App registrations → this app → API permissions.";
    } else if (errorCode.indexOf("50105") !== -1 || /not assigned/i.test(errorMessage)) {
      hint = " This is expected, not a bug: this account hasn't been assigned access yet — ask your Entra ID admin to add it under Enterprise Applications → this app → Users and groups.";
    } else if (/uninitialized_public_client_application/i.test(errorCode) || /initialize/i.test(errorMessage)) {
      hint = " MSAL's own startup step failed silently earlier — check the browser console right after the page loads for a \"MSAL initialize() failed\" message.";
    }

    if (typeof showToast === "function") {
      showToast("Sign-in failed (" + errorCode + "): " + errorMessage + hint, "error", 0);
    }
  }
}

async function signOut() {
  if (!msalInstance) return;
  // Deliberately local-only: does NOT call logoutPopup()/logoutRedirect(),
  // which hand off to Microsoft's own server-side logout page — that page
  // is the actual source of the "which account do you want to sign out
  // of?" confirmation screen, and per Microsoft's own support engineers
  // this is confirmed, by-design behavior of their v2.0 logout endpoint
  // with no ETA to change (see: learn.microsoft.com/answers, "Bypass the
  // account selection screen while logout"). Passing logoutHint (kept
  // below for reference) only narrows it to the right account — it can't
  // remove the confirmation click itself.
  //
  // Instead, this clears the signed-in account directly from MSAL's own
  // cache — the person is fully signed out of THIS app immediately, with
  // no extra screen at all. Tradeoff: their broader Microsoft session
  // (Outlook, Teams, etc.) stays active elsewhere — only this app's
  // session ends, not their whole Microsoft sign-in.
  try {
    await msalReady;
    const account = getActiveAccount();
    if (account) {
      msalInstance.setActiveAccount(null);
      try {
        const cache = msalInstance.getTokenCache ? msalInstance.getTokenCache() : null;
        if (cache && typeof cache.removeAccount === "function") {
          await cache.removeAccount(account);
        }
      } catch (e2) { /* best-effort — the sessionStorage clear below is the real guarantee */ }
    }
    // Guaranteed cleanup: our cache is configured to live in sessionStorage
    // (see MSAL_CONFIG above), so clearing it directly is a hard guarantee
    // no MSAL session state lingers, regardless of any edge case in the
    // removeAccount() call above.
    try { sessionStorage.clear(); } catch (e3) { /* ignore */ }
  } catch (e) {
    console.error("Sign-out failed:", e);
  }
  onSignedOut();
}

// ---- Group-based role assignment ----
// Fill these in with your two Entra ID security groups' real Object IDs —
// find each one under Entra ID → Groups → (click the group) → Object Id.
// Until both are filled in, everyone defaults to Viewer (the safer option),
// so this can be deployed before the group IDs are known without
// accidentally granting Moderator access to everyone.
const MODERATOR_GROUP_ID = "6a773ce8-d3f3-4ab1-9129-608524cbb9e9"; // TREAT-COE-Site Portal Moderators
const VIEWER_GROUP_ID = "9220b96b-9ed3-4fcc-b1cf-064752309e98"; // TREAT-COE-Site Portal Viewers

/**
 * Reads real group membership from the signed-in account's ID token. This
 * requires a one-time Entra ID configuration step (not code): the app
 * registration's Token configuration needs a "groups" optional claim added
 * to the ID token (Entra ID → App registrations → this app → Token
 * configuration → Add groups claim → check ID → Security groups). Once
 * that's done, group Object IDs the signed-in user belongs to show up
 * automatically in account.idTokenClaims.groups — no extra Graph
 * permission or API call needed for this specific check.
 *
 * Note: if someone belongs to more than ~200 groups, Azure AD omits this
 * claim entirely (a documented "groups overage" case) — uncommon for an
 * individual user, but if it ever comes up, this quietly falls back to
 * Viewer rather than breaking.
 */
function determineRoleFromGroups(account) {
  const claims = account && account.idTokenClaims;
  const groups = (claims && claims.groups) || [];
  const moderatorConfigured = MODERATOR_GROUP_ID.indexOf("PASTE-") !== 0;
  const viewerConfigured = VIEWER_GROUP_ID.indexOf("PASTE-") !== 0;

  if (moderatorConfigured && groups.includes(MODERATOR_GROUP_ID)) {
    return "contributor"; // internal role name kept as-is — reuses all existing .contributor-only gating
  }
  if (viewerConfigured && groups.includes(VIEWER_GROUP_ID)) {
    return "viewer";
  }
  // Default to the safer option: someone not recognized in either group
  // (or the groups claim isn't configured yet) sees view-only access, not
  // edit access, by default.
  return "viewer";
}

/**
 * Console-callable diagnostic: run debugMyGroups() in the browser console
 * while signed in to see exactly what group info (if any) is actually on
 * your current sign-in token, and which role that resolves to.
 */
function debugMyGroups() {
  const account = getActiveAccount();
  if (!account) {
    console.log("Not signed in.");
    return;
  }
  const claims = account.idTokenClaims || {};
  console.log("Groups on this token:", claims.groups || "(none present)");
  console.log("Groups overage indicator present:", !!(claims._claim_names && claims._claim_names.groups));
  console.log("Configured Moderator group ID:", MODERATOR_GROUP_ID);
  console.log("Configured Viewer group ID:", VIEWER_GROUP_ID);
  console.log("Role this resolves to:", determineRoleFromGroups(account));
}

function onSignedIn(account) {
  const displayName = account.name || account.username || "Signed-in user";
  if (typeof setUserName === "function") setUserName(displayName);
  const role = determineRoleFromGroups(account);
  if (typeof setRole === "function") setRole(role);
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
    const role = (typeof getRole === "function") ? getRole() : "viewer";
    const roleLabel = role === "contributor" ? "Moderator" : "Viewer";
    const roleClass = role === "contributor" ? "role-badge-moderator" : "role-badge-viewer";
    box.innerHTML = `
      <div class="identity-card">
        <span class="role-badge ${roleClass}">${roleLabel}</span>
        <span class="signed-in-name">👤 ${escapeHtmlAuth(account.name || account.username)}</span>
      </div>
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

/**
 * True when this window IS the transient popup MSAL opened for sign-in
 * (or sign-out), briefly loading the redirect URI just to hand the auth
 * response back to the window that opened it. If we let our own app logic
 * (in particular, the "redirect to login.html if not signed in yet" gate
 * below) run in that popup, it can navigate the popup away before MSAL's
 * own polling in the *opener* window gets a chance to read the response
 * and close the popup — which looks like sign-in "looping" back to the
 * login page instead of completing. So: if we're in that popup, do nothing
 * at all here and let MSAL handle it.
 */
function isMsalPopup() {
  try {
    return !!(window.opener && window.opener !== window && !window.opener.closed);
  } catch (e) {
    return false;
  }
}

/**
 * Inactivity-based auto sign-out — separate from (and in addition to) the
 * sessionStorage cache change above. sessionStorage already handles "the
 * browser was closed"; this handles the other real scenario: a browser
 * tab left open and unattended for a long time without being closed. After
 * SESSION_INACTIVITY_TIMEOUT_MS with no clicks/keystrokes/scrolling, the
 * person is signed out automatically, same as clicking Log Out themselves.
 */
const SESSION_INACTIVITY_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes
const LAST_ACTIVITY_KEY = "eatcoe_last_activity_ts";

function recordActivity() {
  try { sessionStorage.setItem(LAST_ACTIVITY_KEY, String(Date.now())); } catch (e) { /* ignore */ }
}

function getLastActivity() {
  try {
    const v = sessionStorage.getItem(LAST_ACTIVITY_KEY);
    return v ? parseInt(v, 10) : Date.now();
  } catch (e) {
    return Date.now();
  }
}

async function checkSessionTimeout() {
  if (!SSO_ENABLED || !msalInstance) return;
  const account = getActiveAccount();
  if (!account) return; // nothing to time out
  const idleMs = Date.now() - getLastActivity();
  if (idleMs >= SESSION_INACTIVITY_TIMEOUT_MS) {
    await signOut();
    window.location.href = "login.html";
  }
}

function initInactivityTracking() {
  recordActivity();
  ["click", "keydown", "mousemove", "scroll", "touchstart"].forEach((evt) => {
    let throttled = false;
    document.addEventListener(evt, () => {
      if (throttled) return;
      throttled = true;
      setTimeout(() => { throttled = false; }, 5000); // at most once every 5s per event type
      recordActivity();
    }, { passive: true });
  });
  setInterval(checkSessionTimeout, 60 * 1000); // check once a minute
}

document.addEventListener("DOMContentLoaded", async () => {
  if (isMsalPopup()) return;

  if (!SSO_ENABLED) {
    renderAuthUI();
    return;
  }
  try {
    await msalReady;
    await msalInstance.handleRedirectPromise();
  } catch (e) {
    console.error("MSAL redirect handling failed:", e);
  }
  const account = getActiveAccount();
  if (account) {
    onSignedIn(account);
    initInactivityTracking();
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
