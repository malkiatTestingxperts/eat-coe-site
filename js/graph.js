/* ==========================================================================
   TREAT COE — Microsoft Graph integration
   Once someone is signed in (see js/auth.js), this module fetches the real
   file listing from the shared SharePoint folder, powers real downloads,
   and background-indexes PDF/.docx files so the main search bar and the
   chatbot can match text found *inside* real SharePoint documents — not
   just their names/tags. If Graph isn't reachable for any reason (scopes
   not consented yet, offline, folder not resolvable), everything falls
   back to the static built-in document list — nothing breaks.
   ========================================================================== */

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
// Files.ReadWrite (no ".All") is a low-privilege delegated scope scoped to
// files the signed-in user can already access — per Microsoft's own
// permissions reference this does NOT require admin consent, unlike
// Sites.ReadWrite.All / Files.ReadWrite.All (tenant-wide, admin consent
// required), which this used to request. It's ReadWrite rather than plain
// Files.Read specifically because Microsoft's own documentation for the
// /shares endpoint (used below to resolve the shared folder link) lists
// Files.ReadWrite as its minimum required delegated permission — Files.Read
// alone was not enough to resolve the share, which is why downloads/views
// kept falling back to the generic SharePoint folder link even after the
// previous fix.
const GRAPH_SCOPES = ["User.Read", "Files.ReadWrite"];
const GRAPH_CACHE_KEY = "eatcoe_graph_folder_cache";
const GRAPH_CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes — avoid hammering Graph on every page load
const MAX_AUTO_INDEX_FILES = 25; // cap background full-text extraction per session

/* ---------------- sharing-link → driveItem (per Microsoft's documented encoding) ---------------- */
function encodeSharingUrl(url) {
  const bytes = new TextEncoder().encode(url);
  let binary = "";
  bytes.forEach(b => { binary += String.fromCharCode(b); });
  const base64 = btoa(binary);
  return "u!" + base64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}

/* ---------------- token + fetch helpers ---------------- */
// Keyed to the specific scope set, not just a flat flag — otherwise, if the
// requested scopes ever change (like just now, dropping the admin-gated
// .All scopes for plain Files.Read), a stale "already tried and failed"
// flag from the OLD scope set would permanently block the NEW one from
// ever being attempted, even though it's a totally different permission.
const GRAPH_CONSENT_ATTEMPTED_KEY =
  "eatcoe_graph_consent_attempted:" + GRAPH_SCOPES.slice().sort().join(",");

/**
 * Token acquisition for Graph calls. Always tries silently first (if a
 * grant already exists, this never shows any UI, on any page, ever).
 *
 * The *very first time* silent acquisition fails (meaning this scope has
 * never been consented on this device at all), it allows exactly ONE
 * interactive popup attempt to complete that initial consent — this is
 * normal and expected for a low-privilege scope like Files.Read, which
 * Microsoft's own docs list as not requiring admin approval, so that one
 * popup should just show a normal "Accept" button, not an approval-request
 * dead end. Either way — success, cancellation, or an unexpected admin
 * approval requirement — that one attempt is remembered permanently
 * (localStorage, not just this tab), so it's never retried automatically
 * again. If something changes (an admin grants access later, permissions
 * are updated, etc.), clear `eatcoe_graph_consent_attempted` from
 * localStorage (or a different browser/profile) to let it try again.
 */
async function getGraphToken() {
  if (typeof msalInstance === "undefined" || !msalInstance) throw new Error("MSAL is not initialized (SSO not configured).");
  if (typeof msalReady !== "undefined") await msalReady;
  const account = (typeof getActiveAccount === "function") ? getActiveAccount() : null;
  if (!account) throw new Error("Not signed in.");
  const request = { scopes: GRAPH_SCOPES, account };

  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (silentError) {
    let alreadyAttempted = false;
    try { alreadyAttempted = localStorage.getItem(GRAPH_CONSENT_ATTEMPTED_KEY) === "1"; } catch (e) { /* ignore */ }

    if (alreadyAttempted) {
      console.warn(
        "Silent Graph token acquisition failed, and the one-time interactive " +
        "consent attempt already happened previously on this device — not " +
        "asking again automatically. Falling back to the built-in document " +
        "list. Clear localStorage's \"" + GRAPH_CONSENT_ATTEMPTED_KEY + "\" " +
        "key to allow one more attempt.", silentError
      );
      throw silentError;
    }

    try { localStorage.setItem(GRAPH_CONSENT_ATTEMPTED_KEY, "1"); } catch (e) { /* ignore */ }
    try {
      const result = await msalInstance.acquireTokenPopup(request);
      return result.accessToken;
    } catch (popupError) {
      console.warn(
        "The one-time interactive Graph consent attempt failed or was " +
        "cancelled — not asking again automatically going forward. Falling " +
        "back to the built-in document list.", popupError
      );
      throw popupError;
    }
  }
}

async function graphFetch(pathOrUrl, options = {}) {
  const token = await getGraphToken();
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH_BASE + pathOrUrl;
  const res = await fetch(url, {
    ...options,
    headers: {
      "Authorization": "Bearer " + token,
      ...(options.headers || {})
    }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph request failed (${res.status} ${res.statusText}): ${text.slice(0, 300)}`);
  }
  return res;
}

/* ---------------- resolving + listing the shared folder ---------------- */
/**
 * Parses a SharePoint sharing URL like
 *   https://tenant.sharepoint.com/:f:/s/SiteName/AbCdEf...
 * or a normal site URL like
 *   https://tenant.sharepoint.com/sites/SiteName/...
 * into { hostname, sitePath } for use with Graph's /sites/{hostname}:{sitePath}
 * endpoint. Returns null if the URL doesn't match either shape.
 */
function parseSharePointSitePath(url) {
  const sharingMatch = /^https:\/\/([^/]+)\/:[a-z]:\/[a-z]\/([^/]+)\//i.exec(url);
  if (sharingMatch) return { hostname: sharingMatch[1], sitePath: "/sites/" + sharingMatch[2] };
  const siteMatch = /^https:\/\/([^/]+)\/sites\/([^/]+)/i.exec(url);
  if (siteMatch) return { hostname: siteMatch[1], sitePath: "/sites/" + siteMatch[2] };
  return null;
}

/**
 * Fallback used when resolving via the sharing link itself (/shares) fails.
 * This can genuinely happen even when the exact same link opens fine in a
 * browser for the same person — some tenants apply stricter rules to
 * API/Graph-based access to a sharing link than to normal interactive
 * access. Resolving the SharePoint SITE directly and using its default
 * document library is a more standard, usually more reliable Graph call
 * that lines up with how interactive site access works.
 *
 * Caveat: this lands on the site's document library ROOT, which may be
 * broader (or in rare cases narrower) than whatever specific subfolder the
 * original sharing link pointed to — it's the best available fallback, not
 * a guaranteed exact match for the original link's target folder.
 */
async function resolveSharePointFolderViaSite() {
  const parsed = parseSharePointSitePath(SHAREPOINT_FOLDER_URL);
  if (!parsed) throw new Error("Could not parse a SharePoint site path from SHAREPOINT_FOLDER_URL to use as a fallback.");
  const siteRes = await graphFetch(`/sites/${parsed.hostname}:${parsed.sitePath}`);
  const site = await siteRes.json();
  const rootRes = await graphFetch(`/sites/${site.id}/drive/root`);
  const rootItem = await rootRes.json();
  return {
    driveId: rootItem.parentReference && rootItem.parentReference.driveId,
    itemId: rootItem.id,
    webUrl: rootItem.webUrl,
    name: rootItem.name
  };
}

async function resolveSharePointFolder() {
  const encoded = encodeSharingUrl(SHAREPOINT_FOLDER_URL);
  try {
    const res = await graphFetch(`/shares/${encoded}/driveItem`);
    const item = await res.json();
    return {
      driveId: item.parentReference && item.parentReference.driveId,
      itemId: item.id,
      webUrl: item.webUrl,
      name: item.name
    };
  } catch (shareError) {
    console.warn(
      "Resolving via the sharing link (/shares) failed — trying the SharePoint " +
      "site directly instead as a fallback. This mismatch (link works in a " +
      "browser, fails via Graph) is a known pattern in some tenants, not a bug " +
      "in this code.", shareError
    );
    try {
      return await resolveSharePointFolderViaSite();
    } catch (siteError) {
      console.error(
        "Both resolution paths failed — the sharing-link lookup (/shares) and " +
        "the direct site lookup. This points to a real access/permissions issue " +
        "for the signed-in user on the SharePoint side, not something a further " +
        "code change here can fix. The exact link involved is:\n  " + SHAREPOINT_FOLDER_URL +
        "\nCheck: does this account have real membership/access on the " +
        "\"DigitalBanking\" SharePoint site itself (not just the sharing link)? " +
        "An admin may need to add them as a site member or grant access directly.",
        siteError
      );
      throw siteError;
    }
  }
}

async function listSharePointChildren(driveId, itemId) {
  const res = await graphFetch(`/drives/${driveId}/items/${itemId}/children?$top=200`);
  const data = await res.json();
  return data.value || [];
}

function inferPillarCodeFromFolderName(name) {
  const m = /^0[1-5]/.exec(name || "");
  return m ? m[0] : null;
}

function mapDriveItemToDoc(item, driveId, pillarCode, storyFolderName) {
  let storyCode = null;
  if (storyFolderName) {
    const m = /^\d+\.\d+/.exec(storyFolderName);
    if (m) storyCode = m[0];
  }
  return {
    id: "graph-" + item.id,
    type: "document", sourceType: "graph",
    name: item.name,
    url: item.webUrl,
    driveId, itemId: item.id,
    downloadUrl: item["@microsoft.graph.downloadUrl"] || null,
    pillarCode: pillarCode || null,
    pillar: (typeof PILLAR_NAMES !== "undefined" && PILLAR_NAMES[pillarCode]) || null,
    storyCode: storyCode,
    tags: [], featured: false, downloads: 0,
    uploadedBy: (item.createdBy && item.createdBy.user && item.createdBy.user.displayName) || "Unknown",
    uploadDate: (item.createdDateTime || "").slice(0, 10),
    lastModifiedBy: (item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName) || "Unknown",
    lastModifiedDate: (item.lastModifiedDateTime || "").slice(0, 10),
    location: "SharePoint (live via Microsoft Graph)",
    fullText: null, fullTextStatus: null,
    size: item.size || 0,
    mimeType: (item.file && item.file.mimeType) || null
  };
}

/**
 * Lists real files under the shared folder, up to two levels deep
 * (pillar folder → story folder → files), matching the docs/01-standards/
 * 1.1 .../ convention already used for the seed documents.
 */
async function listAllSharePointFiles(driveId, rootItemId) {
  const results = [];
  const topLevel = await listSharePointChildren(driveId, rootItemId);

  for (const entry of topLevel) {
    if (entry.folder) {
      const pillarCode = inferPillarCodeFromFolderName(entry.name);
      const children = await listSharePointChildren(driveId, entry.id);
      for (const child of children) {
        if (!child.folder) {
          results.push(mapDriveItemToDoc(child, driveId, pillarCode));
        } else {
          const grandchildren = await listSharePointChildren(driveId, child.id);
          grandchildren.filter(g => !g.folder).forEach(g => {
            results.push(mapDriveItemToDoc(g, driveId, pillarCode, child.name));
          });
        }
      }
    } else {
      results.push(mapDriveItemToDoc(entry, driveId, null));
    }
  }
  return results;
}

/* ---------------- session cache (avoid re-listing on every page load) ---------------- */
function getGraphCacheEntry() {
  try {
    const raw = sessionStorage.getItem(GRAPH_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (Date.now() - parsed.ts > GRAPH_CACHE_TTL_MS) return null;
    return parsed;
  } catch (e) { return null; }
}
function setGraphCacheEntry(folder, docs) {
  try {
    sessionStorage.setItem(GRAPH_CACHE_KEY, JSON.stringify({ ts: Date.now(), folder, docs }));
  } catch (e) { /* storage full or unavailable — just skip caching */ }
}

async function loadGraphCatalog({ forceRefresh = false } = {}) {
  if (!forceRefresh) {
    const cached = getGraphCacheEntry();
    if (cached) {
      setGraphDocuments(cached.docs);
      indexGraphDocumentsInBackground(cached.docs);
      return cached.docs;
    }
  }
  const folder = await resolveSharePointFolder();
  const docs = await listAllSharePointFiles(folder.driveId, folder.itemId);
  setGraphCacheEntry(folder, docs);
  setGraphDocuments(docs);
  indexGraphDocumentsInBackground(docs);
  return docs;
}

/**
 * Given a placeholder document (seed catalog entry, no real driveId/itemId
 * yet), tries to find the matching real file in the live SharePoint
 * listing by name, loading that listing first if it isn't already cached.
 * Returns the real Graph-sourced document object (with a working
 * downloadUrl/driveId/itemId) if found, otherwise null — callers should
 * fall back to the old generic-folder-link behavior in that case.
 */
async function findRealDocumentByName(name) {
  let docs = (typeof GRAPH_DOCS !== "undefined" && GRAPH_DOCS && GRAPH_DOCS.length) ? GRAPH_DOCS : null;
  if (!docs) {
    docs = await loadGraphCatalog();
  }
  const target = (name || "").trim().toLowerCase();
  return docs.find(d => (d.name || "").trim().toLowerCase() === target) || null;
}

/* ---------------- real download ---------------- */
async function downloadGraphItem(doc) {
  let url = doc.downloadUrl;
  if (!url) {
    const res = await graphFetch(`/drives/${doc.driveId}/items/${doc.itemId}`);
    const item = await res.json();
    url = item["@microsoft.graph.downloadUrl"];
  }
  if (!url) throw new Error("No download URL available for " + doc.name);
  const res = await fetch(url); // pre-authenticated URL — no bearer token needed
  if (!res.ok) throw new Error("Download failed: " + res.status);
  const blob = await res.blob();
  const objectUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = objectUrl;
  a.download = doc.name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 4000);
}

/**
 * Navigates an already-open window straight to the document's real,
 * pre-authenticated content URL — used as a fallback when we don't already
 * have doc.downloadUrl cached (see openDocument() in site.js, which handles
 * the common case of an already-cached URL by opening it directly, with no
 * async step at all).
 *
 * This deliberately does NOT fetch the bytes itself and build a blob: URL.
 * That was tried first, but browsers (Chromium at least) silently refuse to
 * navigate another window via `.location.href` once the user-gesture that
 * opened it has been "spent" by an intervening cross-origin fetch() — the
 * navigation call doesn't error, it just does nothing. Handing the browser
 * the real URL directly and letting its own navigation fetch it sidesteps
 * that restriction entirely (this is exactly what a normal link click
 * does), at the cost of not being able to guarantee inline rendering vs.
 * download the way a blob: URL would — that now depends on whatever
 * Content-Disposition SharePoint's own download URL returns for the file
 * type, same as clicking a raw file link anywhere else.
 *
 * IMPORTANT: `targetWindow` must already be open (e.g. `window.open("",
 * "_blank")`, called synchronously by the caller, before any `await`).
 */
async function viewGraphItem(doc, targetWindow) {
  if (!targetWindow) throw new Error("No target window provided — it must be opened synchronously by the caller before calling this.");
  const res = await graphFetch(`/drives/${doc.driveId}/items/${doc.itemId}`);
  const item = await res.json();
  const url = item["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("No content URL available for " + doc.name);
  targetWindow.location.href = url;
}

/* ---------------- real upload ----------------
 * Simple upload (fine up to 4MB — larger files need Graph's chunked
 * upload-session API instead). Uploads to the pillar subfolder if one
 * already exists with a matching name, otherwise to the shared folder root.
 * Not currently wired to a button (Register a Document intentionally sends
 * an email per the current requirements) — available for whenever real
 * in-browser upload is wanted instead of/alongside the email flow.
 *
 * NOTE: GRAPH_SCOPES already includes Files.ReadWrite, so the token this
 * function gets does have write access — it's just not wired to any button
 * yet (Register a Document intentionally emails instead, per current
 * requirements).
 */
async function uploadFileToSharePoint(file, pillarCode) {
  const folder = await resolveSharePointFolder();
  let targetItemId = folder.itemId;

  if (pillarCode && typeof PILLAR_FOLDERS !== "undefined" && PILLAR_FOLDERS[pillarCode]) {
    const wantedName = PILLAR_FOLDERS[pillarCode].split("/").pop();
    const children = await listSharePointChildren(folder.driveId, folder.itemId);
    const match = children.find(c => c.folder && c.name.toLowerCase() === wantedName.toLowerCase());
    if (match) targetItemId = match.id;
  }

  const bytes = await file.arrayBuffer();
  const res = await graphFetch(
    `/drives/${folder.driveId}/items/${targetItemId}:/${encodeURIComponent(file.name)}:/content`,
    { method: "PUT", headers: { "Content-Type": "application/octet-stream" }, body: bytes }
  );
  return res.json();
}

/* ---------------- background full-text indexing ---------------- */
async function indexGraphDocumentsInBackground(docs) {
  const candidates = docs.filter(d => /\.(pdf|docx|txt|md)$/i.test(d.name)).slice(0, MAX_AUTO_INDEX_FILES);
  for (const doc of candidates) {
    try {
      let url = doc.downloadUrl;
      if (!url) {
        const res = await graphFetch(`/drives/${doc.driveId}/items/${doc.itemId}`);
        const item = await res.json();
        url = item["@microsoft.graph.downloadUrl"];
      }
      if (!url) continue;
      const res = await fetch(url);
      if (!res.ok) continue;
      const blob = await res.blob();
      const asFile = new File([blob], doc.name, { type: doc.mimeType || blob.type });
      const { text, status } = await extractFullText(asFile);
      doc.fullText = text;
      doc.fullTextStatus = status;
    } catch (e) {
      console.warn("Background full-text indexing failed for", doc.name, e);
    }
    // Re-render progressively so results improve as indexing completes,
    // instead of waiting for every file before showing anything.
    if (typeof refreshDocViews === "function") refreshDocViews();
  }
}

/**
 * Entry point — called once someone is signed in (see auth.js). Fails
 * silently and leaves the static seed catalog in place if Graph scopes
 * weren't granted yet, the folder can't be resolved, or the network is
 * unavailable — the rest of the site keeps working either way.
 */
async function initGraphCatalog() {
  try {
    await loadGraphCatalog();
  } catch (e) {
    console.warn("Could not load the live SharePoint catalog via Microsoft Graph — using the built-in document list instead.", e);
  }
}

/**
 * Diagnostic helper, callable from the browser console:
 *   retryGraphConnection()
 * Clears the cached "already tried" flag and the folder/file listing
 * cache, then immediately retries connecting to SharePoint and reports
 * exactly what happened — useful for checking the real cause without
 * waiting for a full page reload or digging through console warnings.
 */
async function retryGraphConnection() {
  try { localStorage.removeItem(GRAPH_CONSENT_ATTEMPTED_KEY); } catch (e) { /* ignore */ }
  try { sessionStorage.removeItem(GRAPH_CACHE_KEY); } catch (e) { /* ignore */ }
  console.log("Retrying Microsoft Graph connection with a clean slate...");
  try {
    const docs = await loadGraphCatalog({ forceRefresh: true });
    console.log("✅ Success — loaded", docs.length, "real document(s) from SharePoint:", docs);
    return docs;
  } catch (e) {
    console.error("❌ Still failed. Full error below — this is the real cause:", e);
    throw e;
  }
}
