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
// Sites.Read.All — this is the permission your admin actually granted (a
// tenant-wide, read-only grant covering every SharePoint site, not just
// this one). Requesting this exact scope name is what matters: MSAL only
// uses a granted permission if the app asks for that specific scope, so
// this has to match whatever your admin consented to, or the token won't
// include it even though it's been granted in Entra ID.
const GRAPH_SCOPES = ["User.Read", "Sites.Read.All"];
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
/**
 * Parses a SharePoint URL into { hostname, sitePath, subPath, isSharingLink }.
 * Handles two shapes:
 *   - A sharing link:  https://tenant.sharepoint.com/:f:/s/SiteName/AbCdEf...
 *   - A normal site/library URL: https://tenant.sharepoint.com/sites/SiteName/Shared Documents
 * `subPath` (e.g. "Shared Documents") is only meaningful for the second
 * shape — sharing links carry an opaque token instead, resolved via /shares.
 */
function parseSharePointSitePath(url) {
  const sharingMatch = /^https:\/\/([^/]+)\/:[a-z]:\/[a-z]\/([^/]+)\//i.exec(url);
  if (sharingMatch) {
    return { hostname: sharingMatch[1], sitePath: "/sites/" + sharingMatch[2], subPath: null, isSharingLink: true };
  }
  const siteMatch = /^https:\/\/([^/]+)\/sites\/([^/]+)\/?(.*)$/i.exec(url);
  if (siteMatch) {
    const subPath = decodeURIComponent(siteMatch[3] || "").replace(/\/+$/, "");
    return { hostname: siteMatch[1], sitePath: "/sites/" + siteMatch[2], subPath: subPath || null, isSharingLink: false };
  }
  return null;
}

/**
 * Resolves a SharePoint site (and, if the URL includes one, a specific
 * library/folder subpath within it) directly via Graph — the reliable path
 * for a normal, navigable SharePoint URL like
 * ".../sites/SiteName/Shared Documents" (as opposed to an opaque sharing
 * link, which needs /shares instead — see resolveSharePointFolder below).
 */
async function resolveSharePointFolderViaSite() {
  const parsed = parseSharePointSitePath(SHAREPOINT_FOLDER_URL);
  if (!parsed) throw new Error("Could not parse a SharePoint site path from SHAREPOINT_FOLDER_URL.");
  const siteRes = await graphFetch(`/sites/${parsed.hostname}:${parsed.sitePath}`);
  const site = await siteRes.json();

  const rootPath = parsed.subPath
    ? `/sites/${site.id}/drive/root:/${parsed.subPath.split("/").map(encodeURIComponent).join("/")}`
    : `/sites/${site.id}/drive/root`;
  const rootRes = await graphFetch(rootPath);
  const rootItem = await rootRes.json();
  return {
    driveId: rootItem.parentReference && rootItem.parentReference.driveId,
    itemId: rootItem.id,
    webUrl: rootItem.webUrl,
    name: rootItem.name
  };
}

/**
 * Resolves SHAREPOINT_FOLDER_URL to a real folder (driveId + itemId).
 * - Opaque sharing links (":f:/s/...") are resolved via Graph's /shares
 *   endpoint, since that's what it's actually for.
 * - Normal, navigable site/library URLs (e.g. ".../sites/SiteName/Shared
 *   Documents") skip /shares entirely and go straight to the more reliable
 *   site-based resolution — /shares is built for opaque share tokens, not
 *   plain URLs, so trying it first for a normal URL would just fail.
 */
async function resolveSharePointFolder() {
  const parsed = parseSharePointSitePath(SHAREPOINT_FOLDER_URL);

  if (parsed && !parsed.isSharingLink) {
    return await resolveSharePointFolderViaSite();
  }

  try {
    const encoded = encodeSharingUrl(SHAREPOINT_FOLDER_URL);
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
        "\nCheck: does this account have real membership/access on this " +
        "SharePoint site itself (not just the sharing link)? An admin may need " +
        "to add them as a site member or grant access directly.",
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
const MAX_FOLDER_DEPTH = 8; // safety cap against runaway recursion, not a real limit in practice

/**
 * Recursively walks every folder under the resolved root, at any depth —
 * not just a fixed 2 levels. This matters because the real folder
 * structure doesn't always match the "pillar folder / story folder / file"
 * convention exactly; there can be extra wrapper folders in between (e.g.
 * "Shared Documents" → "EAT" → "01-standards" → "1.1 Story Name" → file)
 * that don't themselves match the pillar/story naming pattern. Whenever a
 * folder name *does* match (e.g. "01-standards" or "1.1 Some Story"), that
 * becomes the pillar/story context for every file found underneath it,
 * however many extra wrapper folders are in between — folders that don't
 * match just pass the current context through unchanged.
 */
async function listAllSharePointFiles(driveId, rootItemId) {
  const results = [];

  async function walk(itemId, pillarCode, storyName, depth) {
    if (depth > MAX_FOLDER_DEPTH) return;
    const children = await listSharePointChildren(driveId, itemId);
    for (const child of children) {
      if (!child.folder) {
        results.push(mapDriveItemToDoc(child, driveId, pillarCode, storyName));
        continue;
      }
      const inferredPillar = inferPillarCodeFromFolderName(child.name);
      const inferredStory = /^\d+\.\d+/.test(child.name) ? child.name : null;
      await walk(
        child.id,
        inferredPillar || pillarCode,
        inferredStory || storyName,
        depth + 1
      );
    }
  }

  await walk(rootItemId, null, null, 0);
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
 * NOTE: GRAPH_SCOPES is Sites.Read.All right now — read-only — so this
 * function will fail with a permissions error if actually called. It needs
 * "Sites.ReadWrite.All" instead (a further, separate admin consent step)
 * before being wired to any button. Also not currently wired to anything
 * (Register a Document intentionally emails instead, per current
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
