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
// Files.Read (no ".All") is a low-privilege delegated scope scoped to files
// the signed-in user can already access — per Microsoft's own permissions
// reference this does NOT require admin consent, unlike Sites.ReadWrite.All /
// Files.ReadWrite.All (tenant-wide, "Yes" for admin consent required), which
// this used to request and which triggered a blocking "Approval required"
// screen on every attempt in tenants that require admin approval for those.
// Everything this file currently does — resolving the shared folder link,
// listing its contents, and downloading files — is read-only, so this is
// enough. Real upload (see uploadFileToSharePoint below) needs write access
// and isn't wired to any button yet; add "Files.ReadWrite" (also no admin
// consent required per Microsoft's docs) to this list if that's ever wired up.
const GRAPH_SCOPES = ["User.Read", "Files.Read"];
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
const GRAPH_CONSENT_ATTEMPTED_KEY = "eatcoe_graph_consent_attempted"; // localStorage: forever, across tabs/sessions

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
async function resolveSharePointFolder() {
  const encoded = encodeSharingUrl(SHAREPOINT_FOLDER_URL);
  const res = await graphFetch(`/shares/${encoded}/driveItem`);
  const item = await res.json();
  return {
    driveId: item.parentReference && item.parentReference.driveId,
    itemId: item.id,
    webUrl: item.webUrl,
    name: item.name
  };
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

/* ---------------- real upload ----------------
 * Simple upload (fine up to 4MB — larger files need Graph's chunked
 * upload-session API instead). Uploads to the pillar subfolder if one
 * already exists with a matching name, otherwise to the shared folder root.
 * Not currently wired to a button (Register a Document intentionally sends
 * an email per the current requirements) — available for whenever real
 * in-browser upload is wanted instead of/alongside the email flow.
 *
 * NOTE: this needs write access, which GRAPH_SCOPES no longer requests (it's
 * read-only now, specifically to avoid the admin-consent prompt — see the
 * comment on GRAPH_SCOPES above). Add "Files.ReadWrite" to GRAPH_SCOPES
 * before wiring this up to anything — that scope is still "No" for admin
 * consent required per Microsoft's docs, same as Files.Read.
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
