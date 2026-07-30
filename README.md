# TREAT COE — Tx's Repository For Enterprise Applications Testing

A static website (HTML/CSS/JS) with real Microsoft SSO and a real, live
Microsoft Graph integration to your SharePoint/Teams document library.

## How access works

- **login.html** is the only entry point for anyone not signed in. Visiting
  *any* other page while signed out redirects here automatically.
- Signing in with a Microsoft work account (via MSAL.js) sends you straight
  to **index.html** (Home) — which has no login button at all, since you can
  only ever reach it already authenticated.
- Everyone who successfully signs in is currently treated as a Contributor
  (can edit tags, update story status). This is **not yet tied to real
  SharePoint permissions** — see "Going further" below if you want that.
- Real credentials are already in place in `js/auth.js`:
  - Client ID: `9a817d03-ec3b-4e4f-8fa6-b7278cab47fe`
  - Tenant ID: `d7e861c9-d924-4413-86db-05780e928657`
  - If sign-in ever needs to be turned off site-wide, set
    `REQUIRE_SIGNIN = false` near the top of `js/auth.js`.

## The live document catalog (Microsoft Graph)

Once signed in, `js/graph.js` does the following automatically:
1. Resolves the shared SharePoint folder link (`SHAREPOINT_FOLDER_URL` in
   `js/site.js`) to a real drive/folder via Graph's `/shares` endpoint.
2. Lists real files under it (up to two folder levels deep, matching the
   `01-standards / 1.1 Story Name / file.docx` convention).
3. Replaces the built-in seed document list with this live listing
   everywhere — Documents page, pillar/story pages, Quick Links, search,
   the chatbot, and Dashboard metrics.
4. In the background, downloads and text-indexes every PDF/.docx/.txt/.md
   file it found (capped at 25 files per session — see
   `MAX_AUTO_INDEX_FILES` in `js/graph.js`) using PDF.js/Mammoth.js, so the
   main search bar and chatbot can match **text inside the documents**, not
   just their file names.
5. If any of this fails for any reason (Graph scopes not consented yet,
   offline, folder not resolvable) it fails silently and the site falls back
   to the built-in seed document list — nothing breaks either way.

This is cached in `sessionStorage` for 5 minutes (`GRAPH_CACHE_TTL_MS`) to
avoid re-listing the whole folder on every page load.

**Buttons now do real things:**
- **⬇ Download** on a real SharePoint document fetches and downloads the
  actual file bytes via Graph.
- **↗ View in SharePoint** opens the file's real SharePoint URL.
- A Graph **upload** function (`uploadFileToSharePoint`) is implemented and
  ready in `js/graph.js`, but nothing calls it yet — see below.

### Required Graph permissions
`js/auth.js` requests `User.Read` at sign-in (always works). `js/graph.js`
separately requests `Sites.ReadWrite.All` + `Files.ReadWrite.All` the first
time it's actually needed (incremental consent) — so a Graph permissions
issue never blocks basic sign-in, only the live document features.

## Register a Document (email-based submission)

This page is intentionally **not** a Graph upload — it composes an email
instead:
1. Your name/email auto-fill from your signed-in Microsoft account.
2. Fill in the document details, click **📎 Choose file** to attach a file
   from your computer, then **📧 Send**.
3. Send opens your email client with the subject/body fully pre-filled,
   including a suggested SharePoint folder computed from the pillar/story
   you picked (e.g. `docs/03-monitoring/3.6 Defect Trend Analysis/`).
4. **You still have to manually attach the file in your email app before
   hitting send there.** Browsers cannot attach files to an email
   automatically — this is a hard security restriction, not a limitation of
   this build. The email body includes an explicit reminder naming the file.
5. Submissions go to **`Treat@testingxperts.com`** (`COE_INTAKE_EMAIL` in
   `js/site.js`) — change it there if needed.
6. While a file is attached (before you send), it's also indexed
   client-side and shows up in search results tagged **📋 pending
   submission** — purely a same-browser-session convenience so you can
   confirm the content matches before sending.

### Going further: a real upload button instead of email
`uploadFileToSharePoint(file, pillarCode)` in `js/graph.js` already does a
real Graph upload (simple upload, good up to 4MB — larger files need Graph's
chunked upload-session API instead). To wire it to the Send button instead
of (or alongside) the email flow, call it from the `rfSendBtn` click handler
in `js/site.js`'s `initRegisterForm()`.

## Roles and real permission enforcement (optional next step)

Right now, signing in = Contributor, for everyone in the tenant. If you want
this tied to real SharePoint permissions:
1. Create Entra ID security groups (e.g. `TREAT-Viewers`, `TREAT-Contributors`)
   and grant them Read / Edit on the SharePoint folder respectively — this
   is where *real* enforcement happens, in SharePoint itself.
2. In Entra ID → Enterprise Applications → this app → set "Assignment
   required = Yes" and assign those groups, so only people in them can even
   sign in.
3. Add a "groups" claim to the ID token (App registration → Token
   configuration) so `js/auth.js` can read the signed-in user's group
   membership and only grant Contributor UI to people in the Contributors
   group — this is UI-only, but ties the initial assumption of the role to
   something more scoped than "your entire tenant."

## Other structural notes

- **Progress/status** lives only on the Dashboard (`dashboard.html`) — a
  story-level table where Contributors can update status inline. Pillar and
  story pages show content/documents only, no status chips.
- **Quick Links** (Home page) show each document with every badge it earns
  (Featured / Most Downloaded / Recently Modified) rather than listing the
  same document three times.
- **Full-text search** works the same way everywhere it appears — main hero
  search bar, the Documents page search, and the chatbot — all powered by
  the same `searchAll()` function in `js/site.js`.
- `documents122.html` in this repo is an old unreferenced backup file, not
  linked from anywhere — safe to delete whenever convenient.

## Hosting

Still a fully static site — GitHub Pages, Netlify, Vercel, or any static
host works with zero configuration. Just make sure the redirect URI
configured on the Entra ID app registration matches wherever it's actually
hosted (currently set for local testing — update `js/auth.js`'s
`redirectUri` is computed automatically from `window.location`, so this
generally does not need manual changes, but double check the Entra ID app
registration's allowed redirect URIs include your real hosted URL).
