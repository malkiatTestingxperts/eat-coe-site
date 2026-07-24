# EAT COE — Enterprise Application Testing Center of Excellence

A fully static website (HTML/CSS/JS, no server, no build step, no signup).
All 29 deliverables plus a document catalog are embedded in `js/site.js`.

## What's new in this version

- **Dropdown navigation**: Home is a direct link; About / Pillars / Documents /
  Dashboard each open a dropdown of related pages.
- **A single "Log In" button** in the header — no manual Viewer/Contributor
  dropdown, no typed-in name. Signed out = read-only Viewer. Signed in with
  Microsoft = Contributor (upload, tag, edit story status). See "SSO login"
  below for exactly how signing in works and what it does/doesn't enforce.
- **Real upload and download, right now, no backend.** "Upload a document"
  (on `documents.html` or any pillar/story page, Contributor only) takes an
  actual file, stores it in your browser (up to 4MB), and the **Download**
  button on that document genuinely downloads those exact bytes back out —
  tested end-to-end. There's also always a **View in SharePoint** button that
  opens the real shared folder:
  ```
  https://txplin.sharepoint.com/:f:/s/DigitalBanking/IgCYGuPe89-RRaiIMVPoFOy-AVkT3yBTWBOcsZ4esISUN84
  ```
  The 8 pre-loaded documents don't have local bytes (they're real existing
  SharePoint files), so their Download button currently opens SharePoint too
  — only documents *uploaded through this site* download for real today.
  Once Graph API is connected (see below), `registerDocument()` and
  `downloadDocument()` in `js/site.js` are the only two functions that need
  to change — from browser storage to a real SharePoint upload/download call
  — the buttons and forms stay exactly as they are.
- **Progress/status now lives only on the Dashboard.** Pillar and story pages
  show story content, owner, acceptance criteria, and linked documents — no
  status chips. The Dashboard has a story-level table where Contributors can
  update status inline.
- **Dashboard metrics**: total documents in the catalog, total opens, active
  contributors, and a 7-day activity tracker + log.
- **Home page** is now a brief intro + Quick Links (Featured / Most Downloaded
  / Recently Modified documents) + the five pillar cards. The old long
  "What EAT is" writeup moved to its own **About** page.
- **Unified search** (hero search bar + chatbot) now searches both stories
  and documents, and results are labeled with a Story/Document badge.

## SSO login (Microsoft sign-in)

One **Log In** button in the header, using Microsoft sign-in (MSAL.js) — no
separate password, no new signup, just an existing Microsoft 365 account.
Signing in does two things: it shows the person's real name instead of
"Guest", and it grants Contributor access (upload/tag/status-edit). Signing
out (or never signing in) leaves someone as a read-only Viewer. There's no
separate role dropdown anymore — Log In *is* the role switch.

### Turning it on
1. Ask your Entra ID admin to register the site:
   - Azure Portal → **Entra ID → App registrations → New registration**
   - Single-tenant, Platform = **Single-page application (SPA)**
   - Redirect URI = the site's real URL (e.g. `https://you.github.io/eat-coe-site/`)
   - **API permissions** → Microsoft Graph → Delegated → `User.Read` → **Grant admin consent**
   - Copy the **Application (client) ID** and **Directory (tenant) ID**
2. Open `js/auth.js` and replace the two placeholders near the top:
   ```js
   clientId: "PASTE-YOUR-CLIENT-ID-HERE",
   authority: "https://login.microsoftonline.com/PASTE-YOUR-TENANT-ID-HERE",
   ```
3. That's it — reload the site. Until both placeholders are replaced, the
   Log In button still shows (so you can see the UI), but clicking it just
   explains what's still needed instead of doing anything — nothing is ever
   half-broken.

### Behavior once enabled
- Visitors see a "Log In" gate before the content is usable (blurred
  background + button). Set `REQUIRE_SIGNIN = false` near the top of
  `js/auth.js` if you'd rather make sign-in optional (a button in the
  header, no gate).
- Only people in your Microsoft tenant can log in at all — anyone else's
  Microsoft account is rejected by Entra ID itself, before your site sees
  anything.

### What this does *not* do yet
- Anyone in the tenant who logs in becomes a Contributor — it isn't yet
  scoped to a specific security group, and it isn't tied to real SharePoint
  permissions. For that, see the security-group + `Sites.Selected` approach
  discussed in the project history — it's a bigger change (real permission
  enforcement happens in SharePoint itself, not in the browser).
- It doesn't yet call Graph to upload into the real SharePoint folder —
  uploads are real, but browser-local, until that's wired in (see above).


## Important limitation (because there's no shared backend)

Everything a visitor adds or changes — role/login, uploaded documents, tags,
story status — is saved with `localStorage`, **in that one browser only**.
Two different people opening the same link will not see each other's
uploads or sign-in. This keeps the site 100% free, but it means it's best for:
- a single person's/browser's working session, or
- a demo of the workflow before wiring up a real shared backend.

### Going further (optional, still free)
If you want everyone to see the same catalog/status/roles:
- **Supabase** or **Firebase** (free tier) — a hosted database + real login,
  ~10–15 minutes to wire in, no server to run.
- **Microsoft Graph API** against your own SharePoint/Teams — since the
  documents already live there, this is the most natural fit; it needs an
  Azure AD app registration from your M365 admin.

Either way, only `js/site.js`'s storage functions (`getUserDocs`,
`getStatusOverrides`, `getRole`, etc.) need to change from `localStorage` to
API calls — the rest of the UI stays the same.

## Hosting for free

### GitHub Pages
1. Create a new **public** repo, upload every file/folder here (keep
   `assets/`, `js/`, `data/`, all `.html` files, and `.nojekyll`).
2. **Settings → Pages** → Source: `Deploy from a branch`, branch `main`,
   folder `/ (root)` → Save.
3. Live in ~1 minute at `https://<your-username>.github.io/<repo>/`.

### Netlify Drop (no GitHub needed)
Go to `app.netlify.com/drop`, drag the folder in, get an instant live URL.

Vercel and Cloudflare Pages work the same way (drag-and-drop or connect a
repo, no build command, output directory `/`).

## Editing content

- **Stories** (the 29 deliverables): edit the `STORIES` array at the top of
  `js/site.js`, or edit `data/documents.json` and re-run
  `python3 build2.py` (requires `pip install beautifulsoup4`) to regenerate
  all pages from the `src/` originals.
- **Seed documents** (the 8 real docs already in the catalog): edit
  `SEED_DOCS` in `js/site.js`, or edit `data/seed_documents.json` and
  re-embed it (see the data-loading block at the top of the script that
  generates `js/site.js`), then re-run `python3 build2.py`.
- **SharePoint folder link**: change `SHAREPOINT_FOLDER_URL` at the top of
  `js/site.js`.
