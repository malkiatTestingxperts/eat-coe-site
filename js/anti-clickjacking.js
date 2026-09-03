/* ==========================================================================
   TREAT COE — Clickjacking detection/mitigation
   Extracted into its own external file specifically so it can run under a
   strict Content-Security-Policy without needing 'unsafe-inline' on
   script-src — an inline <script> block, however well-intentioned, is
   exactly the kind of execution CSP is designed to block by default.
   ========================================================================== */

// Modern browsers (Chrome/Firefox/Safari) now block automatic
// cross-origin frame-to-parent navigation as an anti-abuse measure,
// similar to a popup blocker -- confirmed by direct testing, so the
// classic "break out of the frame" trick alone is not reliable anymore.
// What still works without needing that permission: if this page detects
// it's framed, immediately blank its own content -- a clickjacking PoC
// specifically relies on the real page being visible (at reduced opacity)
// inside the attacker's frame for the illusion to work; with nothing
// rendered, there's nothing left to trick someone into clicking. The
// real, complete fix is still the CSP frame-ancestors / X-Frame-Options
// header — this is a partial, static-file-only mitigation, not a
// substitute for it.
if (window.top !== window.self) {
  document.documentElement.innerHTML =
    '<body style="font-family:sans-serif;padding:60px;text-align:center;color:#333">' +
    '<h2>This page cannot be displayed in a frame.</h2>' +
    '<p>For your security, please open it directly instead:</p>' +
    '<a href="' + window.self.location.href + '" target="_top">' + window.self.location.href + '</a>' +
    '</body>';
  try { window.top.location = window.self.location; } catch (e) { /* ignore -- likely blocked by the browser */ }
  // window.stop() halts the parser immediately -- without this, the
  // browser keeps parsing and running the rest of the original page's
  // scripts (auth.js, site.js) right after this point, which would
  // otherwise silently re-populate the real content moments later and
  // undo the blanking above entirely. Confirmed this exact failure mode
  // directly during testing.
  window.stop();
}
