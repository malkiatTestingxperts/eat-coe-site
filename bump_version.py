#!/usr/bin/env python3
"""
Bumps the cache-busting version query string (?v=...) on every page's
<script src="js/site.js">, auth.js, and graph.js tags, all in one go.

WHY THIS EXISTS:
Browsers (and GitHub Pages' CDN) can keep serving someone a cached copy of
site.js/auth.js/graph.js from before your latest change, unless the *URL*
itself changes. Appending ?v=<version> makes that happen — but only if the
version string actually changes every time you push a real code update.
Forgetting to bump it (which is exactly what happened before) means people
keep silently running old code indefinitely.

USAGE:
Run this from the root of the site (same folder as index.html) any time
you change js/site.js, js/auth.js, or js/graph.js, right before you commit
and push:

    python3 bump_version.py

Then commit and push as normal. Every page's script tags will point to a
new, never-before-seen URL, so every browser is forced to fetch the real,
current files instead of reusing anything cached.
"""

import re
import sys
from datetime import datetime, timezone
from pathlib import Path

PAGES = [
    "index.html", "about.html", "documents.html", "register-document.html",
    "dashboard.html", "sitemap.html", "login.html",
    "01-standards.html", "02-tools.html", "03-monitoring.html",
    "04-capability.html", "05-innovation.html",
]

def main():
    root = Path(__file__).parent
    new_version = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")

    pattern = re.compile(r'(src="js/(?:site|auth|graph)\.js)(?:\?v=[^"]*)?(")')

    changed = 0
    missing = []
    for name in PAGES:
        path = root / name
        if not path.exists():
            missing.append(name)
            continue
        html = path.read_text(encoding="utf-8")
        new_html, n = pattern.subn(rf'\1?v={new_version}\2', html)
        if n:
            path.write_text(new_html, encoding="utf-8")
            changed += 1

    print(f"Bumped {changed} page(s) to ?v={new_version}")
    if missing:
        print("Note: these expected pages were not found (skipped):", ", ".join(missing))
    print("\nNext step: commit and push these changes as normal.")

if __name__ == "__main__":
    sys.exit(main())
