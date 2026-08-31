/* Same-origin loader for the vendored PDF.js (patched, CVE-2024-4367 fixed,
   version 4.10.38). PDF.js 4.x ships as an ES module only — no classic
   UMD/global build exists anymore — so this small shim imports it properly
   and exposes it the same way the rest of the codebase already expects
   (a global `pdfjsLib`), via an explicit readiness signal rather than
   assuming script-execution order, since module scripts are deferred and
   would otherwise race against the code that uses them. */
import * as pdfjsLib from "./vendor/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "js/vendor/pdf.worker.min.mjs";
window.pdfjsLib = pdfjsLib;
window.__pdfjsReadyPromise = Promise.resolve(pdfjsLib);
document.dispatchEvent(new CustomEvent("pdfjs-ready"));
