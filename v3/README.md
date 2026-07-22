# Sales Dashboard

A single-page sales-reporting tool: two modes — **Monthly Report** (KPIs,
MoM, model-grade summaries, per-company Excel export) and **Last Selling
Price** (Last ASP, Sales-by-Model drill-down). All processing runs
client-side in the browser.

## Files

| File | Purpose |
|------|---------|
| `index.html` | Page structure + CDN library tags. Loads `style.css` and `app.min.js`. |
| `style.css` | Stylesheet (readable). |
| `app.js` | Application logic — **the readable source of truth**. |
| `app.min.js` | Minified build product. The page loads this at runtime. |
| `build.js` | Minify `app.js` → `app.min.js` using [terser](https://github.com/terser/terser). |
| `package.json` | npm scripts + terser dev dependency. |
| `employer.companyId253A722650.png` | Header logo. |
| `node_modules/` | terser (dev only — do not commit; see .gitignore). |

## Running locally

Just open `index.html` in a browser. First load needs an internet
connection to fetch SheetJS / Chart.js / JSZip from the CDN; they're
cached after that.

## Editing the code

1. Edit `app.js` (the readable source).
2. Rebuild the minified bundle the page loads:

   ```sh
   npm install
   npm run build
   ```

   This regenerates `app.min.js`. Commit both files.

## Why a minified bundle?

The page loads `app.min.js`, not `app.js`. This adds mild friction to
casual "View Source" copying — the live code is squashed and hard to
read. The readable `app.js` is committed alongside it so the project's
behavior is transparent and the license is clear.

Honest note: minification is **not** copy protection. Anyone determined
can deobfuscate or "Save Page As". For real protection you'd need
server-side logic — out of scope for this build.

## Dependencies (CDN)

- [SheetJS / xlsx](https://sheetjs.com/) 0.18.5 — Excel parse + write
- [Chart.js](https://www.chartjs.org/) 4.4.3 — charts
- [JSZip](https://stuk.github.io/jszip/) 3.10.1 — multi-file export

## License

All rights reserved. See the project owner for licensing terms.
