Margin Sheet Tool — v2 (UI redesign)
=====================================

Redesigned presentation layer for the Margin Sheet Tool. Business logic is
unchanged — aggregator.js, generator.js, and app.js are copied verbatim from
the v1 folder. All visual/UX improvements live in the new index.html, the
new CSS files, and one new additive module (ui.js).

What changed vs v1
------------------
- Polished enterprise visual design: neutral palette, one restrained blue
  accent, semantic success/warning/danger colors, light + dark themes.
- Step 1 / Step 2 cards with live status pills (Not ready / Ready / Building /
  Complete / Error).
- Styled file-upload cards replace native file inputs (the native input is
  hidden but still receives the change event, so all logic is preserved).
- Compact "Margin Summary Ready" result card with stat tiles (rows, regions,
  sheets, alias expansions) replaces the giant auto-rendered table.
- The full summary table is hidden by default inside a collapsible accordion
  ("Preview Margin Summary"). It remains in the DOM so the SheetJS export
  still reads it live and exports the complete dataset.
- Source-files-changed clearly invalidates the previous summary (card flips
  to a "rebuild required" warning, Export disabled).
- Summary Source in Step 2 is a segmented radio showing which source is active.
- Optional theme toggle (top right). Light is default; respects OS preference
  and the toggle choice is remembered in localStorage.
- Accessible: real labels, visible focus rings, aria-live status regions,
  keyboard-operable accordion. Respects prefers-reduced-motion.

Folder structure
----------------
margin_sheet_v2/
  index.html
  README.txt
  assets/
    css/
      variables.css     design tokens (palette, type, spacing, motion, dark theme)
      components.css    buttons, fields, cards, tables, status blocks, accordion, etc.
      app.css           page layout, grids, responsive, imports the two above
    js/
      aggregator.js     (copied unchanged from v1)
      generator.js      (copied unchanged from v1)
      app.js            (copied unchanged from v1)
      ui.js             NEW — result card, accordion, step pills, file-card visuals

Running it
----------
This tool loads Excel libraries (SheetJS, ExcelJS) from a CDN, so it needs
internet on first open. The spreadsheets themselves are processed in-browser.

Recommended: serve over http to avoid browser file:// restrictions. From this
folder:

    python -m http.server 8000

Then open http://localhost:8000/ in your browser.

(You can also double-click index.html, but some browsers restrict things like
live FX fetch under file:// origin — use the local server if anything seems off.)

Preserved behavior (do not regress)
-----------------------------------
All business logic is untouched: model normalization, pricing SKU mapping,
alias expansion + conflict guard, Last ASP overlay, FX conversion (live fetch
with fallback + manual override), region mapping, summary aggregation,
C+Load / ASP / Last Sold Day logic, request matching, Price Lock, margin /
total / GP% / Dif formulas (with cached values so they preview in viewers),
Foreign FX offer conversion, drop-no-C+Load, JPY conversion, and the
downloaded workbook structure including the Alias_Expansion_Log sheet.

Required DOM hooks that must remain in index.html (consumed by the JS):
sharedMappingFile, mappingStatus, step1Status, generatedSummaryHint, fxStatus,
updateFxBtn, foreignFxCheckbox, fxRate, processBtn, summaryGenerated,
downloadBtn, mappingTableBody, skipRowsCheckbox, aspMethod, fxRatesInput,
mappingSection, aliasExpansionCheckbox, aliasConflictGuardCheckbox, output,
dropZone, fileInput, lastAspInput, status, reqFile, priceFile, summaryFile,
loadBtn, generateBtn, resetBtn, regionCol, modelCol, qtyCol, offerCol,
defaultRegion, compareRegions, compareLastSold, reserveOriginal, jpy,
dropNoCloadCheckbox, regionChoices, availableRegions, preview,
droppedRowsPanel, droppedRowsTable — plus the two summarySource radios
(generated / uploaded). The load-bearing classes (status-box, ok, error,
hidden, dragover, output-table, include-checkbox, type-select, region-select,
region-option, muted, preview-wrap) must keep those exact names.
