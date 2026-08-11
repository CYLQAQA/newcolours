Margin Sheet Tool - Combined Version

Workflow
1. Load ONE shared Mapping File. It is used by both the summary builder and margin generator.
2. Optional Step 1: add source margin files and build a Margin Summary. The generated summary is kept in memory and automatically selected for Step 2. You can still download it.
3. Step 2: either use the generated Step 1 summary or choose Upload existing Summary File.
4. Load the Request File (and optional Price Lock), verify detected columns, then Generate Margin Sheet.

Files
- index.html
- assets/css/app.css
- assets/js/aggregator.js
- assets/js/generator.js
- assets/js/app.js

Notes
- Existing aggregator behavior is retained, including Type 1/Type 2 handling, Last ASP overlay, alias expansion and Summary download.
- Existing Part 2 workbook generation logic is retained.
- Generated summary is passed in memory to Step 2; it is not downloaded and re-read.
- Existing summary upload remains supported.
