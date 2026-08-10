<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Margin Sheet Generator (Part 2)</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
<style>
  :root { color-scheme: light; --bg:#f5f6f8; --card:#fff; --text:#1d2433; --muted:#687083; --line:#d8dde7; --accent:#2457d6; --danger:#b42318; --ok:#067647; }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:var(--text); font-family:Arial,Helvetica,sans-serif; }
  .app { width:min(1180px, calc(100% - 28px)); margin:20px auto 40px; }
  h1 { margin:0 0 6px; font-size:24px; }
  .sub { color:var(--muted); margin-bottom:16px; font-size:13px; }
  fieldset { border:1px solid var(--line); border-radius:10px; background:var(--card); margin:0 0 12px; padding:14px; }
  legend { padding:0 7px; font-weight:700; }
  .grid { display:grid; grid-template-columns: 220px minmax(0,1fr); gap:9px 12px; align-items:center; }
  .grid4 { display:grid; grid-template-columns: 220px minmax(180px,1fr) 230px minmax(180px,1fr); gap:9px 12px; align-items:center; }
  label { font-size:13px; }
  input[type=file], select { width:100%; border:1px solid var(--line); border-radius:7px; padding:8px; background:#fff; }
  button { border:0; border-radius:8px; padding:9px 14px; font-weight:700; cursor:pointer; background:var(--accent); color:#fff; }
  button.secondary { background:#eef2fb; color:#214da8; border:1px solid #cbd7f6; }
  button:disabled { opacity:.5; cursor:not-allowed; }
  .actions { display:flex; gap:9px; flex-wrap:wrap; align-items:center; }
  .opts { display:grid; grid-template-columns:repeat(2,minmax(240px,1fr)); gap:10px; }
  .check { display:flex; gap:8px; align-items:flex-start; font-size:13px; }
  .preview-wrap { overflow:auto; max-height:280px; border:1px solid var(--line); border-radius:8px; }
  table { border-collapse:collapse; width:max-content; min-width:100%; font-size:12px; background:#fff; }
  th, td { border:1px solid #e3e6ed; padding:6px 8px; white-space:nowrap; text-align:left; }
  th { background:#f2f4f8; position:sticky; top:0; z-index:1; }
  .regions { display:grid; grid-template-columns:repeat(auto-fill,minmax(110px,1fr)); gap:6px 10px; max-height:160px; overflow:auto; border:1px solid var(--line); border-radius:8px; padding:10px; }
  .region-option { display:flex; gap:6px; align-items:center; font-size:12px; }
  .muted { color:var(--muted); font-size:12px; }
  #status { min-height:38px; border-radius:8px; padding:10px 12px; background:#eef2f7; font-size:13px; white-space:pre-wrap; }
  #status.error { background:#fef3f2; color:var(--danger); }
  #status.ok { background:#ecfdf3; color:var(--ok); }
  .footer-note { margin-top:10px; color:var(--muted); font-size:11px; }
  @media (max-width:850px) { .grid4 { grid-template-columns:1fr; } .grid { grid-template-columns:1fr; } .opts { grid-template-columns:1fr; } }
</style>
</head>
<body>
<div class="app">
  <h1>Margin Sheet Generator (Part 2)</h1>
  <div class="sub">Browser version of the Python/Tkinter tool. Files are processed locally in this browser.</div>

  <fieldset>
    <legend>Files</legend>
    <div class="grid">
      <label for="refFile">Reference Map</label><input id="refFile" type="file" accept=".xlsx,.xls,.xlsm" />
      <label for="summaryFile">Summary File</label><input id="summaryFile" type="file" accept=".xlsx,.xls,.xlsm" />
      <label for="reqFile">Request File</label><input id="reqFile" type="file" accept=".xlsx,.xls,.xlsm" />
      <label for="priceFile">Price Lock File (optional)</label><input id="priceFile" type="file" accept=".xlsx,.xls,.xlsm,.csv" />
      <div></div><div class="actions"><button id="loadBtn">Load Files</button></div>
    </div>
  </fieldset>

  <fieldset>
    <legend>Select columns</legend>
    <div class="grid4">
      <label for="regionCol">Region Column (optional)</label><select id="regionCol"><option value=""></option></select>
      <label for="defaultRegion">Default Region if no Region column</label><select id="defaultRegion"><option value=""></option></select>
      <label for="modelCol">Model Grade Column</label><select id="modelCol"></select>
      <span></span><span></span>
      <label for="qtyCol">Quantity Column</label><select id="qtyCol"></select>
      <label for="offerCol">Offer Column</label><select id="offerCol"></select>
    </div>
  </fieldset>

  <fieldset>
    <legend>Request Preview (first 5 rows)</legend>
    <div id="preview" class="preview-wrap"><div class="muted" style="padding:12px">Load files to preview the request.</div></div>
  </fieldset>

  <fieldset>
    <legend>Options</legend>
    <div class="opts">
      <label class="check"><input id="jpy" type="checkbox" /> <span>Japanese yen</span></label>
      <label class="check"><input id="compareRegions" type="checkbox" /> <span>Compare other regions (adds optional ASP comparison columns)</span></label>
      <label class="check"><input id="reserveOriginal" type="checkbox" /> <span>Reserve original request columns in output</span></label>
      <label class="check"><input id="compareLastSold" type="checkbox" checked /> <span>Add Last Sold Day for compared regions</span></label>
    </div>
  </fieldset>

  <fieldset>
    <legend>Regions to compare</legend>
    <div class="muted" id="availableRegions">Available summary regions: -</div>
    <div id="regionChoices" class="regions" style="margin-top:8px"><span class="muted">Load files first.</span></div>
  </fieldset>

  <div class="actions">
    <button id="generateBtn" disabled>Generate Margin Sheet</button>
    <button id="resetBtn" class="secondary">Reset</button>
  </div>
  <div id="status" style="margin-top:12px">Ready.</div>
  <div class="footer-note">If this page is opened without internet access, the Excel libraries may not load. The selected spreadsheets themselves are not uploaded by this page.</div>
</div>

<script>
(() => {
  'use strict';

  const UNI_DASHES = "\u2010\u2011\u2012\u2013\u2014\u2212";
  const GRADE_SUFFIXES = new Set(["W","A","B","Z","C","NB","N","AS","CP","CX","D","DX","F","WP","RR"]);
  const state = { refMap:null, summaryByRegion:null, req:null, priceLockMap:new Map(), requestFileName:'output.xlsx' };

  const $ = id => document.getElementById(id);
  const statusEl = $('status');
  function setStatus(msg, type='') { statusEl.textContent = msg; statusEl.className = type; }
  function normalizeRegion(s) { return (s == null ? '' : String(s)).trim().toUpperCase(); }
  function collapseUnicodeDashes(s) { let out = String(s == null ? '' : s); for (const ch of UNI_DASHES) out = out.split(ch).join('-'); return out; }
  function tidySpaces(s) { return String(s == null ? '' : s).trim().replace(/\s+/g,' '); }
  function normKey(s) { return tidySpaces(collapseUnicodeDashes(s)).replace(/[\s-]+/g,'').toLowerCase(); }
  function normalizeSyn(s) { return String(s == null ? '' : s).replace(/[\s-]+/g,'').toLowerCase(); }
  function isMissing(v) { return v == null || (typeof v === 'string' && v.trim()==='') || (typeof v === 'number' && Number.isNaN(v)); }
  function safeFloat(v) {
    if (isMissing(v)) return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    let s = String(v).trim(); let neg = false;
    if (s.startsWith('(') && s.endsWith(')')) { neg = true; s = s.slice(1,-1); }
    s = s.replace(/,/g,''); const x = Number(s); return Number.isFinite(x) ? (neg ? -x : x) : null;
  }
  function splitModelGrade(modelGrade) {
    const mg = tidySpaces(collapseUnicodeDashes(modelGrade));
    const pos = mg.lastIndexOf('-');
    if (pos >= 0) return [tidySpaces(mg.slice(0,pos)), tidySpaces(mg.slice(pos+1))];
    return [mg,''];
  }
  function canonicalizeBaseModel(baseModel, refMap) { return tidySpaces(refMap.get(normKey(baseModel)) || baseModel); }
  function stripGradeIfPresent(model) { const [base, grade] = splitModelGrade(model); return grade && GRADE_SUFFIXES.has(grade.toUpperCase()) ? base : tidySpaces(collapseUnicodeDashes(model)); }

  function makeUniqueHeaders(rawHeaders) {
    const seen = new Map();
    return rawHeaders.map((h, i) => {
      let base = tidySpaces(h == null ? '' : String(h));
      if (!base) base = `Column ${i+1}`;
      const n = (seen.get(base) || 0) + 1; seen.set(base,n);
      return n === 1 ? base : `${base} (${n})`;
    });
  }

  async function readTableFile(file) {
    if (!window.XLSX) throw new Error('Excel reader library did not load. Check internet access and reopen the HTML.');
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type:'array', cellDates:true, raw:true });
    if (!wb.SheetNames.length) throw new Error(`${file.name}: workbook has no sheets.`);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const arr = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:true, blankrows:false });
    if (!arr.length) return { headers:[], rows:[] };
    const headers = makeUniqueHeaders(arr[0]);
    const rows = arr.slice(1).map(row => {
      const obj = {}; headers.forEach((h,i) => obj[h] = i < row.length ? row[i] : ''); return obj;
    });
    return { headers, rows };
  }

  function findFirstColumn(table, acceptedNames, containsAny=[]) {
    const exactLookup = new Map(table.headers.map(c => [normalizeSyn(c), c]));
    for (const name of acceptedNames) { const col = exactLookup.get(normalizeSyn(name)); if (col != null) return col; }
    for (const c of table.headers) { const lc = String(c).trim().toLowerCase(); if (containsAny.some(t => lc.includes(t))) return c; }
    return null;
  }

  function loadReferenceMap(table) {
    const canonicalCol = table.headers.find(c => ['en name','en_name','enname'].includes(String(c).trim().toLowerCase()));
    if (!canonicalCol) throw new Error("Reference file missing 'en name' column");
    const accepted = new Set(['modelcapcity','model capcity','model capacity','bi name','pricing sku','pricingsku','sku','en name','en_name','enname']);
    const synonymCols = table.headers.filter(c => accepted.has(String(c).trim().toLowerCase()));
    const map = new Map();
    for (const row of table.rows) {
      const raw = row[canonicalCol]; if (isMissing(raw)) continue;
      const canonical = tidySpaces(raw); map.set(normKey(canonical), canonical);
      for (const col of synonymCols) { const val = row[col]; if (isMissing(val)) continue; const key = normKey(val); if (key) map.set(key, canonical); }
    }
    return map;
  }

  function parseSummary(table) {
    const mgCol = table.headers.find(c => ['model/grade','modelgrade','model_grade'].includes(String(c).trim().toLowerCase()));
    if (!mgCol) throw new Error("Cannot find 'Model/Grade' column in summary.");
    const byRegion = new Map();
    for (const row of table.rows) {
      const mgVal = row[mgCol]; if (isMissing(mgVal)) continue;
      const keyExact = tidySpaces(collapseUnicodeDashes(mgVal)); const keyNorm = normKey(keyExact);
      for (const col of table.headers) {
        if (col === mgCol) continue; const val = row[col]; if (isMissing(val)) continue;
        const mReg = String(col).match(/\(([^)]+)\)\s*$/); if (!mReg) continue;
        const region = normalizeRegion(mReg[1]); if (!byRegion.has(region)) byRegion.set(region,new Map());
        const rmap = byRegion.get(region); let info = rmap.get(keyNorm) || rmap.get(keyExact); if (!info) info = {};
        rmap.set(keyExact, info); rmap.set(keyNorm, info);
        const name = String(col);
        if (name.startsWith('C+Load')) {
          const num = safeFloat(val); if (num != null) info.costUSD = num;
        } else if (name.startsWith('Last ASP')) {
          const mCurr = name.match(/^Last ASP\s*\(([^)]+)\)/); const currencyPart = mCurr ? mCurr[1].trim() : 'USD'; const num = safeFloat(val); if (num == null) continue;
          if (currencyPart.toUpperCase().includes('FX:')) {
            const parts = currencyPart.split(/FX:/i); const fx = parts.length > 1 ? Number(parts[1]) : NaN;
            info.aspUSD = num; if (Number.isFinite(fx)) info.fxRate = fx; if (!info.aspCurrency) info.aspCurrency = 'USD';
          } else { info.aspLocal = num; info.aspCurrency = currencyPart ? currencyPart.toUpperCase() : 'USD'; }
        } else if (name.startsWith('Last Sold Day')) {
          let d = val;
          if (d instanceof Date && !isNaN(d)) d = d.toISOString().slice(0,10); else d = String(d).trim().split(' ')[0];
          if (d && String(d).toUpperCase() !== 'NAT') info.lastSoldDay = d;
        }
      }
    }
    return byRegion;
  }

  function loadPriceLockMap(table, refMap) {
    if (!table.rows.length) throw new Error('Price Lock file appears to be empty.');
    const modelCol = findFirstColumn(table, new Set(['model','model capacity','modelcapcity','model/capacity','name','product']), ['model']);
    const priceCol = findFirstColumn(table, new Set(['price','price lock','pricelock','lock price','locked price','usd price']), ['price']);
    if (!modelCol || !priceCol) throw new Error('Price Lock file must contain Model and Price columns.');
    const map = new Map();
    for (const row of table.rows) {
      const mr = row[modelCol], pr = row[priceCol]; if (isMissing(mr) || isMissing(pr)) continue; const price = safeFloat(pr); if (price == null) continue;
      const modelOnly = stripGradeIfPresent(mr); const canonicalBase = canonicalizeBaseModel(modelOnly, refMap); if (!canonicalBase) continue;
      map.set(canonicalBase, price); map.set(normKey(canonicalBase), price); map.set(normKey(modelOnly), price);
    }
    if (!map.size) throw new Error('No usable Model/Price rows found in Price Lock file.');
    return map;
  }

  function getPriceLockPrice(map, canonicalBase) { if (!map || !canonicalBase) return null; return map.has(canonicalBase) ? map.get(canonicalBase) : (map.has(normKey(canonicalBase)) ? map.get(normKey(canonicalBase)) : null); }
  function detectRequestColumns(table) {
    const d = {region:null, model:null, qty:null, offer:null};
    for (const col of table.headers) {
      const lc = col.trim().toLowerCase();
      if (!d.region && ['region','warehouse','warehouse code'].includes(lc)) d.region = col;
      else if (!d.model && ['model grade','model_grade','modelgrade','model','name','product'].includes(lc)) d.model = col;
      else if (!d.qty && ['qty','quantity','qty.'].includes(lc)) d.qty = col;
      else if (!d.offer && ['offer','usd offer','offer usd','price'].includes(lc)) d.offer = col;
    }
    return d;
  }
  function getRegionInfo(summary, region, key) { const rmap = summary.get(normalizeRegion(region)); return rmap ? (rmap.get(key) || rmap.get(normKey(key)) || null) : null; }
  function buildRegionMetrics(info) {
    let cost=null, aspLocal=null, aspCurrency=null, aspUSD=null, fxRate=null, lastSold=null;
    if (info) { cost=info.costUSD ?? null; aspLocal=info.aspLocal ?? null; aspCurrency=info.aspCurrency ?? null; aspUSD=info.aspUSD ?? null; fxRate=info.fxRate ?? null; lastSold=info.lastSoldDay ?? null; }
    if (aspUSD == null && aspLocal != null) { if ((aspCurrency || 'USD').toUpperCase()==='USD') aspUSD=aspLocal; else if (fxRate) aspUSD=aspLocal*fxRate; }
    return {costUSD:cost,aspLocal,aspCurrency,aspUSD,fxRate,lastSoldDay:lastSold};
  }

  function selectedCompareRegions() { return [...document.querySelectorAll('#regionChoices input[type=checkbox]:checked')].map(x=>x.value); }
  function processRecords() {
    const regionCol = $('regionCol').value, modelCol=$('modelCol').value, qtyCol=$('qtyCol').value, offerCol=$('offerCol').value;
    if (!modelCol || !qtyCol || !offerCol) throw new Error('Please select Model Grade, Quantity and Offer columns.');
    const defaultRegion = normalizeRegion($('defaultRegion').value); if (!regionCol && !defaultRegion) throw new Error('Please select a Region column, or choose a Default Region to apply to all rows.');
    const compareRegions = $('compareRegions').checked ? selectedCompareRegions().map(normalizeRegion) : [];
    const out=[];
    for (const row of state.req.rows) {
      const region = regionCol ? normalizeRegion(row[regionCol]) : defaultRegion; if (!region) throw new Error('No region found in one or more request rows.');
      const mgStr = tidySpaces(collapseUnicodeDashes(row[modelCol])); const [baseModel, gradeSuffix] = splitModelGrade(mgStr); const canonicalBase = canonicalizeBaseModel(baseModel,state.refMap); const canonicalKey = canonicalBase + (gradeSuffix ? '-' + gradeSuffix : '');
      const priceLock = getPriceLockPrice(state.priceLockMap,canonicalBase); const info = getRegionInfo(state.summaryByRegion,region,canonicalKey); const cur=buildRegionMetrics(info); const compareData={};
      for (const cr of compareRegions) { if (cr===region) continue; const ci=getRegionInfo(state.summaryByRegion,cr,canonicalKey); const cm=buildRegionMetrics(ci); compareData[cr]={aspUSD:cm.aspUSD,lastSoldDay:cm.lastSoldDay,lookupFound:!!ci}; }
      const original={}; state.req.headers.forEach(c=> original[c]=isMissing(row[c]) ? '' : row[c]);
      out.push({ OriginalData:original, Region:region, 'Model Grade':mgStr, Grade:gradeSuffix, qty:row[qtyCol], Offer:row[offerCol], 'C+Load':cur.costUSD ?? '', 'ASP Local':cur.aspLocal ?? '', 'ASP Local Currency':cur.aspCurrency ?? '', 'ASP USD':cur.aspUSD ?? '', 'FX Rate':cur.fxRate, 'Last Sold Day':cur.lastSoldDay ?? '', 'Price Lock':priceLock ?? '', 'Price Lock Lookup Found':priceLock != null, 'Canonical Key':canonicalKey, 'Lookup Found':!!info, CompareData:compareData });
    }
    if ($('jpy').checked) {
      for (const rec of out) {
        for (const k of ['C+Load','ASP Local','ASP USD','Price Lock']) { const v=rec[k]; if (!isMissing(v)) { const n=Number(v); if (Number.isFinite(n)) rec[k]=n/0.0068; } }
        for (const comp of Object.values(rec.CompareData)) { const v=comp.aspUSD; if (!isMissing(v)) { const n=Number(v); if (Number.isFinite(n)) comp.aspUSD=n/0.0068; } }
      }
    }
    return { records:out, compareRegions };
  }

  function colLetter(n) { let s=''; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }
  function excelValue(v) { return v instanceof Date ? v : (v == null ? '' : v); }
  function applyBorder(cell) { cell.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}; }

  async function writeOutputExcel(records, compareRegions) {
    if (!window.ExcelJS) throw new Error('Excel writer library did not load. Check internet access and reopen the HTML.');
    if (!records.length) throw new Error('No request rows to export.');
    const includeLast=$('compareLastSold').checked, reserve=$('reserveOriginal').checked, includePL=state.priceLockMap.size>0;
    const ordered = records.filter(r=>!isMissing(r['C+Load'])).concat(records.filter(r=>isMissing(r['C+Load'])));
    let gen=['Region','Model Grade','Grade','qty','Offer','C+Load','Margin','Total Offer','Total Margin','GP%','','Last ASP','Last ASP(USD)','Dif','Last Sold Day'];
    if (includePL) gen.push('Price Lock','Price Lock Dif');
    for (const reg of compareRegions) { gen.push(`Last ASP(USD) - ${reg}`,`Dif - ${reg}`); if(includeLast) gen.push(`Last Sold Day - ${reg}`); }
    const headers = reserve ? [...state.req.headers,'',...gen] : [...gen];
    const generatedStart = reserve ? state.req.headers.length + 2 : 1;
    const gi={}; gen.forEach((h,i)=>{ if(h) gi[h]=generatedStart+i; });
    const wb=new ExcelJS.Workbook(); wb.creator='Margin Sheet Generator HTML'; wb.calcProperties.fullCalcOnLoad=true; wb.calcProperties.forceFullCalc=true; wb.calcProperties.calcMode='auto';
    const ws=wb.addWorksheet('Margin'); ws.addRow(headers);
    ws.getRow(1).font={bold:true};
    const currency='"$"#,##0.00', percent='0.00%';
    const qL=colLetter(gi.qty), oL=colLetter(gi.Offer), cL=colLetter(gi['C+Load']), mL=colLetter(gi.Margin), toL=colLetter(gi['Total Offer']), tmL=colLetter(gi['Total Margin']), aspL=colLetter(gi['Last ASP(USD)']);
    let r=2;
    for (const rec of ordered) {
      if (reserve) state.req.headers.forEach((c,i)=>ws.getCell(r,i+1).value=excelValue(rec.OriginalData[c]));
      ws.getCell(r,gi.Region).value=rec.Region; ws.getCell(r,gi['Model Grade']).value=rec['Model Grade']; ws.getCell(r,gi.Grade).value=rec.Grade; ws.getCell(r,gi.qty).value=excelValue(rec.qty); ws.getCell(r,gi.Offer).value=excelValue(rec.Offer); ws.getCell(r,gi['C+Load']).value=excelValue(rec['C+Load']);
      ws.getCell(r,gi.Margin).value={formula:`IF(${cL}${r}="","",${oL}${r}-${cL}${r})`};
      ws.getCell(r,gi['Total Offer']).value={formula:`IF(${cL}${r}="","",${qL}${r}*${oL}${r})`};
      ws.getCell(r,gi['Total Margin']).value={formula:`IF(${cL}${r}="","",${qL}${r}*${mL}${r})`};
      ws.getCell(r,gi['GP%']).value={formula:`IF(${cL}${r}="","",IFERROR(${tmL}${r}/${toL}${r},""))`};
      ws.getCell(r,gi['Last ASP']).value=excelValue(rec['ASP Local']); ws.getCell(r,gi['Last ASP(USD)']).value=excelValue(rec['ASP USD']); ws.getCell(r,gi.Dif).value={formula:`IF(${aspL}${r}="","",${oL}${r}-${aspL}${r})`}; ws.getCell(r,gi['Last Sold Day']).value=excelValue(rec['Last Sold Day']);
      if(includePL){ ws.getCell(r,gi['Price Lock']).value=excelValue(rec['Price Lock']); const plL=colLetter(gi['Price Lock']); ws.getCell(r,gi['Price Lock Dif']).value={formula:`IF(${plL}${r}="","",${oL}${r}-${plL}${r})`}; ws.getCell(r,gi['Price Lock']).numFmt=currency; ws.getCell(r,gi['Price Lock Dif']).numFmt=currency; }
      for(const reg of compareRegions){ const ah=`Last ASP(USD) - ${reg}`, dh=`Dif - ${reg}`, sh=`Last Sold Day - ${reg}`; const comp=rec.CompareData[reg]||{}; ws.getCell(r,gi[ah]).value=excelValue(comp.aspUSD ?? ''); const al=colLetter(gi[ah]); ws.getCell(r,gi[dh]).value={formula:`IF(${al}${r}="","",${oL}${r}-${al}${r})`}; ws.getCell(r,gi[ah]).numFmt=currency; ws.getCell(r,gi[dh]).numFmt=currency; if(includeLast) ws.getCell(r,gi[sh]).value=excelValue(comp.lastSoldDay ?? ''); }
      [gi.Offer,gi['C+Load'],gi.Margin,gi['Total Offer'],gi['Total Margin'],gi['Last ASP(USD)'],gi.Dif].forEach(c=>ws.getCell(r,c).numFmt=currency); ws.getCell(r,gi['GP%']).numFmt=percent; r++;
    }
    const lastData=r-1, totalRow=r;
    ws.getCell(totalRow,gi.Region).value='Total'; ws.getCell(totalRow,gi.qty).value={formula:`SUM(${qL}2:${qL}${lastData})`}; ws.getCell(totalRow,gi['Total Offer']).value={formula:`SUM(${toL}2:${toL}${lastData})`}; ws.getCell(totalRow,gi['Total Margin']).value={formula:`SUM(${tmL}2:${tmL}${lastData})`}; ws.getCell(totalRow,gi['GP%']).value={formula:`IFERROR(${tmL}${totalRow}/${toL}${totalRow},"")`}; ws.getRow(totalRow).font={bold:true};
    [gi.Offer,gi['C+Load'],gi.Margin,gi['Total Offer'],gi['Total Margin'],gi['Last ASP(USD)'],gi.Dif].forEach(c=>ws.getCell(totalRow,c).numFmt=currency); ws.getCell(totalRow,gi['GP%']).numFmt=percent; if(includePL){ws.getCell(totalRow,gi['Price Lock']).numFmt=currency;ws.getCell(totalRow,gi['Price Lock Dif']).numFmt=currency;}
    const regions=[]; for(const rec of ordered) if(!regions.includes(rec.Region)) regions.push(rec.Region);
    const ss=totalRow+2; ws.getCell(ss,gi.Region).value='Region'; ws.getCell(ss,gi['Total Offer']).value='Total Offer'; ws.getCell(ss,gi['Total Margin']).value='Total Margin'; ws.getCell(ss,gi['GP%']).value='GP%'; [gi.Region,gi['Total Offer'],gi['Total Margin'],gi['GP%']].forEach(c=>ws.getCell(ss,c).font={bold:true});
    const regionL=colLetter(gi.Region); let sr=ss+1;
    for(const reg of regions){ ws.getCell(sr,gi.Region).value=reg; ws.getCell(sr,gi['Total Offer']).value={formula:`SUMIF($${regionL}$2:$${regionL}$${lastData},"${String(reg).replace(/"/g,'""')}",$${toL}$2:$${toL}$${lastData})`}; ws.getCell(sr,gi['Total Margin']).value={formula:`SUMIF($${regionL}$2:$${regionL}$${lastData},"${String(reg).replace(/"/g,'""')}",$${tmL}$2:$${tmL}$${lastData})`}; ws.getCell(sr,gi['GP%']).value={formula:`IFERROR(${tmL}${sr}/${toL}${sr},"")`}; ws.getCell(sr,gi['Total Offer']).numFmt=currency; ws.getCell(sr,gi['Total Margin']).numFmt=currency; ws.getCell(sr,gi['GP%']).numFmt=percent; sr++; }
    const summaryEnd=sr-1;
    const widths={'Region':15,'Model Grade':40,'Grade':6,'qty':12,'Offer':15,'C+Load':15,'Margin':15,'Total Offer':18,'Total Margin':18,'GP%':10,'Last ASP':18,'Last ASP(USD)':18,'Dif':15,'Last Sold Day':14,'Price Lock':15,'Price Lock Dif':15};
    if(reserve) state.req.headers.forEach((h,i)=>ws.getColumn(i+1).width=Math.min(Math.max(String(h).length+2,12),35));
    for(const [h,w] of Object.entries(widths)) if(gi[h]) ws.getColumn(gi[h]).width=w;
    for(const reg of compareRegions){ws.getColumn(gi[`Last ASP(USD) - ${reg}`]).width=18;ws.getColumn(gi[`Dif - ${reg}`]).width=15;if(includeLast)ws.getColumn(gi[`Last Sold Day - ${reg}`]).width=16;}
    const maxCol=ws.columnCount; for(let rr=1;rr<=totalRow;rr++) for(let cc=1;cc<=maxCol;cc++) applyBorder(ws.getCell(rr,cc)); for(let rr=ss;rr<=summaryEnd;rr++) [gi.Region,gi['Total Offer'],gi['Total Margin'],gi['GP%']].forEach(cc=>applyBorder(ws.getCell(rr,cc)));
    ws.views=[{state:'frozen',ySplit:1}];
    const base=(state.requestFileName||'output.xlsx').replace(/\.[^.]+$/,''); const filename=`${base}_ms.xlsx`; const buffer=await wb.xlsx.writeBuffer(); const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000); return filename;
  }

  function fillSelect(select, values, chosen='', allowBlank=false) { select.innerHTML=''; if(allowBlank){const o=document.createElement('option');o.value='';o.textContent='';select.appendChild(o);} for(const v of values){const o=document.createElement('option');o.value=v;o.textContent=v;select.appendChild(o);} if(chosen && values.includes(chosen)) select.value=chosen; else if(allowBlank) select.value=''; }
  function renderPreview(table){ const wrap=$('preview'); wrap.innerHTML=''; const t=document.createElement('table'), head=document.createElement('thead'), hr=document.createElement('tr'); table.headers.forEach(h=>{const th=document.createElement('th');th.textContent=h;hr.appendChild(th);});head.appendChild(hr);t.appendChild(head);const body=document.createElement('tbody');table.rows.slice(0,5).forEach(row=>{const tr=document.createElement('tr');table.headers.forEach(h=>{const td=document.createElement('td');let v=row[h];if(v instanceof Date&&!isNaN(v))v=v.toISOString().slice(0,10);td.textContent=v==null?'':String(v);tr.appendChild(td);});body.appendChild(tr);});t.appendChild(body);wrap.appendChild(t); }
  function renderRegions(regions){ $('availableRegions').textContent='Available summary regions: '+(regions.length?regions.join(', '):'-'); const box=$('regionChoices'); box.innerHTML=''; if(!regions.length){box.innerHTML='<span class="muted">No regions found.</span>';return;} for(const reg of regions){const lab=document.createElement('label');lab.className='region-option';const cb=document.createElement('input');cb.type='checkbox';cb.value=reg;const s=document.createElement('span');s.textContent=reg;lab.append(cb,s);box.appendChild(lab);} }

  $('loadBtn').addEventListener('click', async()=>{
    try {
      const ref=$('refFile').files[0], sum=$('summaryFile').files[0], req=$('reqFile').files[0], pl=$('priceFile').files[0]; if(!ref||!sum||!req) throw new Error('Please select Reference Map, Summary File and Request File before loading.');
      $('loadBtn').disabled=true; $('generateBtn').disabled=true; setStatus('Loading and parsing files...');
      const [refT,sumT,reqT,plT]=await Promise.all([readTableFile(ref),readTableFile(sum),readTableFile(req),pl?readTableFile(pl):Promise.resolve(null)]);
      if(!reqT.rows.length) throw new Error('The request file appears to be empty.');
      const refMap=loadReferenceMap(refT), summary=parseSummary(sumT); let plMap=new Map(); if(plT) plMap=loadPriceLockMap(plT,refMap);
      state.refMap=refMap; state.summaryByRegion=summary; state.req=reqT; state.priceLockMap=plMap; state.requestFileName=req.name;
      const det=detectRequestColumns(reqT), cols=reqT.headers, regions=[...summary.keys()].sort();
      fillSelect($('regionCol'),cols,det.region||'',true); fillSelect($('modelCol'),cols,det.model||''); fillSelect($('qtyCol'),cols,det.qty||''); fillSelect($('offerCol'),cols,det.offer||''); fillSelect($('defaultRegion'),regions,regions.length===1?regions[0]:'',true); renderPreview(reqT); renderRegions(regions);
      $('generateBtn').disabled=false; setStatus(`Files loaded. ${reqT.rows.length} request row(s), ${regions.length} summary region(s).${plMap.size?' Price Lock matching enabled.':''}\nPlease verify the selected columns before generating.`, 'ok');
    } catch(e) { setStatus(e.message||String(e),'error'); } finally { $('loadBtn').disabled=false; }
  });

  $('generateBtn').addEventListener('click', async()=>{
    try { $('generateBtn').disabled=true; setStatus('Processing request and building Excel...'); const {records,compareRegions}=processRecords(); const file=await writeOutputExcel(records,compareRegions); const missing=records.filter(r=>!r['Lookup Found']).length; setStatus(`Done. Downloaded ${file}.\nRows: ${records.length}. Summary lookups not found: ${missing}.`, 'ok'); }
    catch(e){ console.error(e); setStatus(e.message||String(e),'error'); }
    finally{$('generateBtn').disabled=false;}
  });

  $('resetBtn').addEventListener('click',()=>location.reload());
})();
</script>
</body>
</html>
