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

  // ----- Foreign FX helpers -----
  function getOfferFxSetup() {
    const on = $('foreignFxCheckbox') && $('foreignFxCheckbox').checked;
    if (!on) return { enabled:false, rate:null, header:'Offer', convert:(v)=>v };
    const raw = $('fxRate').value;
    const rate = parseFloat(raw);
    if (!raw || !Number.isFinite(rate) || rate <= 0) throw new Error('Foreign FX is enabled but FX rate is missing or invalid.');
    return {
      enabled:true,
      rate,
      header:`Offer(Fx:${rate})`,
      convert:(v)=> v == null || !Number.isFinite(v) ? v : v / rate
    };
  }

  // ----- Formula-as-cached-value helper -----
  // Writes { formula, result } so the cached value is visible in Excel/Sheets/Numbers previews,
  // and the formula still recomputes when the user edits inputs.
  function setFormulaCell(ws, rowIdx, colIdx, formula, result, numFmt) {
    const cell = ws.getCell(rowIdx, colIdx);
    cell.value = { formula, result };
    if (numFmt) cell.numFmt = numFmt;
    return cell;
  }

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
    const offerFx = getOfferFxSetup(); // validates FX rate if checkbox on
    const dropNoCload = $('dropNoCloadCheckbox') && $('dropNoCloadCheckbox').checked;
    const out=[];
    const dropped=[];
    for (const row of state.req.rows) {
      const region = regionCol ? normalizeRegion(row[regionCol]) : defaultRegion; if (!region) throw new Error('No region found in one or more request rows.');
      const mgStr = tidySpaces(collapseUnicodeDashes(row[modelCol])); const [baseModel, gradeSuffix] = splitModelGrade(mgStr); const canonicalBase = canonicalizeBaseModel(baseModel,state.refMap); const canonicalKey = canonicalBase + (gradeSuffix ? '-' + gradeSuffix : '');
      const priceLock = getPriceLockPrice(state.priceLockMap,canonicalBase); const info = getRegionInfo(state.summaryByRegion,region,canonicalKey); const cur=buildRegionMetrics(info); const compareData={};
      for (const cr of compareRegions) { if (cr===region) continue; const ci=getRegionInfo(state.summaryByRegion,cr,canonicalKey); const cm=buildRegionMetrics(ci); compareData[cr]={aspUSD:cm.aspUSD,lastSoldDay:cm.lastSoldDay,lookupFound:!!ci}; }
      const original={}; state.req.headers.forEach(c=> original[c]=isMissing(row[c]) ? '' : row[c]);
      const rawOffer = safeFloat(row[offerCol]);
      const offerOut = offerFx.convert(rawOffer);
      const noCload = isMissing(cur.costUSD);
      const rec = {
        OriginalData:original, Region:region, 'Model Grade':mgStr, Grade:gradeSuffix,
        qty:row[qtyCol], Offer:offerOut,
        'C+Load':cur.costUSD ?? '', 'ASP Local':cur.aspLocal ?? '', 'ASP Local Currency':cur.aspCurrency ?? '',
        'ASP USD':cur.aspUSD ?? '', 'FX Rate':cur.fxRate, 'Last Sold Day':cur.lastSoldDay ?? '',
        'Price Lock':priceLock ?? '', 'Price Lock Lookup Found':priceLock != null,
        'Canonical Key':canonicalKey, 'Lookup Found':!!info, CompareData:compareData,
        _noCload: noCload
      };
      if (dropNoCload && noCload) {
        dropped.push({ Region:region, 'Model Grade':mgStr, Reason:'No C+Load' });
      } else {
        out.push(rec);
      }
    }
    if ($('jpy').checked) {
      for (const rec of out) {
        for (const k of ['C+Load','ASP Local','ASP USD','Price Lock']) { const v=rec[k]; if (!isMissing(v)) { const n=Number(v); if (Number.isFinite(n)) rec[k]=n/0.0068; } }
        for (const comp of Object.values(rec.CompareData)) { const v=comp.aspUSD; if (!isMissing(v)) { const n=Number(v); if (Number.isFinite(n)) comp.aspUSD=n/0.0068; } }
      }
    }
    return { records:out, compareRegions, dropped, offerFx };
  }

  function colLetter(n) { let s=''; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }
  function excelValue(v) { return v instanceof Date ? v : (v == null ? '' : v); }
  function applyBorder(cell) { cell.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}}; }

  // Builds one "Margin"-style worksheet (headers, rows, totals, per-region summary, formatting).
  function buildMarginSheet(wb, sheetName, ordered, compareRegions, offerFx) {
    const includeLast=$('compareLastSold').checked, reserve=$('reserveOriginal').checked, includePL=state.priceLockMap.size>0;
    const offerHeader = (offerFx && offerFx.header) || 'Offer';
    let gen=['Region','Model Grade','Grade','qty',offerHeader,'C+Load','Margin','Total Offer','Total Margin','GP%','','Last ASP','Last ASP(USD)','Dif','Last Sold Day'];
    if (includePL) gen.push('Price Lock','Price Lock Dif');
    for (const reg of compareRegions) { gen.push(`Last ASP(USD) - ${reg}`,`Dif - ${reg}`); if(includeLast) gen.push(`Last Sold Day - ${reg}`); }
    const headers = reserve ? [...state.req.headers,'',...gen] : [...gen];
    const generatedStart = reserve ? state.req.headers.length + 2 : 1;
    const gi={}; gen.forEach((h,i)=>{ if(h) gi[h]=generatedStart+i; });
    const ws=wb.addWorksheet(sheetName); ws.addRow(headers);
    ws.getRow(1).font={bold:true};
    const currency='"$"#,##0.00', percent='0.00%';
    const qL=colLetter(gi.qty), oL=colLetter(gi[offerHeader]), cL=colLetter(gi['C+Load']), mL=colLetter(gi.Margin), toL=colLetter(gi['Total Offer']), tmL=colLetter(gi['Total Margin']), aspL=colLetter(gi['Last ASP(USD)']);
    let r=2;
    for (const rec of ordered) {
      if (reserve) state.req.headers.forEach((c,i)=>ws.getCell(r,i+1).value=excelValue(rec.OriginalData[c]));
      ws.getCell(r,gi.Region).value=rec.Region; ws.getCell(r,gi['Model Grade']).value=rec['Model Grade']; ws.getCell(r,gi.Grade).value=rec.Grade; ws.getCell(r,gi.qty).value=excelValue(rec.qty); ws.getCell(r,gi[offerHeader]).value=excelValue(rec.Offer); ws.getCell(r,gi['C+Load']).value=excelValue(rec['C+Load']);

      // Pre-compute formula results so cached values are visible in Excel/Sheets/Numbers previews.
      const offerV = isMissing(rec.Offer) ? null : Number(rec.Offer);
      const costV = isMissing(rec['C+Load']) ? null : Number(rec['C+Load']);
      const qtyV = isMissing(rec.qty) ? null : Number(rec.qty);
      const marginRes = (offerV != null && costV != null) ? (offerV - costV) : '';
      const totalOfferRes = (offerV != null && qtyV != null) ? (offerV * qtyV) : '';
      const totalMarginRes = (marginRes !== '' && qtyV != null) ? (marginRes * qtyV) : '';
      const gpRes = (totalOfferRes !== '' && totalOfferRes !== 0 && totalMarginRes !== '') ? (totalMarginRes / totalOfferRes) : '';

      setFormulaCell(ws, r, gi.Margin, `IF(${cL}${r}="","",${oL}${r}-${cL}${r})`, marginRes, currency);
      setFormulaCell(ws, r, gi['Total Offer'], `IF(${cL}${r}="","",${qL}${r}*${oL}${r})`, totalOfferRes, currency);
      setFormulaCell(ws, r, gi['Total Margin'], `IF(${cL}${r}="","",${qL}${r}*${mL}${r})`, totalMarginRes, currency);
      setFormulaCell(ws, r, gi['GP%'], `IF(${cL}${r}="","",IFERROR(${tmL}${r}/${toL}${r},""))`, gpRes, percent);

      const aspV = isMissing(rec['ASP USD']) ? null : Number(rec['ASP USD']);
      const difRes = (offerV != null && aspV != null) ? (offerV - aspV) : '';
      ws.getCell(r,gi['Last ASP']).value=excelValue(rec['ASP Local']); ws.getCell(r,gi['Last ASP(USD)']).value=excelValue(rec['ASP USD']);
      setFormulaCell(ws, r, gi.Dif, `IF(${aspL}${r}="","",${oL}${r}-${aspL}${r})`, difRes, currency);
      ws.getCell(r,gi['Last Sold Day']).value=excelValue(rec['Last Sold Day']);
      if(includePL){
        ws.getCell(r,gi['Price Lock']).value=excelValue(rec['Price Lock']);
        const plL=colLetter(gi['Price Lock']);
        const plV = isMissing(rec['Price Lock']) ? null : Number(rec['Price Lock']);
        const plDifRes = (offerV != null && plV != null) ? (offerV - plV) : '';
        setFormulaCell(ws, r, gi['Price Lock Dif'], `IF(${plL}${r}="","",${oL}${r}-${plL}${r})`, plDifRes, currency);
        ws.getCell(r,gi['Price Lock']).numFmt=currency; ws.getCell(r,gi['Price Lock Dif']).numFmt=currency;
      }
      for(const reg of compareRegions){ const ah=`Last ASP(USD) - ${reg}`, dh=`Dif - ${reg}`, sh=`Last Sold Day - ${reg}`; const comp=rec.CompareData[reg]||{}; const compAsp = isMissing(comp.aspUSD) ? null : Number(comp.aspUSD); const compDif = (offerV != null && compAsp != null) ? (offerV - compAsp) : ''; ws.getCell(r,gi[ah]).value=excelValue(comp.aspUSD ?? ''); const al=colLetter(gi[ah]); setFormulaCell(ws, r, gi[dh], `IF(${al}${r}="","",${oL}${r}-${al}${r})`, compDif, currency); ws.getCell(r,gi[ah]).numFmt=currency; ws.getCell(r,gi[dh]).numFmt=currency; if(includeLast) ws.getCell(r,gi[sh]).value=excelValue(comp.lastSoldDay ?? ''); }
      [gi[offerHeader],gi['C+Load'],gi.Margin,gi['Total Offer'],gi['Total Margin'],gi['Last ASP(USD)'],gi.Dif].forEach(c=>ws.getCell(r,c).numFmt=currency); ws.getCell(r,gi['GP%']).numFmt=percent; r++;
    }
    const lastData=r-1, totalRow=r;
    // Compute total row values from the row values we already wrote
    let sumQty=0, sumTO=0, sumTM=0; let anyQty=false, anyTO=false, anyTM=false;
    for (let rr=2; rr<=lastData; rr++) {
      const qv = ws.getCell(rr, gi.qty).value; const qn = (typeof qv === 'object' && qv && 'result' in qv) ? Number(qv.result) : Number(qv);
      const tv = ws.getCell(rr, gi['Total Offer']).value; const tn = (typeof tv === 'object' && tv && 'result' in tv) ? Number(tv.result) : Number(tv);
      const mv = ws.getCell(rr, gi['Total Margin']).value; const mn = (typeof mv === 'object' && mv && 'result' in mv) ? Number(mv.result) : Number(mv);
      if (Number.isFinite(qn)) { sumQty+=qn; anyQty=true; }
      if (Number.isFinite(tn)) { sumTO+=tn; anyTO=true; }
      if (Number.isFinite(mn)) { sumTM+=mn; anyTM=true; }
    }
    const totalGp = (anyTO && sumTO!==0 && anyTM) ? (sumTM/sumTO) : '';
    ws.getCell(totalRow,gi.Region).value='Total';
    setFormulaCell(ws, totalRow, gi.qty, `SUM(${qL}2:${qL}${lastData})`, anyQty?sumQty:'', null);
    setFormulaCell(ws, totalRow, gi['Total Offer'], `SUM(${toL}2:${toL}${lastData})`, anyTO?sumTO:'', currency);
    setFormulaCell(ws, totalRow, gi['Total Margin'], `SUM(${tmL}2:${tmL}${lastData})`, anyTM?sumTM:'', currency);
    setFormulaCell(ws, totalRow, gi['GP%'], `IFERROR(${tmL}${totalRow}/${toL}${totalRow},"")`, totalGp, percent);
    ws.getRow(totalRow).font={bold:true};
    [gi[offerHeader],gi['C+Load'],gi.Margin,gi['Total Offer'],gi['Total Margin'],gi['Last ASP(USD)'],gi.Dif].forEach(c=>ws.getCell(totalRow,c).numFmt=currency); ws.getCell(totalRow,gi['GP%']).numFmt=percent; if(includePL){ws.getCell(totalRow,gi['Price Lock']).numFmt=currency;ws.getCell(totalRow,gi['Price Lock Dif']).numFmt=currency;}
    const regions=[]; for(const rec of ordered) if(!regions.includes(rec.Region)) regions.push(rec.Region);
    const ss=totalRow+2; ws.getCell(ss,gi.Region).value='Region'; ws.getCell(ss,gi['Total Offer']).value='Total Offer'; ws.getCell(ss,gi['Total Margin']).value='Total Margin'; ws.getCell(ss,gi['GP%']).value='GP%'; [gi.Region,gi['Total Offer'],gi['Total Margin'],gi['GP%']].forEach(c=>ws.getCell(ss,c).font={bold:true});
    const regionL=colLetter(gi.Region); let sr=ss+1;
    // Pre-compute per-region totals by scanning column values
    const regionTotals={};
    for (let rr=2; rr<=lastData; rr++) {
      const rv = ws.getCell(rr, gi.Region).value; if (typeof rv !== 'string') continue;
      const tv = ws.getCell(rr, gi['Total Offer']).value; const tn = (typeof tv === 'object' && tv && 'result' in tv) ? Number(tv.result) : Number(tv);
      const mv = ws.getCell(rr, gi['Total Margin']).value; const mn = (typeof mv === 'object' && mv && 'result' in mv) ? Number(mv.result) : Number(mv);
      if (!regionTotals[rv]) regionTotals[rv]={to:0,tm:0,hasTo:false,hasTm:false};
      if (Number.isFinite(tn)) { regionTotals[rv].to+=tn; regionTotals[rv].hasTo=true; }
      if (Number.isFinite(mn)) { regionTotals[rv].tm+=mn; regionTotals[rv].hasTm=true; }
    }
    for(const reg of regions){
      const rt = regionTotals[reg] || {to:0,tm:0,hasTo:false,hasTm:false};
      const gp = (rt.hasTo && rt.to!==0 && rt.hasTm) ? (rt.tm/rt.to) : '';
      ws.getCell(sr,gi.Region).value=reg;
      setFormulaCell(ws, sr, gi['Total Offer'], `SUMIF($${regionL}$2:$${regionL}$${lastData},"${String(reg).replace(/"/g,'""')}",$${toL}$2:$${toL}$${lastData})`, rt.hasTo?rt.to:'', currency);
      setFormulaCell(ws, sr, gi['Total Margin'], `SUMIF($${regionL}$2:$${regionL}$${lastData},"${String(reg).replace(/"/g,'""')}",$${tmL}$2:$${tmL}$${lastData})`, rt.hasTm?rt.tm:'', currency);
      setFormulaCell(ws, sr, gi['GP%'], `IFERROR(${tmL}${sr}/${toL}${sr},"")`, gp, percent);
      ws.getCell(sr,gi['Total Offer']).numFmt=currency; ws.getCell(sr,gi['Total Margin']).numFmt=currency; ws.getCell(sr,gi['GP%']).numFmt=percent; sr++;
    }
    const summaryEnd=sr-1;
    const widths={'Region':15,'Model Grade':40,'Grade':6,'qty':12,'C+Load':15,'Margin':15,'Total Offer':18,'Total Margin':18,'GP%':10,'Last ASP':18,'Last ASP(USD)':18,'Dif':15,'Last Sold Day':14,'Price Lock':15,'Price Lock Dif':15};
    widths[offerHeader]=offerHeader.startsWith('Offer(Fx:') ? 22 : 15;
    if(reserve) state.req.headers.forEach((h,i)=>ws.getColumn(i+1).width=Math.min(Math.max(String(h).length+2,12),35));
    for(const [h,w] of Object.entries(widths)) if(gi[h]) ws.getColumn(gi[h]).width=w;
    for(const reg of compareRegions){ws.getColumn(gi[`Last ASP(USD) - ${reg}`]).width=18;ws.getColumn(gi[`Dif - ${reg}`]).width=15;if(includeLast)ws.getColumn(gi[`Last Sold Day - ${reg}`]).width=16;}
    const maxCol=ws.columnCount; for(let rr=1;rr<=totalRow;rr++) for(let cc=1;cc<=maxCol;cc++) applyBorder(ws.getCell(rr,cc)); for(let rr=ss;rr<=summaryEnd;rr++) [gi.Region,gi['Total Offer'],gi['Total Margin'],gi['GP%']].forEach(cc=>applyBorder(ws.getCell(rr,cc)));
    ws.views=[{state:'frozen',ySplit:1}];
    return ws;
  }

  async function writeOutputExcel(records, compareRegions, dropped, offerFx) {
    if (!window.ExcelJS) throw new Error('Excel writer library did not load. Check internet access and reopen the HTML.');
    if (!records.length) throw new Error('No request rows to export.');
    const ordered = records.filter(r=>!isMissing(r['C+Load'])).concat(records.filter(r=>isMissing(r['C+Load'])));
    const wb=new ExcelJS.Workbook(); wb.creator='Margin Sheet Generator HTML'; wb.calcProperties.fullCalcOnLoad=true; wb.calcProperties.forceFullCalc=true; wb.calcProperties.calcMode='auto';
    buildMarginSheet(wb, 'Margin', ordered, compareRegions, offerFx);

    // Optional: split by region — one extra tab per region with only that region's rows.
    const splitEl = $('splitByRegionCheckbox');
    if (splitEl && splitEl.checked) {
      const regions=[]; for(const rec of ordered) if(!regions.includes(rec.Region)) regions.push(rec.Region);
      for (const reg of regions) {
        const rows = ordered.filter(r=>r.Region===reg);
        if (!rows.length) continue;
        // ExcelJS sheet names: max 31 chars, no []:*?/\
        const safe = String(reg).replace(/[[\]:*?/\\]/g,'-').slice(0,31) || 'Region';
        let name=safe, n=2; while (wb.worksheets.some(s=>s.name===name)) name=`${safe.slice(0,28)}_${n++}`;
        buildMarginSheet(wb, name, rows, compareRegions, offerFx);
      }
    }

    // Render dropped-rows preview panel in HTML (no workbook sheet)
    renderDroppedRowsPanel(dropped || []);

    const base=(state.requestFileName||'output.xlsx').replace(/\.[^.]+$/,''); const filename=`${base}_ms.xlsx`; const buffer=await wb.xlsx.writeBuffer(); const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000); return filename;
  }

  function renderDroppedRowsPanel(dropped) {
    const panel = $('droppedRowsPanel');
    const wrap = $('droppedRowsTable');
    if (!panel || !wrap) return;
    if (!dropped || !dropped.length) { panel.classList.add('hidden'); wrap.innerHTML='<div class="muted" style="padding:12px">No rows were dropped.</div>'; return; }
    wrap.innerHTML='';
    const t = document.createElement('table');
    const head = document.createElement('thead'); const hr = document.createElement('tr');
    ['Region','Model Grade','Reason'].forEach(h=>{ const th=document.createElement('th'); th.textContent=h; hr.appendChild(th); });
    head.appendChild(hr); t.appendChild(head);
    const body = document.createElement('tbody');
    dropped.forEach(d=>{
      const tr = document.createElement('tr');
      [d.Region, d['Model Grade'], d.Reason].forEach(v=>{ const td=document.createElement('td'); td.textContent = v == null ? '' : String(v); tr.appendChild(td); });
      body.appendChild(tr);
    });
    t.appendChild(body); wrap.appendChild(t);
    panel.classList.remove('hidden');
  }

  function fillSelect(select, values, chosen='', allowBlank=false) { select.innerHTML=''; if(allowBlank){const o=document.createElement('option');o.value='';o.textContent='';select.appendChild(o);} for(const v of values){const o=document.createElement('option');o.value=v;o.textContent=v;select.appendChild(o);} if(chosen && values.includes(chosen)) select.value=chosen; else if(allowBlank) select.value=''; }
  function renderPreview(table){ const wrap=$('preview'); wrap.innerHTML=''; const t=document.createElement('table'), head=document.createElement('thead'), hr=document.createElement('tr'); table.headers.forEach(h=>{const th=document.createElement('th');th.textContent=h;hr.appendChild(th);});head.appendChild(hr);t.appendChild(head);const body=document.createElement('tbody');table.rows.slice(0,5).forEach(row=>{const tr=document.createElement('tr');table.headers.forEach(h=>{const td=document.createElement('td');let v=row[h];if(v instanceof Date&&!isNaN(v))v=v.toISOString().slice(0,10);td.textContent=v==null?'':String(v);tr.appendChild(td);});body.appendChild(tr);});t.appendChild(body);wrap.appendChild(t); }
  function renderRegions(regions){ $('availableRegions').textContent='Available summary regions: '+(regions.length?regions.join(', '):'-'); const box=$('regionChoices'); box.innerHTML=''; if(!regions.length){box.innerHTML='<span class="muted">No regions found.</span>';return;} for(const reg of regions){const lab=document.createElement('label');lab.className='region-option';const cb=document.createElement('input');cb.type='checkbox';cb.value=reg;const s=document.createElement('span');s.textContent=reg;lab.append(cb,s);box.appendChild(lab);} }

  let sharedMappingReady=false;
  let generatedSummaryObject=null;

  function generatedSummaryToMap(obj) {
    const byRegion=new Map();
    if(!obj || typeof obj!=='object') return byRegion;
    for(const [regionRaw, rows] of Object.entries(obj)) {
      const region=normalizeRegion(regionRaw);
      if(!region || !rows || typeof rows!=='object') continue;
      const rmap=new Map();
      for(const [storedKey, recRaw] of Object.entries(rows)) {
        const rec=recRaw || {};
        const exact=tidySpaces(collapseUnicodeDashes(rec.originalKey || storedKey));
        const info={
          costUSD:rec.costUSD ?? null,
          aspLocal:rec.aspLocal ?? null,
          aspCurrency:rec.aspCurrency ?? null,
          aspUSD:rec.aspUSD ?? null,
          fxRate:rec.fxRate ?? null,
          lastSoldDay:rec.lastSoldDay ?? null
        };
        if(exact) {
          rmap.set(exact,info);
          rmap.set(normKey(exact),info);
        }
        if(storedKey) rmap.set(storedKey,info);
      }
      byRegion.set(region,rmap);
    }
    return byRegion;
  }

  function getSummarySource() {
    const checked=document.querySelector('input[name="summarySource"]:checked');
    return checked ? checked.value : 'generated';
  }

  function setSummarySourceUi() {
    const uploaded=getSummarySource()==='uploaded';
    $('summaryFile').disabled=!uploaded;
    const generatedRadio=$('summaryGenerated');
    if(generatedRadio) generatedRadio.disabled=!generatedSummaryObject;
  }

  async function setSharedMappingFile(file) {
    if(!file) throw new Error('No mapping file selected.');
    const table=await readTableFile(file);
    state.refMap=loadReferenceMap(table);
    sharedMappingReady=true;
    return state.refMap.size;
  }

  function setGeneratedSummary(summaryObj) {
    generatedSummaryObject=summaryObj || null;
    const radio=$('summaryGenerated');
    if(radio) {
      radio.disabled=!generatedSummaryObject;
      if(generatedSummaryObject) radio.checked=true;
    }
    setSummarySourceUi();
  }

  async function loadGeneratorInputs() {
    if(!sharedMappingReady || !state.refMap) throw new Error('Please load the shared Mapping File first.');
    const req=$('reqFile').files[0], pl=$('priceFile').files[0];
    if(!req) throw new Error('Please select a Request File.');

    let summary;
    const source=getSummarySource();
    if(source==='generated') {
      if(!generatedSummaryObject) throw new Error('No generated summary is available yet. Build Step 1 first, or choose an existing Summary File.');
      summary=generatedSummaryToMap(generatedSummaryObject);
    } else {
      const sum=$('summaryFile').files[0];
      if(!sum) throw new Error('Please select an existing Summary File.');
      const sumT=await readTableFile(sum);
      summary=parseSummary(sumT);
    }

    const [reqT,plT]=await Promise.all([
      readTableFile(req),
      pl ? readTableFile(pl) : Promise.resolve(null)
    ]);
    if(!reqT.rows.length) throw new Error('The request file appears to be empty.');

    let plMap=new Map();
    if(plT) plMap=loadPriceLockMap(plT,state.refMap);
    state.summaryByRegion=summary;
    state.req=reqT;
    state.priceLockMap=plMap;
    state.requestFileName=req.name;

    const det=detectRequestColumns(reqT), cols=reqT.headers, regions=[...summary.keys()].sort();
    fillSelect($('regionCol'),cols,det.region||'',true);
    fillSelect($('modelCol'),cols,det.model||'');
    fillSelect($('qtyCol'),cols,det.qty||'');
    fillSelect($('offerCol'),cols,det.offer||'');
    fillSelect($('defaultRegion'),regions,regions.length===1?regions[0]:'',true);
    renderPreview(reqT);
    renderRegions(regions);
    $('generateBtn').disabled=false;
    setStatus(`Loaded ${reqT.rows.length} request row(s) with ${regions.length} summary region(s) from ${source==='generated'?'Step 1 generated summary':'uploaded summary file'}.${plMap.size?' Price Lock matching enabled.':''}\nPlease verify the selected columns before generating.`,'ok');
    return {rows:reqT.rows.length,regions:regions.length,source};
  }

  $('loadBtn').addEventListener('click', async()=>{
    try {
      $('loadBtn').disabled=true; $('generateBtn').disabled=true; setStatus('Loading request and summary...');
      await loadGeneratorInputs();
    } catch(e) { setStatus(e.message||String(e),'error'); }
    finally { $('loadBtn').disabled=false; }
  });

  $('generateBtn').addEventListener('click', async()=>{
    try{
      $('generateBtn').disabled=true; setStatus('Generating workbook...');
      const {records,compareRegions,dropped,offerFx}=processRecords();
      await writeOutputExcel(records,compareRegions,dropped,offerFx);
      const missing=records.filter(r=>!r['Lookup Found']).length;
      const plMissing=state.priceLockMap.size ? records.filter(r=>!r['Price Lock Lookup Found']).length : 0;
      const dropMsg = dropped && dropped.length ? ` Dropped ${dropped.length} row(s) with no C+Load (see preview panel).` : '';
      setStatus(`Done. ${records.length} row(s) exported.${dropMsg} Summary lookup missing: ${missing}.${state.priceLockMap.size?` Price Lock missing: ${plMissing}.`:''}`,'ok');
    }catch(e){setStatus(e.message||String(e),'error');}
    finally{$('generateBtn').disabled=false;}
  });

  document.querySelectorAll('input[name="summarySource"]').forEach(el=>el.addEventListener('change',setSummarySourceUi));
  $('resetBtn').addEventListener('click',()=>location.reload());
  setSummarySourceUi();

  window.MarginGeneratorAPI={
    setSharedMappingFile,
    setGeneratedSummary,
    loadGeneratorInputs,
    generatedSummaryToMap,
    isMappingReady(){return sharedMappingReady;}
  };
})();
