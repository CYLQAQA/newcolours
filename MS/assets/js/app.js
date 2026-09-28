(() => {
  'use strict';
  const $=id=>document.getElementById(id);
  const mappingInput=$('sharedMappingFile');
  const mappingStatus=$('mappingStatus');
  const step1Status=$('step1Status');
  const sourceHint=$('generatedSummaryHint');
  const fxStatus=$('fxStatus');

  // All rates in this module are FOREIGN -> USD (e.g. HKD=0.13 means 1 HKD = 0.13 USD),
  // which is the direction the aggregator multiplies with. USD itself is fixed at 1.
  const FALLBACK_RATES = {HKD:0.127478046,JPY:0.006349635,AUD:0.701366543,EUR:1.138490544,GBP:1.323462071,CAD:0.706786707,SGD:0.782339163,NZD:0.565720612};
  const FALLBACK_RATES_DATE = '2026-09-28';
  const FX_CACHE_PREFIX = 'marginFx_';
  const FX_SOURCES = [
    'https://open.er-api.com/v6/latest/USD',
    'https://api.frankfurter.dev/v1/latest?base=USD'
  ];

  function setBox(el,msg,type=''){
    el.textContent=msg;
    el.className='status-box'+(type?` ${type}`:'');
  }

  function todayKey(){
    const d=new Date();
    const pad=n=>String(n).padStart(2,'0');
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  }

  const fxState={rates:{USD:1, ...FALLBACK_RATES}, source:'fallback'};

  function setFxStatus(){
    if(!fxStatus) return;
    if(fxState.source==='live'){
      fxStatus.textContent=`FX: live rates as of ${todayKey()}. Click Update FX to refresh.`;
    } else if(fxState.source==='cache'){
      const cachedDate=fxState.cacheDate || todayKey();
      const isToday=cachedDate===todayKey();
      fxStatus.textContent=`FX: cached rates from ${cachedDate}${isToday?'':' (older)'}. Click Update FX to refresh.`;
    } else {
      fxStatus.textContent=`FX: using built-in fallback table (as of ${FALLBACK_RATES_DATE}). Click Update FX to fetch live rates.`;
    }
  }

  // Live APIs return USD -> foreign (1 USD = N foreign). Normalize any of them
  // into the FOREIGN -> USD direction this tool uses. Returns null on failure.
  function normalizeLiveRates(data){
    let usdToForeign=null;
    if(data && data.rates && typeof data.rates==='object'){
      usdToForeign=data.rates;
    } else if(data && typeof data==='object'){
      // er-api-style body also carries rates at top level under .rates; frankfurter v1 nests nothing.
      usdToForeign=null;
    }
    if(!usdToForeign) return null;
    const out={USD:1};
    for(const [code,usdToCur] of Object.entries(usdToForeign)){
      const n=Number(usdToCur);
      const C=String(code).toUpperCase();
      if(C==='USD'){ out.USD=1; continue; }
      if(Number.isFinite(n) && n>0) out[C]=1/n;
    }
    return Object.keys(out).length>1 ? out : null;
  }

  async function fetchLiveFx(){
    if(fxStatus) fxStatus.textContent='FX: refreshing...';
    let lastErr=null;
    for(const url of FX_SOURCES){
      try{
        const controller=new AbortController();
        const timeoutId=setTimeout(()=>controller.abort(),10000);
        let resp;
        try{
          resp=await fetch(url,{cache:'no-store',signal:controller.signal});
        } finally {
          clearTimeout(timeoutId);
        }
        if(!resp.ok) throw new Error('HTTP '+resp.status);
        const data=await resp.json();
        const normalized=normalizeLiveRates(data);
        if(!normalized) throw new Error('Bad payload');
        fxState.rates=normalized;
        fxState.source='live';
        fxState.cacheDate=todayKey();
        try{ localStorage.setItem(FX_CACHE_PREFIX+todayKey(), JSON.stringify({rates:normalized, providerUrl:url, fetchedAt:Date.now()})); }catch(e){}
        setFxStatus();
        try{ window.dispatchEvent(new CustomEvent('margin-fx-updated')); }catch(e){}
        return;
      }catch(err){
        lastErr=err;
      }
    }
    // Both sources failed: keep whatever rates are currently loaded rather than
    // overwriting them with the stale fallback table.
    fxState.source = fxState.source==='live' || fxState.source==='cache' ? fxState.source : 'fallback';
    if(fxStatus) fxStatus.textContent=`FX: live fetch failed (${(lastErr&&lastErr.message)||lastErr}). Keeping current rates — click Update FX to retry.`;
    try{ window.dispatchEvent(new CustomEvent('margin-fx-updated')); }catch(e){}
  }

  function loadCachedFx(){
    const todayKeyStr=FX_CACHE_PREFIX+todayKey();
    // Caches now store normalized foreign->USD rates. Entries written before the
    // direction fix hold raw USD->foreign values; looksLikeForeignToUsd rejects
    // those so the tool falls back to (or refetches) correct rates.
    try{
      // Candidate keys, newest first: today's entry, then prior days descending.
      const candidates=[todayKeyStr];
      const others=[];
      for(let i=0;i<localStorage.length;i++){
        const k=localStorage.key(i);
        if(k && k.startsWith(FX_CACHE_PREFIX) && k!==todayKeyStr) others.push(k);
      }
      others.sort((a,b)=>{
        const da=a.slice(FX_CACHE_PREFIX.length), db=b.slice(FX_CACHE_PREFIX.length);
        return da<db?1:da>db?-1:0;
      });
      candidates.push(...others);

      for(const k of candidates){
        const parsed=JSON.parse(localStorage.getItem(k)||'null');
        if(parsed && parsed.rates && typeof parsed.rates==='object' && looksLikeForeignToUsd(parsed.rates)){
          fxState.rates={USD:1, ...FALLBACK_RATES, ...parsed.rates};
          fxState.source='cache';
          fxState.cacheDate=k.slice(FX_CACHE_PREFIX.length);
          return true;
        }
      }
    }catch(e){}
    return false;
  }

  // Heuristic: cached rates are foreign->USD only if a sample of major currencies
  // has values < 1. Old-bug caches (USD->foreign) have HKD/JPY/KRW-style values >> 1.
  function looksLikeForeignToUsd(rates){
    const samples=['HKD','JPY','AUD','EUR','GBP','CAD','SGD','NZD'].filter(c=>rates[c]!==undefined && rates[c]!==null);
    if(!samples.length) return false;
    const wrong=samples.filter(c=>Number(rates[c])>1.2 && (c!=='EUR' && c!=='GBP' && c!=='AUD' && c!=='NZD'));
    const strongWrong=samples.filter(c=>Number(rates[c])>2);
    // HKD/JPY/SGD/CAD must be < 1 in the correct direction; any value > 2 is definitely USD->foreign.
    return strongWrong.length===0;
  }

  // Initialize FX on load: try cache first (today's, then most recent); no network unless user clicks Update.
  if(loadCachedFx()){
    setFxStatus();
  } else {
    fxState.rates={USD:1, ...FALLBACK_RATES};
    fxState.source='fallback';
    setFxStatus();
  }

  $('updateFxBtn').addEventListener('click',()=>fetchLiveFx());

  const resetFxFallbackBtn=$('resetFxFallbackBtn');
  if(resetFxFallbackBtn) resetFxFallbackBtn.addEventListener('click',()=>{
    window.MarginFxAPI.resetToFallback();
    // Clear any manual override so the fallback table is what actually gets used.
    const input=$('fxRatesInput');
    if(input) input.value='';
  });

  // ---------- Rates panel: show exactly what a Build Summary will use ----------
  // Mirrors the aggregator's priority: a typed override in #fxRatesInput wins;
  // otherwise the live/cache/fallback table held in fxState is used.
  const FX_CHIP_ORDER=['HKD','JPY','AUD','EUR','GBP','CAD','SGD','NZD'];
  function formatFxRate(n){
    if(!Number.isFinite(n)) return '?';
    if(n===1) return '1';
    const s=n.toFixed(9).replace(/0+$/,'').replace(/\.$/,'');
    return s.length>12 ? n.toExponential(6) : s;
  }
  function resolveEffectiveFx(){
    const input=$('fxRatesInput');
    const raw=input ? input.value.trim() : '';
    if(raw){
      const parsed=window.MarginAggregatorAPI && window.MarginAggregatorAPI.parseFxRatesForDisplay
        ? window.MarginAggregatorAPI.parseFxRatesForDisplay(raw)
        : parseFxRatesLocally(raw);
      if(parsed && Object.keys(parsed).length){
        // The aggregator merges a partial override with its fallback defaults
        // for currencies not listed — mirror that here so the panel is exact.
        return {rates:{USD:1, ...FALLBACK_RATES, ...parsed}, source:'override'};
      }
    }
    return {rates:fxState.rates, source:fxState.source};
  }
  function parseFxRatesLocally(input){
    const rates={};
    String(input||'').split(/[,;\n]+/).forEach(p=>{
      const parts=p.split('=');
      if(parts.length===2){
        const code=parts[0].trim();
        const rate=parseFloat(parts[1]);
        if(code && !isNaN(rate)) rates[code.toUpperCase()]=rate;
      }
    });
    return rates;
  }
  function renderFxPanel(){
    const panel=document.getElementById('fxRatesPanel');
    if(!panel) return;
    const sourceLine=document.getElementById('fxSourceLine');
    const sourceText=document.getElementById('fxSourceText');
    const chips=document.getElementById('fxChips');
    const {rates, source}=resolveEffectiveFx();

    const sourceLabels={
      live:'Live rates fetched today',
      cache:`Cached live rates${fxState.cacheDate?` from ${fxState.cacheDate}`:''}`,
      override:'Manual override in the field above',
      fallback:`Built-in fallback table (as of ${FALLBACK_RATES_DATE})`
    };
    if(sourceLine) sourceLine.dataset.source=source;
    if(sourceText){
      const when=source==='live'?` as of ${todayKey()}`:'';
      sourceText.textContent=`FX source: ${sourceLabels[source]||source}${when} — these are the rates Build Summary will use.`;
    }
    if(chips){
      // Show the 8 known currencies in fixed order, then any extras alphabetically.
      const codes=[...FX_CHIP_ORDER.filter(c=>rates[c]!==undefined),
                   ...Object.keys(rates).filter(c=>!FX_CHIP_ORDER.includes(c) && c!=='USD').sort()];
      chips.innerHTML='';
      if(!codes.length){
        const none=document.createElement('span');
        none.className='source-hint';
        none.textContent='No rates loaded.';
        chips.appendChild(none);
      }
      codes.forEach(code=>{
        const chip=document.createElement('span');
        chip.className='fx-chip'+(source==='override'?' is-override':'');
        const codeEl=document.createElement('code');
        codeEl.textContent=code+'→USD';
        const val=document.createElement('span');
        val.textContent=formatFxRate(rates[code]);
        chip.appendChild(codeEl);
        chip.appendChild(val);
        chips.appendChild(chip);
      });
    }
  }
  // Re-render whenever anything relevant changes.
  const fxInputForPanel=$('fxRatesInput');
  if(fxInputForPanel) fxInputForPanel.addEventListener('input',renderFxPanel);
  window.addEventListener('margin-fx-updated',renderFxPanel);
  renderFxPanel();

  window.MarginFxAPI={
    getRates(){ return fxState.rates; },
    getSource(){ return fxState.source; },
    refresh(){ return fetchLiveFx(); },
    resetToFallback(){
      fxState.rates={USD:1, ...FALLBACK_RATES};
      fxState.source='fallback';
      delete fxState.cacheDate;
      setFxStatus();
      try{ window.dispatchEvent(new CustomEvent('margin-fx-updated')); }catch(e){}
    }
  };

  // Foreign FX checkbox: enable/disable the FX rate input
  $('foreignFxCheckbox').addEventListener('change',e=>{
    $('fxRate').disabled=!e.target.checked;
  });

  async function loadSharedMapping(file){
    if(!file) return;
    window.MarginAggregatorAPI.clearGeneratedSummary();
    window.MarginGeneratorAPI.setGeneratedSummary(null);
    sourceHint.textContent='Mapping changed. Rebuild Step 1, or use an existing Summary File.';
    setBox(mappingStatus,'Loading shared mapping...');
    try{
      const [aggInfo, genCount]=await Promise.all([
        window.MarginAggregatorAPI.loadSharedMapping(file),
        window.MarginGeneratorAPI.setSharedMappingFile(file)
      ]);
      setBox(mappingStatus,`Mapping loaded: ${file.name}. ${genCount} normalization key(s); ${aggInfo.pricingSkuGroups} pricing SKU alias group(s).`,'ok');
    }catch(err){
      setBox(mappingStatus,'Mapping load failed: '+(err.message||String(err)),'error');
    }
  }

  mappingInput.addEventListener('change',e=>loadSharedMapping(e.target.files&&e.target.files[0]));

  // Auto-load: if namematching.xlsx sits next to index.html, load it on startup.
  // Falls back silently to manual input when not found (404 / file:// fetch blocked).
  (async function autoLoadNameMatching(){
    if(location.protocol==='file:'){
      if(mappingStatus) setBox(mappingStatus,'Mapping: use manual input, or place namematching.xlsx next to the page and serve over http for auto-load.');
      return;
    }
    try{
      const resp=await fetch('namematching.xlsx',{cache:'no-store'});
      if(!resp.ok) return; // not found -> manual input
      const blob=await resp.blob();
      const type=blob.type||'';
      // Some servers return the 404 page as HTML with status 200; guard against that.
      if(type.includes('text/html')) return;
      const file=new File([blob],'namematching.xlsx',{type:type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      // Reflect the auto-loaded file in the native input so the UI (file card) shows it.
      try{ const dt=new DataTransfer(); dt.items.add(file); mappingInput.files=dt.files; }catch(e){}
      await loadSharedMapping(file); // sets "Mapping loaded: namematching.xlsx. N key(s)…" status
    }catch(e){
      // network/parse failure -> manual input remains available
    }
  })();

  window.addEventListener('margin-summary-generated',e=>{
    const summary=e.detail&&e.detail.summaryByRegion;
    window.MarginGeneratorAPI.setGeneratedSummary(summary);
    const regions=summary?Object.keys(summary):[];
    let uniqueRows=0;
    if(summary){
      const keys=new Set();
      regions.forEach(r=>Object.keys(summary[r]||{}).forEach(k=>keys.add(k)));
      uniqueRows=keys.size;
    }
    sourceHint.textContent=`Ready: ${regions.length} region(s), ${uniqueRows} unique Model/Grade row(s). Step 2 switched to this generated summary.`;
    setBox(step1Status,`Summary built successfully. ${regions.length} region(s), ${uniqueRows} unique Model/Grade row(s).`,'ok');
  });

  window.addEventListener('margin-summary-invalidated',()=>{
    window.MarginGeneratorAPI.setGeneratedSummary(null);
    sourceHint.textContent='Step 1 source files changed. Rebuild the summary before using the generated-summary option.';
    setBox(step1Status,'Loading Step 1 source files...');
  });

  window.addEventListener('margin-source-files-loaded',e=>{
    const d=e.detail||{};
    setBox(step1Status,`Ready to build: ${d.fileCount||0} source file(s), ${d.sheetCount||0} sheet(s) loaded. Review the detected Type and select a Region for sheets that require one, then click Build Summary.`,'ok');
  });

  window.addEventListener('margin-source-files-error',e=>{
    setBox(step1Status,(e.detail&&e.detail.message)||'Could not read the Step 1 source files.','error');
  });

  window.addEventListener('margin-summary-build-error',e=>{
    setBox(step1Status,(e.detail&&e.detail.message)||'Step 1 could not build the summary.','error');
  });

  $('processBtn').addEventListener('click',()=>{
    if(!window.MarginGeneratorAPI.isMappingReady()){
      setBox(step1Status,'Warning: no shared Mapping File is loaded. The summary can still process, but model normalization/alias expansion may be incomplete.','error');
    } else {
      setBox(step1Status,'Building summary...');
    }
  },true);

  $('summaryGenerated').addEventListener('change',()=>{
    if($('summaryGenerated').checked && !window.MarginAggregatorAPI.getGeneratedSummary()){
      sourceHint.textContent='No generated summary yet. Run Step 1 first.';
    }
  });
})();
