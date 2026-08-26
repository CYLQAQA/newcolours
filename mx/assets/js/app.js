(() => {
  'use strict';
  const $=id=>document.getElementById(id);
  const mappingInput=$('sharedMappingFile');
  const mappingStatus=$('mappingStatus');
  const step1Status=$('step1Status');
  const sourceHint=$('generatedSummaryHint');
  const fxStatus=$('fxStatus');

  const FALLBACK_RATES = {HKD:0.1276,JPY:0.00616,AUD:0.6939,EUR:1.1432,GBP:1.3410,CAD:0.7056,SGD:0.7737,NZD:0.5703};
  const FX_CACHE_PREFIX = 'marginFx_';

  function setBox(el,msg,type=''){
    el.textContent=msg;
    el.className='status-box'+(type?` ${type}`:'');
  }

  function todayKey(){
    const d=new Date();
    const pad=n=>String(n).padStart(2,'0');
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  }

  const fxState={rates:{...FALLBACK_RATES}, source:'fallback'};

  function setFxStatus(){
    if(!fxStatus) return;
    if(fxState.source==='live'){
      fxStatus.textContent=`FX: live rates from frankfurter.app as of ${todayKey()}. Click Update FX to refresh.`;
    } else if(fxState.source==='cache'){
      const cachedDate=fxState.cacheDate || todayKey();
      const isToday=cachedDate===todayKey();
      fxStatus.textContent=`FX: cached rates from ${cachedDate}${isToday?'':' (older)'}. Click Update FX to refresh.`;
    } else {
      fxStatus.textContent=`FX: using built-in fallback table. Click Update FX to fetch live rates.`;
    }
  }

  async function fetchLiveFx(){
    if(fxStatus) fxStatus.textContent='FX: refreshing...';
    try{
      const resp=await fetch('https://api.frankfurter.app/latest?from=USD');
      if(!resp.ok) throw new Error('HTTP '+resp.status);
      const data=await resp.json();
      if(!data.rates || typeof data.rates!=='object') throw new Error('Bad payload');
      const merged={USD:1, ...FALLBACK_RATES, ...data.rates};
      // frankfurter returns rates where 1 USD = N foreign. We want USD->foreign, which is what we have.
      fxState.rates=merged;
      fxState.source='live';
      fxState.cacheDate=todayKey();
      try{ localStorage.setItem(FX_CACHE_PREFIX+todayKey(), JSON.stringify({rates:data.rates, fetchedAt:Date.now()})); }catch(e){}
      setFxStatus();
    }catch(err){
      fxState.rates={...FALLBACK_RATES};
      fxState.source='fallback';
      if(fxStatus) fxStatus.textContent=`FX: live fetch failed (${err.message||err}). Using built-in fallback table.`;
    }
  }

  function loadCachedFx(){
    const key=FX_CACHE_PREFIX+todayKey();
    try{
      const raw=localStorage.getItem(key);
      if(raw){
        const parsed=JSON.parse(raw);
        if(parsed && parsed.rates && typeof parsed.rates==='object'){
          fxState.rates={USD:1, ...FALLBACK_RATES, ...parsed.rates};
          fxState.source='cache';
          fxState.cacheDate=todayKey();
          return true;
        }
      }
      // Also try the most recent cache from prior days
      let best=null;
      for(let i=0;i<localStorage.length;i++){
        const k=localStorage.key(i);
        if(k && k.startsWith(FX_CACHE_PREFIX)){
          const date=k.slice(FX_CACHE_PREFIX.length);
          if(!best || date>best.date) best={date,k};
        }
      }
      if(best){
        const parsed=JSON.parse(localStorage.getItem(best.k));
        if(parsed && parsed.rates){
          fxState.rates={USD:1, ...FALLBACK_RATES, ...parsed.rates};
          fxState.source='cache';
          fxState.cacheDate=best.date;
          return true;
        }
      }
    }catch(e){}
    return false;
  }

  // Initialize FX on load: try cache first (today's, then most recent); no network unless user clicks Update.
  if(loadCachedFx()){
    setFxStatus();
  } else {
    fxState.rates={...FALLBACK_RATES};
    fxState.source='fallback';
    setFxStatus();
  }

  $('updateFxBtn').addEventListener('click',()=>fetchLiveFx());

  window.MarginFxAPI={
    getRates(){ return fxState.rates; },
    getSource(){ return fxState.source; },
    refresh(){ return fetchLiveFx(); }
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
