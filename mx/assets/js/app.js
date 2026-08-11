(() => {
  'use strict';
  const $=id=>document.getElementById(id);
  const mappingInput=$('sharedMappingFile');
  const mappingStatus=$('mappingStatus');
  const step1Status=$('step1Status');
  const sourceHint=$('generatedSummaryHint');

  function setBox(el,msg,type=''){
    el.textContent=msg;
    el.className='status-box'+(type?` ${type}`:'');
  }

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
    setBox(step1Status,'Source files loaded/changed. Build Summary to refresh Step 1 output.');
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
