// ---------- Helpers ----------
    function normalizeKey(str){
      if(!str) return '';
      return String(str).trim().replace(/\s+/g,' ').replace(/\s+-\s+/g,'-').toLowerCase();
    }
    function toNumber(value){
      if(value===undefined||value===null) return NaN;
      let s=String(value).trim();
      if(!s) return NaN;
      let neg=false;
      if(s.startsWith('(')&&s.endsWith(')')){neg=true; s=s.slice(1,-1);}
      s=s.replace(/,/g,'');
      let n=parseFloat(s);
      if(isNaN(n)) return NaN;
      return neg?-n:n;
    }
    function parseFxRates(input){
      const rates={};
      if(input){
        input.split(/[,;\n]+/).forEach(p=>{
          const parts=p.split('=');
          if(parts.length===2){
            const code=parts[0].trim();
            const rate=parseFloat(parts[1]);
            if(code && !isNaN(rate)) rates[code.toUpperCase()]=rate;
          }
        });
      }
      const defaults={HKD:0.1276,JPY:0.00616,AUD:0.6939,EUR:1.1432,GBP:1.3410,CAD:0.7056,SGD:0.7737,NZD:0.5703};
      Object.keys(defaults).forEach(k=>{ if(!rates[k]) rates[k]=defaults[k]; });
      return rates;
    }

    // Country Type -> Region mapping for Type 2 master (with IE)
    const countryTypeToRegionMap = {
      'US':'US',
      'IE MVA':'IE',
      'CA':'CA',
      'UK MVA':'UK',
      'UK SVA':'UK',
      'HK':'HK',
      'SCHJP':'JP',
      'DUB':'DUB',
      'AU':'AU',
      'IE SVA':'IE',
      'JP':'JP',
      'DE':'IE',
      'AT':'IGNORE',
      'CH':'IGNORE',
      'HK IE':'IGNORE',
      'NL IE':'IGNORE',
      'SCHKR':'IGNORE',
      'SCHTW':'IGNORE',
      'AU IE':'IGNORE'
    };
    function mapCountryTypeToRegion(val){
      if(val===undefined || val===null) return null;
      const key=String(val).trim().toUpperCase();
      if(!key) return null;
      const mapped=countryTypeToRegionMap[key];
      if(!mapped || mapped.toUpperCase()==='IGNORE') return null;
      return mapped;
    }

    // Sheet name -> Region mapping for Last ASP master (your mapping)
    const sheetNameToRegionMap = {
      'JAPAN':'JP',
      'HONG KONG':'HK',
      'HONGKONG':'HK',
      'AUSTRALIA':'AU',
      'UNITED STATES':'US',
      'USA':'US',
      'US':'US',
      'IRELAND':'IE',
      'UNITED KINGDOM':'UK',
      'UK':'UK',
      'UNITED ARAB EMIRATES':'DUB',
      'UAE':'DUB',
      'DUBAI':'DUB',
      'DUBLIN':'DUB',
      'CANADA':'CA',
      'GERMANY':'IE'   // DE bucket into IRE
    };
    function mapSheetNameToRegion(sheetName){
      if(!sheetName) return null;
      const s=String(sheetName).trim();
      if(!s) return null;
      const up=s.toUpperCase();
      if(sheetNameToRegionMap[up]) return sheetNameToRegionMap[up];
      const simple=up.replace(/\s+/g,'');
      const known=['AU','CA','US','UK','IE','JP','HK','DUB'];
      if(known.includes(simple)) return simple;
      return null;
    }

    const sheetsInfo=[]; // {fileName, sheetName, data}
    const referenceMap={}; let referenceLoaded=false;
    const nameMapping={}; let nameMappingLoaded=false;
    const nameToPricingSkus={}; // normalized base name -> pricing_sku list, for Type 2 rows without pricing_sku
    const aliasExpansionLog=[];

    // Last ASP master: region -> key -> {aspUSD, lastSoldDay}
    const lastAspByRegion={}; let lastAspLoaded=false;
    let generatedSummaryByRegion=null;
    let generatedFxRates=null;

    function normalizeForRef(str){
      if(!str) return '';
      return String(str).trim().replace(/\s+/g,' ').replace(/\s*-\s*/g,'-').toLowerCase();
    }

    function splitModelGradeString(value){
      if(value===undefined || value===null) return {base:'', grade:''};
      const raw=String(value).trim().replace(/\s+/g,' ').replace(/\s+-\s+/g,'-');
      const hy=raw.lastIndexOf('-');
      if(hy<=0) return {base:raw, grade:''};
      return {base:raw.slice(0,hy).trim(), grade:raw.slice(hy+1).trim()};
    }

    function buildCanonicalModelGrade(rawValue){
      if(rawValue===undefined || rawValue===null || rawValue==='') return '';
      const rawModel=(typeof rawValue==='string'?rawValue.trim():String(rawValue)).replace(/\s+/g,' ');
      if(!rawModel) return '';
      const parts=splitModelGradeString(rawModel);
      let canonBase=parts.base;
      if(referenceLoaded){
        const n=normalizeForRef(parts.base);
        if(n && referenceMap[n]) canonBase=referenceMap[n];
      }
      return parts.grade?`${canonBase}-${parts.grade}`:canonBase;
    }

    function normalizeCapacityToken(name){
      const caps=[...String(name||'').matchAll(/(\d+(?:\.\d+)?)\s*(tb|gb)\b/ig)].map(m=>{
        const n=parseFloat(m[1]);
        const unit=m[2].toLowerCase();
        return unit==='tb' ? String(n*1024)+'GB' : String(n)+'GB';
      });
      return [...new Set(caps)].sort().join('|');
    }

    function modelSignature(name){
      const s=String(name||'').toLowerCase().replace(/\s+/g,' ').trim();

      // iPhone, including SE generation so SE 2nd/3rd Gen are never treated as the same model.
      let m=s.match(/\biphone\s*se(?:\s*(\d+)(?:st|nd|rd|th)?\s*gen)?/i);
      if(m) return 'iphone-se-'+(m[1]||'unknown');
      m=s.match(/\biphone\s*(x|xr|xs|\d{1,2})(?:\s*(pro\s*max|pro|max|mini|plus))?/i);
      if(m) return 'iphone-'+m[1].toLowerCase()+'-'+(m[2]||'').replace(/\s+/g,'');

      // Apple Watch aliases such as "Apple Watch Ultra 2" and "Watch Series Ultra 2".
      m=s.match(/\b(?:apple\s+)?watch(?:\s+series)?\s+ultra\s*(\d+)?(?:\s+(\d{2})\s*mm)?/i);
      if(m) return 'watch-ultra-'+(m[1]||'')+'-'+(m[2]||'');
      m=s.match(/\b(?:apple\s+)?watch(?:\s+series)?\s*(\d+|se)(?:\s+(\d{2})\s*mm)?/i);
      if(m) return 'watch-series-'+(m[1]||'')+'-'+(m[2]||'');

      m=s.match(/\b(?:samsung\s+)?(?:galaxy\s+)?s(\d{1,2})(?:\s*(ultra|plus|\+|fe))?\b/i);
      if(m) return 'galaxy-s'+m[1]+'-'+((m[2]||'').replace('+','plus'));
      m=s.match(/\bpixel\s*(\d+)?(?:\s*(xl|pro|a))?\b/i);
      if(m && (m[1] || m[2])) return 'pixel-'+(m[1]||'')+'-'+(m[2]||'');
      m=s.match(/\bipad(?:\s+(air|mini|pro))?\s*(\d+)?/i);
      if(m && (m[1] || m[2])) return 'ipad-'+(m[1]||'base')+'-'+(m[2]||'');
      return '';
    }

    function aliasGroupLooksSafe(baseNames){
      const names=(baseNames||[]).map(v=>String(v||'').trim()).filter(Boolean);
      if(names.length<=1) return true;
      const capacities=[...new Set(names.map(normalizeCapacityToken).filter(Boolean))];
      if(capacities.length>1) return false;
      const sigs=[...new Set(names.map(modelSignature).filter(Boolean))];
      if(sigs.length>1) return false;
      return true;
    }

    function uniqueArray(values){
      return [...new Set((values||[]).map(v=>String(v||'').trim()).filter(Boolean))];
    }

    function addUnique(target, value){
      if(!value) return;
      if(!target.includes(value)) target.push(value);
    }

    function mergeUniqueArrays(a,b){
      const out=uniqueArray(a);
      uniqueArray(b).forEach(v=>addUnique(out,v));
      return out;
    }

    function pricingSkuConnectivityType(pricingSku){
      const ps=String(pricingSku||'').toUpperCase();
      if(/-WIF(?:-|$)/.test(ps)) return 'WIF';
      if(/-(?:FAC|LBT|LBX)(?:-|$)/.test(ps)) return 'FAC';
      return '';
    }

    function resolvePricingSkuCandidatesForName(name){
      if(name===undefined || name===null) return [];
      const raw=String(name).trim().replace(/\s+/g,' ');
      if(!raw) return [];

      const found=[];
      const direct=normalizeKey(raw);
      if(nameToPricingSkus[direct]) nameToPricingSkus[direct].forEach(ps=>addUnique(found,ps));

      if(referenceLoaded){
        const refKey=normalizeForRef(raw);
        const refName=refKey && referenceMap[refKey] ? referenceMap[refKey] : '';
        if(refName){
          const refNorm=normalizeKey(refName);
          if(nameToPricingSkus[refNorm]) nameToPricingSkus[refNorm].forEach(ps=>addUnique(found,ps));
        }
      }

      if(found.length<=1) return found;

      // Connectivity-aware ambiguity resolution.
      // A generic iPad name can exist under both WIF and FAC in Name Mapping.
      // Explicit "WiFi & Cellular" / "Cellular" belongs to FAC; explicit WiFi belongs to WIF.
      // If connectivity is omitted and exactly one WIF candidate exists alongside FAC, prefer WIF
      // (e.g. "iPad Mini 6 64GB" is the short alias of the WiFi pricing SKU).
      const low=raw.toLowerCase();
      const hasCellular=/\bcellular\b/.test(low);
      const hasWifi=/\bwi\s*-?\s*fi\b|\bwifi\b/.test(low);
      const wif=found.filter(ps=>pricingSkuConnectivityType(ps)==='WIF');
      const fac=found.filter(ps=>pricingSkuConnectivityType(ps)==='FAC');

      if(hasCellular && fac.length===1) return fac;
      if(hasWifi && !hasCellular && wif.length===1) return wif;
      if(!hasWifi && !hasCellular && wif.length===1 && fac.length>=1) return wif;

      // Anything still ambiguous is intentionally blocked from reverse expansion.
      return [];
    }

    function findPricingSkusForBaseNames(baseNames){
      const found=[];
      (baseNames||[]).forEach(name=>{
        resolvePricingSkuCandidatesForName(name).forEach(ps=>addUnique(found,ps));
      });
      return found;
    }

    function hasNumber(value){
      return value!==undefined && value!==null && !isNaN(value);
    }

    function fieldIsBlank(value){
      return value===undefined || value===null || value==='' || (typeof value==='number' && isNaN(value));
    }

    function snapshotRecord(rec){
      const out=Object.assign({}, rec||{});
      out.sourcePricingSkus=uniqueArray((rec&&rec.sourcePricingSkus)||[]);
      return out;
    }

    function chooseCostAliasSource(candidates){
      let best=null;
      (candidates||[]).forEach(c=>{
        if(!c || !c.rec || !hasNumber(c.rec.costUSD)) return;
        if(!best || c.rec.costUSD>best.rec.costUSD) best=c;
      });
      return best;
    }

    function soldFreshness(rec){
      if(!rec || rec.lastSoldDay===undefined || rec.lastSoldDay===null || rec.lastSoldDay==='') return null;
      const v=rec.lastSoldDay;
      if(typeof v==='number' && !isNaN(v)) return {kind:'days', value:v}; // lower = fresher
      const t=Date.parse(String(v));
      if(!isNaN(t)) return {kind:'date', value:t}; // higher = fresher
      return null;
    }

    function chooseAspAliasSource(candidates){
      const usable=(candidates||[]).filter(c=>c && c.rec && hasNumber(c.rec.aspUSD));
      if(usable.length===0) return null;

      // Last ASP master is the strongest ASP source when present.
      const master=usable.filter(c=>c.rec.aspFromLastMaster===true);
      const pool=master.length>0 ? master : usable;
      let best=null;
      pool.forEach(c=>{
        if(!best){ best=c; return; }
        const cf=soldFreshness(c.rec), bf=soldFreshness(best.rec);
        if(cf && bf && cf.kind===bf.kind){
          if(cf.kind==='days' && cf.value<bf.value){ best=c; return; }
          if(cf.kind==='date' && cf.value>bf.value){ best=c; return; }
        } else if(cf && !bf){
          best=c; return;
        }
        if(c.rec.aspUSD>best.rec.aspUSD) best=c;
      });
      return best;
    }

    function createAliasRecord(aliasModelGrade, pricingSku, costSource, aspSource){
      const rec={
        costUSD:undefined, costCurrency:'USD', costVal:undefined,
        aspUSD:undefined, aspLocal:undefined, aspCurrency:'USD',
        lastSoldDay:undefined,
        originalKey:aliasModelGrade,
        aliasExpanded:true,
        aliasSourcePricingSku:pricingSku,
        sourcePricingSkus:[pricingSku]
      };
      if(costSource && costSource.rec){
        rec.costUSD=costSource.rec.costUSD;
        rec.costCurrency=costSource.rec.costCurrency;
        rec.costVal=costSource.rec.costVal;
      }
      if(aspSource && aspSource.rec){
        rec.aspUSD=aspSource.rec.aspUSD;
        rec.aspLocal=aspSource.rec.aspLocal;
        rec.aspCurrency=aspSource.rec.aspCurrency;
        rec.lastSoldDay=aspSource.rec.lastSoldDay;
        rec.aspFromLastMaster=aspSource.rec.aspFromLastMaster===true;
      }
      return rec;
    }

    function fillMissingAliasFields(targetRec, pricingSku, costSource, aspSource){
      const filled=[];
      if(costSource && costSource.rec && fieldIsBlank(targetRec.costUSD) && hasNumber(costSource.rec.costUSD)){
        targetRec.costUSD=costSource.rec.costUSD;
        targetRec.costCurrency=costSource.rec.costCurrency;
        targetRec.costVal=costSource.rec.costVal;
        filled.push('C+Load');
      }
      if(aspSource && aspSource.rec){
        if(fieldIsBlank(targetRec.aspUSD) && hasNumber(aspSource.rec.aspUSD)){
          targetRec.aspUSD=aspSource.rec.aspUSD;
          targetRec.aspLocal=aspSource.rec.aspLocal;
          targetRec.aspCurrency=aspSource.rec.aspCurrency;
          if(aspSource.rec.aspFromLastMaster===true) targetRec.aspFromLastMaster=true;
          filled.push('Last ASP');
        }
        if(fieldIsBlank(targetRec.lastSoldDay) && !fieldIsBlank(aspSource.rec.lastSoldDay)){
          targetRec.lastSoldDay=aspSource.rec.lastSoldDay;
          filled.push('Last Sold Day');
        }
      }
      targetRec.sourcePricingSkus=mergeUniqueArrays(targetRec.sourcePricingSkus||[], [pricingSku]);
      return filled;
    }

    function applyAliasExpansion(summaryByRegion, enableExpansion, guardConflicts){
      aliasExpansionLog.length=0;
      if(!enableExpansion || !nameMappingLoaded) return;

      Object.keys(summaryByRegion).forEach(region=>{
        const regionMap=summaryByRegion[region];
        if(!regionMap) return;

        // IMPORTANT: sourceSnapshot is immutable for this expansion pass.
        // Newly created aliases and fields filled during this function can never become sources
        // for another pricing SKU, which prevents cascading A->B->C alias propagation.
        const sourceSnapshot={};
        Object.keys(regionMap).forEach(k=>{ sourceSnapshot[k]=snapshotRecord(regionMap[k]); });

        // Current output index is used only to find rows to fill / avoid duplicates.
        const currentBaseGradeIndex={};
        Object.keys(regionMap).forEach(k=>{
          const rec=regionMap[k]||{};
          const parts=splitModelGradeString(rec.originalKey || k);
          if(!parts.base || !parts.grade) return;
          const b=normalizeKey(parts.base), g=normalizeKey(parts.grade);
          if(!currentBaseGradeIndex[b]) currentBaseGradeIndex[b]={};
          currentBaseGradeIndex[b][g]=k;
        });

        // Build real-source candidates by confirmed pricing SKU + grade.
        const sourceByPricingSkuGrade={};
        Object.keys(sourceSnapshot).forEach(k=>{
          const rec=sourceSnapshot[k]||{};
          const parts=splitModelGradeString(rec.originalKey || k);
          if(!parts.base || !parts.grade) return;
          const gradeNorm=normalizeKey(parts.grade);

          let psList=uniqueArray(rec.sourcePricingSkus||[]);
          if(rec.aliasSourcePricingSku) addUnique(psList, String(rec.aliasSourcePricingSku).trim().toUpperCase());
          if(psList.length===0) psList=findPricingSkusForBaseNames([parts.base]);

          psList.forEach(ps=>{
            const psKey=String(ps||'').trim().toUpperCase();
            if(!psKey) return;
            if(!sourceByPricingSkuGrade[psKey]) sourceByPricingSkuGrade[psKey]={};
            if(!sourceByPricingSkuGrade[psKey][gradeNorm]) sourceByPricingSkuGrade[psKey][gradeNorm]=[];
            sourceByPricingSkuGrade[psKey][gradeNorm].push({key:k, alias:parts.base, rec});
          });
        });

        Object.keys(nameMapping).forEach(pricingSku=>{
          const allAliases=uniqueArray(nameMapping[pricingSku]);

          // An alias appearing under multiple pricing SKUs is included only when our resolver
          // confirms this pricing SKU. This is what stops generic iPad names bridging WIF and FAC.
          const aliases=allAliases.filter(alias=>resolvePricingSkuCandidatesForName(alias).includes(pricingSku));
          if(aliases.length<2) return;

          if(guardConflicts && !aliasGroupLooksSafe(aliases)){
            const hasAnySource=sourceByPricingSkuGrade[pricingSku] && Object.keys(sourceByPricingSkuGrade[pricingSku]).length>0;
            if(hasAnySource){
              aliasExpansionLog.push({
                Action:'Skipped - possible conflict', Region:region, Pricing_SKU:pricingSku, Grade:'',
                Alias_Name:'', Filled_From:'', Cost_USD:'', Fields_Filled:'', Names:aliases.join(' | ')
              });
            }
            return;
          }

          const byGrade=sourceByPricingSkuGrade[pricingSku]||{};
          Object.keys(byGrade).forEach(gradeNorm=>{
            const sources=byGrade[gradeNorm]||[];
            if(sources.length===0) return;

            const costSource=chooseCostAliasSource(sources);
            const aspSource=chooseAspAliasSource(sources);
            if(!costSource && !aspSource) return;

            const gradeSource=(costSource||aspSource||sources[0]);
            const gradeParts=splitModelGradeString((gradeSource.rec&&gradeSource.rec.originalKey) || gradeSource.alias+'-'+gradeNorm);
            const gradeDisplay=gradeParts.grade || gradeNorm;

            aliases.forEach(alias=>{
              const baseNorm=normalizeKey(alias);
              const existingKey=currentBaseGradeIndex[baseNorm] && currentBaseGradeIndex[baseNorm][gradeNorm];

              if(existingKey && regionMap[existingKey]){
                const filled=fillMissingAliasFields(regionMap[existingKey], pricingSku, costSource, aspSource);
                if(filled.length>0){
                  aliasExpansionLog.push({
                    Action:'Filled missing alias fields', Region:region, Pricing_SKU:pricingSku, Grade:gradeDisplay,
                    Alias_Name:regionMap[existingKey].originalKey || `${alias}-${gradeDisplay}`,
                    Filled_From:[costSource&&costSource.rec ? (costSource.rec.originalKey||costSource.alias) : '', aspSource&&aspSource.rec ? (aspSource.rec.originalKey||aspSource.alias) : ''].filter(Boolean).filter((v,i,a)=>a.indexOf(v)===i).join(' | '),
                    Cost_USD:(costSource&&hasNumber(costSource.rec.costUSD)) ? Number(costSource.rec.costUSD).toFixed(2) : '',
                    Fields_Filled:filled.join(' | '), Names:aliases.join(' | ')
                  });
                }
                return;
              }

              const aliasModelGrade=`${alias}-${gradeDisplay}`;
              const newKey=normalizeKey(aliasModelGrade);
              if(regionMap[newKey]) return;
              regionMap[newKey]=createAliasRecord(aliasModelGrade, pricingSku, costSource, aspSource);

              if(!currentBaseGradeIndex[baseNorm]) currentBaseGradeIndex[baseNorm]={};
              currentBaseGradeIndex[baseNorm][gradeNorm]=newKey;

              aliasExpansionLog.push({
                Action:'Added missing alias', Region:region, Pricing_SKU:pricingSku, Grade:gradeDisplay,
                Alias_Name:aliasModelGrade,
                Filled_From:[costSource&&costSource.rec ? (costSource.rec.originalKey||costSource.alias) : '', aspSource&&aspSource.rec ? (aspSource.rec.originalKey||aspSource.alias) : ''].filter(Boolean).filter((v,i,a)=>a.indexOf(v)===i).join(' | '),
                Cost_USD:(costSource&&hasNumber(costSource.rec.costUSD)) ? Number(costSource.rec.costUSD).toFixed(2) : '',
                Fields_Filled:['C+Load','Last ASP','Last Sold Day'].filter(f=>{
                  if(f==='C+Load') return costSource&&hasNumber(costSource.rec.costUSD);
                  if(f==='Last ASP') return aspSource&&hasNumber(aspSource.rec.aspUSD);
                  return aspSource&&!fieldIsBlank(aspSource.rec.lastSoldDay);
                }).join(' | '),
                Names:aliases.join(' | ')
              });
            });
          });
        });
      });
    }

    function loadReference(file){
      return new Promise((resolve,reject)=>{
        const reader=new FileReader();
        reader.onload=e=>{
          try{
            const data=new Uint8Array(e.target.result);
            const wb=XLSX.read(data,{type:'array'});
            const ws=wb.Sheets[wb.SheetNames[0]];
            const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
            rows.forEach(row=>{
              let canonical='';
              Object.keys(row).forEach(col=>{
                const low=col.toLowerCase();
                if(low==='en name'||low==='en_name'||low==='enname'){ canonical=row[col]; }
              });
              if(!canonical) return;
              const normCanon=normalizeForRef(canonical);
              if(normCanon) referenceMap[normCanon]=canonical;
              const fields=['modelcapcity','model capacity','bi name','en name','en_name','enname','pricing sku','pricingsku','sku'];
              fields.forEach(field=>{
                Object.keys(row).forEach(col=>{
                  if(col.toLowerCase().replace(/\s+/g,'')===field.replace(/\s+/g,'')){
                    const v=row[col];
                    const n=normalizeForRef(v);
                    if(n) referenceMap[n]=canonical;
                  }
                });
              });
            });
            referenceLoaded=true; resolve();
          }catch(err){ reject(err); }
        };
        reader.onerror=reject;
        reader.readAsArrayBuffer(file);
      });
    }

    function loadNameMapping(file){
      return new Promise((resolve,reject)=>{
        const reader=new FileReader();
        reader.onload=e=>{
          try{
            const data=new Uint8Array(e.target.result);
            const wb=XLSX.read(data,{type:'array'});
            const ws=wb.Sheets[wb.SheetNames[0]];
            const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
            Object.keys(nameMapping).forEach(k=>{ delete nameMapping[k]; });
            Object.keys(nameToPricingSkus).forEach(k=>{ delete nameToPricingSkus[k]; });
            rows.forEach(row=>{
              let skuVal=''; let enVal='';
              Object.keys(row).forEach(col=>{
                const low=col.toLowerCase().replace(/\s+/g,'');
                if(low==='pricing_sku' || low==='pricingsku'){
                  skuVal=row[col];
                }
                if(low==='enname' || low==='en_name' || low==='en name'){
                  enVal=row[col];
                }
              });
              if(skuVal!==undefined && skuVal!==null && skuVal!==''){
                const key=String(skuVal).trim().toUpperCase();
                if(key){
                  const nameStr=(enVal!==undefined && enVal!==null)?String(enVal).trim():'';
                  if(nameStr){
                    if(nameMapping[key]){
                      const arr=nameMapping[key];
                      if(!arr.includes(nameStr)) arr.push(nameStr);
                    } else {
                      nameMapping[key]=[nameStr];
                    }
                    const nameNorm=normalizeKey(nameStr);
                    if(nameNorm){
                      if(!nameToPricingSkus[nameNorm]) nameToPricingSkus[nameNorm]=[];
                      if(!nameToPricingSkus[nameNorm].includes(key)) nameToPricingSkus[nameNorm].push(key);
                    }
                  }
                }
              }
            });
            nameMappingLoaded=true;
            resolve();
          }catch(err){ reject(err); }
        };
        reader.onerror=reject;
        reader.readAsArrayBuffer(file);
      });
    }

    // Last ASP master: each tab:
    // col0 = Model/Grade, col1 = Last ASP (USD), col3 = Last Sold Day (string)
    function loadLastAspMaster(file){
      return new Promise((resolve,reject)=>{
        const reader=new FileReader();
        reader.onload=e=>{
          try{
            Object.keys(lastAspByRegion).forEach(r=>{ delete lastAspByRegion[r]; });
            const data=new Uint8Array(e.target.result);
            const wb=XLSX.read(data,{type:'array'});

            wb.SheetNames.forEach(sheetName=>{
              const region=mapSheetNameToRegion(sheetName);
              if(!region) return;

              const ws=wb.Sheets[sheetName];
              const arr=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
              if(!arr || arr.length===0) return;

              const dataRows=arr.slice(1);
              if(!lastAspByRegion[region]) lastAspByRegion[region]={};

              dataRows.forEach(rvals=>{
                const mg = rvals[0];
                const aspVal = toNumber(rvals[1]);

                // Last Sold Day – keep as full string (e.g. "2025-09-15")
                let lastSold;
                if (rvals[3] !== undefined && rvals[3] !== null) {
                  const raw = String(rvals[3]).trim();
                  lastSold = raw === '' ? undefined : raw;
                }

                if (!mg || isNaN(aspVal)) return;

                let rawMG = (typeof mg === 'string' ? mg.trim() : String(mg)).replace(/\s+/g,' ');
                let base = rawMG, gradePart = '';
                const hy = rawMG.lastIndexOf('-');
                if (hy > 0) {
                  base = rawMG.slice(0, hy).trim();
                  gradePart = rawMG.slice(hy + 1).trim();
                }
                let canonicalBase = base;
                if (referenceLoaded) {
                  const n = normalizeForRef(base);
                  if (n && referenceMap[n]) canonicalBase = referenceMap[n];
                }
                const canonicalName = gradePart ? `${canonicalBase}-${gradePart}` : canonicalBase;
                const key = normalizeKey(canonicalName);

                const existing = lastAspByRegion[region][key];
                if (!existing || aspVal > existing.aspUSD) {
                  lastAspByRegion[region][key] = {
                    aspUSD: aspVal,
                    lastSoldDay: lastSold
                  };
                }
              });
            });

            lastAspLoaded=true;
            resolve();
          }catch(err){ reject(err); }
        };
        reader.onerror=reject;
        reader.readAsArrayBuffer(file);
      });
    }

    function handleFiles(files){
      const fileList=Array.from(files||[]);
      generatedSummaryByRegion=null;
      generatedFxRates=null;
      window._generatedTable=null;
      const exportBtn=document.getElementById('downloadBtn');
      const processBtn=document.getElementById('processBtn');
      if(exportBtn) exportBtn.disabled=true;
      if(processBtn) processBtn.disabled=true;
      window.dispatchEvent(new CustomEvent('margin-summary-invalidated',{detail:{reason:'source-change'}}));
      sheetsInfo.length=0; const tasks=[];
      fileList.forEach(file=>{
        tasks.push(new Promise((resolve,reject)=>{
          const reader=new FileReader();
          reader.onload=e=>{
            try{
              const data=new Uint8Array(e.target.result);
              const wb=XLSX.read(data,{type:'array'});
              wb.SheetNames.forEach(sheetName=>{
                const ws=wb.Sheets[sheetName];
                const arr=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
                sheetsInfo.push({fileName:file.name,sheetName,data:arr});
              });
              resolve();
            }catch(err){ reject(err); }
          };
          reader.onerror=reject;
          reader.readAsArrayBuffer(file);
        }));
      });
      Promise.all(tasks).then(()=>{
        buildMappingTable();
        if(processBtn) processBtn.disabled=sheetsInfo.length===0;
        window.dispatchEvent(new CustomEvent('margin-source-files-loaded',{
          detail:{fileCount:fileList.length,sheetCount:sheetsInfo.length}
        }));
      }).catch(err=>{
        const msg='Error reading source files: '+(err.message||String(err));
        window.dispatchEvent(new CustomEvent('margin-source-files-error',{detail:{message:msg}}));
        alert(msg);
      });
    }

    function buildMappingTable(){
      const tbody=document.getElementById('mappingTableBody'); tbody.innerHTML='';
      const skipFirstTwo=document.getElementById('skipRowsCheckbox').checked;
      sheetsInfo.forEach((item,idx)=>{
        const headerRow=skipFirstTwo?item.data[2]:item.data[0];
        const headers=(headerRow||[]).map(h=>h!==undefined&&h!==null?String(h).trim():'');
        let detectedType='';
        const lower=headers.map(h=>h.toLowerCase().replace(/\s+/g,''));
        if(lower.includes('model/grade') && lower.some(h=>/cpu\+?load/.test(h)) && lower.some(h=>/lastasp/.test(h))){
          detectedType='Type 1';
        }
        else if(lower.includes('modelcapacity') && lower.includes('grade') && lower.some(h=>/netcpu\+load/.test(h))){
          detectedType='Type 2';
        }
        else {
          detectedType='Unknown';
        }
        const tr=document.createElement('tr'); tr.dataset.index=idx;
        tr.innerHTML=`
          <td class="px-3 py-2 text-sm">${item.fileName}</td>
          <td class="px-3 py-2 text-sm">${item.sheetName}</td>
          <td class="px-3 py-2 text-sm">${detectedType}</td>
          <td class="px-3 py-2 text-sm">
            <select class="type-select border-gray-300 rounded p-1">
              <option value="Type 1" ${detectedType==='Type 1'?'selected':''}>Type 1</option>
              <option value="Type 2" ${detectedType==='Type 2'?'selected':''}>Type 2</option>
              <option value="Unknown" ${detectedType==='Unknown'?'selected':''}>Unknown</option>
            </select>
          </td>
          <td class="px-3 py-2 text-sm">
            <select class="region-select border-gray-300 rounded p-1">
              <option value="">Select</option>
              <option value="HK">HK</option>
              <option value="JP">JP</option>
              <option value="US">US</option>
              <option value="AU">AU</option>
              <option value="IE">IE</option>
              <option value="UK">UK</option>
              <option value="DUB">DUB</option>
              <option value="CA">CA</option>
            </select>
          </td>
          <td class="px-3 py-2 text-center"><input type="checkbox" class="include-checkbox" checked></td>`;
        tbody.appendChild(tr);
      });
      document.getElementById('mappingSection').classList.remove('hidden');
    }

    // ---------- Main processing ----------
    document.getElementById('processBtn').addEventListener('click',()=>{
      try {
      if(sheetsInfo.length===0){
        const msg='Please upload at least one source file first.';
        window.dispatchEvent(new CustomEvent('margin-summary-build-error',{detail:{message:msg}}));
        alert(msg);
        return;
      }
      const skipFirstTwo=document.getElementById('skipRowsCheckbox').checked;
      const aspAggMethod=document.getElementById('aspMethod').value || 'max';
      const fxInputVal=document.getElementById('fxRatesInput').value.trim();
      const fxRates=fxInputVal ? parseFxRates(fxInputVal) : (window.MarginFxAPI ? window.MarginFxAPI.getRates() : parseFxRates(''));

      const rows=[...document.querySelectorAll('#mappingTableBody tr')];
      const includedRows=rows.filter(row=>row.querySelector('.include-checkbox').checked);
      const missingRegions=[];
      const unknownTypes=[];

      includedRows.forEach(row=>{
        const idx=parseInt(row.dataset.index,10);
        const item=sheetsInfo[idx];
        const type=row.querySelector('.type-select').value;
        const region=row.querySelector('.region-select').value;
        const headerRow=skipFirstTwo?item.data[2]:item.data[0];
        const headers=(headerRow||[]).map(h=>h!==undefined&&h!==null?String(h).trim():'');
        const hasCountryType=headers.some(h=>h && h.toLowerCase().replace(/\s+/g,'')==='countrytype');
        const label=`${item.fileName} / ${item.sheetName}`;
        if(type==='Unknown') unknownTypes.push(label);
        else if((type!=='Type 2' || !hasCountryType) && !region) missingRegions.push(label);
      });

      if(includedRows.length===0){
        const msg='No sheets are included. Tick at least one Include checkbox before building.';
        window.dispatchEvent(new CustomEvent('margin-summary-build-error',{detail:{message:msg}}));
        alert(msg);
        return;
      }
      if(unknownTypes.length){
        const msg=`${unknownTypes.length} included sheet(s) still have Type = Unknown. Select Type 1/Type 2 or untick Include:\n- ${unknownTypes.slice(0,8).join('\n- ')}${unknownTypes.length>8?'\n- ...':''}`;
        window.dispatchEvent(new CustomEvent('margin-summary-build-error',{detail:{message:msg}}));
        alert(msg);
        return;
      }
      if(missingRegions.length){
        const msg=`Select a Region for ${missingRegions.length} included sheet(s) before building:\n- ${missingRegions.slice(0,8).join('\n- ')}${missingRegions.length>8?'\n- ...':''}`;
        window.dispatchEvent(new CustomEvent('margin-summary-build-error',{detail:{message:msg}}));
        alert(msg);
        return;
      }

      const summaryByRegion={};
      const aggregatorType1={};

      rows.forEach(row=>{
        const idx=parseInt(row.dataset.index,10);
        const include=row.querySelector('.include-checkbox').checked;
        if(!include) return;

        const type=row.querySelector('.type-select').value;
        const regionSel=row.querySelector('.region-select').value;

        const item=sheetsInfo[idx];
        const data=item.data;
        const headerRow=skipFirstTwo?data[2]:data[0];
        const dataRows=skipFirstTwo?data.slice(3):data.slice(1);
        const headers=(headerRow||[]).map(h=>h!==undefined&&h!==null?String(h).trim():'');
        const headerMap={}; headers.forEach((h,i)=>headerMap[h]=i);

        const hasCountryType=headers.some(h=>h && h.toLowerCase().replace(/\s+/g,'')==='countrytype');

        let region=regionSel;

        if(type!=='Type 2' || !hasCountryType){
          if(!region) return;
          if(!summaryByRegion[region]) summaryByRegion[region]={};
        }

        if(type==='Type 1'){
          let modelIdx=-1,costIdx=-1,aspIdx=-1,soldIdx=-1;
          let pricingSkuIdx=-1, gradeTransIdx=-1;
          let costCurrency='USD',aspCurrency='USD';
          let fallbackCostIdx=-1,fallbackCostCurrency='USD';
          headers.forEach((h,i)=>{
            const lower=h.toLowerCase();
            const lowerNoSpace=lower.replace(/\s+/g,'');
            if(lower==='model/grade') modelIdx=i;
            if(/cpu\s*\+?\s*load/i.test(lower)){
              let curr='USD'; const m=h.match(/\(([^)]+)\)/); if(m) curr=m[1].toUpperCase().trim();
              if(curr==='USD'){ costIdx=i; costCurrency=curr; }
              else if(fallbackCostIdx<0){ fallbackCostIdx=i; fallbackCostCurrency=curr; }
            }
            if(/last\s*asp/i.test(lower)){ aspIdx=i; const m2=h.match(/\(([^)]+)\)/); if(m2) aspCurrency=m2[1].toUpperCase().trim(); }
            if(/last\s*sold/i.test(lower)||/sold\s*\(days\)/i.test(lower)) soldIdx=i;
            if(lowerNoSpace==='pricing_sku' || lowerNoSpace==='pricingsku') pricingSkuIdx=i;
            if(lowerNoSpace==='gradetranslation' || lowerNoSpace==='grade translation' || lowerNoSpace==='gradetrans') gradeTransIdx=i;
          });
          if(costIdx<0 && fallbackCostIdx>=0){ costIdx=fallbackCostIdx; costCurrency=fallbackCostCurrency; }

          dataRows.forEach(rvals=>{
            // Build only the original row name here.
            // Missing aliases are expanded later, after all real costs are collected,
            // so existing names with their own costs are never overwritten.
            let canonicalNames=[];
            const originalModelName = modelIdx>=0 ? buildCanonicalModelGrade(rvals[modelIdx]) : '';
            if(originalModelName){
              canonicalNames=[originalModelName];
            }

            // Fallback only: if the row has no usable Model/Grade, use the first mapped name.
            // Do NOT expand all mapped names at this stage.
            if(canonicalNames.length===0 && nameMappingLoaded && pricingSkuIdx>=0){
              const psVal=rvals[pricingSkuIdx];
              if(psVal!==undefined && psVal!==null && psVal!==''){
                const baseList=nameMapping[String(psVal).trim().toUpperCase()];
                if(baseList && Array.isArray(baseList) && baseList.length>0){
                  let gradePartVal='';
                  if(gradeTransIdx>=0){
                    const gt=rvals[gradeTransIdx];
                    if(gt!==undefined && gt!==null) gradePartVal=String(gt).trim();
                  }
                  const baseName=String(baseList[0]||'').trim();
                  if(baseName) canonicalNames=[gradePartVal?`${baseName}-${gradePartVal}`:baseName];
                }
              }
            }
            if(canonicalNames.length===0) return;

            let rowSourcePricingSkus=[];
            if(pricingSkuIdx>=0){
              const psVal=rvals[pricingSkuIdx];
              if(psVal!==undefined && psVal!==null && psVal!=='') addUnique(rowSourcePricingSkus, String(psVal).trim().toUpperCase());
            }
            // Explicit Pricing SKU is authoritative. Reverse lookup is fallback-only.
            if(rowSourcePricingSkus.length===0){
              if(modelIdx>=0){
                const rawParts=splitModelGradeString(rvals[modelIdx]);
                rowSourcePricingSkus=mergeUniqueArrays(rowSourcePricingSkus, findPricingSkusForBaseNames([rawParts.base]));
              }
              canonicalNames.forEach(cn=>{
                const p=splitModelGradeString(cn);
                rowSourcePricingSkus=mergeUniqueArrays(rowSourcePricingSkus, findPricingSkusForBaseNames([p.base]));
              });
            }

            const costValRaw=toNumber(rvals[costIdx]);
            const costDefined=!isNaN(costValRaw);
            let aspVal=toNumber(rvals[aspIdx]);
            const lastSold=soldIdx>=0?toNumber(rvals[soldIdx]):undefined;
            if(isNaN(aspVal)) aspVal=undefined;

            let aspUSD;
            if(aspVal!==undefined){
              if(aspCurrency && aspCurrency!=='USD'){
                const rate=fxRates[aspCurrency];
                if(rate) aspUSD=aspVal*rate;
              } else {
                aspUSD=aspVal;
              }
            }
            let costUSD;
            if(costDefined){
              costUSD=costValRaw;
              if(costCurrency && costCurrency!=='USD'){
                const cRate=fxRates[costCurrency];
                if(cRate) costUSD=costValRaw*cRate;
              }
            }

            if(!aggregatorType1[region]) aggregatorType1[region]={};
            const regAgg=aggregatorType1[region];

            canonicalNames.forEach(cn=>{
              const key=normalizeKey(cn);
              if(!regAgg[key]){
                regAgg[key]={
                  maxCostUSD:costUSD,
                  maxCostVal:costValRaw,
                  costCurrency:costCurrency||'USD',
                  aspSumUSD:(aspUSD!==undefined)?aspUSD:0,
                  aspCount:(aspUSD!==undefined)?1:0,
                  aspMaxUSD:aspUSD,
                  aspCurrency:aspCurrency||'USD',
                  lastSoldDay:(!isNaN(lastSold)&&lastSold!==undefined)?lastSold:undefined,
                  originalKey:cn,
                  sourcePricingSkus:rowSourcePricingSkus
                };
              } else {
                const ag=regAgg[key];
                ag.sourcePricingSkus=mergeUniqueArrays(ag.sourcePricingSkus || [], rowSourcePricingSkus);
                if(costUSD!==undefined){
                  if(ag.maxCostUSD===undefined || costUSD>ag.maxCostUSD){
                    ag.maxCostUSD=costUSD;
                    ag.maxCostVal=costValRaw;
                    ag.costCurrency=costCurrency||ag.costCurrency;
                  }
                }
                if(aspUSD!==undefined){
                  ag.aspSumUSD+=aspUSD;
                  ag.aspCount+=1;
                  if(ag.aspMaxUSD===undefined || aspUSD>ag.aspMaxUSD){
                    ag.aspMaxUSD=aspUSD;
                    ag.aspCurrency=aspCurrency||ag.aspCurrency;
                  }
                }
                if(!isNaN(lastSold) && lastSold!==undefined){
                  if(ag.lastSoldDay===undefined || lastSold<ag.lastSoldDay){
                    ag.lastSoldDay=lastSold;
                  }
                }
              }
            });
          });
        } else if(type==='Type 2'){
          const hasCountryTypeHere = hasCountryType;

          if(hasCountryTypeHere){
            let modelCapIdx = headerMap['ModelCapacity']!==undefined ? headerMap['ModelCapacity'] : -1;
            let gradeIdx     = headerMap['Grade']!==undefined          ? headerMap['Grade']          : -1;
            let availIdx     = headerMap['Available']!==undefined      ? headerMap['Available']      : -1;

            let countryTypeIdx=-1;
            headers.forEach((h,i)=>{
              if(h && h.toLowerCase().replace(/\s+/g,'')==='countrytype') countryTypeIdx=i;
            });

            let costIdx2=-1;
            headers.forEach((h,i)=>{
              const lower=h.toLowerCase();
              if(costIdx2<0 && /net\s*cpu\s*\+\s*load/i.test(lower)) costIdx2=i;
            });
            let aspIdx2 = headerMap['Last 4w Gross ASP USD']!==undefined ? headerMap['Last 4w Gross ASP USD'] : -1;

            if(modelCapIdx<0 || gradeIdx<0 || countryTypeIdx<0){
              return;
            }

            const aggregatorByRegion={};

            dataRows.forEach(rvals=>{
              const modelCap=rvals[modelCapIdx];
              const grade=rvals[gradeIdx];
              const cTypeVal = countryTypeIdx>=0 ? rvals[countryTypeIdx] : undefined;

              if(!modelCap || !grade || cTypeVal===undefined || cTypeVal===null || cTypeVal==='') return;

              const regionKey=mapCountryTypeToRegion(cTypeVal);
              if(!regionKey) return;

              const baseModelRaw=(typeof modelCap==='string'?modelCap.trim():String(modelCap)).replace(/\s+/g,' ');
              const gradeRaw=(typeof grade==='string'?grade.trim():String(grade));
              let canonicalBase=baseModelRaw;
              if(referenceLoaded){
                const n=normalizeForRef(baseModelRaw);
                if(n && referenceMap[n]) canonicalBase=referenceMap[n];
              }
              const canonicalName=`${canonicalBase}-${gradeRaw}`;
              const key=normalizeKey(canonicalName);
              const sourcePricingSkus=findPricingSkusForBaseNames([baseModelRaw, canonicalBase]);

              const avail=toNumber(availIdx>=0 ? rvals[availIdx] : undefined);
              const cost =toNumber(costIdx2>=0 ? rvals[costIdx2] : undefined);
              const asp  =toNumber(aspIdx2>=0  ? rvals[aspIdx2]  : undefined);

              if(!aggregatorByRegion[regionKey]) aggregatorByRegion[regionKey]={};
              const regionAgg=aggregatorByRegion[regionKey];

              if(!regionAgg[key]) regionAgg[key]={ costSum:0,costAvailSum:0,totalAvail:0,costMax:undefined,aspMax:undefined,aspSumRaw:0,aspCount:0,canonicalName,sourcePricingSkus };
              const ag=regionAgg[key];
              ag.sourcePricingSkus=mergeUniqueArrays(ag.sourcePricingSkus || [], sourcePricingSkus);

              if(!isNaN(avail) && avail>0){
                ag.totalAvail += avail;
                if(!isNaN(cost)){ ag.costSum += avail*cost; ag.costAvailSum += avail; }
              }
              if(!isNaN(cost)){
                if(ag.costMax===undefined || cost>ag.costMax) ag.costMax=cost;
              }
              if(!isNaN(asp)){
                if(aspAggMethod==='avg'){ ag.aspSumRaw += asp; ag.aspCount += 1; }
                else { if(ag.aspMax===undefined || asp>ag.aspMax) ag.aspMax=asp; }
              }
            });

            Object.keys(aggregatorByRegion).forEach(regionKey=>{
              if(!summaryByRegion[regionKey]) summaryByRegion[regionKey]={};
              const regionAgg=aggregatorByRegion[regionKey];
              Object.keys(regionAgg).forEach(cKey=>{
                const ag=regionAgg[cKey];
                let costUSD;
                if(ag.costAvailSum>0) costUSD=ag.costSum/ag.costAvailSum;
                else if(ag.costMax!==undefined) costUSD=ag.costMax;

                let aspUSD;
                if(aspAggMethod==='avg') aspUSD = ag.aspCount>0 ? (ag.aspSumRaw/ag.aspCount) : undefined;
                else aspUSD = ag.aspMax!==undefined ? ag.aspMax : undefined;

                if(costUSD!==undefined || aspUSD!==undefined){
                  const rec={ costUSD:costUSD, costCurrency:'USD', costVal:costUSD, aspUSD:aspUSD, aspLocal:aspUSD, aspCurrency:'USD', lastSoldDay:undefined, originalKey:ag.canonicalName, sourcePricingSkus:ag.sourcePricingSkus || [] };
                  const existing=summaryByRegion[regionKey][cKey];
                  let update=false;
                  if(!existing) update=true;
                  else {
                    if(existing.costUSD===undefined && rec.costUSD!==undefined) update=true;
                    else if(existing.costUSD!==undefined && rec.costUSD!==undefined && rec.costUSD>existing.costUSD) update=true;
                    else if(existing.costUSD===undefined && rec.costUSD===undefined){
                      if((existing.aspUSD===undefined && rec.aspUSD!==undefined) ||
                         (rec.aspUSD!==undefined && existing.aspUSD!==undefined && rec.aspUSD>existing.aspUSD)) update=true;
                    }
                  }
                  if(update) summaryByRegion[regionKey][cKey]=rec;
                }
              });
            });

          } else {
            let modelCapIdx = headerMap['ModelCapacity']!==undefined ? headerMap['ModelCapacity'] : -1;
            let gradeIdx     = headerMap['Grade']!==undefined          ? headerMap['Grade']          : -1;
            let availIdx     = headerMap['Available']!==undefined      ? headerMap['Available']      : -1;

            let costIdx2=-1;
            headers.forEach((h,i)=>{
              const lower=h.toLowerCase();
              if(costIdx2<0 && /net\s*cpu\s*\+\s*load/i.test(lower)) costIdx2=i;
            });
            let aspIdx2 = headerMap['Last 4w Gross ASP USD']!==undefined ? headerMap['Last 4w Gross ASP USD'] : -1;

            const aggregator={};
            dataRows.forEach(rvals=>{
              const modelCap=rvals[modelCapIdx]; const grade=rvals[gradeIdx];
              if(!modelCap||!grade) return;
              const baseModelRaw=(typeof modelCap==='string'?modelCap.trim():String(modelCap)).replace(/\s+/g,' ');
              const gradeRaw=(typeof grade==='string'?grade.trim():String(grade));
              let canonicalBase=baseModelRaw;
              if(referenceLoaded){
                const n=normalizeForRef(baseModelRaw);
                if(n&&referenceMap[n]) canonicalBase=referenceMap[n];
              }
              const canonicalName=`${canonicalBase}-${gradeRaw}`;
              const key=normalizeKey(canonicalName);
              const sourcePricingSkus=findPricingSkusForBaseNames([baseModelRaw, canonicalBase]);

              const avail=toNumber(availIdx>=0 ? rvals[availIdx] : undefined);
              const cost =toNumber(costIdx2>=0 ? rvals[costIdx2] : undefined);
              const asp  =toNumber(aspIdx2>=0  ? rvals[aspIdx2]  : undefined);

              if(!aggregator[key]) aggregator[key]={ costSum:0,costAvailSum:0,totalAvail:0,costMax:undefined,aspMax:undefined,aspSumRaw:0,aspCount:0,canonicalName,sourcePricingSkus };
              const ag=aggregator[key];
              ag.sourcePricingSkus=mergeUniqueArrays(ag.sourcePricingSkus || [], sourcePricingSkus);

              if(!isNaN(avail) && avail>0){
                ag.totalAvail += avail;
                if(!isNaN(cost)){ ag.costSum += avail*cost; ag.costAvailSum += avail; }
              }
              if(!isNaN(cost)){
                if(ag.costMax===undefined || cost>ag.costMax) ag.costMax=cost;
              }
              if(!isNaN(asp)){
                if(aspAggMethod==='avg'){ ag.aspSumRaw += asp; ag.aspCount += 1; }
                else { if(ag.aspMax===undefined || asp>ag.aspMax) ag.aspMax=asp; }
              }
            });

            Object.keys(aggregator).forEach(cKey=>{
              const ag=aggregator[cKey];
              let costUSD;
              if(ag.costAvailSum>0) costUSD=ag.costSum/ag.costAvailSum;
              else if(ag.costMax!==undefined) costUSD=ag.costMax;

              let aspUSD;
              if(aspAggMethod==='avg') aspUSD = ag.aspCount>0 ? (ag.aspSumRaw/ag.aspCount) : undefined;
              else aspUSD = ag.aspMax!==undefined ? ag.aspMax : undefined;

              if(costUSD!==undefined || aspUSD!==undefined){
                const rec={ costUSD:costUSD, costCurrency:'USD', costVal:costUSD, aspUSD:aspUSD, aspLocal:aspUSD, aspCurrency:'USD', lastSoldDay:undefined, originalKey:ag.canonicalName, sourcePricingSkus:ag.sourcePricingSkus || [] };
                const existing=summaryByRegion[region][cKey];
                let update=false;
                if(!existing) update=true;
                else {
                  if(existing.costUSD===undefined && rec.costUSD!==undefined) update=true;
                  else if(existing.costUSD!==undefined && rec.costUSD!==undefined && rec.costUSD>existing.costUSD) update=true;
                  else if(existing.costUSD===undefined && rec.costUSD===undefined){
                    if((existing.aspUSD===undefined && rec.aspUSD!==undefined) ||
                       (rec.aspUSD!==undefined && existing.aspUSD!==undefined && rec.aspUSD>existing.aspUSD)) update=true;
                  }
                }
                if(update) summaryByRegion[region][cKey]=rec;
              }
            });
          }
        }
      });

      // Finalize Type 1 aggregations
      Object.keys(aggregatorType1).forEach(region=>{
        const regAgg=aggregatorType1[region];
        if(!summaryByRegion[region]) summaryByRegion[region]={};
        Object.keys(regAgg).forEach(key=>{
          const ag=regAgg[key];
          const costUSD=ag.maxCostUSD;
          const costVal=ag.maxCostVal;
          const costCurrency=ag.costCurrency||'USD';
          let aspUSD;
          if(aspAggMethod==='avg') aspUSD=ag.aspCount>0?(ag.aspSumUSD/ag.aspCount):undefined;
          else aspUSD=(ag.aspMaxUSD!==undefined)?ag.aspMaxUSD:undefined;
          const aspCurrencyFinal=ag.aspCurrency||'USD';
          let aspLocal;
          if(aspUSD!==undefined){
            if(aspCurrencyFinal!=='USD'){
              const _fxInput=document.getElementById('fxRatesInput').value.trim();
              const _fxTable=_fxInput ? parseFxRates(_fxInput) : (window.MarginFxAPI ? window.MarginFxAPI.getRates() : {});
              const rate=_fxTable[aspCurrencyFinal];
              aspLocal=rate? aspUSD/rate : undefined;
            } else {
              aspLocal=aspUSD;
            }
          }
          const rec={ costUSD:costUSD, costCurrency:costCurrency, costVal:costVal, aspUSD:aspUSD, aspLocal:aspLocal, aspCurrency:aspCurrencyFinal, lastSoldDay:ag.lastSoldDay, originalKey:ag.originalKey, sourcePricingSkus:ag.sourcePricingSkus || [] };
          const existing=summaryByRegion[region][key];
          let update=false;
          if(!existing) update=true;
          else {
            if(existing.costUSD===undefined||existing.costUSD===null){
              if(rec.costUSD!==undefined) update=true;
              else if(existing.aspUSD===undefined && rec.aspUSD!==undefined) update=true;
              else if(rec.aspUSD!==undefined && existing.aspUSD!==undefined && rec.aspUSD>existing.aspUSD) update=true;
            } else if(rec.costUSD!==undefined && rec.costUSD>existing.costUSD) update=true;
          }
          if(update) summaryByRegion[region][key]=rec;
        });
      });

      // Overlay Last ASP master if loaded: ASP & Last Sold Day override everything
      if(lastAspLoaded){
        Object.keys(lastAspByRegion).forEach(region=>{
          const aspMap=lastAspByRegion[region];
          if(!aspMap) return;
          if(!summaryByRegion[region]) summaryByRegion[region]={};
          Object.keys(aspMap).forEach(key=>{
            const info=aspMap[key];
            const aspUSD=info.aspUSD;
            const lastSold=info.lastSoldDay;
            if(aspUSD===undefined || aspUSD===null || isNaN(aspUSD)) return;
            const existing=summaryByRegion[region][key];
            const rec = existing || {
              costUSD:undefined,
              costCurrency:'USD',
              costVal:undefined,
              lastSoldDay:undefined,
              originalKey:key
            };
            rec.aspUSD = aspUSD;
            rec.aspLocal = aspUSD;        // master is USD
            rec.aspCurrency = 'USD';
            rec.aspFromLastMaster = true;
            if(lastSold!==undefined) rec.lastSoldDay = lastSold;
            summaryByRegion[region][key]=rec;
          });
        });
      }

      const expandAliases = document.getElementById('aliasExpansionCheckbox') ? document.getElementById('aliasExpansionCheckbox').checked : true;
      const guardConflicts = document.getElementById('aliasConflictGuardCheckbox') ? document.getElementById('aliasConflictGuardCheckbox').checked : true;
      applyAliasExpansion(summaryByRegion, expandAliases, guardConflicts);

      const builtRegions=Object.keys(summaryByRegion).filter(r=>Object.keys(summaryByRegion[r]||{}).length>0);
      if(builtRegions.length===0){
        const msg='Build finished with no usable summary rows. Check the selected Type, Region, header-row setting, and source columns.';
        window.dispatchEvent(new CustomEvent('margin-summary-build-error',{detail:{message:msg}}));
        alert(msg);
        return;
      }

      generatedSummaryByRegion=summaryByRegion;
      generatedFxRates=fxRates;
      window.dispatchEvent(new CustomEvent('margin-summary-generated',{
        detail:{ summaryByRegion, fxRates }
      }));

      buildOutputTable(summaryByRegion, fxRates);
      } catch(err) {
        const msg = 'Build Summary failed: ' + (err && err.message ? err.message : String(err));
        console.error(err);
        window.dispatchEvent(new CustomEvent('margin-summary-build-error',{detail:{message:msg}}));
        alert(msg);
      }
    });

    function buildOutputTable(summaryByRegion, fxRates){
      const output=document.getElementById('output'); output.innerHTML='';
      const regions=Object.keys(summaryByRegion);
      const allKeys=new Set();
      regions.forEach(r=>Object.keys(summaryByRegion[r]).forEach(k=>allKeys.add(k)));
      const keys=[...allKeys].sort();
      const displayKey={};
      regions.forEach(r=>{
        Object.keys(summaryByRegion[r]).forEach(k=>{
          const row=summaryByRegion[r][k];
          if(row&&row.originalKey && !displayKey[k]) displayKey[k]=row.originalKey;
        });
      });

      const table=document.createElement('table'); table.className='output-table';
      const thead=document.createElement('thead'); const hr=document.createElement('tr');
      const th0=document.createElement('th'); th0.textContent='Model/Grade'; hr.appendChild(th0);

      const regionMeta={};
      regions.forEach(r=>{
        regionMeta[r]={};
        keys.forEach(k=>{
          const row=summaryByRegion[r][k];
          if(row && !regionMeta[r].aspCurrency) regionMeta[r].aspCurrency=row.aspCurrency;
        });
      });

      regions.forEach(r=>{
        const cTh=document.createElement('th'); cTh.textContent=`C+Load(USD)(${r})`; hr.appendChild(cTh);
        const aspCurr=(regionMeta[r].aspCurrency||'USD').toUpperCase();
        const aTh=document.createElement('th'); aTh.textContent=`Last ASP(${aspCurr})(${r})`; hr.appendChild(aTh);
        if(aspCurr!=='USD'){
          const rate=fxRates[aspCurr];
          const usdTh=document.createElement('th'); usdTh.textContent=`Last ASP(USD FX:${rate||''})(${r})`; hr.appendChild(usdTh);
        }
        const sTh=document.createElement('th'); sTh.textContent=`Last Sold Day(${r})`; hr.appendChild(sTh);
      });
      thead.appendChild(hr); table.appendChild(thead);

      const tbody=document.createElement('tbody');
      keys.forEach(k=>{
        const tr=document.createElement('tr');
        const td0=document.createElement('td'); td0.textContent=displayKey[k]||k; tr.appendChild(td0);
        regions.forEach(r=>{
          const row=summaryByRegion[r][k];
          const cTd=document.createElement('td'); cTd.textContent=(row&&row.costUSD!==undefined)?row.costUSD.toFixed(2):''; tr.appendChild(cTd);
          const aTd=document.createElement('td'); aTd.textContent=(row&&row.aspLocal!==undefined)?row.aspLocal.toFixed(2):''; tr.appendChild(aTd);
          const aspCurr=(regionMeta[r].aspCurrency||'USD').toUpperCase();
          if(aspCurr!=='USD'){
            const uTd=document.createElement('td'); uTd.textContent=(row&&row.aspUSD!==undefined)?row.aspUSD.toFixed(2):''; tr.appendChild(uTd);
          }
          const sTd=document.createElement('td'); sTd.textContent=(row&&row.lastSoldDay!==undefined)?String(row.lastSoldDay):''; tr.appendChild(sTd);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody); output.appendChild(table);

      if(aliasExpansionLog.length>0){
        const added=aliasExpansionLog.filter(r=>r.Action==='Added missing alias').length;
        const filled=aliasExpansionLog.filter(r=>r.Action==='Filled missing alias fields').length;
        const skipped=aliasExpansionLog.filter(r=>String(r.Action).indexOf('Skipped')===0).length;
        const note=document.createElement('div');
        note.className='mt-3 text-sm text-gray-700 bg-white border border-gray-200 rounded p-3';
        note.textContent=`Alias expansion: ${added} missing alias row(s) added; ${filled} existing alias row(s) completed with missing fields. ${skipped} alias group check(s) skipped for possible conflict. Details will be included in the downloaded Excel.`;
        output.appendChild(note);
      }

      document.getElementById('downloadBtn').disabled = keys.length===0;
      if (keys.length===0) {
        const note=document.createElement('div');
        note.className='mt-3 text-sm text-red-700 bg-white border border-red-200 rounded p-3';
        note.textContent='Build Summary produced no rows to export. Check the source sheet headers (Model/Grade column, Last 4w Gross ASP USD column for Type 2) and confirm the Region selection.';
        output.appendChild(note);
      }
      window._generatedTable=table;
    }

    document.getElementById('downloadBtn').addEventListener('click',()=>{
      const table=window._generatedTable; if(!table) return;
      const wb=XLSX.utils.book_new();
      const ws=XLSX.utils.table_to_sheet(table);
      XLSX.utils.book_append_sheet(wb,ws,'Summary');
      if(aliasExpansionLog.length>0){
        const logRows=aliasExpansionLog.map(r=>({
          Action:r.Action||'',
          Region:r.Region||'',
          Pricing_SKU:r.Pricing_SKU||'',
          Grade:r.Grade||'',
          Alias_Name:r.Alias_Name||'',
          Filled_From:r.Filled_From||'',
          Cost_USD:r.Cost_USD||'',
          Fields_Filled:r.Fields_Filled||'',
          Names:r.Names||''
        }));
        const logWs=XLSX.utils.json_to_sheet(logRows);
        XLSX.utils.book_append_sheet(wb,logWs,'Alias_Expansion_Log');
      }
      const now=new Date();
      const tzString=now.toLocaleString('en-US',{ timeZone:'Asia/Taipei' });
      const tzDate=new Date(tzString);
      const pad=n=>String(n).padStart(2,'0');
      const mm=pad(tzDate.getMonth()+1), dd=pad(tzDate.getDate()), hh=pad(tzDate.getHours());
      const name=`${mm}${dd}${hh}_MarginSummary.xlsx`;
      XLSX.writeFile(wb,name);
    });

    // ---------- UI wiring ----------
    const dropZone=document.getElementById('dropZone');
    const fileInput=document.getElementById('fileInput');
    dropZone.addEventListener('click',()=>fileInput.click());
    fileInput.addEventListener('change',e=>handleFiles(e.target.files));
    dropZone.addEventListener('dragover',e=>{ e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave',()=>dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop',e=>{ e.preventDefault(); dropZone.classList.remove('dragover'); handleFiles(e.dataTransfer.files); });

    window.MarginAggregatorAPI={
      async loadSharedMapping(file){
        Object.keys(referenceMap).forEach(k=>delete referenceMap[k]);
        Object.keys(nameMapping).forEach(k=>delete nameMapping[k]);
        Object.keys(nameToPricingSkus).forEach(k=>delete nameToPricingSkus[k]);
        referenceLoaded=false;
        nameMappingLoaded=false;
        await loadReference(file);
        await loadNameMapping(file);
        return {
          referenceEntries:Object.keys(referenceMap).length,
          pricingSkuGroups:Object.keys(nameMapping).length
        };
      },
      getGeneratedSummary(){ return generatedSummaryByRegion; },
      getGeneratedFxRates(){ return generatedFxRates; },
      clearGeneratedSummary(){ generatedSummaryByRegion=null; generatedFxRates=null; },
      rebuildSheetMapping(){ buildMappingTable(); }
    };

    document.getElementById('lastAspInput').addEventListener('change', async (e) => {
      const f=e.target.files && e.target.files[0];
      if(!f) { lastAspLoaded=false; return; }
      try{
        await loadLastAspMaster(f);
        alert('Last ASP master loaded (ASP + Last Sold Day overlay).');
      } catch(err){
        alert('Failed to load Last ASP master: '+err.message);
        lastAspLoaded=false;
      }
    });
