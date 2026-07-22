/*
  Sales Dashboard — application logic
  This is the readable source. The page loads app.min.js (a terser-generated minified build).
  To rebuild app.min.js after editing this file:  npm install && npm run build
  See README.md for license and terms.
*/

/* ============================================================
   Sales Dashboard v2 — single-file, double-click HTML.
   Per-company drill-down, resizable filter panel, snapshot export,
   Excel export matching the original Python script's layout.
   ============================================================ */

// ---- Snapshot bootstrap (runs before anything else) ----------
// If this page was loaded as a snapshot, it carries a <script id="__snapshotData">.
const SNAPSHOT_DATA = (() => {
  const el = document.getElementById('__snapshotData');
  if (!el) return null;
  try { return JSON.parse(el.textContent); } catch (e) { return null; }
})();
const IS_SNAPSHOT = !!SNAPSHOT_DATA;

// ---- Constants matching the Python script -------------------
const WORKING_GRADES = ['A','B','Z','C','W','NB','N','AS','CP','WP'];
const DAMAGED_GRADES = ['Cx','D','Dx','F'];
const BACKUP_RATES = {
  USD:1.0, AUD:0.63, JPY:0.0069, EUR:1.12, GBP:1.33, AED:0.27, HKD:0.128,
};
// Rates from the Last Selling Price tool (Oct 2025) — editable from the UI in pricing mode
const RATES = { USD:1.0, AUD:0.6500, EUR:1.1628, AED:0.2723, GBP:1.3514, JPY:0.00661, CAD:0.7143 };
// Category code lookup (from the Last Selling Price tool) — derived from SKU's first segment
const CATEGORY_LOOKUP = { AC:'Accessories', EP:'Environment Products', CP:'Computers', LP:'Laptops', PH:'Smartphones', WB:'Wearables', TB:'Tablets' };
// Excluded company-name patterns. A row is dropped if its CompanyName contains any
// of these substrings (case-insensitive). Editable from the filter panel.
let EXCLUDED_PATTERNS = ['dummy phone store', 'alchemy staff account', 'Alchemy Telco Solutions'];
const REQUIRED_COLUMNS = [
  'ID','CreatedAt','CompanyName','OrderByName','OrderCreatedByUserType',
  'SKU','Quantity','Total','FloorTotal','Currency','Location',
  'PaymentDate','ShippedDate','OrderStatus','Country','SaleType'
];

// ---- Global state -------------------------------------------
let RAW = [];
let SKU_NAME_MAP = {};
let FILTERS = {
  // null = no filter (all selected). A Set = strict whitelist.
  // Default Condition to all 3 selected (matching v1 behavior so nothing is filtered out unexpectedly).
  mg: null,
  mgMode: 'model',
  grade: null,
  condition: new Set(['Working','Damaged','Unknown']),
  status: null,
  country: null,
};
let PRICING_FILTERS = {
  tokens: [],
  categories: new Set(['All']),
  dateRange: 'all',
  priceLock: 'all',
};
let SELECTED_COMPANY = '';
let MODE = 'monthly';   // 'monthly' | 'pricing'
let RAW_PRE_EXCLUSION = [];   // cached copy of pre-exclusion RAW so the editor can re-filter live
let chartModels = null, chartRegion = null;

// ---- Library bootstrap --------------------------------------
window.addEventListener('load', () => {
  setTimeout(() => {
    const xlsxOk = typeof XLSX !== 'undefined';
    const chartOk = typeof Chart !== 'undefined';
    const jszipOk = typeof JSZip !== 'undefined';
    if (!xlsxOk || !chartOk || !jszipOk) {
      document.getElementById('lib-error').classList.add('show');
    }
  }, 1500);
});

// ---- File inputs (disabled in snapshot mode) ---------------
const salesInput = document.getElementById('salesFile');
const mappingInput = document.getElementById('mappingFile');
if (IS_SNAPSHOT) {
  // Hide the file pickers; data is embedded
  salesInput.closest('label').style.display = 'none';
  mappingInput.closest('label').style.display = 'none';
  document.getElementById('status').textContent = 'Snapshot loaded.';
} else {
  salesInput.addEventListener('change', e => loadFile(e.target.files[0], 'sales'));
  mappingInput.addEventListener('change', e => loadFile(e.target.files[0], 'mapping'));
}

let salesBuf = null, mappingBuf = null;

async function loadFile(file, kind) {
  if (!file) return;
  const buf = await file.arrayBuffer();
  if (kind === 'sales') salesBuf = buf;
  else mappingBuf = buf;
  document.getElementById('status').textContent =
    (salesBuf ? '✓ sales' : '… sales') + '  ·  ' + (mappingBuf ? '✓ mapping' : '… mapping') + '   (waiting for both)';

  if (salesBuf && mappingBuf) {
    showOverlay('Parsing data…');
    await new Promise(r => setTimeout(r, 30));
    try { await buildAll(); }
    catch (err) { console.error(err); alert('Failed to load data:\n' + err.message); }
    hideOverlay();
  }
}

// ---- Parsing ------------------------------------------------
async function buildAll(opts) {
  // If opts has embedded data (snapshot mode), use it. Otherwise parse from buffers.
  if (opts && opts.embedded) {
    RAW = opts.rows;
    SKU_NAME_MAP = opts.skuNameMap;
    // Apply state
    if (opts.state) applyState(opts.state);
  } else {
    SKU_NAME_MAP = parseMapping(mappingBuf);
    RAW = parseSales(salesBuf);
    enrichRows(RAW);
  }
  populateFilters();
  // Company selection: single dropdown, default to first company
  const companies = uniqueSorted(RAW.map(r=>r.CompanyName).filter(Boolean));
  if (!companies.includes(SELECTED_COMPANY)) SELECTED_COMPANY = companies[0] || '';
  populateCompanySelect();
  setDefaultMonths();
  renderAll();
  // Capture the post-build HTML for snapshot use
  capturePageHtml();
  document.getElementById('exportBtn').disabled = false;
  document.getElementById('exportAllBtn').disabled = false;
  document.getElementById('snapshotBtn').disabled = false;
  // Per-tab export buttons
  for (const id of ['exportAspAll','exportAspSplit','exportSalesModelAll','exportSalesModelSplit','exportRaw','exportCountry','exportModelGrade','exportExcelLayout']) {
    const el = document.getElementById(id);
    if (el) el.disabled = false;
  }
  document.getElementById('placeholder').style.display = 'none';
  document.getElementById('appBody').style.display = 'block';
  document.getElementById('filterPanel').style.display = 'block';
  if (!IS_SNAPSHOT) {
    document.getElementById('status').textContent =
      `Loaded ${RAW.length.toLocaleString()} rows · ${companies.length} companies`;
  }
  // Restore filter width/collapse from localStorage
  restoreFilterPrefs();
}

function applyState(state) {
  if (state.selectedCompany) SELECTED_COMPANY = state.selectedCompany;
  FILTERS.mgMode = state.mgMode || 'model';
  // Accept either null (no filter) or an array (whitelist). Old snapshots with [] mean "all".
  FILTERS.grade = state.grade == null ? null : new Set(state.grade);
  FILTERS.condition = state.condition == null ? null : new Set(state.condition);
  FILTERS.status = state.status == null ? null : new Set(state.status);
  FILTERS.country = state.country == null ? null : new Set(state.country);
  FILTERS.mg = state.mg == null ? null : new Set(state.mg);
  if (state.currMonth) document.getElementById('currMonth').value = state.currMonth;
  if (state.prevMonth) document.getElementById('prevMonth').value = state.prevMonth;
  document.getElementById('useRange').checked = !!state.useRange;
  if (state.rangeStart) document.getElementById('rangeStart').value = state.rangeStart;
  if (state.rangeEnd) document.getElementById('rangeEnd').value = state.rangeEnd;
}

function snapshotState() {
  return {
    selectedCompany: SELECTED_COMPANY,
    mgMode: FILTERS.mgMode,
    // null means "no filter"; we persist it as null rather than [] to preserve intent.
    grade: FILTERS.grade == null ? null : [...FILTERS.grade],
    condition: FILTERS.condition == null ? null : [...FILTERS.condition],
    status: FILTERS.status == null ? null : [...FILTERS.status],
    country: FILTERS.country == null ? null : [...FILTERS.country],
    mg: FILTERS.mg == null ? null : [...FILTERS.mg],
    currMonth: document.getElementById('currMonth').value,
    prevMonth: document.getElementById('prevMonth').value,
    useRange: document.getElementById('useRange').checked,
    rangeStart: document.getElementById('rangeStart').value,
    rangeEnd: document.getElementById('rangeEnd').value,
  };
}

function parseMapping(buf) {
  const wb = XLSX.read(buf, {type:'array', cellDates:true});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:''});
  if (!rows.length) throw new Error('Name-mapping file is empty.');
  const keys = Object.keys(rows[0]).map(k => ({raw:k, low:k.trim().toLowerCase()}));
  const skuCol = keys.find(k => k.low==='sku')?.raw;
  const nameCol = keys.find(k => ['en_name','model','name','product_name'].includes(k.low))?.raw;
  if (!skuCol || !nameCol) throw new Error("Name-mapping file must contain 'sku' and 'en_name' columns.");
  const map = {};
  for (const r of rows) {
    const sku = String(r[skuCol]??'').trim();
    const name = String(r[nameCol]??'').trim();
    if (sku && name && !map[sku]) map[sku] = name;
  }
  return map;
}

function parseSales(buf) {
  const wb = XLSX.read(buf, {type:'array', cellDates:true});
  let sheetName = wb.SheetNames.includes('sales_line_item_detail') ? 'sales_line_item_detail' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, {defval:''});
}

function splitSkuGrade(sku) {
  sku = (sku==null?'':String(sku)).trim();
  if (!sku || !sku.includes('-')) return [sku, ''];
  const idx = sku.lastIndexOf('-');
  return [sku.slice(0,idx).trim(), sku.slice(idx+1).trim()];
}
function gradeCondition(grade) {
  if (WORKING_GRADES.includes(grade)) return 'Working';
  if (DAMAGED_GRADES.includes(grade)) return 'Damaged';
  return 'Unknown';
}
function getRate(currency) {
  currency = String(currency||'USD').trim().toUpperCase();
  return BACKUP_RATES[currency] ?? 1.0;
}
function enrichRows(rows) {
  // Enrich ALL rows first (so derived fields exist), then filter out excluded companies.
  const excluded = (EXCLUDED_PATTERNS || []).map(p => String(p).toLowerCase()).filter(Boolean);
  const cnMap = new Map();
  for (const r of rows) {
    const cn = String(r.CompanyName ?? '').trim();
    cnMap.set(r, cn);
    const [base, grade] = splitSkuGrade(r.SKU);
    r.BaseSKU = base;
    r.Grade = grade;
    const modelName = SKU_NAME_MAP[base] || base;
    r.ModelName = modelName;
    r.ModelGrade = grade ? `${modelName}-${grade}` : modelName;
    r.Condition = gradeCondition(grade);
    r.Quantity = num(r.Quantity);
    r.Total = num(r.Total);
    const firstSeg = (r.SKU || '').split('-')[0] || '';
    r.Category = CATEGORY_LOOKUP[firstSeg] || 'Other';
    const curr = String(r.Currency || 'USD').trim().toUpperCase();
    const rate = (RATES[curr] ?? BACKUP_RATES[curr] ?? 1.0);
    r.rate = rate;
    r.USD_Total = r.Total * rate;
    r.perUnit = r.Quantity > 0 ? r.Total / r.Quantity : 0;
    r.perUnitUsd = r.perUnit * rate;
    const d = r.CreatedAt;
    r._date = d instanceof Date ? d : (d ? new Date(d) : null);
    r.CompanyName = cn;
    r.Country = String(r.Country ?? '').trim();
    r.OrderStatus = String(r.OrderStatus ?? '').trim();
  }

  // Cache the fully-enriched pre-exclusion rows so the editor can re-filter live.
  RAW_PRE_EXCLUSION = rows.map(r => ({...r}));

  // Filter out excluded companies in place.
  if (excluded.length) {
    const kept = rows.filter(r => !excluded.some(p => r.CompanyName.toLowerCase().includes(p)));
    rows.length = 0;
    for (const r of kept) rows.push(r);
  }
}
function num(v){ const n = parseFloat(v); return isNaN(n) ? 0 : n; }

// ---- Searchable company picker (combobox) -------------------
function allCompanies() { return uniqueSorted(RAW.map(r => r.CompanyName).filter(Boolean)); }
function populateCompanySelect() {
  const inp = document.getElementById('comboInput');
  inp.disabled = false;
  if (SELECTED_COMPANY) inp.value = SELECTED_COMPANY;
  if (typeof window.renderComboList === 'function') window.renderComboList('');
}

function setupCompanyCombo() {
  const inp = document.getElementById('comboInput');
  const list = document.getElementById('comboList');
  let activeIdx = -1;

  function currentMatches() {
    const q = inp.value.trim().toLowerCase();
    return allCompanies().filter(c => !q || c.toLowerCase().includes(q));
  }
  window.renderComboList = function (q) {
    const matches = allCompanies().filter(c => !q || c.toLowerCase().includes(q.toLowerCase()));
    activeIdx = matches.length ? 0 : -1;
    if (!matches.length) {
      list.innerHTML = `<div class="opt disabled">${inp.value.trim() ? 'No matches' : 'No companies loaded'}</div>`;
    } else {
      list.innerHTML = matches.map((c,i) => `<div class="opt${i===0?' active':''}" data-co="${escapeHtml(c)}">${escapeHtml(c)}</div>`).join('');
    }
    list.classList.add('show');
  };
  function hideList() { list.classList.remove('show'); activeIdx = -1; }
  function setActive(idx) {
    const opts = [...list.querySelectorAll('.opt:not(.disabled)')];
    if (!opts.length) return;
    opts.forEach(o => o.classList.remove('active'));
    const wrap = idx < 0 ? idx + opts.length : idx % opts.length;
    opts[wrap].classList.add('active');
    activeIdx = wrap;
    opts[wrap].scrollIntoView({block:'nearest'});
  }
  function pick(c) {
    if (!c) return;
    SELECTED_COMPANY = c;
    inp.value = c;
    hideList();
    renderAll();
  }
  inp.addEventListener('focus', () => renderComboList(inp.value));
  inp.addEventListener('input', () => renderComboList(inp.value));
  inp.addEventListener('keydown', e => {
    const matches = currentMatches();
    if (e.key === 'ArrowDown') { e.preventDefault(); setActive(activeIdx + 1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(activeIdx - 1); }
    else if (e.key === 'Enter') {
      e.preventDefault();
      const opt = list.querySelector('.opt.active');
      if (opt) pick(opt.dataset.co);
      else if (matches.length) pick(matches[0]);
    }
    else if (e.key === 'Escape') { hideList(); inp.blur(); }
  });
  list.addEventListener('mousedown', e => {
    const opt = e.target.closest('.opt');
    if (opt && !opt.classList.contains('disabled')) pick(opt.dataset.co);
  });
  document.addEventListener('click', e => {
    if (!document.getElementById('companyCombo').contains(e.target)) hideList();
  });
}
setupCompanyCombo();

// "All" button next to the company picker — clears SELECTED_COMPANY so all companies show.
document.getElementById('companyAllBtn')?.addEventListener('click', () => {
  SELECTED_COMPANY = '';
  const inp = document.getElementById('comboInput');
  if (inp) inp.value = '';
  renderAll();
});

function escapeHtml(s){ return String(s).replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch])); }

function populateFilters() {
  renderChecklist('grade', uniqueSorted(RAW.map(r=>r.Grade).filter(Boolean)));
  renderChecklist('status', uniqueSorted(RAW.map(r=>r.OrderStatus).filter(Boolean)));
  renderChecklist('country', uniqueSorted(RAW.map(r=>r.Country).filter(Boolean)));
  renderChecklist('condition', ['Working','Damaged','Unknown']);
  refreshModelGradeChecklist();
}
function uniqueSorted(arr){ return [...new Set(arr)].sort((a,b)=>a.localeCompare(b)); }

function renderChecklist(name, values) {
  const el = document.getElementById('cl-'+name);
  el.innerHTML = '';
  const selected = FILTERS[name];
  for (const v of values) {
    const lab = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = selected == null || selected.has(v);
    cb.dataset.field = name;
    cb.dataset.value = v;
    cb.addEventListener('change', onFilterChange);
    lab.appendChild(cb);
    lab.appendChild(document.createTextNode(' ' + (v===''?'(blank)':v)));
    el.appendChild(lab);
  }
}

function refreshModelGradeChecklist() {
  const field = FILTERS.mgMode === 'model' ? 'ModelName' : 'ModelGrade';
  const values = uniqueSorted(RAW.map(r => r[field]).filter(v => v && v!==''));
  renderChecklist('mg', values);
}

function onFilterChange(e) {
  const field = e.target.dataset.field;
  if (!field) return;
  rebuildSetFromDom(field);
  renderAll();
}

// Rebuild FILTERS[field] from ALL checkboxes (visible AND search-hidden).
// This is the source of truth — not the Set. The Set is just a cache.
function rebuildSetFromDom(field) {
  const el = document.getElementById('cl-'+field);
  if (!el) return;
  const checkboxes = [...el.querySelectorAll('input[type=checkbox]')];
  const checked = checkboxes.filter(cb => cb.checked).map(cb => cb.dataset.value);
  // If every checkbox is checked → treat as "no filter" (null).
  // Otherwise store the explicit checked set.
  if (checked.length === checkboxes.length) {
    FILTERS[field] = null;
  } else {
    FILTERS[field] = new Set(checked);
  }
}

// search boxes
document.querySelectorAll('.search').forEach(inp => {
  inp.addEventListener('input', e => {
    const target = e.target.dataset.target;
    const q = e.target.value.toLowerCase();
    const el = document.getElementById('cl-'+target);
    el.querySelectorAll('label').forEach(lab => {
      lab.style.display = lab.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
  });
});
// all/none helpers (used by the buttons and by tests)
function selectAll(field) {
  document.querySelectorAll(`#cl-${field} input[type=checkbox]`).forEach(cb => cb.checked = true);
  FILTERS[field] = null;  // null = no filter (all selected)
  renderAll();
}
function selectNone(field) {
  document.querySelectorAll(`#cl-${field} input[type=checkbox]`).forEach(cb => cb.checked = false);
  FILTERS[field] = new Set();  // empty Set = explicit whitelist of zero items
  renderAll();
}
// all/none buttons
document.querySelectorAll('[data-checkall]').forEach(a => {
  a.addEventListener('click', () => selectAll(a.dataset.checkall));
});
document.querySelectorAll('[data-checknone]').forEach(a => {
  a.addEventListener('click', () => selectNone(a.dataset.checknone));
});

// Reset all filters to "no filter" (all selected). Does NOT clear pricing-mode fields,
// the company picker, the period pickers, or the excluded-company patterns — those
// are config, not session filters.
function resetAllFilters() {
  // Monthly filter sets: null = no filter
  FILTERS.mg = null;
  FILTERS.grade = null;
  FILTERS.status = null;
  FILTERS.country = null;
  // Condition is special: default to all 3 conditions selected (preserves v1 behavior).
  FILTERS.condition = new Set(['Working','Damaged','Unknown']);
  // Pricing filter sets
  PRICING_FILTERS.tokens = [];
  PRICING_FILTERS.categories = new Set(['All']);
  PRICING_FILTERS.dateRange = 'all';
  PRICING_FILTERS.priceLock = 'all';
  // Clear search inputs in the filter panel
  document.querySelectorAll('.filters input.search').forEach(inp => inp.value = '');
  // Restore chip-group active states for pricing filters
  document.querySelectorAll('#dateChips .chip').forEach(c => c.classList.toggle('active', c.dataset.d === 'all'));
  document.querySelectorAll('#priceChips .chip').forEach(c => c.classList.toggle('active', c.dataset.pl === 'all'));
  document.querySelectorAll('#categoryChips .chip').forEach(c => {
    if (c.dataset.cat === 'All') c.classList.add('active'); else c.classList.remove('active');
  });
  // Re-render all checklists so checkboxes + search visibility reset
  populateFilters();
  if (typeof renderModelTokens === 'function') renderModelTokens();
  renderAll();
}
document.getElementById('resetFilters')?.addEventListener('click', resetAllFilters);
// model/modelgrade toggle
document.querySelectorAll('[data-mgmode]').forEach(b => {
  b.addEventListener('click', () => {
    document.querySelectorAll('[data-mgmode]').forEach(x => {
      x.classList.remove('active');
      x.style.background = '#fff';
      x.style.borderColor = 'var(--border)';
      x.style.color = 'var(--text)';
    });
    b.classList.add('active');
    b.style.background = 'var(--accent-soft)';
    b.style.borderColor = 'var(--accent)';
    b.style.color = 'var(--accent)';
    FILTERS.mgMode = b.dataset.mgmode;
    FILTERS.mg = null;  // switching model↔model-grade resets the model filter
    refreshModelGradeChecklist();
    renderAll();
  });
});
// period + range inputs (company selection is handled by the chip input)
['currMonth','prevMonth','useRange','rangeStart','rangeEnd'].forEach(id => {
  document.getElementById(id).addEventListener('change', renderAll);
});

function setDefaultMonths() {
  const dates = RAW.map(r=>r._date).filter(Boolean).sort((a,b)=>a-b);
  if (!dates.length) return;
  if (document.getElementById('currMonth').value) return; // state already set
  const latest = dates[dates.length-1];
  const curr = new Date(latest.getFullYear(), latest.getMonth(), 1);
  const prev = new Date(curr); prev.setMonth(prev.getMonth()-1);
  document.getElementById('currMonth').value = fmtMonth(curr);
  document.getElementById('prevMonth').value = fmtMonth(prev);
}
function fmtMonth(d){ return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; }

// ---- Filtering core -----------------------------------------
// opts.ignoreCompany: when true, skip the SELECTED_COMPANY narrowing (used by ASP "all companies" scope).
function activeRows(opts) {
  const ignoreCompany = !!(opts && opts.ignoreCompany);
  // Date range (pricing-mode): anchored to latest data date, not wall clock
  let minDate = null;
  if (PRICING_FILTERS.dateRange !== 'all') {
    const days = parseInt(PRICING_FILTERS.dateRange, 10);
    const latest = RAW.map(r=>r._date).reduce((a,b)=> (a && a>b) ? a : b, null);
    if (latest) {
      minDate = new Date(latest); minDate.setHours(0,0,0,0);
      minDate.setDate(minDate.getDate() - (days - 1));
    }
  }

  return RAW.filter(r => {
    if (FILTERS.grade != null && !FILTERS.grade.has(r.Grade)) return false;
    if (FILTERS.status != null && !FILTERS.status.has(r.OrderStatus)) return false;
    if (FILTERS.country != null && !FILTERS.country.has(r.Country)) return false;
    if (FILTERS.condition != null && !FILTERS.condition.has(r.Condition)) return false;
    if (FILTERS.mg != null) {
      const key = FILTERS.mgMode === 'model' ? r.ModelName : r.ModelGrade;
      if (!FILTERS.mg.has(key)) return false;
    }
    if (!ignoreCompany && SELECTED_COMPANY && r.CompanyName !== SELECTED_COMPANY) return false;
    // Pricing-mode filters
    if (!PRICING_FILTERS.categories.has('All')) {
      if (!PRICING_FILTERS.categories.has(r.Category)) return false;
    }
    if (minDate && r._date && r._date < minDate) return false;
    if (PRICING_FILTERS.priceLock === 'exclude' && r.SaleType === 'Pre-Order') return false;
    if (PRICING_FILTERS.priceLock === 'only' && r.SaleType !== 'Pre-Order') return false;
    if (PRICING_FILTERS.tokens.length > 0) {
      const nm = (r.ModelName || '').toLowerCase();
      if (!PRICING_FILTERS.tokens.some(t => nm.includes(t.toLowerCase()))) return false;
    }
    return true;
  });
}

function monthBounds(yearMonth) {
  const [y,m] = yearMonth.split('-').map(Number);
  const start = new Date(y, m-1, 1);
  const end = new Date(y, m, 0, 23, 59, 59);
  return [start, end];
}

function splitByPeriod(rows) {
  const useRange = document.getElementById('useRange').checked;
  if (useRange) {
    const s = document.getElementById('rangeStart').value;
    const e = document.getElementById('rangeEnd').value;
    if (!s || !e) return {curr:[], prev:[], currLabel:'(no range)', prevLabel:''};
    const sd = new Date(s+'T00:00:00'), ed = new Date(e+'T23:59:59');
    const span = ed - sd;
    const psd = new Date(sd.getTime() - span - 86400000);
    const ped = new Date(sd.getTime() - 86400000);
    const curr = rows.filter(r => r._date && r._date>=sd && r._date<=ed);
    const prev = rows.filter(r => r._date && r._date>=psd && r._date<=ped);
    return {curr, prev, currLabel:`${s} → ${e}`, prevLabel:'prior window'};
  }
  const cm = document.getElementById('currMonth').value;
  const pm = document.getElementById('prevMonth').value;
  const [cs,ce] = cm ? monthBounds(cm) : [null,null];
  const [ps,pe] = pm ? monthBounds(pm) : [null,null];
  const curr = cs ? rows.filter(r => r._date && r._date>=cs && r._date<=ce) : [];
  const prev = ps ? rows.filter(r => r._date && r._date>=ps && r._date<=pe) : [];
  return {curr, prev, currLabel: cm||'(none)', prevLabel: pm||'(none)'};
}

// ---- Aggregations -------------------------------------------
function sumQty(rows){ return rows.reduce((s,r)=>s+r.Quantity,0); }
function sumUSD(rows){ return rows.reduce((s,r)=>s+r.USD_Total,0); }
function qtyByCond(rows, cond){ return rows.filter(r=>r.Condition===cond).reduce((s,r)=>s+r.Quantity,0); }

function groupBy(rows, keyFn) {
  const m = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return m;
}

function momPct(delta, prev){ return prev ? (delta/prev*100) : null; }
function fmtPct(p){ return p==null ? 'N/A' : `${p>=0?'+':''}${p.toFixed(1)}%`; }
function fmtUSD(v){ return '$'+v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }
function fmtNum(v){ return v.toLocaleString(); }
function safeFilename(s){ return String(s).replace(/[^A-Za-z0-9\-_ ]/g,'_'); }

// ---- Render -------------------------------------------------
function renderAll() {
  if (RAW.length === 0) return;
  const rows = activeRows();
  const {curr, prev, currLabel, prevLabel} = splitByPeriod(rows);
  // update breadcrumb with the selected company
  const label = SELECTED_COMPANY || 'All companies';
  document.getElementById('breadcrumb').innerHTML =
    `<b>${escapeHtml(label)}</b> · ${currLabel} vs ${prevLabel}`;
  if (MODE === 'monthly') {
    renderKPIs(curr, prev);
    renderCharts(curr, prev);
    renderCountryTable(curr, prev);
    renderModelGradeTable(curr, prev, currLabel, prevLabel);
    renderExcelLayout(curr, prev, currLabel, prevLabel);
  } else {
    renderASP();
    renderSalesByModel();
  }
  renderRaw(rows);
}

function renderASP() {
  const host = document.getElementById('aspContent');
  const scopeLabel = document.getElementById('aspScope');
  // ASP follows the global company picker directly: empty picker → all companies,
  // a company selected → narrow to that company. No tab-local override.
  const scoped = activeRows();
  if (scopeLabel) scopeLabel.textContent = SELECTED_COMPANY || 'all companies';
  if (!scoped.length) { host.innerHTML = '<div style="color:var(--muted)">No results for the selected filters.</div>'; return; }
  // Group by country × model grade, keep the most recent sale per group
  const groups = {};
  for (const r of scoped) {
    const c = r.Country || 'Unknown', mg = r.ModelGrade;
    if (!groups[c]) groups[c] = {};
    const existing = groups[c][mg];
    if (!existing || r._date > existing._date) groups[c][mg] = r;
  }
  let html = '';
  for (const c of Object.keys(groups).sort()) {
    html += `<div class="asp-country"><h3>${escapeHtml(c)}</h3>`;
    html += '<div class="table-wrap"><table><thead><tr><th>Model Grade</th><th class="num">Last ASP (USD)</th><th>Company</th><th>Date</th></tr></thead><tbody>';
    const entries = Object.values(groups[c]).sort((a,b) => a.ModelGrade.localeCompare(b.ModelGrade));
    for (const e of entries) {
      html += `<tr><td>${escapeHtml(e.ModelGrade)}</td><td class="num">${fmtUSD(e.perUnitUsd)}</td><td>${escapeHtml(e.CompanyName)}</td><td>${e._date.toISOString().split('T')[0]}</td></tr>`;
    }
    html += '</tbody></table></div></div>';
  }
  host.innerHTML = html;
}

function renderSalesByModel() {
  const host = document.getElementById('salesModelContent');
  const rows = activeRows();
  if (!rows.length) { host.innerHTML = '<div style="color:var(--muted)">No results for the selected filters.</div>'; return; }
  const byCountry = {};
  for (const r of rows) {
    const c = r.Country || 'Unknown';
    if (!byCountry[c]) byCountry[c] = {};
    if (!byCountry[c][r.ModelGrade]) byCountry[c][r.ModelGrade] = [];
    byCountry[c][r.ModelGrade].push(r);
  }
  const countries = Object.keys(byCountry).sort();
  let tabs = '<div class="country-tabs" style="display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap">';
  countries.forEach((c, i) => { tabs += `<button class="country-tab-btn btn-mini${i===0?' active':''}" data-cc="${escapeHtml(c)}">${escapeHtml(c)}</button>`; });
  tabs += '</div>';
  let tables = '';
  countries.forEach((c, i) => {
    tables += `<div class="country-table-area" data-cc="${escapeHtml(c)}" style="${i===0?'':'display:none'}">`;
    tables += '<div class="table-wrap"><table><thead><tr><th>Model Grade</th><th class="num">ASP (USD)</th><th class="num">Qty</th><th>Company</th><th>Date</th></tr></thead><tbody>';
    const grades = Object.keys(byCountry[c]).sort();
    for (const g of grades) {
      const items = byCountry[c][g].sort((a,b) => b._date - a._date);
      tables += `<tr class="asp-grade" style="background:#f3f4f6;font-weight:600"><td colspan="5">${escapeHtml(g)}</td></tr>`;
      for (const it of items) {
        tables += `<tr><td class="indent" style="padding-left:24px">${escapeHtml(it.SKU)}</td><td class="num">${fmtUSD(it.perUnitUsd)}</td><td class="num">${fmtNum(it.Quantity)}</td><td>${escapeHtml(it.CompanyName)}</td><td>${it._date.toISOString().split('T')[0]}</td></tr>`;
      }
    }
    tables += '</tbody></table></div></div>';
  });
  host.innerHTML = tabs + tables;
  host.querySelectorAll('.country-tab-btn').forEach(b => b.addEventListener('click', () => {
    const cc = b.dataset.cc;
    host.querySelectorAll('.country-tab-btn').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    host.querySelectorAll('.country-table-area').forEach(t => t.style.display = t.dataset.cc === cc ? '' : 'none');
  }));
}

function renderRaw(rows) {
  const tbl = document.getElementById('tblRaw');
  if (!tbl) return;
  if (!rows.length) { tbl.innerHTML = '<tbody><tr><td style="color:var(--muted)">No data.</td></tr></tbody>'; return; }
  const cols = ['ID','CreatedAt','CompanyName','Country','SKU','ModelGrade','Category','Condition','OrderStatus','SaleType','Currency','Quantity','Total','perUnitUsd','USD_Total'];
  let h = '<thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead><tbody>';
  for (const r of rows) {
    h += '<tr>' + cols.map(c => {
      let v = r[c];
      if (v instanceof Date) v = v.toISOString().split('T')[0];
      else if (typeof v === 'number') v = ['perUnitUsd','USD_Total','perUnit','rate'].includes(c) ? v.toFixed(2) : v;
      const isNum = ['Quantity','Total','perUnitUsd','USD_Total','perUnit','rate'].includes(c);
      return `<td class="${isNum?'num':''}">${escapeHtml(String(v ?? ''))}</td>`;
    }).join('') + '</tr>';
  }
  h += '</tbody>';
  tbl.innerHTML = h;
}

function renderKPIs(curr, prev) {
  const cq = sumQty(curr), cu = sumUSD(curr);
  const pq = sumQty(prev), pu = sumUSD(prev);
  const wq = qtyByCond(curr,'Working'), dq = qtyByCond(curr,'Damaged');
  const kpis = [
    {label:'Total Qty', value: fmtNum(cq), delta: cq-pq, prev: pq},
    {label:'Working Qty', value: fmtNum(wq), delta: wq-qtyByCond(prev,'Working'), prev: qtyByCond(prev,'Working')},
    {label:'Damaged Qty', value: fmtNum(dq), delta: dq-qtyByCond(prev,'Damaged'), prev: qtyByCond(prev,'Damaged')},
    {label:'Total USD', value: fmtUSD(cu), delta: cu-pu, prev: pu, money:true},
  ];
  const grid = document.getElementById('kpiGrid');
  grid.innerHTML = kpis.map(k => {
    const pct = momPct(k.delta, k.prev);
    const cls = k.delta>0?'up':(k.delta<0?'down':'flat');
    const arrow = k.delta>0?'▲':(k.delta<0?'▼':'→');
    const deltaStr = `${arrow} ${k.money?fmtUSD(Math.abs(k.delta)):fmtNum(Math.abs(Math.round(k.delta)))} (${fmtPct(pct)})`;
    return `<div class="kpi"><div class="label">${k.label}</div><div class="value">${k.value}</div><div class="delta ${cls}">${deltaStr}</div></div>`;
  }).join('');
}

function renderCharts(curr, prev) {
  const topN = 10;
  // Top models
  const mField = FILTERS.mgMode === 'model' ? 'ModelName' : 'ModelGrade';
  const modCurr = groupBy(curr, r=>r[mField]||'(blank)');
  const modPrev = groupBy(prev, r=>r[mField]||'(blank)');
  const modKeys = [...new Set([...modCurr.keys(), ...modPrev.keys()])]
    .sort((a,b)=> (sumUSD(modCurr.get(b)||[])) - (sumUSD(modCurr.get(a)||[])))
    .slice(0, topN);
  chartModels = drawBar('chartModels', chartModels, modKeys,
    modKeys.map(k=>sumUSD(modCurr.get(k)||[])),
    modKeys.map(k=>sumUSD(modPrev.get(k)||[])), 'USD', true);

  // Region (country) distribution by USD — pie chart
  const regG = groupBy(curr, r=>r.Country||'(blank)');
  const regEntries = [...regG.entries()].sort((a,b)=>sumUSD(b[1])-sumUSD(a[1]));
  const regLabels = regEntries.map(e=>e[0]);
  const regData = regEntries.map(e=>sumUSD(e[1]));
  const regColors = regLabels.map((_,i) => `hsl(${(i*47)%360} 65% 55%)`);
  const regEl = document.getElementById('chartRegion').getContext('2d');
  if (chartRegion) chartRegion.destroy();
  chartRegion = new Chart(regEl, {
    type: 'pie',
    data: { labels: regLabels, datasets: [{ data: regData, backgroundColor: regColors }] },
    options: { responsive:true, maintainAspectRatio:false, plugins:{ legend:{position:'right'} } }
  });
}

function drawBar(canvasId, existing, labels, currData, prevData, label, rotateLabels=false) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  if (existing) existing.destroy();
  const ch = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {label:'Current', data:currData, backgroundColor:'#2563eb'},
        {label:'Compare', data:prevData, backgroundColor:'#cbd5e1'},
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'bottom'} },
      scales:{
        x:{ ticks:{ autoSkip:false, maxRotation: rotateLabels?45:0, minRotation: rotateLabels?45:0 } },
        y:{ ticks:{ callback:v=>'$'+v.toLocaleString() } }
      }
    }
  });
  return ch;
}

function renderCountryTable(curr, prev) {
  const currG = groupBy(curr, r=>r.Country||'(blank)');
  const prevG = groupBy(prev, r=>r.Country||'(blank)');
  const keys = [...new Set([...currG.keys(), ...prevG.keys()])].sort();
  const totalUSD = sumUSD(curr);
  const rows = keys.map(k => {
    const c = currG.get(k)||[], p = prevG.get(k)||[];
    const cUSD = sumUSD(c), pUSD = sumUSD(p);
    const cQ = sumQty(c), pQ = sumQty(p);
    return {country:k, cQ, pQ, cUSD, pUSD, pct: totalUSD?cUSD/totalUSD*100:0,
      qd:cQ-pQ, ud:cUSD-pUSD, qp:momPct(cQ-pQ,pQ), up:momPct(cUSD-pUSD,pUSD)};
  });
  const tot = rows.reduce((a,r)=>({cQ:a.cQ+r.cQ,pQ:a.pQ+r.pQ,cUSD:a.cUSD+r.cUSD,pUSD:a.pUSD+r.pUSD}),{cQ:0,pQ:0,cUSD:0,pUSD:0});
  const cols = [
    {h:'Country', k:'country'},
    {h:'Qty (Curr)', k:'cQ', num:true, fmt:fmtNum},
    {h:'Qty (Prev)', k:'pQ', num:true, fmt:fmtNum},
    {h:'Qty Δ', k:'qd', num:true, fmt:fmtNum},
    {h:'Qty %', k:'qp', num:true, fmt:fmtPct},
    {h:'USD (Curr)', k:'cUSD', num:true, fmt:fmtUSD},
    {h:'USD (Prev)', k:'pUSD', num:true, fmt:fmtUSD},
    {h:'USD Δ', k:'ud', num:true, fmt:fmtUSD},
    {h:'USD %', k:'up', num:true, fmt:fmtPct},
    {h:'% of total', k:'pct', num:true, fmt:v=>v.toFixed(2)+'%'},
  ];
  let html = `<thead><tr>${cols.map(c=>`<th class="${c.num?'num':''}">${c.h}</th>`).join('')}</tr></thead><tbody>`;
  for (const r of rows) {
    html += `<tr>${cols.map(c=>`<td class="${c.num?'num':''}">${c.fmt?c.fmt(r[c.k]):r[c.k]}</td>`).join('')}</tr>`;
  }
  html += `<tr class="total-row"><td>Total</td><td class="num">${fmtNum(tot.cQ)}</td><td class="num">${fmtNum(tot.pQ)}</td><td class="num">${fmtNum(tot.cQ-tot.pQ)}</td><td class="num">${fmtPct(momPct(tot.cQ-tot.pQ,tot.pQ))}</td><td class="num">${fmtUSD(tot.cUSD)}</td><td class="num">${fmtUSD(tot.pUSD)}</td><td class="num">${fmtUSD(tot.cUSD-tot.pUSD)}</td><td class="num">${fmtPct(momPct(tot.cUSD-tot.pUSD,tot.pUSD))}</td><td class="num"></td></tr>`;
  html += '</tbody>';
  document.getElementById('tblCountry').innerHTML = html;
  attachSort('tblCountry', cols, rows);
}

function renderModelGradeTable(curr, prev, currLabel, prevLabel) {
  const field = FILTERS.mgMode === 'model' ? 'ModelName' : 'ModelGrade';
  const mgLabel = FILTERS.mgMode === 'model' ? 'Model' : 'Model Grade';
  // Update the section labels with the period names
  const elCur = document.getElementById('mgCurrentLabel');
  const elPrev = document.getElementById('mgPrevLabel');
  if (elCur) elCur.textContent = currLabel || 'current';
  if (elPrev) elPrev.textContent = prevLabel || 'previous';
  const tL = document.getElementById('mgLeftTitle'); if (tL) tL.textContent = `Current month — ${currLabel || ''}`;
  const tR = document.getElementById('mgRightTitle'); if (tR) tR.textContent = `Previous month — ${prevLabel || ''}`;

  function buildOne(rowsArr, prefix) {
    const g = groupBy(rowsArr, r=>r[field]||'(blank)');
    const keys = [...g.keys()].sort((a,b)=>sumQty(g.get(b))-sumQty(g.get(a)));
    const data = keys.map(k=>{ const v=g.get(k); return {mg:k, qty:sumQty(v), usd:sumUSD(v)}; });
    const tot = data.reduce((a,r)=>({qty:a.qty+r.qty,usd:a.usd+r.usd}),{qty:0,usd:0});
    const cols = [
      {h: mgLabel, k:'mg'},
      {h:'Quantity', k:'qty', num:true, fmt:fmtNum},
      {h:'Total (USD)', k:'usd', num:true, fmt:fmtUSD},
    ];
    let html = `<thead><tr>${cols.map(c=>`<th class="${c.num?'num':''}">${c.h}</th>`).join('')}</tr></thead><tbody>`;
    for (const r of data) html += `<tr>${cols.map(c=>`<td class="${c.num?'num':''}">${c.fmt?c.fmt(r[c.k]):r[c.k]}</td>`).join('')}</tr>`;
    html += `<tr class="total-row"><td>Total</td><td class="num">${fmtNum(tot.qty)}</td><td class="num">${fmtUSD(tot.usd)}</td></tr>`;
    html += '</tbody>';
    document.getElementById(prefix).innerHTML = html;
    attachSort(prefix, cols, data);
  }
  buildOne(curr, 'tblModelGradeCurr');
  buildOne(prev, 'tblModelGradePrev');
}

function attachSort(tableId, cols, rows) {
  const table = document.getElementById(tableId);
  const ths = table.querySelectorAll('thead th');
  ths.forEach((th, i) => {
    th.onclick = () => {
      const col = cols[i];
      const dir = th.dataset.dir === 'asc' ? -1 : 1;
      th.dataset.dir = dir === 1 ? 'asc' : 'desc';
      rows.sort((a,b)=>{
        const av = a[col.k], bv = b[col.k];
        if (typeof av === 'number') return (av-bv)*dir;
        return String(av).localeCompare(String(bv))*dir;
      });
      const tbody = table.querySelector('tbody');
      const totalRow = tbody.querySelector('tr.total-row');
      [...tbody.querySelectorAll('tr:not(.total-row)')].forEach(r=>r.remove());
      const frag = document.createDocumentFragment();
      for (const r of rows) {
        const tr = document.createElement('tr');
        tr.innerHTML = cols.map(c=>`<td class="${c.num?'num':''}">${c.fmt?c.fmt(r[c.k]):r[c.k]}</td>`).join('');
        frag.appendChild(tr);
      }
      tbody.insertBefore(frag, totalRow);
    };
  });
}

// ---- Excel Layout tab (matches Python script output) -------
function renderExcelLayout(curr, prev, currLabel, prevLabel) {
  const host = document.getElementById('excelLayout');
  let html = '';

  // Section 1: All-company MoM summary across the selected companies
  html += renderAllCompanyMoMTable(curr, prev, currLabel, prevLabel);

  // Section 2: Block for the selected company (matches Python script layout)
  if (SELECTED_COMPANY) {
    const currByCo = groupBy(curr, r=>r.CompanyName||'(blank)');
    const prevByCo = groupBy(prev, r=>r.CompanyName||'(blank)');
    html += renderCompanyBlock(SELECTED_COMPANY, currByCo.get(SELECTED_COMPANY)||[], prevByCo.get(SELECTED_COMPANY)||[], currLabel, prevLabel);
  }
  host.innerHTML = html;
}

function renderAllCompanyMoMTable(curr, prev, currLabel, prevLabel) {
  const compCurr = groupBy(curr, r=>r.CompanyName||'(blank)');
  const compPrev = groupBy(prev, r=>r.CompanyName||'(blank)');
  const keys = [...new Set([...compCurr.keys(), ...compPrev.keys()])].sort();
  let tQp=0,tUp=0,tQc=0,tUc=0;
  let body = '';
  for (const k of keys) {
    const c=compCurr.get(k)||[], p=compPrev.get(k)||[];
    const cQ=sumQty(c),pQ=sumQty(p),cU=sumUSD(c),pU=sumUSD(p);
    tQp+=pQ;tUp+=pU;tQc+=cQ;tUc+=cU;
    body += `<tr><td>${escapeHtml(k)}</td><td class="num">${fmtNum(pQ)}</td><td class="num">${fmtUSD(pU)}</td><td class="num">${fmtNum(cQ)}</td><td class="num">${fmtUSD(cU)}</td><td class="num">${fmtNum(cQ-pQ)}</td><td class="num">${fmtPct(momPct(cQ-pQ,pQ))}</td><td class="num">${fmtUSD(cU-pU)}</td><td class="num">${fmtPct(momPct(cU-pU,pU))}</td></tr>`;
  }
  const total = `<tr class="total-row"><td>Total</td><td class="num">${fmtNum(tQp)}</td><td class="num">${fmtUSD(tUp)}</td><td class="num">${fmtNum(tQc)}</td><td class="num">${fmtUSD(tUc)}</td><td class="num">${fmtNum(tQc-tQp)}</td><td class="num">${fmtPct(momPct(tQc-tQp,tQp))}</td><td class="num">${fmtUSD(tUc-tUp)}</td><td class="num">${fmtPct(momPct(tUc-tUp,tUp))}</td></tr>`;
  return `
  <div class="excel-block">
    <h3>📊 All-Company MoM Summary</h3>
    <div class="subtitle">Current: ${escapeHtml(currLabel)} · Compare: ${escapeHtml(prevLabel)}</div>
    <div class="table-wrap"><table>
      <thead><tr>
        <th>Company</th><th class="num">Qty (Prev)</th><th class="num">$ (Prev)</th>
        <th class="num">Qty (Curr)</th><th class="num">$ (Curr)</th>
        <th class="num">Qty Δ</th><th class="num">Qty %</th>
        <th class="num">$ Δ</th><th class="num">$ %</th>
      </tr></thead>
      <tbody>${body}${total}</tbody>
    </table></div>
  </div>`;
}

function renderCompanyBlock(company, cc, cp, currLabel, prevLabel) {
  // Group A (left): orders by ID/SaleType/Location/CompanyName/CreatedAt/OrderStatus
  const groupCols = ['ID','SaleType','Location','CreatedAt','OrderStatus'];
  const G_curr = groupBy(cc, r=>groupCols.map(k=>r[k]??'').join('||'));
  const G_prev = groupBy(cp, r=>groupCols.map(k=>r[k]??'').join('||'));
  const gKeys = [...G_curr.keys()].sort();
  const groupRows = gKeys.map(k=>{
    const cs = G_curr.get(k)||[], ps = G_prev.get(k)||[];
    return {
      parts: k.split('||'),
      cW: qtyByCond(cs,'Working'),
      cD: qtyByCond(cs,'Damaged'),
      cT: sumQty(cs),
      cU: sumUSD(cs),
      pT: sumQty(ps),
      pU: sumUSD(ps),
    };
  });
  const gTot = groupRows.reduce((a,r)=>({
    cW:a.cW+r.cW,cD:a.cD+r.cD,cT:a.cT+r.cT,cU:a.cU+r.cU,pT:a.pT+r.pT,pU:a.pU+r.pU
  }),{cW:0,cD:0,cT:0,cU:0,pT:0,pU:0});
  const gA = `
    <table>
      <thead><tr>
        <th>ID</th><th>SaleType</th><th>Location</th><th>CreatedAt</th><th>OrderStatus</th>
        <th class="num">Working(Qty)</th><th class="num">Damaged(Qty)</th>
        <th class="num">Total(Qty)</th><th class="num">Total(USD)</th>
      </tr></thead>
      <tbody>
        ${groupRows.map(r=>`<tr>
          ${r.parts.map(p=>`<td>${escapeHtml(p)}</td>`).join('')}
          <td class="num">${fmtNum(r.cW)}</td><td class="num">${fmtNum(r.cD)}</td>
          <td class="num">${fmtNum(r.cT)}</td><td class="num">${fmtUSD(r.cU)}</td>
        </tr>`).join('')}
        <tr class="total-row">
          <td colspan="5">Total</td>
          <td class="num">${fmtNum(gTot.cW)}</td><td class="num">${fmtNum(gTot.cD)}</td>
          <td class="num">${fmtNum(gTot.cT)}</td><td class="num">${fmtUSD(gTot.cU)}</td>
        </tr>
      </tbody>
    </table>`;

  // Group B (right top): country breakdown (current month)
  const ccG = groupBy(cc, r=>r.Country||'(blank)');
  const totalU = sumUSD(cc);
  const countryRows = [...ccG.keys()].sort().map(ct=>{
    const c=ccG.get(ct)||[];
    return {Country:ct, Quantity:sumQty(c), 'Total(USD)':sumUSD(c), pct: totalU?sumUSD(c)/totalU*100:0};
  });
  const cTot = countryRows.reduce((a,r)=>({Quantity:a.Quantity+r.Quantity, USD:a['Total(USD)']+r['Total(USD)']}),{Quantity:0,USD:0});
  const gB = `
    <table>
      <thead><tr><th>Country</th><th class="num">Quantity</th><th class="num">Total(USD)</th><th class="num">% of total</th></tr></thead>
      <tbody>
        ${countryRows.map(r=>`<tr><td>${escapeHtml(r.Country)}</td><td class="num">${fmtNum(r.Quantity)}</td><td class="num">${fmtUSD(r['Total(USD)'])}</td><td class="num">${r.pct.toFixed(2)}%</td></tr>`).join('')}
        <tr class="total-row"><td>Total</td><td class="num">${fmtNum(cTot.Quantity)}</td><td class="num">${fmtUSD(cTot.USD)}</td><td class="num"></td></tr>
      </tbody>
    </table>`;

  // Group C (right middle): country MoM
  const cpG = groupBy(cp, r=>r.Country||'(blank)');
  const momCkeys = [...new Set([...ccG.keys(), ...cpG.keys()])].sort();
  const momCountryRows = momCkeys.map(ct=>{
    const c=ccG.get(ct)||[], p=cpG.get(ct)||[];
    const cQ=sumQty(c),pQ=sumQty(p),cU=sumUSD(c),pU=sumUSD(p);
    return {Country:ct,pQ,pU,cQ,cU,
      qd:cQ-pQ, qp:momPct(cQ-pQ,pQ),
      ud:cU-pU, up:momPct(cU-pU,pU)};
  });
  const mcTot = momCountryRows.reduce((a,r)=>({pQ:a.pQ+r.pQ,pU:a.pU+r.pU,cQ:a.cQ+r.cQ,cU:a.cU+r.cU}),{pQ:0,pU:0,cQ:0,cU:0});
  const gC = `
    <h4 style="font-size:13px;margin:14px 0 6px">Country MoM</h4>
    <table>
      <thead><tr>
        <th>Country</th>
        <th class="num">Qty (Prev)</th><th class="num">$ (Prev)</th>
        <th class="num">Qty (Curr)</th><th class="num">$ (Curr)</th>
        <th class="num">Qty Δ</th><th class="num">Qty %</th>
        <th class="num">$ Δ</th><th class="num">$ %</th>
      </tr></thead>
      <tbody>
        ${momCountryRows.map(r=>`<tr>
          <td>${escapeHtml(r.Country)}</td>
          <td class="num">${fmtNum(r.pQ)}</td><td class="num">${fmtUSD(r.pU)}</td>
          <td class="num">${fmtNum(r.cQ)}</td><td class="num">${fmtUSD(r.cU)}</td>
          <td class="num">${fmtNum(r.qd)}</td><td class="num">${fmtPct(r.qp)}</td>
          <td class="num">${fmtUSD(r.ud)}</td><td class="num">${fmtPct(r.up)}</td>
        </tr>`).join('')}
        <tr class="total-row">
          <td>Total</td>
          <td class="num">${fmtNum(mcTot.pQ)}</td><td class="num">${fmtUSD(mcTot.pU)}</td>
          <td class="num">${fmtNum(mcTot.cQ)}</td><td class="num">${fmtUSD(mcTot.cU)}</td>
          <td class="num">${fmtNum(mcTot.cQ-mcTot.pQ)}</td><td class="num">${fmtPct(momPct(mcTot.cQ-mcTot.pQ,mcTot.pQ))}</td>
          <td class="num">${fmtUSD(mcTot.cU-mcTot.pU)}</td><td class="num">${fmtPct(momPct(mcTot.cU-mcTot.pU,mcTot.pU))}</td>
        </tr>
      </tbody>
    </table>`;

  // Group D + E: model-grade summary, current and previous (corrected layout: title above header)
  const mgField = FILTERS.mgMode === 'model' ? 'ModelName' : 'ModelGrade';
  const mgLabel = FILTERS.mgMode === 'model' ? 'Model' : 'Model Grade';
  const buildMG = (rows) => {
    const g = groupBy(rows, r=>r[mgField]||'(blank)');
    return [...g.keys()]
      .sort((a,b)=>sumQty(g.get(b))-sumQty(g.get(a)))
      .map(k=>{ const v=g.get(k); return {label:k, Quantity:sumQty(v), USD:sumUSD(v)}; });
  };
  const mgCurr = buildMG(cc);
  const mgPrev = buildMG(cp);
  const mgTotals = (arr) => arr.reduce((a,r)=>({Quantity:a.Quantity+r.Quantity, USD:a.USD+r.USD}),{Quantity:0,USD:0});
  const mgCTot = mgTotals(mgCurr), mgPTot = mgTotals(mgPrev);
  const renderMGTable = (rows, tot) => `
    <table>
      <thead><tr><th>${mgLabel}</th><th class="num">Quantity</th><th class="num">Total(USD)</th></tr></thead>
      <tbody>
        ${rows.map(r=>`<tr><td>${escapeHtml(r.label)}</td><td class="num">${fmtNum(r.Quantity)}</td><td class="num">${fmtUSD(r.USD)}</td></tr>`).join('')}
        <tr class="total-row"><td>Total</td><td class="num">${fmtNum(tot.Quantity)}</td><td class="num">${fmtUSD(tot.USD)}</td></tr>
      </tbody>
    </table>`;

  return `
  <div class="excel-block">
    <h3>🏢 ${escapeHtml(company)}</h3>
    <div class="subtitle">Current: ${escapeHtml(currLabel)} · Compare: ${escapeHtml(prevLabel)}</div>
    <h4 style="font-size:13px;margin:0 0 6px">Order-level Summary</h4>
    ${gA}
    <h4 style="font-size:13px;margin:18px 0 6px">Country Breakdown — ${escapeHtml(currLabel)}</h4>
    <div class="cols-2">
      <div>${gB}</div>
      <div>${gC}</div>
    </div>
    <h4 style="font-size:13px;margin:18px 0 6px">Model Grade Summary</h4>
    <div class="cols-2">
      <div>
        <div class="subtitle" style="margin:0 0 6px">Current — ${escapeHtml(currLabel)}</div>
        ${renderMGTable(mgCurr, mgCTot)}
      </div>
      <div>
        <div class="subtitle" style="margin:0 0 6px">Previous — ${escapeHtml(prevLabel)}</div>
        ${renderMGTable(mgPrev, mgPTot)}
      </div>
    </div>
  </div>`;
}

// ---- Tabs ---------------------------------------------------
document.querySelectorAll('.tabs button').forEach(b => {
  b.addEventListener('click', () => {
    document.querySelectorAll('.tabs button').forEach(x=>x.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(x=>x.classList.remove('active'));
    b.classList.add('active');
    document.getElementById('tab-'+b.dataset.tab).classList.add('active');
    if (b.dataset.tab==='dashboard') { chartModels?.resize(); chartRegion?.resize(); }
  });
});

// ---- Filter panel: collapse + resize -----------------------
function setCollapsed(collapsed) {
  document.body.classList.toggle('filters-collapsed', collapsed);
  localStorage.setItem('filterCollapsed', collapsed ? '1' : '0');
  // Chart.js canvases need a resize after their container width changes
  setTimeout(() => { chartModels?.resize(); chartRegion?.resize(); }, 200);
}
document.getElementById('filterCollapseBtn').addEventListener('click', () => setCollapsed(true));
document.getElementById('filterExpandBtn').addEventListener('click', () => setCollapsed(false));

(function setupResizer(){
  const resizer = document.getElementById('resizer');
  let dragging = false, startX = 0, startW = 0;
  resizer.addEventListener('mousedown', e => {
    dragging = true;
    document.body.classList.add('resizing');
    startX = e.clientX;
    const cur = getComputedStyle(document.documentElement).getPropertyValue('--filter-width').trim();
    startW = parseInt(cur) || 280;
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const w = Math.max(180, Math.min(800, startW + (e.clientX - startX)));
    document.documentElement.style.setProperty('--filter-width', w + 'px');
  });
  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    document.body.classList.remove('resizing');
    const w = getComputedStyle(document.documentElement).getPropertyValue('--filter-width').trim();
    localStorage.setItem('filterWidth', w);
  });
})();

function restoreFilterPrefs() {
  const raw = localStorage.getItem('filterWidth');
  if (raw) {
    const n = parseInt(raw, 10);
    if (!isNaN(n)) document.documentElement.style.setProperty('--filter-width', Math.max(180, Math.min(800, n)) + 'px');
  }
  if (localStorage.getItem('filterCollapsed') === '1') setCollapsed(true);
}

// ---- Mode toggle + mode visibility --------------------------
document.querySelectorAll('#modeToggle button').forEach(b => b.addEventListener('click', () => {
  MODE = b.dataset.mode;
  document.querySelectorAll('#modeToggle button').forEach(x => x.classList.remove('active'));
  b.classList.add('active');
  applyModeVisibility();
  renderAll();
}));
function applyModeVisibility() {
  // Show/hide filter groups by mode
  document.querySelectorAll('[data-mode-show]').forEach(el => {
    const modes = el.dataset.modeShow.split(',');
    el.style.display = modes.includes(MODE) ? '' : 'none';
  });
  // Show/hide tab buttons by mode
  document.querySelectorAll('.tabs button[data-mode-show]').forEach(b => {
    const modes = b.dataset.modeShow.split(',');
    b.style.display = modes.includes(MODE) ? '' : 'none';
  });
  // Switch to first visible tab in the active mode
  const visibleTab = [...document.querySelectorAll('.tabs button[data-mode-show]')]
    .find(b => b.dataset.modeShow.split(',').includes(MODE));
  if (visibleTab) visibleTab.click();
}

// ---- Currency rates editor ---------------------------------
function populateRateGrid() {
  const grid = document.getElementById('rateGrid');
  if (!grid) return;
  grid.innerHTML = '';
  for (const [k, v] of Object.entries(RATES)) {
    const lab = document.createElement('label');
    lab.innerHTML = `<span>${k}</span>`;
    const inp = document.createElement('input');
    inp.type = 'number'; inp.step = '0.0001'; inp.min = '0'; inp.value = v; inp.dataset.curr = k;
    inp.addEventListener('change', () => {
      const n = parseFloat(inp.value);
      if (!isNaN(n) && n > 0) {
        RATES[k] = n;
        for (const r of RAW) r.rate = (RATES[(r.Currency||'USD').toUpperCase()] ?? BACKUP_RATES[(r.Currency||'USD').toUpperCase()] ?? 1);
        recomputeDerived(); renderAll();
      }
    });
    lab.appendChild(inp);
    grid.appendChild(lab);
  }
}
function recomputeDerived() {
  for (const r of RAW) {
    const rate = r.rate ?? ((RATES[(r.Currency||'USD').toUpperCase()] ?? BACKUP_RATES[(r.Currency||'USD').toUpperCase()] ?? 1));
    r.USD_Total = r.Total * rate;
    r.perUnitUsd = r.perUnit * rate;
  }
}
document.getElementById('updateRates')?.addEventListener('click', () => {
  recomputeDerived(); renderAll();
});
document.getElementById('resetRates')?.addEventListener('click', e => {
  e.preventDefault();
  Object.assign(RATES, { USD:1.0, AUD:0.6500, EUR:1.1628, AED:0.2723, GBP:1.3514, JPY:0.00661, CAD:0.7143 });
  populateRateGrid();
  for (const r of RAW) r.rate = (RATES[(r.Currency||'USD').toUpperCase()] ?? 1);
  recomputeDerived(); renderAll();
});

// ---- Excluded company patterns ------------------------------
function renderExcludedChips() {
  const host = document.getElementById('excludedChips');
  if (!host) return;
  host.innerHTML = '';
  for (const p of EXCLUDED_PATTERNS) {
    const chip = document.createElement('span');
    chip.className = 'chip active';
    chip.textContent = p + ' ';
    const x = document.createElement('span');
    x.textContent = '×';
    x.style.cursor = 'pointer';
    x.style.marginLeft = '4px';
    x.onclick = () => { EXCLUDED_PATTERNS = EXCLUDED_PATTERNS.filter(x => x !== p); renderExcludedChips(); };
    chip.appendChild(x);
    host.appendChild(chip);
  }
  // Re-filter the data immediately when patterns change
  const before = RAW.length;
  // Apply exclusions to current RAW (in-place) — rows that were previously filtered are NOT recoverable here,
  // but newly added patterns will drop rows going forward. To get a clean re-filter we rebuild from the
  // last-loaded source if available.
  if (typeof rebuildFilteredRaw === 'function') {
    rebuildFilteredRaw();
    renderAll();
  }
}
document.getElementById('addExclusion')?.addEventListener('click', () => {
  const inp = document.getElementById('exclusionInput');
  const v = inp.value.trim();
  if (v && !EXCLUDED_PATTERNS.some(p => p.toLowerCase() === v.toLowerCase())) {
    EXCLUDED_PATTERNS.push(v);
    inp.value = '';
    renderExcludedChips();
  }
});
document.getElementById('exclusionInput')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); document.getElementById('addExclusion').click(); }
});

// ---- Pricing-mode chip filters ------------------------------
document.querySelectorAll('#dateChips .chip').forEach(c => c.addEventListener('click', () => {
  document.querySelectorAll('#dateChips .chip').forEach(x => x.classList.remove('active'));
  c.classList.add('active');
  PRICING_FILTERS.dateRange = c.dataset.d;
  renderAll();
}));
document.querySelectorAll('#priceChips .chip').forEach(c => c.addEventListener('click', () => {
  document.querySelectorAll('#priceChips .chip').forEach(x => x.classList.remove('active'));
  c.classList.add('active');
  PRICING_FILTERS.priceLock = c.dataset.pl;
  renderAll();
}));
document.querySelectorAll('#categoryChips .chip').forEach(c => c.addEventListener('click', () => {
  const v = c.dataset.cat;
  if (v === 'All') {
    document.querySelectorAll('#categoryChips .chip').forEach(x => x.classList.remove('active'));
    c.classList.add('active');
    PRICING_FILTERS.categories = new Set(['All']);
  } else {
    document.querySelector('#categoryChips .chip[data-cat="All"]').classList.remove('active');
    if (c.classList.contains('active')) { c.classList.remove('active'); PRICING_FILTERS.categories.delete(v); }
    else { c.classList.add('active'); PRICING_FILTERS.categories.add(v); }
    if (PRICING_FILTERS.categories.size === 0) {
      document.querySelector('#categoryChips .chip[data-cat="All"]').classList.add('active');
      PRICING_FILTERS.categories = new Set(['All']);
    }
  }
  renderAll();
}));
const modelInputEl = document.getElementById('modelInput');
modelInputEl?.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    const v = modelInputEl.value.trim();
    if (v && !PRICING_FILTERS.tokens.includes(v)) PRICING_FILTERS.tokens.push(v);
    modelInputEl.value = '';
    renderModelTokens();
    renderAll();
  }
});
function renderModelTokens() {
  const host = document.getElementById('modelTokens');
  if (!host) return;
  host.innerHTML = '';
  for (const t of PRICING_FILTERS.tokens) {
    const chip = document.createElement('span');
    chip.className = 'chip active';
    chip.textContent = t + ' ';
    const x = document.createElement('span');
    x.textContent = '×';
    x.style.cursor = 'pointer';
    x.style.marginLeft = '4px';
    x.onclick = () => {
      PRICING_FILTERS.tokens = PRICING_FILTERS.tokens.filter(x => x !== t);
      renderModelTokens(); renderAll();
    };
    chip.appendChild(x);
    host.appendChild(chip);
  }
}
renderExcludedChips();

// ---- Rebuild RAW from saved pre-filter source for exclusion edits ----
// Cache the post-enrichment RAW rows BEFORE exclusions are applied so the user can
// edit patterns and re-filter without re-uploading. Cleared whenever new data is loaded.
function rebuildFilteredRaw() {
  if (!RAW_PRE_EXCLUSION.length) return;
  const excluded = EXCLUDED_PATTERNS.map(p => p.toLowerCase());
  // Restore fully-enriched rows from the cache and re-derive anything that needs
  // the live rate table.
  RAW = RAW_PRE_EXCLUSION
    .filter(r => !excluded.some(p => r.CompanyName.toLowerCase().includes(p)))
    .map(r => {
      const copy = {...r};
      // Recompute perUnit / perUnitUsd in case rates have been edited since cache
      copy.perUnit = copy.Quantity > 0 ? copy.Total / copy.Quantity : 0;
      const rate = copy.rate ?? ((RATES[(copy.Currency||'USD').toUpperCase()] ?? BACKUP_RATES[(copy.Currency||'USD').toUpperCase()] ?? 1));
      copy.rate = rate;
      copy.perUnitUsd = copy.perUnit * rate;
      copy.USD_Total = copy.Total * rate;
      return copy;
    });
}

// ---- Per-tab exports ---------------------------------------
// Each export reads the live `activeRows()` so the file matches what's on screen,
// and respects the current mode + period + filters.

// Last ASP — follows the global company picker (no tab-local scope override).
// splitByRegion: false → one sheet with Country as a column;
//                true  → one sheet per country (no Country column).
function buildAspWorkbook(splitByRegion) {
  const scoped = activeRows();
  const groups = {};
  for (const r of scoped) {
    const c = r.Country || 'Unknown', mg = r.ModelGrade;
    if (!groups[c]) groups[c] = {};
    if (!groups[c][mg] || r._date > groups[c][mg]._date) groups[c][mg] = r;
  }
  const wb = XLSX.utils.book_new();
  const countries = Object.keys(groups).sort();
  if (splitByRegion) {
    // One sheet per country
    for (const c of countries) {
      const entries = Object.values(groups[c]).sort((a,b) => a.ModelGrade.localeCompare(b.ModelGrade));
      const data = entries.map(e => ({
        ModelGrade: e.ModelGrade,
        'Last ASP (USD)': +e.perUnitUsd.toFixed(2),
        Company: e.CompanyName, Date: e._date.toISOString().split('T')[0]
      }));
      const ws = XLSX.utils.json_to_sheet(data);
      ws['!cols'] = [{wch:24},{wch:14},{wch:18},{wch:12}];
      XLSX.utils.book_append_sheet(wb, ws, c.substring(0, 31));
    }
  } else {
    // Single sheet, Country as a column
    const rowsArr = [];
    for (const c of countries) {
      const entries = Object.values(groups[c]).sort((a,b) => a.ModelGrade.localeCompare(b.ModelGrade));
      for (const e of entries) {
        rowsArr.push({
          Country: c, ModelGrade: e.ModelGrade,
          'Last ASP (USD)': +e.perUnitUsd.toFixed(2),
          Company: e.CompanyName, Date: e._date.toISOString().split('T')[0]
        });
      }
    }
    const ws = XLSX.utils.json_to_sheet(rowsArr);
    ws['!cols'] = [{wch:12},{wch:24},{wch:14},{wch:18},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, 'LastASP');
  }
  return wb;
}

// Sales by Model — flat rows.
// splitByRegion: false → one sheet with Country as a column;
//                true  → one sheet per country (no Country column).
function buildSalesByModelWorkbook(splitByRegion) {
  const rows = activeRows();
  const byCountry = {};
  for (const r of rows) {
    const c = r.Country || 'Unknown';
    if (!byCountry[c]) byCountry[c] = [];
    byCountry[c].push(r);
  }
  const wb = XLSX.utils.book_new();
  const sortFn = (a,b) => {
    const m = a.ModelGrade.localeCompare(b.ModelGrade);
    if (m !== 0) return m;
    return b._date - a._date;
  };
  if (splitByRegion) {
    for (const c of Object.keys(byCountry).sort()) {
      const items = byCountry[c].slice().sort(sortFn);
      const sheet = XLSX.utils.json_to_sheet(items.map(r => ({
        ModelGrade: r.ModelGrade,
        Date: r._date.toISOString().split('T')[0],
        'ASP (USD)': +r.perUnitUsd.toFixed(2),
        Quantity: r.Quantity,
        Company: r.CompanyName
      })));
      sheet['!cols'] = [{wch:24},{wch:12},{wch:12},{wch:10},{wch:18}];
      XLSX.utils.book_append_sheet(wb, sheet, c.substring(0, 31));
    }
  } else {
    // Single sheet with Country as a column
    const all = [];
    for (const c of Object.keys(byCountry).sort()) {
      const items = byCountry[c].slice().sort(sortFn);
      for (const r of items) {
        all.push({
          Country: c,
          ModelGrade: r.ModelGrade,
          Date: r._date.toISOString().split('T')[0],
          'ASP (USD)': +r.perUnitUsd.toFixed(2),
          Quantity: r.Quantity,
          Company: r.CompanyName
        });
      }
    }
    const sheet = XLSX.utils.json_to_sheet(all);
    sheet['!cols'] = [{wch:12},{wch:24},{wch:12},{wch:12},{wch:10},{wch:18}];
    XLSX.utils.book_append_sheet(wb, sheet, 'SalesByModel');
  }
  return wb;
}

// Raw data — every column the app shows.
function buildRawWorkbook() {
  const rows = activeRows();
  const cols = ['ID','CreatedAt','CompanyName','Country','SKU','ModelGrade','Category','Condition','OrderStatus','SaleType','Currency','Quantity','Total','perUnitUsd','USD_Total'];
  const data = rows.map(r => {
    const o = {};
    for (const c of cols) {
      let v = r[c];
      if (v instanceof Date) v = v.toISOString().split('T')[0];
      else if (typeof v === 'number') v = ['perUnitUsd','USD_Total','perUnit','rate'].includes(c) ? +v.toFixed(2) : v;
      o[c] = v ?? '';
    }
    return o;
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = cols.map(c => ({wch: ['CompanyName','ModelGrade','Category','SaleType'].includes(c) ? 22 : (c==='SKU'?28:14)}));
  XLSX.utils.book_append_sheet(wb, ws, 'Raw');
  return wb;
}

// Country breakdown — current/compare MoM by country.
function buildCountryBreakdownWorkbook() {
  const rows = activeRows();
  const {curr, prev, currLabel, prevLabel} = splitByPeriod(rows);
  const cG = groupBy(curr, r => r.Country || '(blank)');
  const pG = groupBy(prev, r => r.Country || '(blank)');
  const keys = [...new Set([...cG.keys(), ...pG.keys()])].sort();
  const data = keys.map(k => {
    const c = cG.get(k) || [], p = pG.get(k) || [];
    const cQ = sumQty(c), pQ = sumQty(p), cU = sumUSD(c), pU = sumUSD(p);
    const total = sumUSD(curr);
    return {
      Country: k,
      'Qty (Curr)': cQ, 'Qty (Prev)': pQ,
      'Qty Δ': cQ - pQ,
      'Qty %': pQ ? (cQ - pQ) / pQ : null,
      'USD (Curr)': +cU.toFixed(2), 'USD (Prev)': +pU.toFixed(2),
      'USD Δ': +(cU - pU).toFixed(2),
      'USD %': pU ? (cU - pU) / pU : null,
      '% of total': total ? cU / total : 0
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:14},{wch:10},{wch:10},{wch:10},{wch:10},{wch:14},{wch:14},{wch:12},{wch:10},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws, 'Country');
  return wb;
}

// Model-grade summary — current and previous as two sheets.
function buildModelGradeWorkbook() {
  const rows = activeRows();
  const {curr, prev, currLabel, prevLabel} = splitByPeriod(rows);
  const field = FILTERS.mgMode === 'model' ? 'ModelName' : 'ModelGrade';
  const lbl = FILTERS.mgMode === 'model' ? 'Model' : 'Model Grade';
  const build = (src, label) => {
    const g = groupBy(src, r => r[field] || '(blank)');
    return [...g.keys()].sort((a,b) => sumQty(g.get(b)) - sumQty(g.get(a)))
      .map(k => ({[lbl]: k, Quantity: sumQty(g.get(k)), 'Total (USD)': +sumUSD(g.get(k)).toFixed(2)}));
  };
  const wb = XLSX.utils.book_new();
  const wsCurr = XLSX.utils.json_to_sheet(build(curr, currLabel));
  wsCurr['!cols'] = [{wch:24},{wch:12},{wch:14}];
  XLSX.utils.book_append_sheet(wb, wsCurr, 'Current');
  const wsPrev = XLSX.utils.json_to_sheet(build(prev, prevLabel));
  wsPrev['!cols'] = [{wch:24},{wch:12},{wch:14}];
  XLSX.utils.book_append_sheet(wb, wsPrev, 'Previous');
  return wb;
}

// Excel layout — the rendered Excel view, exported as a single sheet that mirrors what's on screen.
function buildExcelLayoutWorkbook() {
  // Reuse the existing buildAllCompanyMoMWorkbook (summary across all companies in active filter)
  // and the per-company block builder for SELECTED_COMPANY (or all if empty).
  const rows = activeRows();
  const {curr, prev, currLabel, prevLabel} = splitByPeriod(rows);
  const wb = XLSX.utils.book_new();
  // MoM summary sheet
  const cG = groupBy(curr, r => r.CompanyName);
  const pG = groupBy(prev, r => r.CompanyName);
  const keys = SELECTED_COMPANY ? [SELECTED_COMPANY] : [...new Set([...cG.keys(),...pG.keys()])].sort();
  const momRows = keys.map(k => {
    const c = cG.get(k) || [], p = pG.get(k) || [];
    return {
      Company: k,
      'Qty (Prev)': sumQty(p), '$ (Prev)': +sumUSD(p).toFixed(2),
      'Qty (Curr)': sumQty(c), '$ (Curr)': +sumUSD(c).toFixed(2),
      'Qty Δ': sumQty(c) - sumQty(p), 'Qty %': sumQty(p) ? (sumQty(c) - sumQty(p)) / sumQty(p) : null,
      '$ Δ': +(sumUSD(c) - sumUSD(p)).toFixed(2), '$ %': sumUSD(p) ? (sumUSD(c) - sumUSD(p)) / sumUSD(p) : null
    };
  });
  const wsMom = XLSX.utils.json_to_sheet(momRows);
  wsMom['!cols'] = [{wch:30},{wch:12},{wch:14},{wch:12},{wch:14},{wch:10},{wch:10},{wch:14},{wch:10}];
  XLSX.utils.book_append_sheet(wb, wsMom, 'MoM');
  // Per-company sheets using the existing builder
  for (const company of keys) {
    const wbCo = buildCompanyWorkbook(company, cG.get(company) || [], pG.get(company) || [], currLabel, prevLabel);
    // Copy the single sheet into the combined workbook
    const sheetName = wbCo.SheetNames[0];
    const sheet = wbCo.Sheets[sheetName];
    XLSX.utils.book_append_sheet(wb, sheet, sheetName.substring(0, 31));
  }
  return wb;
}

function downloadWorkbook(wb, baseName) {
  const stamp = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `${baseName}_${stamp}.xlsx`);
}

// Per-tab export button listeners
// ASP + Sales-by-Model export buttons (each offers all-in-one-tab OR split-by-region).
// Filename scope tag follows the global company picker: empty → "all-companies".
function aspScopeTag() {
  return SELECTED_COMPANY ? safeFilename(SELECTED_COMPANY) : 'all-companies';
}
document.getElementById('exportAspAll')?.addEventListener('click', () => {
  if (!RAW.length) return;
  downloadWorkbook(buildAspWorkbook(false), `LastASP_${aspScopeTag()}_all-regions`);
});
document.getElementById('exportAspSplit')?.addEventListener('click', () => {
  if (!RAW.length) return;
  downloadWorkbook(buildAspWorkbook(true), `LastASP_${aspScopeTag()}_by-region`);
});
document.getElementById('exportSalesModelAll')?.addEventListener('click', () => {
  if (!RAW.length) return;
  downloadWorkbook(buildSalesByModelWorkbook(false), 'SalesByModel_all-regions');
});
document.getElementById('exportSalesModelSplit')?.addEventListener('click', () => {
  if (!RAW.length) return;
  downloadWorkbook(buildSalesByModelWorkbook(true), 'SalesByModel_by-region');
});
document.getElementById('exportRaw')?.addEventListener('click', () => {
  if (!RAW.length) return;
  downloadWorkbook(buildRawWorkbook(), 'RawData');
});
document.getElementById('exportCountry')?.addEventListener('click', () => {
  if (!RAW.length) return;
  const scope = SELECTED_COMPANY ? safeFilename(SELECTED_COMPANY) : 'all-companies';
  downloadWorkbook(buildCountryBreakdownWorkbook(), `Country_${scope}`);
});
document.getElementById('exportModelGrade')?.addEventListener('click', () => {
  if (!RAW.length) return;
  const scope = SELECTED_COMPANY ? safeFilename(SELECTED_COMPANY) : 'all-companies';
  downloadWorkbook(buildModelGradeWorkbook(), `ModelGrade_${scope}`);
});
document.getElementById('exportExcelLayout')?.addEventListener('click', () => {
  if (!RAW.length) return;
  const scope = SELECTED_COMPANY ? safeFilename(SELECTED_COMPANY) : 'all-companies';
  downloadWorkbook(buildExcelLayoutWorkbook(), `ExcelLayout_${scope}`);
});

// ---- Export to Excel (.zip) --------------------------------
document.getElementById('exportBtn').addEventListener('click', exportZip);

async function exportZip() {
  showOverlay('Building Excel files…');
  await new Promise(r=>setTimeout(r,30));
  try {
    const rows = activeRows();
    const {curr, prev, currLabel, prevLabel} = splitByPeriod(rows);
    const zip = new JSZip();
    const folder = zip.folder('CompanyReports');

    // 1) All-company MoM summary
    const allMoM = buildAllCompanyMoMWorkbook(curr, prev, currLabel, prevLabel);
    folder.file(`All_Company_MoM_Summary_${currLabel}_vs_${prevLabel}.xlsx`, XLSX.write(allMoM,{type:'array',bookType:'xlsx'}));

    // 2) Per-company files (matches Python block layout)
    const compCurr = groupBy(curr, r=>r.CompanyName||'(blank)');
    const compPrev = groupBy(prev, r=>r.CompanyName||'(blank)');
    const keys = [...new Set([...compCurr.keys(), ...compPrev.keys()])].sort();
    for (const company of keys) {
      const safe = safeFilename(company);
      const wb = buildCompanyWorkbook(company, compCurr.get(company)||[], compPrev.get(company)||[], currLabel, prevLabel);
      folder.file(`${safe}_Summary_${currLabel}_vs_${prevLabel}.xlsx`, XLSX.write(wb,{type:'array',bookType:'xlsx'}));
    }

    const blob = await zip.generateAsync({type:'blob'});
    downloadBlob(blob, `CompanyReports_${currLabel}_vs_${prevLabel}.zip`);
  } catch (err) {
    console.error(err);
    alert('Export failed:\n'+err.message);
  } finally {
    hideOverlay();
  }
}

function buildAllCompanyMoMWorkbook(curr, prev, currLabel, prevLabel) {
  const compCurr = groupBy(curr, r=>r.CompanyName||'(blank)');
  const compPrev = groupBy(prev, r=>r.CompanyName||'(blank)');
  const keys = [...new Set([...compCurr.keys(), ...compPrev.keys()])].sort();
  const data = keys.map(k=>{
    const c=compCurr.get(k)||[], p=compPrev.get(k)||[];
    return {
      CompanyName:k,
      'Quantity (Prev)':sumQty(p), '$ (Prev)':round2(sumUSD(p)),
      'Quantity (Curr)':sumQty(c), '$ (Curr)':round2(sumUSD(c)),
      'Qty MoM Δ':sumQty(c)-sumQty(p),
      'Qty MoM %':sumQty(p) ? (sumQty(c)-sumQty(p))/sumQty(p) : null,
      '$ MoM Δ':round2(sumUSD(c)-sumUSD(p)),
      '$ MoM %':sumUSD(p) ? (sumUSD(c)-sumUSD(p))/sumUSD(p) : null,
    };
  });
  // total row
  data.push({
    CompanyName:'Total',
    'Quantity (Prev)':data.reduce((s,r)=>s+(r['Quantity (Prev)']||0),0),
    '$ (Prev)':round2(data.reduce((s,r)=>s+(r['$ (Prev)']||0),0)),
    'Quantity (Curr)':data.reduce((s,r)=>s+(r['Quantity (Curr)']||0),0),
    '$ (Curr)':round2(data.reduce((s,r)=>s+(r['$ (Curr)']||0),0)),
    'Qty MoM Δ':data.reduce((s,r)=>s+(r['Qty MoM Δ']||0),0),
    'Qty MoM %': null,
    '$ MoM Δ':round2(data.reduce((s,r)=>s+(r['$ MoM Δ']||0),0)),
    '$ MoM %': null,
  });
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:30},{wch:14},{wch:14},{wch:14},{wch:14},{wch:12},{wch:12},{wch:14},{wch:12}];
  // number formats
  for (let r=1; r<=data.length; r++) {
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:1}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:2}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:3}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:4}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:5}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:6}), '0.0%');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:7}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:8}), '0.0%');
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Summary');
  return wb;
}

function buildCompanyWorkbook(company, cc, cp, currLabel, prevLabel) {
  const wb = XLSX.utils.book_new();
  const ws = {};
  const mgField = FILTERS.mgMode === 'model' ? 'ModelName' : 'ModelGrade';
  const mgLabel = FILTERS.mgMode === 'model' ? 'Model' : 'Model Grade';

  // Block A: order-level summary at (col=0, row=0). Header at row 0, data from row 1.
  const aHeader = ['ID','SaleType','Location','CreatedAt','OrderStatus','Working(Qty)','Damaged(Qty)','Total(Qty)','Total(USD)'];
  const G_curr = groupBy(cc, r=>['ID','SaleType','Location','CreatedAt','OrderStatus'].map(k=>r[k]??'').join('||'));
  const aRows = [...G_curr.keys()].sort().map(k=>{
    const cs = G_curr.get(k)||[];
    const parts = k.split('||');
    return [...parts, qtyByCond(cs,'Working'), qtyByCond(cs,'Damaged'), sumQty(cs), round2(sumUSD(cs))];
  });
  // total
  const aTot = aRows.reduce((acc,r)=>({
    cW: acc.cW + (r[5]||0),
    cD: acc.cD + (r[6]||0),
    cT: acc.cT + (r[7]||0),
    cU: acc.cU + (r[8]||0),
  }), {cW:0,cD:0,cT:0,cU:0});
  const aRowsWithTotal = [...aRows, ['','','','','Total', aTot.cW, aTot.cD, aTot.cT, round2(aTot.cU)]];
  putAOA(ws, [aHeader, ...aRowsWithTotal], 0, 0);

  // Block B: country breakdown at (col=0, row= aRows.length+4 (3-row gap + 1))
  const bHeader = ['Country','Quantity','Total(USD)','% of total'];
  const ccG = groupBy(cc, r=>r.Country||'(blank)');
  const totalU = sumUSD(cc);
  const bRows = [...ccG.keys()].sort().map(ct=>{
    const c = ccG.get(ct)||[];
    return [ct, sumQty(c), round2(sumUSD(c)), totalU?sumUSD(c)/totalU:0];
  });
  const bTot = bRows.reduce((acc,r)=>({q:acc.q+(r[1]||0), u:acc.u+(r[2]||0)}), {q:0,u:0});
  const bRowsWithTotal = [...bRows, ['Total', bTot.q, round2(bTot.u), '']];
  const bStartRow = aRowsWithTotal.length + 3; // 3-row gap
  putAOA(ws, [bHeader, ...bRowsWithTotal], bStartRow, 0);

  // Block C: country MoM at (col=5 (3-col gap), row= bStartRow)
  const cHeader = ['Country','Quantity (Prev)','$ (Prev)','Quantity (Curr)','$ (Curr)','Qty MoM Δ','Qty MoM %','$ MoM Δ','$ MoM %'];
  const cpG = groupBy(cp, r=>r.Country||'(blank)');
  const cKeys = [...new Set([...ccG.keys(), ...cpG.keys()])].sort();
  const cRows = cKeys.map(ct=>{
    const c=ccG.get(ct)||[], p=cpG.get(ct)||[];
    const cQ=sumQty(c),pQ=sumQty(p),cU=sumUSD(c),pU=sumUSD(p);
    return [ct, pQ, round2(pU), cQ, round2(cU), cQ-pQ, pQ?(cQ-pQ)/pQ:null, round2(cU-pU), pU?(cU-pU)/pU:null];
  });
  const cTot = cRows.reduce((acc,r)=>({pQ:acc.pQ+(r[1]||0),pU:acc.pU+(r[2]||0),cQ:acc.cQ+(r[3]||0),cU:acc.cU+(r[4]||0)}), {pQ:0,pU:0,cQ:0,cU:0});
  cRows.push(['Total', cTot.pQ, round2(cTot.pU), cTot.cQ, round2(cTot.cU), cTot.cQ-cTot.pQ, cTot.pQ?(cTot.cQ-cTot.pQ)/cTot.pQ:null, round2(cTot.cU-cTot.pU), cTot.pU?(cTot.cU-cTot.pU)/cTot.pU:null]);
  putAOA(ws, [cHeader, ...cRows], bStartRow, 5);

  // Block D: model grade summary current at (col=0, row= bStartRow)
  const dHeader = [mgLabel, 'Quantity', 'Total(USD)'];
  const dRows = [...ccG.keys()].length ? [] : [];  // placeholder
  const mgCG = groupBy(cc, r=>r[mgField]||'(blank)');
  const dData = [...mgCG.keys()].sort((a,b)=>sumQty(mgCG.get(b))-sumQty(mgCG.get(a)))
    .map(k=>{ const v=mgCG.get(k); return [k, sumQty(v), round2(sumUSD(v))]; });
  const dTot = dData.reduce((acc,r)=>({q:acc.q+(r[1]||0),u:acc.u+(r[2]||0)}),{q:0,u:0});
  dData.push(['Total', dTot.q, round2(dTot.u)]);
  putAOA(ws, [dHeader, ...dData], bStartRow, 0);

  // Block E: model grade summary previous, 4 rows below D
  const eStartRow = bStartRow + dData.length + 4;
  const mgPG = groupBy(cp, r=>r[mgField]||'(blank)');
  const eData = [...mgPG.keys()].sort((a,b)=>sumQty(mgPG.get(b))-sumQty(mgPG.get(a)))
    .map(k=>{ const v=mgPG.get(k); return [k, sumQty(v), round2(sumUSD(v))]; });
  const eTot = eData.reduce((acc,r)=>({q:acc.q+(r[1]||0),u:acc.u+(r[2]||0)}),{q:0,u:0});
  eData.push(['Total', eTot.q, round2(eTot.u)]);
  putAOA(ws, [dHeader, ...eData], eStartRow, 0);

  // Compute final range
  const lastRow = Math.max(aRowsWithTotal.length, bStartRow + bRowsWithTotal.length, eStartRow + eData.length + 1);
  ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:lastRow, c:Math.max(8, 5+cHeader.length-1)}});
  // column widths
  ws['!cols'] = [
    {wch:16}, {wch:12}, {wch:14}, {wch:14}, {wch:14}, {wch:12}, {wch:12}, {wch:12}, {wch:14},
    {wch:14}, {wch:12}, {wch:12}, {wch:12}, {wch:12}
  ];
  // number formats for blocks A, B, C
  for (let r=1; r<aRowsWithTotal.length+1; r++) {
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:5}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:6}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:7}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:8}), '"$"#,##0.00');
  }
  for (let r=bStartRow+1; r<bStartRow+1+bRowsWithTotal.length; r++) {
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:1}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:2}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:3}), '0.00%');
  }
  for (let r=bStartRow+1; r<bStartRow+1+cRows.length; r++) {
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:1}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:2}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:3}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:4}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:5}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:6}), '0.0%');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:7}), '"$"#,##0.00');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:8}), '0.0%');
  }
  // D and E number formats
  for (let r=bStartRow+1; r<bStartRow+1+dData.length; r++) {
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:1}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:2}), '"$"#,##0.00');
  }
  for (let r=eStartRow+1; r<eStartRow+1+eData.length; r++) {
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:1}), '#,##0');
    setCellNumFmt(ws, XLSX.utils.encode_cell({r:r,c:2}), '"$"#,##0.00');
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Summary');
  return wb;
}

function putAOA(ws, rows, startRow, startCol) {
  for (let r=0; r<rows.length; r++) {
    for (let c=0; c<rows[r].length; c++) {
      const addr = XLSX.utils.encode_cell({r:startRow+r, c:startCol+c});
      ws[addr] = { v: rows[r][c] ?? '', t: typeof rows[r][c] === 'number' ? 'n' : 's' };
    }
  }
}
function setCellNumFmt(ws, addr, fmt) {
  if (ws[addr]) ws[addr].z = fmt;
}
function round2(v){ return Math.round(v*100)/100; }

// ---- Snapshot -----------------------------------------------
let PAGE_HTML = null;
function capturePageHtml() {
  if (PAGE_HTML) return;
  PAGE_HTML = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
}

document.getElementById('snapshotBtn').addEventListener('click', downloadSnapshot);

function downloadSnapshot() {
  if (!PAGE_HTML) capturePageHtml();
  showOverlay('Building snapshot…');
  // Allow paint
  setTimeout(async () => {
    try {
      const data = {
        rows: RAW,
        skuNameMap: SKU_NAME_MAP,
        state: snapshotState(),
      };
      const json = JSON.stringify(data);
      const dataScript = `<script id="__snapshotData" type="application/json">${json.replace(/<\/script>/gi, '<\\/script>')}<\/script>`;
      // Inject the data script right after the <head> opening tag.
      let html = PAGE_HTML.replace(/<head>/i, '<head>\n' + dataScript + '\n');
      const company = SELECTED_COMPANY || 'company';
      const cm = document.getElementById('currMonth').value;
      const pm = document.getElementById('prevMonth').value;
      const fname = `snapshot_${safeFilename(company)}_${cm}_vs_${pm}.html`;
      const blob = new Blob([html], {type:'text/html'});
      downloadBlob(blob, fname);
    } catch (err) {
      console.error(err);
      alert('Snapshot failed:\n'+err.message);
    } finally {
      hideOverlay();
    }
  }, 30);
}

// ---- Bootstrap ---------------------------------------------
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 1000);
}

function showOverlay(msg){ document.getElementById('overlayMsg').textContent = msg; document.getElementById('overlay').classList.add('show'); }
function hideOverlay(){ document.getElementById('overlay').classList.remove('show'); }

// ---- Snapshot auto-bootstrap --------------------------------
if (IS_SNAPSHOT) {
  // data is already parsed (or needs to be re-parsed from JSON)
  // RAW rows are stored as-is; we need to make sure _date is a Date object
  if (RAW.length && typeof RAW[0]._date === 'string') {
    for (const r of RAW) {
      if (r._date && typeof r._date === 'string') r._date = new Date(r._date);
    }
  }
  buildAll({ embedded: true, rows: RAW, skuNameMap: SKU_NAME_MAP, state: SNAPSHOT_DATA.state });
}
