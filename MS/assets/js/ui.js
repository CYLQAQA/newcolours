/* ui.js — Margin Sheet Tool v2 presentation layer.
 * Purely additive: listens to the existing CustomEvents already dispatched by
 * aggregator.js / app.js and reflects state into the new UI chrome. It does NOT
 * change any business logic, file parsing, or Excel export. The two logic files
 * are untouched. This module only reads already-computed results and re-renders.
 */
(() => {
  'use strict';
  const $ = (id) => document.getElementById(id);
  if (!window.MarginUI) window.MarginUI = {};

  /* ---------- Theme toggle ---------- */
  const themeBtn = $('themeToggle');
  const storedTheme = (() => { try { return localStorage.getItem('mst-theme'); } catch (e) { return null; } })();
  if (storedTheme === 'light' || storedTheme === 'dark') document.documentElement.setAttribute('data-theme', storedTheme);
  if (themeBtn) themeBtn.addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try { localStorage.setItem('mst-theme', next); } catch (e) {}
  });

  /* ---------- File-card visuals ---------- */
  // Map of native input id -> { card, name, meta, replace } for styled file cards.
  const FILE_CARDS = ['sharedMappingFile', 'lastAspInput', 'summaryFile', 'reqFile', 'priceFile'];

  function humanSize(bytes) {
    if (!bytes && bytes !== 0) return '';
    const u = ['B', 'KB', 'MB', 'GB'];
    let i = 0, v = bytes;
    while (v >= 1024 && i < u.length - 1) { v /= 1024; i++; }
    return (v >= 10 || i === 0 ? Math.round(v) : v.toFixed(1)) + ' ' + u[i];
  }

  function setCardLoaded(id, file, metaOverride) {
    const input = $(id); if (!input) return;
    const card = document.querySelector(`.file-card[for="${id}"]`);
    if (!card) return;
    const nameEl = card.querySelector('.fc-name');
    const metaEl = card.querySelector('.fc-meta');
    const repEl = card.querySelector('.fc-replace');
    if (file) {
      card.classList.add('is-loaded');
      if (nameEl) nameEl.textContent = file.name;
      if (metaEl) metaEl.textContent = metaOverride || humanSize(file.size);
      if (repEl) repEl.style.display = '';
    } else {
      card.classList.remove('is-loaded');
      if (repEl) repEl.style.display = 'none';
    }
  }

  FILE_CARDS.forEach((id) => {
    const input = $(id);
    if (!input) return;
    input.addEventListener('change', () => {
      const f = input.files && input.files[0];
      if (f) setCardLoaded(id, f);
    });
  });

  // The mapping file's meta (key/group counts) comes from the status text app.js writes.
  // Mirror it into the card meta line for a richer loaded state.
  const mappingStatus = $('mappingStatus');
  if (mappingStatus && $('sharedMappingFile')) {
    const obs = new MutationObserver(() => {
      const f = $('sharedMappingFile').files && $('sharedMappingFile').files[0];
      if (f && mappingStatus.classList.contains('ok')) {
        setCardLoaded('sharedMappingFile', f, mappingStatus.textContent);
      }
    });
    obs.observe(mappingStatus, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] });
  }

  // Last ASP master — aggregator alerts on load; we can't easily read the counts, so show file size + a friendly note.
  // (Kept simple to avoid duplicating parsing logic.)

  /* ---------- Step status pills ---------- */
  function setPill(el, state, text) {
    if (!el) return;
    el.className = 'status-pill ' + (state || '');
    const t = el.querySelector('.pill-text');
    if (t && text) t.textContent = text;
  }
  const step1Pill = $('step1Pill'), step2Pill = $('step2Pill');
  const step1Card = $('step1Card'), step2Card = $('step2Card');

  function step1State(state, text) {
    setPill(step1Pill, state, text);
    step1Card.classList.toggle('is-complete', state === 'complete');
    step1Card.classList.toggle('is-error', state === 'error');
  }
  function step2State(state, text) {
    setPill(step2Pill, state, text);
    step2Card.classList.toggle('is-complete', state === 'complete');
    step2Card.classList.toggle('is-error', state === 'error');
  }

  /* ---------- Summary result card ---------- */
  const resultMount = $('summaryResultMount');
  const accordion = $('outputAccordion');
  const accordionHead = $('outputAccordionHead');
  const accordionMeta = $('outputAccordionMeta');

  function countSummary(summaryByRegion) {
    const regions = summaryByRegion ? Object.keys(summaryByRegion) : [];
    const keys = new Set();
    regions.forEach((r) => Object.keys(summaryByRegion[r] || {}).forEach((k) => keys.add(k)));
    return { regionCount: regions.length, rowCount: keys.size, regions: regions.slice().sort() };
  }

  function readAliasStatsFromDom() {
    // aggregator.js appends a note div to #output summarising alias expansion.
    const note = document.querySelector('#output .mt-3.text-sm');
    if (!note) return null;
    return note.textContent;
  }

  function renderResultCard(summaryByRegion) {
    if (!resultMount) return;
    const { regionCount, rowCount, regions } = countSummary(summaryByRegion);
    const aliasNote = readAliasStatsFromDom();

    const chips = regions.map((r) => `<span class="region-chip">${escapeHtml(r)}</span>`).join('');
    resultMount.innerHTML = `
      <div class="summary-result">
        <div class="sr-head">
          <svg class="check-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>
          Margin Summary Ready
        </div>
        <div class="sr-desc">The source files were processed successfully. The generated summary is now available in Step 2 — no need to export and re-upload.</div>
        <div class="sr-stats">
          <div class="stat-tile"><div class="st-val">${rowCount.toLocaleString()}</div><div class="st-lbl">Model/Grade rows</div></div>
          <div class="stat-tile"><div class="st-val">${regionCount.toLocaleString()}</div><div class="st-lbl">Regions</div></div>
          <div class="stat-tile"><div class="st-val">${getSheetCount()}</div><div class="st-lbl">Sheets processed</div></div>
          <div class="stat-tile"><div class="st-val">${getAliasShort(aliasNote)}</div><div class="st-lbl">Alias expansions</div></div>
        </div>
        <div class="sr-regions">${chips}</div>
        ${aliasNote ? `<div class="source-hint" style="margin-top:2px">${escapeHtml(aliasNote)}</div>` : ''}
      </div>`;
  }

  function renderInvalidatedCard() {
    if (!resultMount) return;
    resultMount.innerHTML = `
      <div class="summary-result is-empty">
        <div class="sr-head">
          <svg class="warn-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
          Source files changed — rebuild required
        </div>
        <div class="sr-desc">The Step 1 source files have changed. The previous summary is no longer valid. Click <strong>Build Summary</strong> again, or use an existing Summary File in Step 2.</div>
      </div>`;
  }

  function getSheetCount() {
    // aggregator.js exposes the count via the source-files-loaded status text, but the
    // simplest reliable source at generated-time is the mapping rows that were built.
    const rows = document.querySelectorAll('#mappingTableBody tr').length;
    return rows ? rows.toLocaleString() : '–';
  }

  function getAliasShort(note) {
    if (!note) return '–';
    const m = note.match(/(\d+)\s+missing alias row\(s\) added/);
    return m ? parseInt(m[1], 10).toLocaleString() : '–';
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  /* ---------- Accordion (collapsible full preview) ---------- */
  // #output keeps receiving the full table from aggregator.js (kept in DOM for SheetJS export).
  // The accordion wraps it so the giant table is hidden until the user expands it.
  if (accordionHead && accordion) {
    accordionHead.addEventListener('click', () => {
      const open = accordion.classList.toggle('open');
      accordionHead.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
  }

  /* ---------- Sheet-mapping count helper ---------- */
  const mappingSection = $('mappingSection');
  const sheetCountEl = $('sheetMappingCount');
  function refreshSheetCount() {
    if (!sheetCountEl) return;
    const n = document.querySelectorAll('#mappingTableBody tr').length;
    sheetCountEl.textContent = n ? `${n} sheet${n === 1 ? '' : 's'} detected` : '';
  }
  if (mappingSection) {
    new MutationObserver(refreshSheetCount).observe(mappingSection, { childList: true, subtree: true });
  }

  /* ---------- Listen to existing CustomEvents ---------- */
  // Step 1 source files loaded
  window.addEventListener('margin-source-files-loaded', (e) => {
    const d = (e && e.detail) || {};
    step1State('ready', `${(d.fileCount || 0)} file(s), ${(d.sheetCount || 0)} sheet(s)`);
    setStep1ActionNote('Verify detected Type & Region per sheet, then click Build Summary.');
    refreshSheetCount();
  });

  window.addEventListener('margin-source-files-error', () => {
    step1State('error', 'Read error');
  });

  window.addEventListener('margin-summary-invalidated', () => {
    if (accordion) accordion.classList.add('hidden');
    renderInvalidatedCard();
    step1State('warning', 'Rebuild required');
  });

  // Build started — app.js processBtn capture listener sets "Building summary..." text.
  // We detect that via step1Status mutation to flip the pill.
  const step1Status = $('step1Status');
  if (step1Status) {
    new MutationObserver(() => {
      const txt = (step1Status.textContent || '').toLowerCase();
      if (txt.includes('building summary')) step1State('building', 'Building…');
    }).observe(step1Status, { childList: true, characterData: true, subtree: true });
  }

  // Build succeeded
  window.addEventListener('margin-summary-generated', (e) => {
    const summary = e && e.detail && e.detail.summaryByRegion;
    // Defer one tick so aggregator's buildOutputTable (#output) has landed — we read alias note from it.
    setTimeout(() => {
      renderResultCard(summary);
      // Show accordion (collapsed) + meta
      if (accordion) {
        accordion.classList.remove('hidden');
        const { rowCount } = countSummary(summary);
        if (accordionMeta) accordionMeta.textContent = `Showing full summary · ${rowCount.toLocaleString()} rows · click to expand`;
      }
      step1State('complete', 'Summary ready');
      setStep1ActionNote('Summary built. Continue to Step 2, or export / preview below.');
      step2State('ready', 'Ready');
    }, 0);
  });

  window.addEventListener('margin-summary-build-error', (e) => {
    step1State('error', 'Error');
    const msg = (e && e.detail && e.detail.message) || 'Step 1 could not build the summary.';
    setStep1ActionNote(msg);
  });

  // Step 2 signals — generator writes into #status, so mirror pill from its class/text.
  const statusEl = $('status');
  if (statusEl) {
    new MutationObserver(() => {
      const cls = statusEl.className || '';
      const txt = (statusEl.textContent || '').toLowerCase();
      if (cls.includes('error')) step2State('error', 'Error');
      else if (txt.startsWith('done')) step2State('complete', 'Generated');
      else if (txt.includes('loaded') || txt.includes('please verify')) step2State('ready', 'Loaded');
      else if (txt.includes('generating') || txt.includes('loading')) step2State('building', 'Working…');
    }).observe(statusEl, { childList: true, characterData: true, subtree: true, attributes: true, attributeFilter: ['class'] });
  }

  function setStep1ActionNote(t) { const el = $('step1ActionNote'); if (el && t) el.textContent = t; }

  // Initial pill state
  step1State('', 'Not ready');
  step2State('', 'Waiting on Step 1');
  refreshSheetCount();
})();
