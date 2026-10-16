/* =========================================================
   VIEW MODE — read-only viewer dengan filter + realtime
   + Sheet tabs + Search + Go To
   ========================================================= */
const ROWS = 10000, COLS = 50;
let ROW_H = 26;
const BUFFER = 20;

const gridWrap = document.getElementById('grid-wrap');
const gridEl = document.getElementById('grid');
const HyperFormulaClass = (window.HyperFormula && window.HyperFormula.HyperFormula) || window.HyperFormula;

let wbData = { sheets: [{ name: '', cells: {}, merged: [], filters: {}, colWidths: {}, images: [] }], activeSheet: 0 };
let activeSheetIdx = 0;
let hf = null;
let rowEls = {};
let visibleStart = -1, visibleEnd = -1, scrollRAF = null;
let cachedRowMap = null;
let filterMode = false;
let filterPopupEl = null;

/* Selection state (untuk search & goto) */
let activeCell = 'A1';
let selection = { start: 'A1', end: 'A1' };

/* Realtime state */
let viewerWs = null;
let viewerWsReconnectTimer = null;
let viewerPingTimer = null;
let viewerUpdateFlashTimer = null;
let viewerWsConnected = false;

/* Search state */
let _searchState = {
  query: '',
  matches: [],
  currentIdx: -1
};

const colName = c => { let s='', n=c+1; while(n>0){const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; };
const addrOf = (r, c) => colName(c) + (r + 1);
function parseAddr(a) { const m=/^([A-Z]+)(\d+)$/.exec(a); if(!m)return null; let c=0; for(const ch of m[1])c=c*26+(ch.charCodeAt(0)-64); return {r:+m[2]-1,c:c-1}; }
function rangeParts(a) { const [s,e]=a.split(':'); return [s,e||s]; }
const escapeHtml = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);

function sheet() { return wbData.sheets[activeSheetIdx]; }

/* =========================================================
   FILTER HELPERS
   ========================================================= */
function getActiveFilters() {
  const f = sheet().filters || {};
  return Object.entries(f).filter(([_, v]) => v && Array.isArray(v.values) && v.values.length);
}
function invalidateRowMap() { cachedRowMap = null; }
function getRowMap() {
  const filters = getActiveFilters();
  if (!filters.length) return null;
  if (cachedRowMap) return cachedRowMap;
  let maxR = 100;
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (p && p.r > maxR) maxR = p.r;
  }
  maxR += 20;
  const out = [];
  for (let r = 0; r <= maxR; r++) {
    let pass = true;
    for (const [col, f] of filters) {
      const c = parseInt(col, 10);
      const raw = sheet().cells[addrOf(r, c)];
      const v = raw ? String(raw.v ?? '') : '';
      if (!f.values.includes(v)) { pass = false; break; }
    }
    if (pass) out.push(r);
  }
  cachedRowMap = out;
  return out;
}

/* =========================================================
   HYPERFORMULA
   ========================================================= */
function resetHF() {
  if (hf) { try { hf.destroy(); } catch {} hf = null; }
  let maxR = 50, maxC = 26;
  const cells = sheet().cells || {};
  for (const a of Object.keys(cells)) {
    const p = parseAddr(a); if (!p) continue;
    if (p.r + 1 > maxR) maxR = p.r + 1;
    if (p.c + 1 > maxC) maxC = p.c + 1;
  }
  maxR = Math.min(maxR + 20, ROWS);
  maxC = Math.min(maxC + 5, COLS);
  const matrix = Array.from({length: maxR}, () => new Array(maxC).fill(null));
  for (const [a, c] of Object.entries(cells)) {
    const p = parseAddr(a); if (!p || p.r >= maxR || p.c >= maxC) continue;
    if (c.f) matrix[p.r][p.c] = c.f;
    else if (c.v !== undefined && c.v !== null && c.v !== '') matrix[p.r][p.c] = c.v;
  }
  try {
    hf = HyperFormulaClass.buildFromArray(matrix, { licenseKey: 'gpl-v3' });
    hf.__maxR = maxR; hf.__maxC = maxC;
  } catch (e) {
    console.error('HF build failed:', e);
    hf = HyperFormulaClass.buildFromArray([[]], { licenseKey: 'gpl-v3' });
    hf.__maxR = 50; hf.__maxC = 26;
  }
}

/* =========================================================
   BUILD GRID
   ========================================================= */
function buildGrid() {
  const grid = document.getElementById('grid');
  grid.innerHTML = '';
  const cg = document.createElement('colgroup');
  cg.appendChild(document.createElement('col'));
  for (let c = 0; c < COLS; c++) {
    const col = document.createElement('col');
    col.dataset.col = c;
    cg.appendChild(col);
  }
  grid.appendChild(cg);

  const thead = document.createElement('thead');
  const htr = document.createElement('tr');
  htr.appendChild(document.createElement('th'));
  for (let c = 0; c < COLS; c++) {
    const th = document.createElement('th');
    th.textContent = colName(c);
    th.dataset.col = c;
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  grid.appendChild(thead);

  const tbody = document.createElement('tbody');
  tbody.id = 'grid-body';
  grid.appendChild(tbody);

  // Scroll handler
  gridWrap.addEventListener('scroll', () => {
    if (scrollRAF) return;
    scrollRAF = requestAnimationFrame(() => { scrollRAF = null; renderVirtual(); });
  }, { passive: true });

  // Click handler untuk select cell
  grid.addEventListener('click', (e) => {
    const td = e.target.closest('td[data-addr]');
    if (!td) return;
    selectCell(td.dataset.addr);
  });
}

function recalcTableWidth() {
  const grid = document.getElementById('grid');
  let total = 48;
  for (let c = 0; c < COLS; c++) total += sheet().colWidths?.[c] || 96;
  grid.style.width = total + 'px';
  grid.style.minWidth = total + 'px';
  grid.style.tableLayout = 'fixed';
}

function renderVirtual() {
  const scrollTop = gridWrap.scrollTop;
  const viewH = gridWrap.clientHeight;
  const rowMap = getRowMap();
  const totalRows = rowMap ? rowMap.length : ROWS;
  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - BUFFER);
  const end = Math.min(totalRows, Math.ceil((scrollTop + viewH) / ROW_H) + BUFFER);
  if (start === visibleStart && end === visibleEnd) return;
  visibleStart = start; visibleEnd = end;

  const tbody = document.getElementById('grid-body');
  if (!tbody) return;
  const frag = document.createDocumentFragment();

  const visibleMerges = [];
  for (const rangeStr of (sheet().merged || [])) {
    const [s, e] = rangeParts(rangeStr);
    const p1 = parseAddr(s), p2 = parseAddr(e); if (!p1 || !p2) continue;
    const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
    if (r0 < end && r1 >= start) {
      visibleMerges.push({
        r0, r1,
        c0: Math.min(p1.c, p2.c), c1: Math.max(p1.c, p2.c)
      });
    }
  }

  if (start > 0) {
    const tr = document.createElement('tr');
    tr.style.height = (start * ROW_H) + 'px';
    const td = document.createElement('td');
    td.colSpan = COLS + 1; td.style.cssText = 'padding:0;border:0';
    tr.appendChild(td); frag.appendChild(tr);
  }

  for (let di = start; di < end; di++) {
    const r = rowMap ? rowMap[di] : di;
    const tr = buildRow(r, visibleMerges);
    frag.appendChild(tr);
  }

  if (end < totalRows) {
    const tr = document.createElement('tr');
    tr.style.height = ((totalRows - end) * ROW_H) + 'px';
    const td = document.createElement('td');
    td.colSpan = COLS + 1; td.style.cssText = 'padding:0;border:0';
    tr.appendChild(td); frag.appendChild(tr);
  }

  tbody.replaceChildren(frag);
  rowEls = {};
  Array.from(tbody.children).forEach(tr => {
    const th = tr.firstElementChild;
    if (th && th.dataset.row !== undefined) rowEls[+th.dataset.row] = tr;
  });
  renderVisibleCells();
  applySelectionStyles();
  highlightAllMatches();
}

function buildRow(r, visibleMerges) {
  const tr = document.createElement('tr');
  tr.style.height = ROW_H + 'px';
  const th = document.createElement('th');
  th.textContent = r + 1;
  th.dataset.row = r;
  tr.appendChild(th);

  let c = 0;
  while (c < COLS) {
    let found = null;
    for (const vm of visibleMerges) {
      if (c >= vm.c0 && c <= vm.c1 && r >= vm.r0 && r <= vm.r1) { found = vm; break; }
    }
    if (found) {
      if (r === found.r0 && c === found.c0) {
        const td = document.createElement('td');
        td.dataset.addr = addrOf(r, c);
        td.colSpan = found.c1 - found.c0 + 1;
        td.rowSpan = found.r1 - found.r0 + 1;
        td.classList.add('merged-tl');
        const inner = document.createElement('div');
        inner.className = 'cell-inner';
        td.appendChild(inner);
        tr.appendChild(td);
        c = found.c1 + 1;
      } else { c = found.c1 + 1; }
    } else {
      const td = document.createElement('td');
      td.dataset.addr = addrOf(r, c);
      const inner = document.createElement('div');
      inner.className = 'cell-inner';
      td.appendChild(inner);
      tr.appendChild(td);
      c++;
    }
  }
  return tr;
}

function renderVisibleCells() {
  const rowMap = getRowMap();
  for (let di = visibleStart; di < visibleEnd; di++) {
    const r = rowMap ? rowMap[di] : di;
    const tr = rowEls[r]; if (!tr) continue;
    tr.querySelectorAll('td[data-addr]').forEach(td => {
      const p = parseAddr(td.dataset.addr); if (!p) return;
      const addr = td.dataset.addr;
      const cl = sheet().cells[addr];
      let raw = '';
      if (cl && cl.f) {
        try {
          const maxR = hf.__maxR || ROWS;
          const maxC = hf.__maxC || COLS;
          if (p.r < maxR && p.c < maxC) {
            raw = hf.getCellValue({ sheet: 0, row: p.r, col: p.c });
          }
        } catch { raw = ''; }
      } else { raw = cl?.v ?? ''; }
      const inner = td.querySelector('.cell-inner');
      if (inner) {
        inner.textContent = cl?.s?.numFmt ? formatNumber(raw, cl.s.numFmt) : displayVal(raw);
        inner.classList.toggle('cell-link', !!cl?.link);
      }
      td.classList.toggle('error', !!(raw && typeof raw === 'object' && raw.type === 'ERROR'));
      applyFormat(td, cl?.s || {});
    });
  }
}

function displayVal(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return 'value' in v ? String(v.value) : '';
  return String(v);
}

function applyFormat(td, s) {
  td.style.fontWeight = s.bold ? '700' : '';
  td.style.fontStyle = s.italic ? 'italic' : '';
  td.style.textDecoration = [s.underline ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean).join(' ') || '';
  td.style.fontFamily = s.fontFamily || '';
  td.style.fontSize = s.fontSize ? s.fontSize + 'px' : '';
  td.style.color = s.color || '';
  td.style.background = s.bg || '';
  td.style.textAlign = s.align || '';
  td.classList.toggle('wrap', !!s.wrap);
  td.style.borderTop = td.style.borderBottom = td.style.borderLeft = td.style.borderRight = '1px solid #e9eef4';
}

function formatNumber(v, fmt) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v); if (isNaN(n)) return String(v);
  switch (fmt) {
    case 'number': return n.toLocaleString('id-ID');
    case 'currency': return 'Rp ' + n.toLocaleString('id-ID');
    case 'percent': return (n * 100).toFixed(2) + '%';
    case 'date': { const d = new Date(n > 1e10 ? n : Date.now()); return isNaN(d) ? String(v) : d.toLocaleDateString('id-ID'); }
    default: return String(v);
  }
}

/* =========================================================
   SELECTION (read-only visual)
   ========================================================= */
function selectCell(addr) {
  activeCell = addr;
  selection = { start: addr, end: addr };

  // Update formula bar
  const ca = document.getElementById('cell-address');
  if (ca) ca.textContent = addr;

  const fi = document.getElementById('formula-input');
  if (fi) {
    const cl = sheet().cells[addr];
    fi.value = cl?.f || cl?.v || '';
  }

  applySelectionStyles();
}

function applySelectionStyles() {
  // Clear
  document.querySelectorAll('#grid tbody td.active, #grid tbody td.selected')
    .forEach(td => td.classList.remove('active', 'selected'));

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  const selR0 = Math.min(p1.r, p2.r), selR1 = Math.max(p1.r, p2.r);
  const selC0 = Math.min(p1.c, p2.c), selC1 = Math.max(p1.c, p2.c);
  const activeP = parseAddr(activeCell);

  document.querySelectorAll('#grid tbody td[data-addr]').forEach(td => {
    const p = parseAddr(td.dataset.addr);
    if (!p) return;
    const tdR1 = p.r + (td.rowSpan || 1) - 1;
    const tdC1 = p.c + (td.colSpan || 1) - 1;
    const overlap = !(tdR1 < selR0 || p.r > selR1 || tdC1 < selC0 || p.c > selC1);
    if (!overlap) return;
    if (activeP && p.r === activeP.r && p.c === activeP.c) {
      td.classList.add('active');
    } else {
      td.classList.add('selected');
    }
  });
}

/* =========================================================
   FILTER UI
   ========================================================= */
document.getElementById('btn-filter-toggle')?.addEventListener('click', () => {
  filterMode = !filterMode;
  const btn = document.getElementById('btn-filter-toggle');
  if (btn) {
    btn.classList.toggle('active', filterMode);
    btn.style.background = filterMode ? 'linear-gradient(135deg,#10b981,#059669)' : '';
    btn.style.color = filterMode ? '#fff' : '';
  }
  renderFilterHeaders();
});

document.getElementById('btn-filter-clear')?.addEventListener('click', () => {
  sheet().filters = {};
  invalidateRowMap();
  renderSheet();
  renderFilterHeaders();
});

function renderFilterHeaders() {
  const thead = document.querySelector('#grid thead tr');
  if (!thead) return;
  thead.querySelectorAll('.fx-filter-btn').forEach(b => b.remove());
  thead.querySelectorAll('th[data-col]').forEach(th => th.classList.remove('filtered', 'filter-active'));
  if (!filterMode) return;

  const filters = sheet().filters || {};
  thead.querySelectorAll('th[data-col]').forEach(th => {
    const c = +th.dataset.col;
    const btn = document.createElement('button');
    btn.className = 'fx-filter-btn';
    btn.innerHTML = '<i class="fa-solid fa-caret-down"></i>';
    btn.addEventListener('mousedown', e => {
      e.stopPropagation(); e.preventDefault();
      openFilterPopup(c, th);
    });
    th.appendChild(btn);
    if (filters[c]?.values?.length) th.classList.add('filtered', 'filter-active');
  });
}

function openFilterPopup(colIdx, thEl) {
  closeFilterPopup();
  const rect = thEl.getBoundingClientRect();
  const values = new Map();
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (!p || p.c !== colIdx) continue;
    const v = sheet().cells[a]?.v;
    if (v === undefined || v === null || v === '') continue;
    const key = String(v);
    values.set(key, (values.get(key) || 0) + 1);
  }
  const sortedKeys = [...values.keys()].sort((a, b) => {
    const na = Number(a), nb = Number(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });

  const active = sheet().filters?.[colIdx]?.values;
  const activeSet = active ? new Set(active) : null;
  const selected = new Set(activeSet ? [...activeSet] : sortedKeys);

  const popup = document.createElement('div');
  popup.className = 'fx-filter-popup';
  popup.style.left = Math.min(rect.left, window.innerWidth - 330) + 'px';
  popup.style.top = (rect.bottom + 4) + 'px';
  popup.innerHTML = `
    <div class="fx-fp-header"><i class="fa-solid fa-filter"></i> Filter Kolom ${colName(colIdx)}</div>
    <div class="fx-fp-search"><input type="text" placeholder="Cari nilai…" id="fx-fp-search"></div>
    <div class="fx-fp-list" id="fx-fp-list"></div>
    <div class="fx-fp-actions">
      <button id="fx-fp-all">Semua</button>
      <button id="fx-fp-none">Kosongkan</button>
      <button id="fx-fp-clear">Clear</button>
      <button class="primary" id="fx-fp-apply">Terapkan</button>
    </div>
  `;
  document.body.appendChild(popup);
  filterPopupEl = popup;

  const list = popup.querySelector('#fx-fp-list');
  const renderList = (q = '') => {
    const ql = q.toLowerCase();
    const f = sortedKeys.filter(k => !ql || k.toLowerCase().includes(ql));
    if (!f.length) { list.innerHTML = '<div class="fx-fp-item empty">Tidak ada</div>'; return; }
    list.innerHTML = f.map(k => `
      <label class="fx-fp-item">
        <input type="checkbox" data-v="${escapeHtml(k)}" ${selected.has(k) ? 'checked' : ''}>
        <span>${escapeHtml(k)}</span>
        <small>${values.get(k)}</small>
      </label>
    `).join('');
    list.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', () => {
        const v = cb.dataset.v;
        if (cb.checked) selected.add(v); else selected.delete(v);
      });
    });
  };
  renderList();

  popup.querySelector('#fx-fp-search').addEventListener('input', e => renderList(e.target.value));
  popup.querySelector('#fx-fp-all').onclick = () => { sortedKeys.forEach(k => selected.add(k)); renderList(popup.querySelector('#fx-fp-search').value); };
  popup.querySelector('#fx-fp-none').onclick = () => { selected.clear(); renderList(popup.querySelector('#fx-fp-search').value); };
  popup.querySelector('#fx-fp-clear').onclick = () => {
    if (sheet().filters) delete sheet().filters[colIdx];
    invalidateRowMap();
    closeFilterPopup();
    renderSheet();
    renderFilterHeaders();
  };
  popup.querySelector('#fx-fp-apply').onclick = () => {
    sheet().filters ||= {};
    if (selected.size === sortedKeys.length) delete sheet().filters[colIdx];
    else sheet().filters[colIdx] = { values: [...selected] };
    invalidateRowMap();
    closeFilterPopup();
    renderSheet();
    renderFilterHeaders();
  };

  setTimeout(() => document.addEventListener('mousedown', outsideFilterHandler, true), 0);
}

function outsideFilterHandler(e) {
  if (filterPopupEl && !filterPopupEl.contains(e.target) && !e.target.closest('.fx-filter-btn')) {
    closeFilterPopup();
  }
}
function closeFilterPopup() {
  if (filterPopupEl) {
    filterPopupEl.remove();
    filterPopupEl = null;
    document.removeEventListener('mousedown', outsideFilterHandler, true);
  }
}

function renderSheet() {
  for (let c = 0; c < COLS; c++) {
    if (sheet().colWidths?.[c]) {
      const cg = document.querySelector('#grid colgroup');
      if (cg && cg.children[c + 1]) cg.children[c + 1].style.width = sheet().colWidths[c] + 'px';
    }
  }
  ROW_H = sheet().rowHeight || 26;
  recalcTableWidth();
  visibleStart = -1; visibleEnd = -1;
  rowEls = {};
  renderVirtual();
  renderViewSheetTabs();
  updateSheetInfo();
}

function updateSheetInfo() {
  const info = document.getElementById('v-info');
  if (!info) return;
  const cellCount = Object.keys(sheet().cells || {}).length;
  info.innerHTML = `<i class="fa-solid fa-table"></i> ${wbData.sheets.length} sheet · ${cellCount.toLocaleString('id-ID')} sel`;
}

/* =========================================================
   ZOOM
   ========================================================= */
document.getElementById('zoom')?.addEventListener('input', e => {
  const z = e.target.value;
  document.getElementById('zoom-label').textContent = z + '%';
  const grid = document.getElementById('grid');
  grid.style.transform = `scale(${z / 100})`;
  grid.style.transformOrigin = 'top left';
});

/* =========================================================
   SHEET TABS
   ========================================================= */
function renderViewSheetTabs() {
  const cont = document.getElementById('v-sheet-tabs');
  if (!cont) return;

  cont.innerHTML = '';

  wbData.sheets.forEach((s, i) => {
    if (s.hidden && i !== activeSheetIdx) return;

    const btn = document.createElement('button');
    btn.className = 'v-sheet-tab' + (i === activeSheetIdx ? ' active' : '');
    btn.textContent = s.name || `Sheet${i + 1}`;
    btn.dataset.idx = i;
    if (s.tabColor) btn.dataset.color = s.tabColor;

    btn.addEventListener('click', () => switchViewSheet(i));

    cont.appendChild(btn);
  });
}

function switchViewSheet(idx) {
  if (idx === activeSheetIdx) return;
  if (!wbData.sheets[idx]) return;

  activeSheetIdx = idx;
  activeCell = 'A1';
  selection = { start: 'A1', end: 'A1' };

  // Reset filter mode
  filterMode = false;
  const btn = document.getElementById('btn-filter-toggle');
  if (btn) {
    btn.classList.remove('active');
    btn.style.background = '';
    btn.style.color = '';
  }

  // Reset HF
  resetHF();

  // Reset cache
  cachedRowMap = null;

  // Scroll ke atas
  gridWrap.scrollTop = 0;

  // Re-render
  renderSheet();
  renderFilterHeaders();

  // Update search — clear
  clearSearch();

  // Update formula bar
  const ca = document.getElementById('cell-address');
  if (ca) ca.textContent = 'A1';
  const fi = document.getElementById('formula-input');
  if (fi) fi.value = '';
}

document.getElementById('v-tab-prev')?.addEventListener('click', () => {
  const visibleIdxs = wbData.sheets
    .map((s, i) => ({ s, i }))
    .filter(x => !x.s.hidden || x.i === activeSheetIdx)
    .map(x => x.i);
  const pos = visibleIdxs.indexOf(activeSheetIdx);
  if (pos > 0) switchViewSheet(visibleIdxs[pos - 1]);
});

document.getElementById('v-tab-next')?.addEventListener('click', () => {
  const visibleIdxs = wbData.sheets
    .map((s, i) => ({ s, i }))
    .filter(x => !x.s.hidden || x.i === activeSheetIdx)
    .map(x => x.i);
  const pos = visibleIdxs.indexOf(activeSheetIdx);
  if (pos >= 0 && pos < visibleIdxs.length - 1) {
    switchViewSheet(visibleIdxs[pos + 1]);
  }
});

/* =========================================================
   SEARCH
   ========================================================= */
function clearSearch() {
  _searchState = { query: '', matches: [], currentIdx: -1 };
  const input = document.getElementById('v-search-input');
  if (input) input.value = '';
  updateSearchCount();
  document.querySelectorAll('#grid td.find-match, #grid td.find-current')
    .forEach(td => td.classList.remove('find-match', 'find-current'));
}

function performSearch(query) {
  _searchState.query = String(query || '').trim();
  _searchState.matches = [];
  _searchState.currentIdx = -1;

  document.querySelectorAll('#grid td.find-match, #grid td.find-current')
    .forEach(td => td.classList.remove('find-match', 'find-current'));

  if (!_searchState.query) {
    updateSearchCount();
    return;
  }

  const q = _searchState.query.toLowerCase();
  const cells = sheet().cells || {};

  for (const [addr, cl] of Object.entries(cells)) {
    if (!cl) continue;

    let val = '';
    if (cl.f) {
      const p = parseAddr(addr);
      if (p) {
        try {
          const maxR = hf.__maxR || ROWS;
          const maxC = hf.__maxC || COLS;
          if (p.r < maxR && p.c < maxC) {
            val = String(hf.getCellValue({ sheet: 0, row: p.r, col: p.c }) ?? '');
          }
        } catch { val = ''; }
      }
    } else {
      val = String(cl.v ?? '');
    }

    if (val.toLowerCase().includes(q)) {
      const p = parseAddr(addr);
      if (p) _searchState.matches.push({ addr, r: p.r, c: p.c });
    }
  }

  _searchState.matches.sort((a, b) => a.r - b.r || a.c - b.c);

  if (_searchState.matches.length > 0) {
    _searchState.currentIdx = 0;
    gotoSearchMatch(0);
  }

  updateSearchCount();
  highlightAllMatches();
}

function gotoSearchMatch(idx) {
  if (idx < 0 || idx >= _searchState.matches.length) return;
  const m = _searchState.matches[idx];
  scrollViewAddrIntoView(m.addr);
  selectCell(m.addr);
  highlightAllMatches();
  updateSearchCount();
}

function nextSearchMatch() {
  if (!_searchState.matches.length) return;
  _searchState.currentIdx = (_searchState.currentIdx + 1) % _searchState.matches.length;
  gotoSearchMatch(_searchState.currentIdx);
}

function prevSearchMatch() {
  if (!_searchState.matches.length) return;
  _searchState.currentIdx = (_searchState.currentIdx - 1 + _searchState.matches.length) % _searchState.matches.length;
  gotoSearchMatch(_searchState.currentIdx);
}

function highlightAllMatches() {
  document.querySelectorAll('#grid td.find-match, #grid td.find-current')
    .forEach(td => td.classList.remove('find-match', 'find-current'));

  if (!_searchState.matches.length) return;

  for (let i = 0; i < _searchState.matches.length; i++) {
    const m = _searchState.matches[i];
    const td = document.querySelector(`#grid td[data-addr="${m.addr}"]`);
    if (!td) continue;
    if (i === _searchState.currentIdx) {
      td.classList.add('find-current');
    } else {
      td.classList.add('find-match');
    }
  }
}

function updateSearchCount() {
  const el = document.getElementById('v-search-count');
  if (!el) return;

  if (!_searchState.query) {
    el.textContent = '';
    return;
  }
  if (!_searchState.matches.length) {
    el.textContent = '0/0';
    el.style.color = '#dc2626';
    return;
  }
  el.textContent = `${_searchState.currentIdx + 1}/${_searchState.matches.length}`;
  el.style.color = '#64748b';
}

function scrollViewAddrIntoView(addr) {
  const p = parseAddr(addr);
  if (!p) return;
  const targetTop = p.r * ROW_H;
  const viewTop = gridWrap.scrollTop;
  const viewBottom = viewTop + gridWrap.clientHeight;

  if (targetTop < viewTop + ROW_H * 2) {
    gridWrap.scrollTop = Math.max(0, targetTop - ROW_H * 2);
  } else if (targetTop + ROW_H > viewBottom - ROW_H * 2) {
    gridWrap.scrollTop = targetTop - gridWrap.clientHeight + ROW_H * 3;
  }
}

/* Bind search */
(function initSearch() {
  const input = document.getElementById('v-search-input');
  const prevBtn = document.getElementById('v-search-prev');
  const nextBtn = document.getElementById('v-search-next');
  if (!input) return;

  let timer = null;
  input.addEventListener('input', () => {
    clearTimeout(timer);
    timer = setTimeout(() => performSearch(input.value), 200);
  });

  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.shiftKey) prevSearchMatch(); else nextSearchMatch();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      clearSearch();
      input.blur();
    }
  });

  prevBtn?.addEventListener('click', prevSearchMatch);
  nextBtn?.addEventListener('click', nextSearchMatch);

  document.addEventListener('keydown', e => {
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      input.focus();
      input.select();
    }
  }, true);
})();

/* =========================================================
   GO TO
   ========================================================= */
function openGoToModal() {
  const modal = document.getElementById('modal-goto');
  if (!modal) return;
  modal.classList.remove('hidden');
  const input = document.getElementById('goto-input');
  if (input) {
    input.value = '';
    setTimeout(() => input.focus(), 50);
  }
}

function closeGoToModal() {
  document.getElementById('modal-goto')?.classList.add('hidden');
}

function applyGoTo() {
  const raw = document.getElementById('goto-input')?.value.trim() || '';
  if (!raw) return;

  const parts = raw.split(':');
  const startAddr = parts[0].toUpperCase();
  const endAddr = (parts[1] || parts[0]).toUpperCase();

  if (!parseAddr(startAddr) || !parseAddr(endAddr)) {
    return alert('Alamat tidak valid');
  }

  closeGoToModal();

  selection = { start: startAddr, end: endAddr };
  activeCell = startAddr;
  scrollViewAddrIntoView(startAddr);
  applySelectionStyles();

  const ca = document.getElementById('cell-address');
  if (ca) ca.textContent = startAddr;
  const fi = document.getElementById('formula-input');
  if (fi) {
    const cl = sheet().cells[startAddr];
    fi.value = cl?.f || cl?.v || '';
  }
}

document.getElementById('btn-goto')?.addEventListener('click', openGoToModal);
document.getElementById('btn-goto-cancel')?.addEventListener('click', closeGoToModal);
document.getElementById('btn-goto-apply')?.addEventListener('click', applyGoTo);

document.getElementById('goto-input')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); applyGoTo(); }
  if (e.key === 'Escape') { e.preventDefault(); closeGoToModal(); }
});

document.getElementById('modal-goto')?.addEventListener('click', e => {
  if (e.target.id === 'modal-goto') closeGoToModal();
});

document.addEventListener('keydown', e => {
  const ctrl = e.ctrlKey || e.metaKey;
  if (ctrl && e.key.toLowerCase() === 'g') {
    e.preventDefault();
    openGoToModal();
  }
}, true);

/* =========================================================
   REALTIME VIEWER — WebSocket
   ========================================================= */
function connectViewerWS(token, sheetId) {
  if (viewerWs && viewerWs.readyState === 1) return;

  const protocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const wsUrl = `${protocol}//${location.host}/ws?token=${encodeURIComponent(token)}&sheet=${sheetId}`;

  try {
    viewerWs = new WebSocket(wsUrl);
  } catch (e) {
    console.error('WS connect error:', e);
    setViewerStatus('offline');
    return;
  }

  viewerWs.addEventListener('open', () => {
    console.log('🔌 Viewer realtime connected');
    viewerWsConnected = true;
    setViewerStatus('live');
    startViewerPing();
  });

  viewerWs.addEventListener('message', (e) => {
    try {
      const msg = JSON.parse(e.data);
      handleViewerMessage(msg);
    } catch (err) {
      console.error('WS msg error:', err);
    }
  });

  viewerWs.addEventListener('close', () => {
    console.log('🔌 Viewer disconnected — reconnecting in 3s');
    viewerWsConnected = false;
    setViewerStatus('offline');
    stopViewerPing();
    if (viewerWsReconnectTimer) clearTimeout(viewerWsReconnectTimer);
    viewerWsReconnectTimer = setTimeout(() => {
      viewerWsReconnectTimer = null;
      connectViewerWS(token, sheetId);
    }, 3000);
  });

  viewerWs.addEventListener('error', () => {});
}

function handleViewerMessage(msg) {
  if (msg.type === 'edit') {
    const { sheetIdx, cells } = msg;
    if (!wbData?.sheets?.[sheetIdx]) return;

    const targetSheet = wbData.sheets[sheetIdx];
    if (!targetSheet.cells) targetSheet.cells = {};

    let changed = 0;
    for (const [addr, cell] of Object.entries(cells || {})) {
      if (cell === null) {
        if (targetSheet.cells[addr]) {
          delete targetSheet.cells[addr];
          changed++;
        }
      } else {
        targetSheet.cells[addr] = cell;
        changed++;
      }

      if (sheetIdx === activeSheetIdx && hf) {
        const p = parseAddr(addr);
        if (p) {
          try {
            const maxR = hf.__maxR || ROWS;
            const maxC = hf.__maxC || COLS;
            if (p.r < maxR && p.c < maxC) {
              hf.setCellContents({ sheet: 0, row: p.r, col: p.c },
                [[cell?.f || cell?.v || null]]);
            } else {
              resetHF();
            }
          } catch (e) {
            console.warn('HF update failed, rebuilding:', e.message);
            try { resetHF(); } catch {}
          }
        }
      }
    }

    if (sheetIdx === activeSheetIdx && changed) {
      renderVisibleCells();
      flashUpdate(changed);
    }
  } else if (msg.type === 'welcome') {
    console.log('👋 Viewer welcome, role:', msg.role);
  } else if (msg.type === 'pong') {
    // heartbeat OK
  }
}

function setViewerStatus(status) {
  const el = document.getElementById('v-status-left');
  if (!el) return;
  if (status === 'live') {
    el.innerHTML = '<i class="fa-solid fa-circle" style="color:#10b981;font-size:8px;margin-right:5px"></i> Live';
    el.style.color = '#10b981';
  } else if (status === 'offline') {
    el.innerHTML = '<i class="fa-solid fa-circle" style="color:#94a3b8;font-size:8px;margin-right:5px"></i> Offline';
    el.style.color = '#94a3b8';
  } else {
    el.innerHTML = '<i class="fa-solid fa-circle-check"></i> Siap';
    el.style.color = '';
  }
}

function flashUpdate(count) {
  const el = document.getElementById('v-status-left');
  if (!el) return;
  el.innerHTML = `<i class="fa-solid fa-bolt" style="color:#fbbf24;margin-right:5px"></i> ${count} sel diperbarui`;
  el.style.color = '#f59e0b';
  if (viewerUpdateFlashTimer) clearTimeout(viewerUpdateFlashTimer);
  viewerUpdateFlashTimer = setTimeout(() => setViewerStatus('live'), 1500);
}

function startViewerPing() {
  stopViewerPing();
  viewerPingTimer = setInterval(() => {
    if (viewerWs && viewerWs.readyState === 1) {
      try { viewerWs.send(JSON.stringify({ type: 'ping' })); } catch {}
    }
  }, 25000);
}

function stopViewerPing() {
  if (viewerPingTimer) { clearInterval(viewerPingTimer); viewerPingTimer = null; }
}

window.addEventListener('beforeunload', () => {
  if (viewerWs) try { viewerWs.close(); } catch {}
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && !viewerWsConnected) {
    if (window._viewToken && window._viewSheetId) {
      connectViewerWS(window._viewToken, window._viewSheetId);
    }
  }
});

/* =========================================================
   LOAD
   ========================================================= */
async function load() {
  const params = new URLSearchParams(location.search);
  const token = params.get('token');
  if (!token) {
    showError('Token tidak ditemukan di URL');
    return;
  }

  try {
    const res = await fetch(`/api/public/view/${encodeURIComponent(token)}`);
    if (!res.ok) {
      let err = 'Link tidak valid';
      try { const j = await res.json(); err = j.error || err; } catch {}
      showError(err);
      return;
    }
    const data = await res.json();
    document.getElementById('view-title').textContent = data.name;
    document.title = data.name + ' — View';

    wbData = data.data || { sheets: [{ name: '', cells: {}, merged: [], filters: {} }], activeSheet: 0 };
    activeSheetIdx = wbData.activeSheet || 0;

    if (!Array.isArray(wbData.sheets)) {
      wbData.sheets = [{ name: 'Sheet1', cells: wbData.sheets || {}, merged: [], filters: {} }];
    }
    wbData.sheets.forEach(s => {
      if (!s.colWidths) s.colWidths = {};
      if (!s.merged) s.merged = [];
      if (!s.filters) s.filters = {};
    });

    window._viewToken = token;
    window._viewSheetId = data.id;

    resetHF();
    buildGrid();
    renderSheet();
    renderFilterHeaders();
    document.getElementById('loading-overlay').classList.add('hidden');

    setViewerStatus('offline');
    connectViewerWS(token, data.id);
  } catch (e) {
    showError('Gagal memuat: ' + e.message);
  }
}

function showError(msg) {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) overlay.classList.add('hidden');

  const state = document.getElementById('error-state');
  const msgEl = document.getElementById('error-message');
  if (msgEl) msgEl.textContent = msg;
  if (state) state.classList.remove('hidden');
}

load();