/* =========================================================
   WEB EXCEL — Frontend v4
   Virtual Scroll + Delta Save + Autosave
   ========================================================= */
const ROWS = 10000;
const COLS = 50;
let ROW_H = 26;
const BUFFER = 20;
const HF_LICENSE = 'gpl-v3';
const HyperFormulaClass =
  (window.HyperFormula && window.HyperFormula.HyperFormula) || window.HyperFormula;
/* ---------- DOM refs ---------- */
const gridWrap = document.getElementById('grid-wrap');
const gridEl   = document.getElementById('grid');

/* ---------- State ---------- */
let token = localStorage.getItem('wex-token') || '';
let currentUser = null;
let wbId = null;
let wbName = 'Untitled';
let wbData = blankWB();
let activeSheetIdx = 0;
let activeCell = 'A1';
let selection = { start: 'A1', end: 'A1' };
let hf = null;
let dirty = false;
let clipboard = null;
let undoStack = [];
let redoStack = [];
let editor = null;
let chartInstances = new Map();

/* Virtual scroll */
let visibleStart = -1;
let visibleEnd = -1;
let rowEls = {};
let scrollRAF = null;

/* Delta save */
let dirtyCells = {};
let dirtyMeta = {};
let saving = false;
let saveAgain = false;
let saveRetryCount = 0;

/* Chart cache per sheet: Map<sheetIdx, Map<chartId, Chart>> */
const chartCache = new Map();

/* Autosave debounce */
let autosaveTimer = null;

/* Formula range warning */
let formulaWarningShown = false;
let cachedRowMap = null;
let filterPopupEl = null;
let filterMode = false;
let currentAccessMode = 'owner';
let shareUserSuggestList = [];
let shareUserSuggestIdx = -1;

function blankWB() {
  return {
    sheets: [{
      name: 'Sheet1', cells: {},
      colWidths: {}, rowHeight: 26,
      cf: [], charts: [], merged: [], filters: {}
    }],
    activeSheet: 0
  };
}
function sheet() { return wbData.sheets[activeSheetIdx]; }
function ensureCell(addr) { return (sheet().cells[addr] ||= {}); }

/* =========================================================
   UTILS
   ========================================================= */
const colName = c => {
  let s = '', n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};
const addrOf = (r, c) => colName(c) + (r + 1);
function parseAddr(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: +m[2] - 1, c: c - 1 };
}
function expandRange(a, b) {
  const p1 = parseAddr(a), p2 = parseAddr(b);
  if (!p1 || !p2) return [];
  const out = [];
  for (let r = Math.min(p1.r, p2.r); r <= Math.max(p1.r, p2.r); r++)
    for (let c = Math.min(p1.c, p2.c); c <= Math.max(p1.c, p2.c); c++)
      out.push(addrOf(r, c));
  return out;
}
function isRange(a) { return a.includes(':'); }
function rangeParts(a) { const [s, e] = a.split(':'); return [s, e || s]; }
const escapeHtml = s => String(s).replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

/* =========================================================
   LOADING SPINNER
   ========================================================= */
function showLoading(text = 'Memuat…') {
  const el = document.getElementById('loading-overlay');
  const txt = document.getElementById('loading-text');
  if (txt) txt.textContent = text;
  if (el) el.classList.remove('hidden');
}
function updateLoading(text) {
  const txt = document.getElementById('loading-text');
  if (txt) txt.textContent = text;
}
function hideLoading() {
  const el = document.getElementById('loading-overlay');
  if (el) el.classList.add('hidden');
}

/* Yield ke browser — biar UI bisa render spinner dulu */
const yieldToUI = () => new Promise(r => setTimeout(r, 0));

/* Warning kalau ada formula dengan range besar */
function warnBigRange() {
  if (formulaWarningShown) return;
  formulaWarningShown = true;
  const w = document.getElementById('formula-warning');
  if (!w) return;
  w.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Formula dengan range > 5.000 sel bisa memperlambat. Gunakan range lebih kecil atau <b>XLOOKUP</b>.';
  w.classList.remove('hidden');
  setTimeout(() => w.classList.add('hidden'), 6000);
}
/* =========================================================
   HYPERFORMULA
   ========================================================= */
/* Hitung used range — helper yang dipakai resetHF */
function computeUsedRange() {
  let maxR = 50, maxC = 26;
  const cells = sheet().cells || {};
  for (const a of Object.keys(cells)) {
    const p = parseAddr(a);
    if (!p) continue;
    if (p.r + 1 > maxR) maxR = p.r + 1;
    if (p.c + 1 > maxC) maxC = p.c + 1;
  }
  maxR = Math.min(maxR + 20, ROWS);
  maxC = Math.min(maxC + 5, COLS);
  return { maxR, maxC };
}

/* Versi SYNC — untuk dipakai saat data kecil & butuh cepat */
function resetHF() {
  if (hf) hf.destroy();
  const { maxR, maxC } = computeUsedRange();
  const cells = sheet().cells || {};
  const matrix = Array.from({ length: maxR }, () => new Array(maxC).fill(null));
  for (const [a, c] of Object.entries(cells)) {
    const p = parseAddr(a);
    if (!p || p.r >= maxR || p.c >= maxC) continue;
    if (c.f) matrix[p.r][p.c] = c.f;
    else if (c.v !== undefined && c.v !== null && c.v !== '') matrix[p.r][p.c] = c.v;
  }
  hf = HyperFormulaClass.buildFromArray(matrix, { licenseKey: HF_LICENSE });
  hf.__maxR = maxR;
  hf.__maxC = maxC;
}

/* Versi ASYNC — pakai spinner + yield ke browser */
async function resetHFLazy(reason = 'Menyiapkan formula engine…') {
  if (hf) hf.destroy();

  const cellCount = Object.keys(sheet().cells || {}).length;

  if (cellCount < 2000) {
    resetHF();
    return;
  }

  showLoading(reason);
  await yieldToUI();
  updateLoading('Membangun formula engine…');
  await yieldToUI();

  try {
    resetHF();
  } catch (e) {
    console.error('HF build failed:', e);
    hf = HyperFormulaClass.buildFromArray([[]], { licenseKey: HF_LICENSE });
    hf.__maxR = 50; hf.__maxC = 26;
  }

  hideLoading();
}

/* =========================================================
   VIRTUAL GRID BUILD
   ========================================================= */
function buildGrid() {
  const grid = document.getElementById('grid');
  grid.innerHTML = '';

  /* ---- COLGROUP: sumber kebenaran untuk lebar kolom ---- */
  const cg = document.createElement('colgroup');
  cg.appendChild(document.createElement('col')); // row header
  for (let c = 0; c < COLS; c++) {
    const col = document.createElement('col');
    col.dataset.col = c;
    cg.appendChild(col);
  }
  grid.appendChild(cg);

  /* ---- THEAD (static) ---- */
  const thead = document.createElement('thead');
  const htr = document.createElement('tr');
  htr.appendChild(document.createElement('th'));
  for (let c = 0; c < COLS; c++) {
    const th = document.createElement('th');
    th.textContent = colName(c);
    th.dataset.col = c;
    th.style.position = 'sticky';
    th.style.overflow = 'visible';
    // Grip resize
    const grip = document.createElement('div');
    grip.className = 'col-resize-grip';
    grip.dataset.col = c;
    Object.assign(grip.style, {
      position: 'absolute',
      top: '0',
      right: '-4px',
      width: '8px',
      height: '100%',
      cursor: 'col-resize',
      zIndex: '50',
      background: 'transparent',
      pointerEvents: 'auto',
      userSelect: 'none'
    });
    grip.addEventListener('mouseenter', () => {
      grip.style.background =
        'linear-gradient(90deg,transparent 35%,#10b981 35%,#10b981 65%,transparent 65%)';
    });
    grip.addEventListener('mouseleave', () => {
      grip.style.background = 'transparent';
    });
    th.appendChild(grip);
    htr.appendChild(th);
  }
  thead.appendChild(htr);
  grid.appendChild(thead);

  /* ---- TBODY (dynamic) ---- */
  const tbody = document.createElement('tbody');
  tbody.id = 'grid-body';
  grid.appendChild(tbody);

  /* ---- Scroll ---- */
  gridWrap.addEventListener('scroll', onGridScroll, { passive: true });
  window.addEventListener('resize', () => { visibleStart = -1; renderVirtual(); });

  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};
  renderVirtual();
}

// function attachColResize(th, c) {
//   const grip = document.createElement('div');
//   grip.style.cssText =
//     'position:absolute;right:0;top:0;width:6px;height:100%;cursor:col-resize;user-select:none;z-index:1';
//   th.style.position = 'sticky';
//   th.appendChild(grip);
//   grip.addEventListener('mousedown', e => {
//     e.preventDefault();
//     const startX = e.clientX;
//     const startW = sheet().colWidths?.[c] || 96;
//     const move = ev => {
//       const w = Math.max(30, startW + ev.clientX - startX);
//       sheet().colWidths ||= {};
//       sheet().colWidths[c] = w;
//       applyColWidth(c, w);
//     };
//     const up = () => {
//       document.removeEventListener('mousemove', move);
//       document.removeEventListener('mouseup', up);
//       markMetaDirty('structure');
//     };
//     document.addEventListener('mousemove', move);
//     document.addEventListener('mouseup', up);
//   });
// }

function applyColWidth(c, w) {
  const cg = document.querySelector('#grid colgroup');
  if (cg && cg.children[c + 1]) cg.children[c + 1].style.width = w + 'px';
  recalcTableWidth();
}

/* =========================================================
   AUTO-FIT COLUMN WIDTH
   ========================================================= */
function autoFitColumn(c) {
  if (!hf) return;

  // Batas minimum & maksimum (dalam pixel)
  const MIN_W = 40;
  const MAX_W = 500;

  // Padding & karakter → pixel
  const PADDING = 16;
  const CHAR_W = 7.2;              // lebar rata-rata 1 karakter @ 12px
  const NUM_W = 7.6;               // angka sedikit lebih lebar (tabular-nums)

  let maxWidth = MIN_W;

  // Cek header kolom
  const headerText = colName(c);
  maxWidth = Math.max(maxWidth, headerText.length * CHAR_W + PADDING);

  // Cek semua sel di kolom c yang punya data
  const cellsInCol = [];
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (p && p.c === c) cellsInCol.push(p);
  }

  for (const p of cellsInCol) {
    const addr = addrOf(p.r, c);
    const cl = sheet().cells[addr];
    let v;
    if (cl.f) {
      try { v = hf.getCellValue({ sheet: 0, row: p.r, col: c }); }
      catch { v = ''; }
    } else {
      v = cl.v ?? '';
    }
    // Format angka untuk perhitungan (Rp, %, dll)
    let text;
    if (cl?.s?.numFmt) text = formatNumber(v, cl.s.numFmt);
    else text = displayVal(v);

    const str = String(text || '');
    if (!str) continue;

    // Handling multi-baris (wrap)
    const isWrap = cl?.s?.wrap;
    const bold = cl?.s?.bold;
    const fontSize = cl?.s?.fontSize || 12;
    const ratio = fontSize / 12;

    if (isWrap) {
      // Kalau wrap, ambil baris terpanjang
      const lines = str.split('\n');
      const longest = lines.reduce((m, l) => Math.max(m, l.length), 0);
      const w = longest * CHAR_W * ratio + PADDING;
      maxWidth = Math.max(maxWidth, Math.min(w, 300));
    } else {
      // Hitung panjang dalam char
      let w;
      const isNumber = !isNaN(Number(v)) && v !== '' && v !== null;
      if (isNumber) w = str.length * NUM_W * ratio + PADDING;
      else w = str.length * CHAR_W * ratio + PADDING;
      if (bold) w *= 1.08;
      maxWidth = Math.max(maxWidth, Math.min(w, MAX_W));
    }
  }

  const finalW = Math.round(Math.max(MIN_W, Math.min(maxWidth, MAX_W)));
  sheet().colWidths ||= {};
  sheet().colWidths[c] = finalW;
  applyColWidth(c, finalW);
  markMetaDirty('structure');
  scheduleSave();
  return finalW;
}

/* Hitung total lebar table = row-header + Σ colWidths */
function recalcTableWidth() {
  const grid = document.getElementById('grid');
  if (!grid) return;
  let total = 48;                          // kolom row-header
  for (let c = 0; c < COLS; c++) {
    total += sheet().colWidths?.[c] || 96; // default 96px
  }
  grid.style.width    = total + 'px';
  grid.style.minWidth = total + 'px';
  grid.style.tableLayout = 'fixed';
}

function createRowEl(r) {
  const tr = document.createElement('tr');
  tr.style.height = ROW_H + 'px';
  const th = document.createElement('th');
  th.textContent = r + 1;
  th.dataset.row = r;
  tr.appendChild(th);
  for (let c = 0; c < COLS; c++) {
    const td = document.createElement('td');
    td.dataset.addr = addrOf(r, c);
    const inner = document.createElement('div');
    inner.className = 'cell-inner';
    td.appendChild(inner);
    tr.appendChild(td);
  }
  return tr;
}

/* =========================================================
   VIRTUAL RENDER
   ========================================================= */
function onGridScroll() {
  if (scrollRAF) return;
  scrollRAF = requestAnimationFrame(() => {
    scrollRAF = null;
    renderVirtual();
  });
}

function renderVirtual() {
  const gridWrapEl = document.getElementById('grid-wrap');
  const scrollTop = gridWrapEl.scrollTop;
  const viewH = gridWrapEl.clientHeight;

  const rowMap = getRowMap();
  const totalRows = rowMap ? rowMap.length : ROWS;

  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - BUFFER);
  const end = Math.min(totalRows, Math.ceil((scrollTop + viewH) / ROW_H) + BUFFER);

  if (start === visibleStart && end === visibleEnd) return;
  visibleStart = start;
  visibleEnd = end;

  const tbody = document.getElementById('grid-body');
  if (!tbody) return;
  const frag = document.createDocumentFragment();

  /* ---- Kumpulkan merge yang TL-nya berada di visible range ---- */
  const visibleMerges = [];
  for (const rangeStr of (sheet().merged || [])) {
    const [s, e] = rangeParts(rangeStr);
    const p1 = parseAddr(s), p2 = parseAddr(e);
    if (!p1 || !p2) continue;
    const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
    const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);
    // Hanya aktif kalau TL row ada di dalam visible range
    if (r0 >= start && r0 < end) {
      visibleMerges.push({ range: rangeStr, r0, r1, c0, c1 });
    }
  }

  /* Top spacer */
  if (start > 0) {
    const tr = document.createElement('tr');
    tr.style.height = (start * ROW_H) + 'px';
    const td = document.createElement('td');
    td.colSpan = COLS + 1;
    td.style.cssText = 'padding:0;border:0';
    tr.appendChild(td);
    frag.appendChild(tr);
  }

  /* Visible rows — bangun ulang tiap render (rowSpan/colSpan kompatibel dengan ini) */
  for (let di = start; di < end; di++) {
    const r = rowMap ? rowMap[di] : di;
    const tr = buildRowWithMerges(r, visibleMerges);
    frag.appendChild(tr);
  }

  /* Bottom spacer */
  if (end < totalRows) {
    const tr = document.createElement('tr');
    tr.style.height = ((totalRows - end) * ROW_H) + 'px';
    const td = document.createElement('td');
    td.colSpan = COLS + 1;
    td.style.cssText = 'padding:0;border:0';
    tr.appendChild(td);
    frag.appendChild(tr);
  }

  tbody.replaceChildren(frag);

  /* Update cache rowEls dari tr yang baru dibuat */
  rowEls = {};
  Array.from(tbody.children).forEach(tr => {
    const th = tr.firstElementChild;
    if (th && th.dataset.row !== undefined) {
      rowEls[+th.dataset.row] = tr;
    }
  });

  renderVisibleCells();
  renderCFVisible();
  applySelectionStyles();
  renderRowHeaderHighlight();
}

/* =========================================================
   BUILD ROW — dengan colSpan & rowSpan native
   ========================================================= */
function buildRowWithMerges(r, visibleMerges) {
  const tr = document.createElement('tr');
  tr.style.height = ROW_H + 'px';

  const th = document.createElement('th');
  th.textContent = r + 1;
  th.dataset.row = r;
  // Grip resize — INLINE STYLE
  const grip = document.createElement('div');
  grip.className = 'row-resize-grip';
  grip.dataset.row = r;
  Object.assign(grip.style, {
    position: 'absolute',
    left: '0',
    bottom: '-4px',
    width: '100%',
    height: '8px',
    cursor: 'row-resize',
    zIndex: '50',
    background: 'transparent',
    pointerEvents: 'auto',
    userSelect: 'none'
  });
  grip.addEventListener('mouseenter', () => {
    grip.style.background =
      'linear-gradient(180deg,transparent 35%,#10b981 35%,#10b981 65%,transparent 65%)';
  });
  grip.addEventListener('mouseleave', () => {
    grip.style.background = 'transparent';
  });
  th.appendChild(grip);
  tr.appendChild(th);

  let c = 0;
  while (c < COLS) {
    // Cari merge yang mencakup (r, c)
    let found = null;
    for (const vm of visibleMerges) {
      if (c >= vm.c0 && c <= vm.c1 && r >= vm.r0 && r <= vm.r1) {
        found = vm;
        break;
      }
    }

    if (found) {
      if (r === found.r0 && c === found.c0) {
        // Top-left — emit td dengan colSpan + rowSpan
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
      } else {
        // Covered — skip kolom ini (ditangani colSpan/rowSpan di atas)
        c = found.c1 + 1;
      }
    } else {
      // Cell normal
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
    const tr = rowEls[r];
    if (!tr) continue;
    // Iterasi td asli (colSpan sudah handle cell yang di-cover)
    const tds = tr.querySelectorAll('td[data-addr]');
    tds.forEach(td => {
      const addr = td.dataset.addr;
      const p = parseAddr(addr);
      if (p) updateCellDisplay(td, p.r, p.c);
    });
  }
}

function forceRenderCell(td, r, c) {
  if (!td) return;
  const addr = td.dataset.addr;
  const cl = sheet().cells[addr];
  let raw = '';
  if (cl && cl.f) {
    const maxR = (hf && hf.__maxR) || ROWS;
    const maxC = (hf && hf.__maxC) || COLS;
    if (r < maxR && c < maxC) {
      try { raw = hf.getCellValue({ sheet: 0, row: r, col: c }); }
      catch { raw = '#REF!'; }
    } else {
      raw = '#REF!';
    }
  } else {
    raw = cl?.v ?? '';
  }

  let inner = td.querySelector('.cell-inner');
  if (!inner) {
    inner = document.createElement('div');
    inner.className = 'cell-inner';
    td.insertBefore(inner, td.firstChild);
  }
  inner.textContent = cl?.s?.numFmt
    ? formatNumber(raw, cl.s.numFmt)
    : displayVal(raw);

  td.classList.toggle('error', !!isErr(raw));
  applyFormat(td, cl?.s || {});
}

function updateCellDisplay(td, r, c) {
  if (!td) return;
  if (td.querySelector('input')) return;
  forceRenderCell(td, r, c);
}

function renderRow(r) {
  if (r < visibleStart || r >= visibleEnd) return;
  const tr = rowEls[r];
  if (!tr) return;
  for (let c = 0; c < COLS; c++) {
    const td = tr.children[c + 1];
    if (td) updateCellDisplay(td, r, c);
  }
  renderCFVisible();
  applySelectionStyles();
}

function renderSheet() {
  for (let c = 0; c < COLS; c++) {
    if (sheet().colWidths?.[c]) applyColWidth(c, sheet().colWidths[c]);
  }
  ROW_H = sheet().rowHeight || 26;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};
  renderVirtual();
  renderCharts();
}

/* Alias — banyak kode lama panggil renderAll() */
function renderAll() {
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual();
}

/* =========================================================
   DISPLAY HELPERS
   ========================================================= */
function displayVal(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return 'value' in v ? String(v.value) : '';
  return String(v);
}
const isErr = v => v && typeof v === 'object' && v.type === 'ERROR';

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
  if (!filters.length) return null;      // null = no filter → identity

  if (cachedRowMap) return cachedRowMap;

  // Cari max row dengan data
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
      const addr = addrOf(r, c);
      const raw = sheet().cells[addr];
      const v = raw ? String(raw.v ?? '') : '';
      if (!f.values.includes(v)) { pass = false; break; }
    }
    if (pass) out.push(r);
  }
  cachedRowMap = out;
  return out;
}

/* =========================================================
   MERGE HELPERS
   ========================================================= */
function isAddrInRange(addr, rangeStr) {
  const [s, e] = rangeParts(rangeStr);
  const p = parseAddr(addr), p1 = parseAddr(s), p2 = parseAddr(e);
  if (!p || !p1 || !p2) return false;
  return p.r >= Math.min(p1.r, p2.r) && p.r <= Math.max(p1.r, p2.r)
      && p.c >= Math.min(p1.c, p2.c) && p.c <= Math.max(p1.c, p2.c);
}

function getMergeAt(addr) {
  const merges = sheet().merged || [];
  for (const rangeStr of merges) {
    if (isAddrInRange(addr, rangeStr)) {
      const [s, e] = rangeParts(rangeStr);
      const p1 = parseAddr(s), p2 = parseAddr(e);
      const tl = addrOf(Math.min(p1.r, p2.r), Math.min(p1.c, p2.c));
      return {
        range: rangeStr,
        tl,
        r0: Math.min(p1.r, p2.r), r1: Math.max(p1.r, p2.r),
        c0: Math.min(p1.c, p2.c), c1: Math.max(p1.c, p2.c)
      };
    }
  }
  return null;
}

function computeMergePixelSize(merge) {
  let totalW = 0, totalH = 0;
  for (let c = merge.c0; c <= merge.c1; c++) {
    totalW += sheet().colWidths?.[c] || 96;
  }
  for (let r = merge.r0; r <= merge.r1; r++) {
    totalH += sheet().rowHeights?.[r] || ROW_H;
  }
  return { totalW, totalH };
}

function applyFormat(td, s) {
  td.style.fontWeight = s.bold ? '700' : '';
  td.style.fontStyle = s.italic ? 'italic' : '';
  td.style.textDecoration =
    [s.underline ? 'underline' : '', s.strike ? 'line-through' : '']
      .filter(Boolean).join(' ') || '';
  td.style.fontFamily = s.fontFamily || '';
  td.style.fontSize = s.fontSize ? s.fontSize + 'px' : '';
  td.style.color = s.color || '';
  td.style.background = s.bg || '';
  td.style.textAlign = s.align || '';
  td.classList.toggle('wrap', !!s.wrap);

  td.style.borderTop = td.style.borderBottom =
    td.style.borderLeft = td.style.borderRight = '1px solid #e9eef4';
  const b = s.borders || {};
  if (b.top) td.style.borderTop = '2px solid #333';
  if (b.bottom) td.style.borderBottom = '2px solid #333';
  if (b.left) td.style.borderLeft = '2px solid #333';
  if (b.right) td.style.borderRight = '2px solid #333';
  if (b.all) {
    td.style.borderTop = td.style.borderBottom = '2px solid #333';
    td.style.borderLeft = td.style.borderRight = '2px solid #333';
  }
}

function formatNumber(v, fmt) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  if (isNaN(n)) return String(v);
  switch (fmt) {
    case 'number': return n.toLocaleString('id-ID');
    case 'currency': return 'Rp ' + n.toLocaleString('id-ID');
    case 'percent': return (n * 100).toFixed(2) + '%';
    case 'date': {
      const d = new Date(n > 1e10 ? n : Date.now());
      return isNaN(d) ? String(v) : d.toLocaleDateString('id-ID');
    }
    case 'text': return String(v);
    default: return String(v);
  }
}

/* =========================================================
   SELECTION
   ========================================================= */
function applySelectionStyles() {
  document.querySelectorAll('#grid td.active').forEach(td => td.classList.remove('active'));
  document.querySelectorAll('#grid td.selected').forEach(td => td.classList.remove('selected'));

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  const selR0 = Math.min(p1.r, p2.r), selR1 = Math.max(p1.r, p2.r);
  const selC0 = Math.min(p1.c, p2.c), selC1 = Math.max(p1.c, p2.c);
  const activeP = parseAddr(activeCell);

  const rowMap = getRowMap();
  for (let di = visibleStart; di < visibleEnd; di++) {
    const r = rowMap ? rowMap[di] : di;
    const tr = rowEls[r];
    if (!tr) continue;
    const tds = tr.querySelectorAll('td[data-addr]');
    tds.forEach(td => {
      const addr = td.dataset.addr;
      const p = parseAddr(addr);
      if (!p) return;
      // Td (TL atau normal) overlap dengan selection?
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
}

function renderRowHeaderHighlight() {
  document.querySelectorAll('#grid tbody th.active').forEach(th => th.classList.remove('active'));
  document.querySelectorAll('#grid thead th.active').forEach(th => th.classList.remove('active'));
  const p = parseAddr(activeCell);
  if (!p) return;
  const tr = rowEls[p.r];
  if (tr) tr.firstElementChild.classList.add('active');
  const cth = document.querySelector(`#grid thead th[data-col="${p.c}"]`);
  if (cth) cth.classList.add('active');
}

/* =========================================================
   MERGE — apply visual overlay (colSpan tidak dipakai karena virtual)
   ========================================================= */
function renderMergeVisuals() {
  // Bersihkan semua style merge
  document.querySelectorAll('#grid td.merged-tl, #grid td.merged-covered')
    .forEach(td => {
      td.classList.remove('merged-tl', 'merged-covered');
      td.style.removeProperty('background');
      td.style.removeProperty('border-color');
      const inner = td.querySelector('.cell-inner');
      if (inner) {
        inner.style.removeProperty('position');
        inner.style.removeProperty('width');
        inner.style.removeProperty('height');
        inner.style.removeProperty('background');
        inner.style.removeProperty('border');
        inner.style.removeProperty('z-index');
        inner.style.removeProperty('top');
        inner.style.removeProperty('left');
      }
    });

  const merges = sheet().merged || [];
  const rowMap = getRowMap();

  for (const rangeStr of merges) {
    const [s, e] = rangeParts(rangeStr);
    const p1 = parseAddr(s), p2 = parseAddr(e);
    if (!p1 || !p2) continue;
    const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
    const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);
    const tl = addrOf(r0, c0);

    // Top-left ada di visible?
    const tlTr = rowEls[r0];
    if (tlTr) {
      const tlTd = tlTr.children[c0 + 1];
      if (tlTd) {
        tlTd.classList.add('merged-tl');
        // Hitung ukuran total
        const { totalW, totalH } = computeMergePixelSize({ c0, c1, r0, r1 });
        const inner = tlTd.querySelector('.cell-inner');
        if (inner) {
          inner.style.position = 'absolute';
          inner.style.top = '0';
          inner.style.left = '0';
          inner.style.width = totalW + 'px';
          inner.style.height = totalH + 'px';
          inner.style.background = '#fff';
          inner.style.border = '1px solid #e9eef4';
          inner.style.zIndex = '3';
        }
      }
    }

    // Covered cells
    for (let r = Math.max(r0, visibleStart); r <= Math.min(r1, visibleEnd - 1); r++) {
      const tr = rowEls[r];
      if (!tr) continue;
      for (let c = c0; c <= c1; c++) {
        if (r === r0 && c === c0) continue;
        const td = tr.children[c + 1];
        if (!td) continue;
        td.classList.add('merged-covered');
      }
    }
  }
}

function updateSelectionUI() {
  document.getElementById('cell-address').textContent = activeCell;
  document.getElementById('formula-input').value = getRaw(activeCell);
  applySelectionStyles();
  renderRowHeaderHighlight();
  updateStatusStats();
}

function getRaw(addr) {
  const c = sheet().cells[addr];
  if (!c) return '';
  return c.f ?? (c.v ?? '');
}

function updateStatusStats() {
  const stat = document.getElementById('status-stat');
  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) { stat.textContent = ''; return; }

  const rMin = Math.min(p1.r, p2.r), rMax = Math.max(p1.r, p2.r);
  const cMin = Math.min(p1.c, p2.c), cMax = Math.max(p1.c, p2.c);
  const total = (rMax - rMin + 1) * (cMax - cMin + 1);

  if (total > 10000) { stat.textContent = total.toLocaleString('id-ID') + ' sel'; return; }

  let sum = 0, count = 0, numCount = 0;
  const hfR = hf.__maxR || ROWS;
  const hfC = hf.__maxC || COLS;
  for (let r = rMin; r <= rMax; r++) {
    for (let c = cMin; c <= cMax; c++) {
      const v = (r < hfR && c < hfC)
        ? hf.getCellValue({ sheet: 0, row: r, col: c })
        : '';
      const n = Number(v);
      if (!isNaN(n) && v !== '' && v !== null) { sum += n; numCount++; }
      if (v !== '' && v !== null && v !== undefined) count++;
    }
  }
  if (numCount > 0)
    stat.textContent = `Sum: ${sum.toLocaleString('id-ID')}  Avg: ${(sum / numCount).toLocaleString('id-ID')}  Count: ${count}`;
  else stat.textContent = count ? `Count: ${count}` : '';
}

/* =========================================================
   DELTA SAVE TRACKING
   ========================================================= */
function markCellDirty(addr) {
  const cur = sheet().cells[addr];
  dirtyCells[addr] = cur ? JSON.parse(JSON.stringify(cur)) : null;
  dirty = true;
  updateSaveIndicator();
}
function markMetaDirty(key) {
  dirtyMeta[key] = true;
  dirty = true;
  updateSaveIndicator();
}
function markDirty() { markMetaDirty('structure'); }
function updateSaveIndicator() {
  if (saving) return;
  const el = document.getElementById('save-status');
  if (el) el.textContent = '● belum disimpan';
}
function clearDelta() {
  dirtyCells = {};
  dirtyMeta = {};
  dirty = false;
  const el = document.getElementById('save-status');
  if (el) el.textContent = '';
}

/* =========================================================
   EDITING
   ========================================================= */
function findTd(addr) {
  return document.querySelector(`#grid td[data-addr="${addr}"]`);
}

function openEditor(addr, initial) {
    if (currentAccessMode === 'view') {
    toast('👁 Mode lihat saja — tidak bisa edit');
    return;
  }
  if (editor) return;
  const p = parseAddr(addr);
  if (!p) return;
  // pastikan baris visible dulu
  scrollAddrIntoView(addr);
  const td = findTd(addr);
  if (!td) return;

  const input = document.createElement('input');
  input.className = 'cell-input';
  input.value = initial !== undefined ? initial : getRaw(addr);
  const inner = td.querySelector('.cell-inner');
  if (inner) inner.textContent = '';
  td.appendChild(input);
  editor = { input, addr };
  input.focus();
  if (initial === undefined) input.select();
  else input.setSelectionRange(input.value.length, input.value.length);

  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
        e.preventDefault();
        const val = input.value;
        commit(addr, val);
        closeEditor();
        if (!editor) moveSel(1, 0);
    } else if (e.key === 'Escape') {
        e.preventDefault();
        closeEditor();
    } else if (e.key === 'Tab') {
        e.preventDefault();
        const val = input.value;
        commit(addr, val);
        closeEditor();
        if (!editor) moveSel(0, e.shiftKey ? -1 : 1);
    }
    });
  input.addEventListener('blur', () => {
    if (!editor || editor.input !== input) return;
    commit(addr, input.value);
    closeEditor();
    });
}

function closeEditor() {
  if (!editor) return;
  const { input, addr } = editor;
  editor = null;

  try {
    if (input && input.parentNode) input.parentNode.removeChild(input);
  } catch (e) {

  }

  const p = parseAddr(addr);
  if (p) {
    const tr = rowEls[p.r];
    if (tr) {
      const td = tr.children[p.c + 1];
      if (td) forceRenderCell(td, p.r, p.c);
    }
  }

  if (p) renderRow(p.r);
  selectCell(addr, false);
}

function commit(addr, value) {
  if (currentAccessMode === 'view') {
    toast('👁 Mode lihat saja');
    return;
  }
  const p = parseAddr(addr);
  if (!p) return;
  pushUndo();
  const raw = String(value);
  if (raw === '') {
    delete sheet().cells[addr];
    hf.setCellContents({ sheet: 0, row: p.r, col: p.c }, [[null]]);
  } else if (raw.startsWith('=')) {
    ensureCell(addr);
    sheet().cells[addr].f = raw;
    delete sheet().cells[addr].v;
    hf.setCellContents({ sheet: 0, row: p.r, col: p.c }, [[raw]]);
  } else {
    const n = Number(raw);
    const v = raw.trim() !== '' && !isNaN(n) ? n : raw;
    ensureCell(addr);
    sheet().cells[addr].v = v;
    delete sheet().cells[addr].f;
    hf.setCellContents({ sheet: 0, row: p.r, col: p.c }, [[v]]);
  }
  markCellDirty(addr);
  if (raw.startsWith('=') && /:\s*[A-Z]+\d{4,}/.test(raw)) {
    warnBigRange();
  }
  renderVisibleCells();
  renderCFVisible();
  applySelectionStyles();
  scheduleSave();
}

/* =========================================================
   SELECTION MOVEMENT
   ========================================================= */
function scrollAddrIntoView(addr) {
  const p = parseAddr(addr);
  if (!p) return;
  const gridWrapEl = document.getElementById('grid-wrap');
  const targetTop = p.r * ROW_H;
  const viewTop = gridWrapEl.scrollTop;
  const viewBottom = viewTop + gridWrapEl.clientHeight;
  if (targetTop < viewTop + ROW_H * 2) {
    gridWrapEl.scrollTop = Math.max(0, targetTop - ROW_H * 2);
  } else if (targetTop + ROW_H > viewBottom - ROW_H * 2) {
    gridWrapEl.scrollTop = targetTop - gridWrapEl.clientHeight + ROW_H * 3;
  }
  // kolom — biarkan browser handle (scrollLeft)
  const td = findTd(addr);
  if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function selectCell(addr, expandRange_ = false) {
  if (expandRange_) selection.end = addr;
  else { selection.start = addr; selection.end = addr; }
  activeCell = expandRange_ ? selection.start : addr;
  scrollAddrIntoView(activeCell);
  updateSelectionUI();
}

function moveSel(dr, dc) {
  const p = parseAddr(activeCell);
  if (!p) return;
  const r = Math.max(0, Math.min(ROWS - 1, p.r + dr));
  const c = Math.max(0, Math.min(COLS - 1, p.c + dc));
  selectCell(addrOf(r, c));
}

/* =========================================================
   COL & ROW RESIZE — Event Delegation (survive virtual scroll)
   ========================================================= */
document.addEventListener('mousedown', e => {
  if (editor) return;                     // sedang edit cell
  if (e.button !== 0) return;             // left click saja

  const colGrip = e.target.closest('.col-resize-grip');
  if (colGrip) {
    e.preventDefault();
    e.stopPropagation();
    startColResize(+colGrip.dataset.col, e.clientX);
    return;
  }

  const rowGrip = e.target.closest('.row-resize-grip');
  if (rowGrip) {
    e.preventDefault();
    e.stopPropagation();
    startRowResize(+rowGrip.dataset.row, e.clientY);
    return;
  }
}, true);   // capture phase

function startColResize(c, startX) {
  const startW = sheet().colWidths?.[c] || 96;
  document.body.style.cursor = 'col-resize';
  document.body.classList.add('resizing');

  const move = ev => {
    const w = Math.max(30, Math.min(600, startW + ev.clientX - startX));
    sheet().colWidths ||= {};
    sheet().colWidths[c] = w;
    applyColWidth(c, w);
  };
  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    document.body.style.cursor = '';
    document.body.classList.remove('resizing');
    markMetaDirty('structure');
    scheduleSave();
  };
  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startRowResize(r, startY) {
  // Row height uniform untuk semua baris (biar virtual scroll tetap konsisten)
  const startH = sheet().rowHeight || ROW_H;
  document.body.style.cursor = 'row-resize';
  document.body.classList.add('resizing');

  const move = ev => {
    const h = Math.max(18, Math.min(200, startH + ev.clientY - startY));
    sheet().rowHeight = h;
    setRowHeight(h);
  };
  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    document.body.style.cursor = '';
    document.body.classList.remove('resizing');
    markMetaDirty('structure');
    scheduleSave();
  };
  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

/* Set uniform row height — update visible rows + re-layout virtual scroll */
function setRowHeight(h) {
  ROW_H = h;
  // Update tinggi semua row yang terlihat
  document.querySelectorAll('#grid tbody tr').forEach(tr => {
    const th = tr.firstElementChild;
    if (th && th.dataset.row !== undefined) {
      tr.style.height = h + 'px';
    }
  });
  // Force re-render virtual (posisi scroll berubah)
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual();
}
/* =========================================================
   CELL SELECTION — mousedown (fix: handler ini hilang)
   ========================================================= */
document.getElementById('grid').addEventListener('mousedown', e => {
  // Skip klik pada grip resize
  if (e.target.closest('.col-resize-grip') || e.target.closest('.row-resize-grip')) return;

  // Kalau sedang edit cell, biarkan (blur akan commit)
  if (editor) return;

  const td = e.target.closest('td[data-addr]');
  if (!td) return;

  const addr = td.dataset.addr;

  // Kalau merged-TL dengan colspan/rowspan > 1 → select seluruh range
  if (td.classList.contains('merged-tl') && (td.colSpan > 1 || td.rowSpan > 1)) {
    const p = parseAddr(addr);
    const r1 = p.r + (td.rowSpan || 1) - 1;
    const c1 = p.c + (td.colSpan || 1) - 1;
    selection = { start: addr, end: addrOf(r1, c1) };
    activeCell = addr;
    scrollAddrIntoView(addr);
    updateSelectionUI();
    return;
  }

  if (e.shiftKey) selectCell(addr, true);
  else selectCell(addr);
});

/* =========================================================
   DOUBLE-CLICK GRIP → AUTO-FIT
   ========================================================= */
document.addEventListener('dblclick', e => {
  // Auto-fit col
  const colGrip = e.target.closest('.col-resize-grip');
  if (colGrip) {
    e.preventDefault();
    e.stopPropagation();
    const c = +colGrip.dataset.col;
    const w = autoFitColumn(c);
    toast(`📐 Kolom ${colName(c)} di-fit (${w}px)`);
    return;
  }

  // Auto-fit row — pakai tinggi baris (uniform) → pakai default 26
  const rowGrip = e.target.closest('.row-resize-grip');
  if (rowGrip) {
    e.preventDefault();
    e.stopPropagation();
    // Karena row height uniform, reset ke 26px
    sheet().rowHeight = 26;
    setRowHeight(26);
    markMetaDirty('structure');
    scheduleSave();
    toast('📐 Tinggi baris di-reset');
    return;
  }
}, true);   // capture: tangkap sebelum dblclick lain

/* Editor saat double-click */
document.getElementById('grid').addEventListener('dblclick', e => {
  if (e.target.closest('.col-resize-grip') || e.target.closest('.row-resize-grip')) {
    return;
  }
  const td = e.target.closest('td[data-addr]');
  if (td) openEditor(td.dataset.addr);
});

document.addEventListener('keydown', e => {
  if (editor) return;
  const tag = (e.target.tagName || '').toUpperCase();
  const inputFocused = tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
  const ctrl = e.ctrlKey || e.metaKey;

  if (currentAccessMode === 'view') {
    const blockedKeys = ['Backspace', 'Delete'];
    if (blockedKeys.includes(e.key)) {
      e.preventDefault();
      return toast('👁 Mode lihat saja');
    }
    if (ctrl && ['x', 'v', 'b', 'i', 'u'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      return toast('👁 Mode lihat saja');
    }
    if (e.key.length === 1 && !ctrl && !e.altKey && !e.metaKey) {
    }
  }

  if (ctrl && e.key.toLowerCase() === 's') { e.preventDefault(); saveWB(); return; }
  if (ctrl && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); return; }
  if (ctrl && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
  if (ctrl && e.key.toLowerCase() === 'b') { e.preventDefault(); toggleFmt('bold'); return; }
  if (ctrl && e.key.toLowerCase() === 'i') { e.preventDefault(); toggleFmt('italic'); return; }
  if (ctrl && e.key.toLowerCase() === 'u') { e.preventDefault(); toggleFmt('underline'); return; }
  if (ctrl && e.key.toLowerCase() === 'c') { if (!inputFocused) { e.preventDefault(); doCopy(); } return; }
  if (ctrl && e.key.toLowerCase() === 'x') { if (!inputFocused) { e.preventDefault(); doCut(); } return; }
  if (ctrl && e.key.toLowerCase() === 'v') { if (!inputFocused) { e.preventDefault(); doPaste(); } return; }

  if (inputFocused) return;

  const p = parseAddr(activeCell);
  if (!p) return;
  if (e.key === 'ArrowUp') { e.preventDefault(); e.shiftKey ? selectCell(addrOf(Math.max(0, p.r - 1), p.c), true) : moveSel(-1, 0); }
  else if (e.key === 'ArrowDown') { e.preventDefault(); e.shiftKey ? selectCell(addrOf(Math.min(ROWS - 1, p.r + 1), p.c), true) : moveSel(1, 0); }
  else if (e.key === 'ArrowLeft') { e.preventDefault(); e.shiftKey ? selectCell(addrOf(p.r, Math.max(0, p.c - 1)), true) : moveSel(0, -1); }
  else if (e.key === 'ArrowRight') { e.preventDefault(); e.shiftKey ? selectCell(addrOf(p.r, Math.min(COLS - 1, p.c + 1)), true) : moveSel(0, 1); }
  else if (e.key === 'Enter') { e.preventDefault(); openEditor(activeCell); }
  else if (e.key === 'F2') { e.preventDefault(); openEditor(activeCell); }
  else if (e.key === 'Tab') { e.preventDefault(); moveSel(0, e.shiftKey ? -1 : 1); }
  else if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    pushUndo();
    const parts = expandRange(selection.start, selection.end);
    for (const a of parts) {
      delete sheet().cells[a];
      const pp = parseAddr(a);
      if (pp) hf.setCellContents({ sheet: 0, row: pp.r, col: pp.c }, [[null]]);
      markCellDirty(a);
    }
    renderVisibleCells(); renderCFVisible(); applySelectionStyles();
  } else if (e.key.length === 1 && !ctrl && !e.altKey) {
    e.preventDefault(); openEditor(activeCell, e.key);
  }
});

document.getElementById('formula-input').addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); commit(activeCell, e.target.value); moveSel(1, 0); }
  else if (e.key === 'Escape') e.target.value = getRaw(activeCell);
});
document.getElementById('formula-input').addEventListener('blur', e => {
  if (e.target.value !== getRaw(activeCell)) commit(activeCell, e.target.value);
});

/* =========================================================
   CLIPBOARD
   ========================================================= */
function doCopy() {
  const parts = expandRange(selection.start, selection.end);
  clipboard = parts.map(a => ({
    addr: a,
    cell: sheet().cells[a] ? JSON.parse(JSON.stringify(sheet().cells[a])) : null
  }));
  toast('📋 ' + parts.length + ' sel dicopy');
}
function doCut() {
  doCopy(); pushUndo();
  for (const { addr } of clipboard) {
    delete sheet().cells[addr];
    const p = parseAddr(addr);
    if (p) hf.setCellContents({ sheet: 0, row: p.r, col: p.c }, [[null]]);
    markCellDirty(addr);
  }
  renderVisibleCells(); renderCFVisible(); applySelectionStyles();
  toast('✂️ cut');
}
function doPaste() {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  if (!clipboard?.length) return;
  const p = parseAddr(activeCell);
  if (!p) return;
  const origin = parseAddr(clipboard[0].addr);
  if (!origin) return;
  pushUndo();
  for (const { addr, cell } of clipboard) {
    const sp = parseAddr(addr);
    const target = addrOf(p.r + (sp.r - origin.r), p.c + (sp.c - origin.c));
    if (cell) sheet().cells[target] = JSON.parse(JSON.stringify(cell));
    else delete sheet().cells[target];
    markCellDirty(target);
  }
  // update HF untuk semua sel yg tersentuh
  for (const { addr } of clipboard) {
    const sp = parseAddr(addr);
    const target = addrOf(p.r + (sp.r - origin.r), p.c + (sp.c - origin.c));
    const tp = parseAddr(target);
    const c = sheet().cells[target];
    if (tp) hf.setCellContents({ sheet: 0, row: tp.r, col: tp.c }, [[c?.f || c?.v || null]]);
  }
  renderVisibleCells(); renderCFVisible(); applySelectionStyles();
  toast('📥 pasted');
}

/* =========================================================
   FORMATTING
   ========================================================= */
function toggleFmt(key) {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  pushUndo();
  const parts = expandRange(selection.start, selection.end);
  const first = sheet().cells[parts[0]];
  const newVal = !first?.s?.[key];
  for (const a of parts) {
    ensureCell(a);
    sheet().cells[a].s ||= {};
    sheet().cells[a].s[key] = newVal;
    markCellDirty(a);
  }
  renderVisibleCells(); renderCFVisible(); applySelectionStyles();
}
function applyToSelection(fn) {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  pushUndo();
  const parts = expandRange(selection.start, selection.end);
  for (const a of parts) {
    ensureCell(a);
    sheet().cells[a].s ||= {};
    fn(sheet().cells[a].s);
    markCellDirty(a);
  }
  renderVisibleCells(); renderCFVisible(); applySelectionStyles();
}

/* Ribbon bindings */
document.getElementById('btn-bold').onclick = () => toggleFmt('bold');
document.getElementById('btn-italic').onclick = () => toggleFmt('italic');
document.getElementById('btn-underline').onclick = () => toggleFmt('underline');
document.getElementById('btn-strike').onclick = () => toggleFmt('strike');
document.getElementById('color-fg').oninput = e => applyToSelection(s => s.color = e.target.value);
document.getElementById('color-bg').oninput = e => applyToSelection(s => s.bg = e.target.value);
document.querySelectorAll('[data-align]').forEach(b =>
  b.onclick = () => applyToSelection(s => s.align = b.dataset.align));
document.getElementById('btn-wrap').onclick = () => toggleFmt('wrap');
document.getElementById('font-family').onchange = e => applyToSelection(s => s.fontFamily = e.target.value);
document.getElementById('font-size').onchange = e => applyToSelection(s => s.fontSize = +e.target.value);
document.getElementById('num-format').onchange = e => applyToSelection(s => s.numFmt = e.target.value);
document.getElementById('border-style').onchange = e => {
  const v = e.target.value;
  applyToSelection(s => {
    if (!v) s.borders = {};
    else if (v === 'all') s.borders = { all: true };
    else s.borders = { [v]: true };
  });
  e.target.value = '';
};
document.getElementById('btn-clear-all').onclick = () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  pushUndo();
  for (const a of expandRange(selection.start, selection.end)) {
    delete sheet().cells[a];
    const p = parseAddr(a);
    if (p) hf.setCellContents({ sheet: 0, row: p.r, col: p.c }, [[null]]);
    markCellDirty(a);
  }
  renderVisibleCells(); renderCFVisible(); applySelectionStyles();
};
document.getElementById('btn-clear-fmt').onclick = () =>
  applyToSelection(s => { for (const k in s) delete s[k]; });

/* =========================================================
   UNDO / REDO
   ========================================================= */
function snapshot() {
  return JSON.stringify({ sheets: wbData.sheets, activeSheet: activeSheetIdx });
}
function pushUndo() {
  const snap = snapshot();
  if (snap.length > 2_000_000) return;
  undoStack.push(snap);
  while (undoStack.length > 20) undoStack.shift();
  redoStack = [];
}
function undo() {
  if (!undoStack.length) return toast('Tidak ada undo');
  redoStack.push(snapshot());
  const prev = JSON.parse(undoStack.pop());
  wbData.sheets = prev.sheets;
  activeSheetIdx = Math.min(prev.activeSheet, wbData.sheets.length - 1);
  resetHF(); renderSheetTabs(); renderSheet(); markDirty(); toast('↶ undo');
}
function redo() {
  if (!redoStack.length) return toast('Tidak ada redo');
  undoStack.push(snapshot());
  const next = JSON.parse(redoStack.pop());
  wbData.sheets = next.sheets;
  activeSheetIdx = Math.min(next.activeSheet, wbData.sheets.length - 1);
  resetHF(); renderSheetTabs(); renderSheet(); markDirty(); toast('↷ redo');
}
document.getElementById('btn-undo').onclick = undo;
document.getElementById('btn-redo').onclick = redo;

/* =========================================================
   SHEET TABS
   ========================================================= */
function renderSheetTabs() {
  const cont = document.getElementById('sheet-tabs');
  cont.innerHTML = '';
  wbData.sheets.forEach((s, i) => {
    const btn = document.createElement('button');
    btn.className = 'sheet-tab' + (i === activeSheetIdx ? ' active' : '');
    btn.textContent = s.name;
    btn.onclick = () => switchSheet(i);
    btn.ondblclick = () => {
      const n = prompt('Nama sheet:', s.name);
      if (n) { s.name = n; renderSheetTabs(); markDirty(); }
    };
    btn.oncontextmenu = e => {
      e.preventDefault();
      if (wbData.sheets.length <= 1) return;
      if (confirm(`Hapus sheet "${s.name}"?`)) {
        wbData.sheets.splice(i, 1);
        activeSheetIdx = Math.min(i, wbData.sheets.length - 1);
        resetHF(); renderSheetTabs(); renderSheet(); markDirty();
      }
    };
    cont.appendChild(btn);
  });
}
async function switchSheet(i) {
  if (i === activeSheetIdx) return;

  if (dirty) { try { await saveWB(); } catch {} }

  destroyChartsForSheet(activeSheetIdx);

  activeSheetIdx = i;
  wbData.activeSheet = i;
  dirtyMeta.activeSheet = true;

  await resetHFLazy('Membuka sheet…');
  renderSheetTabs();
  renderSheet();

  activeCell = 'A1';
  selection = { start: 'A1', end: 'A1' };
  updateSelectionUI();
}
document.getElementById('btn-add-sheet').onclick = () => {
  wbData.sheets.push({
    name: 'Sheet' + (wbData.sheets.length + 1),
    cells: {},
    colWidths: {}, rowHeight: 26,
    cf: [], charts: [], merged: [], filters: {}
  });
  switchSheet(wbData.sheets.length - 1);
  markDirty();
};

/* =========================================================
   RIBBON TABS
   ========================================================= */
document.querySelectorAll('#ribbon-tabs button').forEach(b => {
  b.onclick = () => {
    document.querySelectorAll('#ribbon-tabs button').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.ribbon-pane').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    document.querySelector(`.ribbon-pane[data-pane="${b.dataset.tab}"]`).classList.add('active');
  };
});

/* =========================================================
   INSERT FUNCTION
   ========================================================= */
document.querySelectorAll('[data-fn]').forEach(b => {
  b.onclick = () => {
    const fn = b.dataset.fn;
    openEditor(activeCell, `=${fn}()`);
    setTimeout(() => {
      if (editor) {
        const val = editor.input.value;
        editor.input.value = val.slice(0, -1);
        editor.input.setSelectionRange(editor.input.value.length, editor.input.value.length);
      }
    }, 0);
  };
});
document.getElementById('btn-fx').onclick = () => openEditor(activeCell, '=');

/* =========================================================
   INSERT / DELETE ROW & COL
   ========================================================= */
document.getElementById('btn-insert-row').onclick = () => {
  const p = parseAddr(activeCell);
  if (!p) return;
  pushUndo();
  const nc = {};
  for (const [a, c] of Object.entries(sheet().cells)) {
    const pp = parseAddr(a);
    if (!pp) continue;
    const t = pp.r >= p.r ? addrOf(pp.r + 1, pp.c) : a;
    nc[t] = c;
  }
  sheet().cells = nc;
  resetHF(); markDirty(); renderSheet();
};
document.getElementById('btn-del-row').onclick = () => {
  const p = parseAddr(activeCell);
  if (!p) return;
  pushUndo();
  const nc = {};
  for (const [a, c] of Object.entries(sheet().cells)) {
    const pp = parseAddr(a);
    if (!pp || pp.r === p.r) continue;
    const t = pp.r > p.r ? addrOf(pp.r - 1, pp.c) : a;
    nc[t] = c;
  }
  sheet().cells = nc;
  resetHF(); markDirty(); renderSheet();
};
document.getElementById('btn-insert-col').onclick = () => {
  const p = parseAddr(activeCell);
  if (!p) return;
  pushUndo();
  const nc = {};
  for (const [a, c] of Object.entries(sheet().cells)) {
    const pp = parseAddr(a);
    if (!pp) continue;
    const t = pp.c >= p.c ? addrOf(pp.r, pp.c + 1) : a;
    nc[t] = c;
  }
  sheet().cells = nc;
  resetHF(); markDirty(); renderSheet();
};
document.getElementById('btn-del-col').onclick = () => {
  const p = parseAddr(activeCell);
  if (!p) return;
  pushUndo();
  const nc = {};
  for (const [a, c] of Object.entries(sheet().cells)) {
    const pp = parseAddr(a);
    if (!pp || pp.c === p.c) continue;
    const t = pp.c > p.c ? addrOf(pp.r, pp.c - 1) : a;
    nc[t] = c;
  }
  sheet().cells = nc;
  resetHF(); markDirty(); renderSheet();
};

/* =========================================================
   SORT
   ========================================================= */
document.getElementById('btn-sort-asc').onclick = () => doSort(true);
document.getElementById('btn-sort-desc').onclick = () => doSort(false);
function doSort(asc) {
  const p = parseAddr(selection.start);
  if (!p) return;
  const col = p.c;

  // Cari rentang baris yg ada datanya
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const v = hf.getCellValue({ sheet: 0, row: r, col });
    if (v === '' || v === null || v === undefined) continue;
    rows.push({ r, v, cells: {} });
  }
  let maxC = 0;
  for (const [a] of Object.entries(sheet().cells)) {
    const pp = parseAddr(a);
    if (pp && pp.c > maxC) maxC = pp.c;
  }
  for (const item of rows) {
    for (let c = 0; c <= maxC; c++) {
      const a = addrOf(item.r, c);
      if (sheet().cells[a]) item.cells[c] = JSON.parse(JSON.stringify(sheet().cells[a]));
    }
  }
  rows.sort((a, b) => {
    const na = Number(a.v), nb = Number(b.v);
    if (!isNaN(na) && !isNaN(nb)) return asc ? na - nb : nb - na;
    return asc ? String(a.v).localeCompare(String(b.v)) : String(b.v).localeCompare(String(a.v));
  });
  pushUndo();
  for (const item of rows)
    for (let c = 0; c <= maxC; c++) delete sheet().cells[addrOf(item.r, c)];
  rows.forEach((item, i) => {
    for (const [c, cell] of Object.entries(item.cells))
      sheet().cells[addrOf(i, +c)] = cell;
  });
  resetHF(); renderSheet(); markDirty(); toast('Sorted');
}

/* =========================================================
   CONDITIONAL FORMATTING
   ========================================================= */
document.getElementById('btn-cf').onclick = () => {
  document.getElementById('cf-range').value = `${selection.start}:${selection.end}`;
  document.getElementById('modal-cf').classList.remove('hidden');
};
document.getElementById('cf-type').onchange = e => {
  document.getElementById('cf-value2-row').style.display =
    e.target.value === 'between' ? 'block' : 'none';
};
document.getElementById('btn-cf-cancel').onclick =
  () => document.getElementById('modal-cf').classList.add('hidden');
document.getElementById('btn-cf-apply').onclick = () => {
  const range = document.getElementById('cf-range').value;
  const type = document.getElementById('cf-type').value;
  const value = document.getElementById('cf-value').value;
  const value2 = document.getElementById('cf-value2').value;
  const bg = document.getElementById('cf-bg').value;
  const color = document.getElementById('cf-color').value;
  sheet().cf ||= [];
  sheet().cf.push({ range, type, value, value2, style: { bg, color } });
  document.getElementById('modal-cf').classList.add('hidden');
  markDirty(); renderCF();
};
document.getElementById('btn-cf-clear').onclick = () => {
  sheet().cf = []; markDirty(); renderCF(); toast('CF cleared');
};

function renderCF() { renderCFVisible(); }
function renderCFVisible() {
  document.querySelectorAll('#grid td.cf').forEach(td => {
    td.classList.remove('cf');
    td.style.background = '';
    td.style.color = '';
  });
  // re-apply format dasar dulu supaya base style pulih
  for (let r = visibleStart; r < visibleEnd; r++) {
    const tr = rowEls[r];
    if (!tr) continue;
    for (let c = 0; c < COLS; c++) {
      const td = tr.children[c + 1];
      if (!td) continue;
      applyFormat(td, sheet().cells[td.dataset.addr]?.s || {});
    }
  }

  for (const rule of (sheet().cf || [])) {
    const [s, e] = rangeParts(rule.range);
    const p1 = parseAddr(s), p2 = parseAddr(e);
    if (!p1 || !p2) continue;
    const rMin = Math.max(visibleStart, Math.min(p1.r, p2.r));
    const rMax = Math.min(visibleEnd - 1, Math.max(p1.r, p2.r));
    const cMin = Math.min(p1.c, p2.c), cMax = Math.max(p1.c, p2.c);

    for (let r = rMin; r <= rMax; r++) {
      const tr = rowEls[r];
      if (!tr) continue;
      for (let c = cMin; c <= cMax; c++) {
        const td = tr.children[c + 1];
        if (!td) continue;
        const v = hf.getCellValue({ sheet: 0, row: r, col: c });
        const n = Number(v);
        let match = false;
        switch (rule.type) {
          case 'greaterThan': match = !isNaN(n) && n > Number(rule.value); break;
          case 'lessThan': match = !isNaN(n) && n < Number(rule.value); break;
          case 'equal': match = String(v) === String(rule.value); break;
          case 'between': match = !isNaN(n) && n >= Number(rule.value) && n <= Number(rule.value2); break;
          case 'contains': match = String(v).toLowerCase().includes(String(rule.value).toLowerCase()); break;
        }
        if (match) {
          td.classList.add('cf');
          td.style.background = rule.style.bg;
          td.style.color = rule.style.color;
        }
      }
    }
  }
}

/* =========================================================
   CHARTS
   ========================================================= */
document.querySelectorAll('[data-chart]').forEach(b => {
  b.onclick = () => showChartPreview(b.dataset.chart);
});
let pendingChart = null;
const palette = i => ['#217346','#f39c12','#3498db','#e74c3c','#9b59b6','#1abc9c','#e67e22','#34495e'][i % 8];

function readChartData() {
  const parts = expandRange(selection.start, selection.end);
  const rows = new Set(), cols = new Set();
  for (const a of parts) {
    const p = parseAddr(a); if (!p) continue;
    rows.add(p.r); cols.add(p.c);
  }
  const rowsArr = [...rows].sort((a, b) => a - b);
  const colsArr = [...cols].sort((a, b) => a - b);
  const labels = [], datasets = [];
  for (const r of rowsArr) {
    const p = parseAddr(addrOf(r, colsArr[0]));
    labels.push(String(hf.getCellValue({ sheet: 0, row: p.r, col: p.c }) ?? ''));
  }
  for (let ci = 1; ci < colsArr.length; ci++) {
    const data = [];
    for (const r of rowsArr) {
      const p = parseAddr(addrOf(r, colsArr[ci]));
      const v = Number(hf.getCellValue({ sheet: 0, row: p.r, col: p.c }));
      data.push(isNaN(v) ? 0 : v);
    }
    datasets.push({ label: 'Series ' + ci, data, backgroundColor: palette(ci - 1) });
  }
  if (!datasets.length) {
    const data = [];
    for (const r of rowsArr) {
      const p = parseAddr(addrOf(r, colsArr[0]));
      const v = Number(hf.getCellValue({ sheet: 0, row: p.r, col: p.c }));
      data.push(isNaN(v) ? 0 : v);
    }
    datasets.push({ label: 'Data', data, backgroundColor: palette(0) });
  }
  return { labels, datasets };
}

function showChartPreview(type) {
  const ctx = document.getElementById('chart-canvas').getContext('2d');
  const data = readChartData();
  if (window._previewChart) window._previewChart.destroy();
  window._previewChart = new Chart(ctx, {
    type, data,
    options: { responsive: false, plugins: { title: { display: true, text: 'Preview' } } }
  });
  pendingChart = { type, range: `${selection.start}:${selection.end}`, data };
  document.getElementById('modal-chart').classList.remove('hidden');
}
document.getElementById('btn-chart-close').onclick = () => {
  document.getElementById('modal-chart').classList.add('hidden');
  if (window._previewChart) { window._previewChart.destroy(); window._previewChart = null; }
};
document.getElementById('btn-chart-pin').onclick = () => {
  if (!pendingChart) return;
  const id = 'c' + Date.now();
  sheet().charts ||= [];
  sheet().charts.push({
    id, type: pendingChart.type, range: pendingChart.range,
    x: 200, y: 200, w: 380, h: 260
  });
  document.getElementById('modal-chart').classList.add('hidden');
  if (window._previewChart) { window._previewChart.destroy(); window._previewChart = null; }
  renderCharts(); markDirty();
};

function renderCharts() {
  // 1. Bersihkan DOM chart window lama
  document.querySelectorAll('.chart-window').forEach(el => el.remove());

  // 2. Hancurkan instance chart untuk sheet AKTIF saja
  const cachedOld = chartCache.get(activeSheetIdx);
  if (cachedOld) {
    cachedOld.forEach(c => { try { c.destroy(); } catch {} });
    chartCache.delete(activeSheetIdx);
  }

  // 3. Rebuild chart untuk sheet AKTIF saja
  const map = new Map();
  for (const ch of (sheet().charts || [])) {
    const el = document.createElement('div');
    el.className = 'chart-window';
    el.style.left = ch.x + 'px'; el.style.top = ch.y + 'px';
    el.style.width = ch.w + 'px'; el.style.height = ch.h + 'px';
    el.innerHTML = `<div class="chart-header"><span class="chart-title">📊 ${ch.type}</span><button class="chart-close">✕</button></div><div class="chart-body"><canvas></canvas></div><div class="resize"></div>`;
    document.body.appendChild(el);

    const [s, e] = rangeParts(ch.range);
    const parts = expandRange(s, e);
    const rows = [...new Set(parts.map(a => parseAddr(a).r))].sort((a, b) => a - b);
    const cols = [...new Set(parts.map(a => parseAddr(a).c))].sort((a, b) => a - b);
    const labels = rows.map(r => String(hf.getCellValue({ sheet: 0, row: r, col: cols[0] }) ?? ''));
    const datasets = [];
    for (let ci = 1; ci < cols.length; ci++) {
      const data = rows.map(r => {
        const v = Number(hf.getCellValue({ sheet: 0, row: r, col: cols[ci] }));
        return isNaN(v) ? 0 : v;
      });
      datasets.push({ label: 'Series ' + ci, data, backgroundColor: palette(ci - 1) });
    }
    if (!datasets.length) {
      const data = rows.map(r => {
        const v = Number(hf.getCellValue({ sheet: 0, row: r, col: cols[0] }));
        return isNaN(v) ? 0 : v;
      });
      datasets.push({ label: 'Data', data, backgroundColor: palette(0) });
    }
    const ctx = el.querySelector('canvas').getContext('2d');
    const inst = new Chart(ctx, {
      type: ch.type, data: { labels, datasets },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
    });
    map.set(ch.id, inst);
    makeDraggable(el, el.querySelector('.chart-header'), ch);
    makeResizable(el, el.querySelector('.resize'), ch);
    el.querySelector('.chart-close').onclick = () => {
      sheet().charts = sheet().charts.filter(x => x.id !== ch.id);
      renderCharts(); markDirty();
    };
  }
  chartCache.set(activeSheetIdx, map);
}

/* Helper: hapus cache & destroy chart sheet tertentu (dipanggil saat hapus sheet) */
function destroyChartsForSheet(idx) {
  const cached = chartCache.get(idx);
  if (cached) {
    cached.forEach(c => { try { c.destroy(); } catch {} });
    chartCache.delete(idx);
  }
}
function makeDraggable(el, handle, model) {
  handle.addEventListener('mousedown', e => {
    if (e.target.classList.contains('chart-close')) return;
    const sx = e.clientX, sy = e.clientY, ox = model.x, oy = model.y;
    const mv = ev => {
      model.x = ox + ev.clientX - sx; model.y = oy + ev.clientY - sy;
      el.style.left = model.x + 'px'; el.style.top = model.y + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      markDirty();
    };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });
}
function makeResizable(el, handle, model) {
  handle.addEventListener('mousedown', e => {
    e.preventDefault();
    const sx = e.clientX, sy = e.clientY, ow = model.w, oh = model.h;
    const mv = ev => {
      model.w = Math.max(240, ow + ev.clientX - sx);
      model.h = Math.max(180, oh + ev.clientY - sy);
      el.style.width = model.w + 'px'; el.style.height = model.h + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      markDirty();
    };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });
}

/* =========================================================
   VIEW TOGGLES
   ========================================================= */
document.getElementById('chk-gridlines').onchange = e => {
  document.querySelectorAll('#grid td').forEach(td =>
    td.style.borderColor = e.target.checked ? '#e9eef4' : 'transparent');
};
document.getElementById('chk-headers').onchange = e => {
  document.querySelectorAll('#grid th').forEach(th =>
    th.style.visibility = e.target.checked ? 'visible' : 'hidden');
};
document.getElementById('zoom').oninput = e => {
  const z = e.target.value;
  document.getElementById('zoom-label').textContent = z + '%';
  document.getElementById('status-zoom').textContent = z + '%';
  const grid = document.getElementById('grid');
  grid.style.transform = `scale(${z / 100})`;
  grid.style.transformOrigin = 'top left';
};

/* =========================================================
   SAVE / LOAD — dengan queue & snapshot untuk hindari race
   ========================================================= */
async function saveWB(forceFull = false) {
  if (!wbId) return;
  if (saving) {
    saveAgain = true;
    return;
  }

  const cellsToSave = { ...dirtyCells };
  const metaToSave = { ...dirtyMeta };

  const hasCells = Object.keys(cellsToSave).length > 0;
  const hasMeta = metaToSave.name || metaToSave.activeSheet || metaToSave.structure;
  if (!hasCells && !hasMeta && !forceFull) return;

  saving = true;
  const statusEl = document.getElementById('save-status');
  if (statusEl) statusEl.textContent = '⟳ menyimpan…';

  try {
    let res;
    const useFull = forceFull || metaToSave.structure;
    const auth = { 'Authorization': 'Bearer ' + token };

    if (useFull) {
      res = await fetch(`/api/sheets/${wbId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json', ...auth },
        body: JSON.stringify({ name: wbName, data: wbData })
      });
    } else {
      res = await fetch(`/api/sheets/${wbId}/delta`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json', ...auth },
        body: JSON.stringify({
          sheetIdx: activeSheetIdx,
          cells: cellsToSave,
          meta: {
            name: metaToSave.name ? wbName : undefined,
            activeSheet: metaToSave.activeSheet ? activeSheetIdx : undefined
          }
        })
      });

      if (res.status === 404) {
        res = await fetch(`/api/sheets/${wbId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', ...auth },
          body: JSON.stringify({ name: wbName, data: wbData })
        });
      }
    }

    if (!res.ok) throw new Error('HTTP ' + res.status);

    for (const k of Object.keys(cellsToSave)) delete dirtyCells[k];
    for (const k of Object.keys(metaToSave)) delete dirtyMeta[k];

    const sisaCells = Object.keys(dirtyCells).length > 0;
    const sisaMeta = Object.keys(dirtyMeta).length > 0;
    dirty = sisaCells || sisaMeta;

    if (statusEl) {
      statusEl.textContent = dirty ? '● masih ada perubahan…' : '✔ tersimpan';
    }
    if (!dirty) {
      setTimeout(() => {
        if (statusEl && statusEl.textContent === '✔ tersimpan') statusEl.textContent = '';
      }, 1800);
    }
    saveRetryCount = 0;
  } catch (err) {
    console.error('save failed:', err);
    if (statusEl) statusEl.textContent = '⚠ gagal save — akan coba lagi';
    dirty = true;
    saveRetryCount = (saveRetryCount || 0) + 1;
    const delay = Math.min(30000, 2000 * saveRetryCount);
    setTimeout(() => { if (dirty && wbId) saveWB(); }, delay);
  } finally {
    saving = false;
    if (saveAgain) {
      saveAgain = false;
      saveWB();
    } else if (Object.keys(dirtyCells).length > 0 || Object.keys(dirtyMeta).length > 0) {
      saveWB();
    }
  }
}

/* =========================================================
   DEBOUNCE — letakkan persis di sini (setelah saveWB)
   ========================================================= */
function scheduleSave() {
  if (autosaveTimer) clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    autosaveTimer = null;
    saveWB();
  }, 1000);
}

setInterval(() => {
  if (dirty && wbId && !saving && !autosaveTimer) saveWB();
}, 10000);

async function openWB(id) {
  document.getElementById('home-screen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');

  if (wbId && dirty) {
    try { await saveWB(); } catch {}
  }

  if (hf) { try { hf.destroy(); } catch {} hf = null; }
  chartCache.forEach((map) => {
    map.forEach(c => { try { c.destroy(); } catch {} });
  });
  chartCache.clear();
  rowEls = {};
  visibleStart = -1;
  visibleEnd = -1;
  undoStack = [];
  redoStack = [];
  if (autosaveTimer) { clearTimeout(autosaveTimer); autosaveTimer = null; }

  // ---- Loading indicator ----
  showLoading('Memuat workbook…');

  let data;
  try {
    const res = await fetch(`/api/sheets/${id}`, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) {
      hideLoading();
      return toast('Gagal memuat');
    }
    data = await res.json();
  } catch (e) {
    hideLoading();
    return toast('Gagal memuat: ' + e.message);
  }

  updateLoading('Menyiapkan tampilan…');
  await yieldToUI();

  wbId = data.id;
  wbName = data.name;
  const wbNameEl = document.getElementById('wb-name');
  if (wbNameEl) wbNameEl.value = wbName;
  wbData = normalizeWB(data.data);
  activeSheetIdx = wbData.activeSheet || 0;
  clearDelta();

  const cellCount = Object.keys(sheet().cells || {}).length;
  if (cellCount > 2000) {
    updateLoading('Membangun formula engine (' + cellCount.toLocaleString('id-ID') + ' sel)…');
  }
  await resetHFLazy('Membangun formula engine…');

  renderSheetTabs();
  const gridWrapEl = document.getElementById('grid-wrap');
  if (gridWrapEl) gridWrapEl.scrollTop = 0;
  renderSheet();
  selectCell('A1');

  document.getElementById('modal-sheets')?.classList.add('hidden');
  hideLoading();

  currentAccessMode = data.access || 'owner';
  applyReadOnlyMode(currentAccessMode === 'view');
  if (data.access === 'view') {
    toast('👁 Mode lihat saja — Anda tidak dapat mengedit');
  }
}

function normalizeWB(d) {
  if (d && Array.isArray(d.sheets)) {
    d.sheets.forEach(s => {
      if (typeof s.rowHeight !== 'number') s.rowHeight = 26;
      if (!s.colWidths) s.colWidths = {};
      if (!s.merged) s.merged = [];
      if (!s.filters) s.filters = {};
    });
    return d;
  }
  return {
    sheets: [{
      name: 'Sheet1', cells: d || {},
      colWidths: {}, rowHeight: 26,
      cf: [], charts: [], merged: [], filters: {}
    }],
    activeSheet: 0
  };
}

/* =========================================================
   NEW WORKBOOK — MODAL
   ========================================================= */
let newWBMode = 'blank';       // 'blank' | 'template'
let newWBTemplateId = null;
let newWBTemplateList = [];

function openNewWBModal() {
  newWBMode = 'blank';
  newWBTemplateId = null;

  // Reset form
  const nameInput = document.getElementById('new-wb-name');
  nameInput.value = '';
  document.getElementById('new-wb-preview-name').textContent = 'Workbook Baru';
  document.getElementById('new-wb-name-hint').textContent = '';
  document.getElementById('new-wb-name-hint').className = 'field-hint';
  document.getElementById('new-wb-templates').classList.add('hidden');
  document.querySelectorAll('.new-wb-option').forEach(b => {
    b.classList.toggle('active', b.dataset.mode === 'blank');
  });

  // Preload templates (async, non-blocking)
  loadQuickTemplates();

  // Tampilkan modal
  document.getElementById('modal-new-wb').classList.remove('hidden');

  // Focus input
  setTimeout(() => nameInput.focus(), 50);
}

function closeNewWBModal() {
  document.getElementById('modal-new-wb').classList.add('hidden');
}

/* Preview nama live */
document.getElementById('new-wb-name')?.addEventListener('input', e => {
  const v = e.target.value.trim();
  const preview = document.getElementById('new-wb-preview-name');
  preview.textContent = v || 'Workbook Baru';

  const hint = document.getElementById('new-wb-name-hint');
  if (!v) {
    hint.textContent = '';
    hint.className = 'field-hint';
  } else if (v.length < 3) {
    hint.textContent = 'Minimal 3 karakter';
    hint.className = 'field-hint err';
  } else if (v.length > 80) {
    hint.textContent = 'Maksimal 80 karakter';
    hint.className = 'field-hint err';
  } else {
    hint.textContent = '✔ Nama sudah bagus';
    hint.className = 'field-hint ok';
  }
});

/* Mode picker */
document.querySelectorAll('.new-wb-option').forEach(btn => {
  btn.addEventListener('click', () => {
    newWBMode = btn.dataset.mode;
    document.querySelectorAll('.new-wb-option').forEach(b => {
      b.classList.toggle('active', b === btn);
    });
    const tplWrap = document.getElementById('new-wb-templates');
    if (newWBMode === 'template') {
      tplWrap.classList.remove('hidden');
    } else {
      tplWrap.classList.add('hidden');
      newWBTemplateId = null;
    }
  });
});

/* Load templates untuk quick-pick */
async function loadQuickTemplates() {
  const box = document.getElementById('new-wb-tpl-list');
  box.innerHTML = `
    <div class="new-wb-tpl-empty">
      <i class="fa-solid fa-spinner fa-spin"></i> Memuat…
    </div>`;

  try {
    const res = await fetch('/api/templates', {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    newWBTemplateList = await res.json();
  } catch (e) {
    box.innerHTML = `<div class="new-wb-tpl-empty">Gagal memuat template</div>`;
    return;
  }

  if (!newWBTemplateList.length) {
    box.innerHTML = `<div class="new-wb-tpl-empty">
      Belum ada template. Buat dari menu <strong>Template</strong>.
    </div>`;
    return;
  }

  box.innerHTML = '';
  newWBTemplateList.slice(0, 8).forEach(t => {
    const item = document.createElement('div');
    item.className = 'new-wb-tpl-item';
    item.dataset.tplId = t.id;
    item.innerHTML = `
      <i class="fa-solid fa-file-lines"></i>
      <div class="tpl-item-name">${escapeHtml(t.name)}</div>
    `;
    item.addEventListener('click', () => {
      newWBTemplateId = t.id;
      box.querySelectorAll('.new-wb-tpl-item').forEach(x => {
        x.classList.toggle('active', x === item);
      });
      // Auto isi nama kalau kosong
      const nameInput = document.getElementById('new-wb-name');
      if (!nameInput.value.trim()) {
        nameInput.value = `${t.name} - Copy`;
        nameInput.dispatchEvent(new Event('input'));
      }
    });
    box.appendChild(item);
  });
}

/* Generate nama unik dengan counter: "Test", "Test (2)", "Test (3)" */
async function generateUniqueName(baseName) {
  try {
    const res = await fetch('/api/sheets', {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) return baseName;
    const list = await res.json();
    const used = new Set(list.map(s => String(s.name || '').toLowerCase()));

    if (!used.has(baseName.toLowerCase())) return baseName;

    // Buang suffix (N) yang sudah ada
    const clean = baseName.replace(/\s*\(\d+\)\s*$/, '');
    let n = 2;
    while (n < 1000) {
      const candidate = `${clean} (${n})`;
      if (!used.has(candidate.toLowerCase())) return candidate;
      n++;
    }
    return `${clean} (${Date.now()})`;
  } catch {
    return baseName;
  }
}
/* Konfirmasi buat */
document.getElementById('btn-new-wb-confirm')?.addEventListener('click', async () => {
  const nameInput = document.getElementById('new-wb-name');
  const name = nameInput.value.trim();

  // Validasi client
  if (!name) {
    nameInput.focus();
    const hint = document.getElementById('new-wb-name-hint');
    hint.textContent = 'Nama wajib diisi';
    hint.className = 'field-hint err';
    nameInput.classList.add('shake');
    setTimeout(() => nameInput.classList.remove('shake'), 400);
    return;
  }
  if (name.length < 3) {
    const hint = document.getElementById('new-wb-name-hint');
    hint.textContent = 'Minimal 3 karakter';
    hint.className = 'field-hint err';
    return;
  }

  if (newWBMode === 'template' && !newWBTemplateId) {
    const box = document.getElementById('new-wb-tpl-list');
    box.classList.add('shake');
    setTimeout(() => box.classList.remove('shake'), 400);
    return;
  }

  // Disable tombol
  const btn = document.getElementById('btn-new-wb-confirm');
  const originalHTML = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Membuat…';

  const hint = document.getElementById('new-wb-name-hint');
  hint.textContent = '';
  hint.className = 'field-hint';

  try {
    let wbId;

    if (newWBMode === 'template' && newWBTemplateId) {
      const res = await fetch(`/api/templates/${newWBTemplateId}/use`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
        body: JSON.stringify({ name })
      });

      if (!res.ok) {
        let errMsg = 'Gagal membuat dari template';
        try {
          const j = await res.json();
          errMsg = j.error || errMsg;
        } catch {}
        if (!res.ok) {
        let errMsg = 'Gagal membuat workbook';
        try {
            const j = await res.json();
            errMsg = j.error || errMsg;
        } catch {}

        // ✅ Kalau error karena duplicate, tawarkan nama unik
        if (res.status === 409) {
            const suggested = await generateUniqueName(name);
            const ok = confirm(
            `${errMsg}\n\nGunakan nama "${suggested}"?`
            );
            if (ok) {
            nameInput.value = suggested;
            nameInput.dispatchEvent(new Event('input'));
            // Retry otomatis
            btn.disabled = false;
            btn.innerHTML = originalHTML;
            btn.click();
            return;
            }
        }
        throw new Error(errMsg);
        }
      }
      const wb = await res.json();
      wbId = wb.id;
    } else {
      // Buat blank
      const res = await fetch('/api/sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
        body: JSON.stringify({ name, data: blankWB() })
      });

      if (!res.ok) {
        // ✅ Baca pesan error dari server (termasuk 409 duplicate)
        let errMsg = 'Gagal membuat workbook';
        try {
          const j = await res.json();
          errMsg = j.error || errMsg;
        } catch {}
        if (!res.ok) {
            let errMsg = 'Gagal membuat workbook';
            try {
                const j = await res.json();
                errMsg = j.error || errMsg;
            } catch {}

            // ✅ Kalau error karena duplicate, tawarkan nama unik
            if (res.status === 409) {
                const suggested = await generateUniqueName(name);
                const ok = confirm(
                `${errMsg}\n\nGunakan nama "${suggested}"?`
                );
                if (ok) {
                nameInput.value = suggested;
                nameInput.dispatchEvent(new Event('input'));
                // Retry otomatis
                btn.disabled = false;
                btn.innerHTML = originalHTML;
                btn.click();
                return;
                }
            }
            throw new Error(errMsg);
            }
      }
      const wb = await res.json();
      wbId = wb.id;
    }

    closeNewWBModal();
    await openWB(wbId);
    toast('✔ Workbook "' + name + '" berhasil dibuat');
  } catch (e) {
    console.error(e);

    hint.textContent = '⚠ ' + e.message;
    hint.className = 'field-hint err';

    // Highlight input nama
    nameInput.classList.add('shake');
    setTimeout(() => nameInput.classList.remove('shake'), 500);
    nameInput.focus();
    nameInput.select();

    toast('⚠ ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalHTML;
  }
});

/* Batal */
document.getElementById('btn-new-wb-cancel')?.addEventListener('click', closeNewWBModal);

/* Enter di nama → langsung submit */
document.getElementById('new-wb-name')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    document.getElementById('btn-new-wb-confirm').click();
  } else if (e.key === 'Escape') {
    closeNewWBModal();
  }
});

/* Escape di modal → close */
document.getElementById('modal-new-wb')?.addEventListener('click', e => {
  if (e.target.id === 'modal-new-wb') closeNewWBModal();
});

/* Ganti fungsi newWB — buka modal, bukan prompt */
function newWB() {
  openNewWBModal();
}

async function listWBs() {
  const res = await fetch('/api/sheets', { headers: { 'Authorization': 'Bearer ' + token } });
  const list = await res.json();
  const cont = document.getElementById('sheets-list');
  cont.innerHTML = '';
  if (!list.length)
    cont.innerHTML = '<div class="muted" style="padding:20px;text-align:center">Belum ada workbook</div>';
  for (const s of list) {
    const row = document.createElement('div');
    row.className = 'sheet-row';
    const badge = s.access === 'owner'
      ? '<span class="share-badge badge-owner">owner</span>'
      : `<span class="share-badge badge-${s.access}">${s.access}</span>`;
    row.innerHTML = `
      <span class="name">${escapeHtml(s.name)} ${badge}</span>
      <span class="date">${new Date(s.updated_at).toLocaleString('id-ID')}</span>
      ${s.access === 'owner' ? '<span class="del" title="Hapus">✕</span>' : ''}
    `;
    row.onclick = async e => {
      if (e.target.classList.contains('del')) {
        if (!confirm(`Hapus workbook "${s.name}"?`)) return;
        await fetch(`/api/sheets/${s.id}`, {
          method: 'DELETE', headers: { 'Authorization': 'Bearer ' + token }
        });
        listWBs();
        return;
      }
      await openWB(s.id);
    };
    cont.appendChild(row);
  }
  document.getElementById('modal-sheets').classList.remove('hidden');
}

document.getElementById('btn-sheets-list').onclick = listWBs;
document.getElementById('btn-new-wb').onclick = newWB;
document.getElementById('btn-close-modal').onclick =
  () => document.getElementById('modal-sheets').classList.add('hidden');
document.getElementById('wb-name').onchange = e => {
  wbName = e.target.value;
  markMetaDirty('name');
  scheduleSave();
};

/* =========================================================
   IMPORT / EXPORT
   ========================================================= */
document.getElementById('btn-import').onclick = () => document.getElementById('file-input').click();
document.getElementById('file-input').onchange = async e => {
  const f = e.target.files[0];
  if (!f) return;
  e.target.value = '';

  if (!wbId) {
    await newWB();
    return;
  }

  const fd = new FormData();
  fd.append('file', f);

  showLoading('Mengimpor file…');

  try {
    const res = await fetch(`/api/sheets/${wbId}/import`, {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token },
      body: fd
    });

    if (!res.ok) {
      let errMsg = 'Import gagal';
      try {
        const j = await res.json();
        errMsg = j.error || errMsg;
      } catch {}
      hideLoading();
      toast('⚠ ' + errMsg);
      return;
    }

    const result = await res.json();
    // Save dulu, baru reload
    clearDelta();
    await openWB(wbId);
    toast(`✔ Import ${result.cells} sel`);
  } catch (err) {
    console.error('Import client error:', err);
    toast('⚠ Import gagal: ' + err.message);
  } finally {
    hideLoading();
  }
};

document.querySelectorAll('[data-export]').forEach(b => b.onclick = () => {
  if (!wbId) return toast('Simpan dulu');
  // Pastikan semua perubahan tersimpan di server dulu
  saveWB(true).then(() => {
    fetch(`/api/sheets/${wbId}/export/${b.dataset.export}`, {
      headers: { 'Authorization': 'Bearer ' + token }
    })
      .then(r => r.blob())
      .then(blob => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = wbName + '.' + b.dataset.export;
        a.click();
        URL.revokeObjectURL(url);
      });
  });
});

/* =========================================================
   TOAST
   ========================================================= */
function toast(msg) {
  const el = document.getElementById('save-status');
  el.textContent = msg;
  setTimeout(() => { if (el.textContent === msg) el.textContent = ''; }, 2200);
}

/* =========================================================
   AUTH
   ========================================================= */
function authMsg(m, type = 'err') {
  const el = document.getElementById('auth-msg');
  el.style.color = type === 'err' ? '#c0392b' : '#1e7d4f';
  el.textContent = m;
}
async function doLogin() {
  const u = document.getElementById('auth-user').value.trim();
  const p = document.getElementById('auth-pass').value;
  if (!u || !p) return authMsg('Isi username & password');
  const res = await fetch('/api/auth/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: u, password: p })
  });
  const data = await res.json();
  if (!res.ok) return authMsg(data.error || 'Gagal');
  token = data.token; currentUser = data.user;
  localStorage.setItem('wex-token', token);
  enterApp();
}
async function doRegister() {
  const u = document.getElementById('auth-user').value.trim();
  const p = document.getElementById('auth-pass').value;
  if (!u || !p) return authMsg('Isi username & password');
  if (p.length < 4) return authMsg('Password minimal 4 karakter');
  const res = await fetch('/api/auth/register', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: u, password: p })
  });
  const data = await res.json();
  if (!res.ok) return authMsg(data.error || 'Gagal');
  token = data.token; currentUser = data.user;
  localStorage.setItem('wex-token', token);
  enterApp();
}
document.getElementById('btn-login').onclick = doLogin;
document.getElementById('btn-register').onclick = doRegister;
document.getElementById('auth-pass').addEventListener('keydown',
  e => e.key === 'Enter' && doLogin());
document.getElementById('btn-logout').onclick = async () => {
  if (dirty) { try { await saveWB(); } catch {} }
  localStorage.removeItem('wex-token');
  location.reload();
};

/* =========================================================
   ENTER APP
   ========================================================= */
async function enterApp() {
  document.getElementById('auth-screen').classList.add('hidden');
  document.getElementById('home-who').textContent = currentUser?.username || '';
  document.getElementById('who').textContent = currentUser?.username || '';
  buildGrid();
  resetHF();
  renderSheetTabs();
  await showHome();
}

/* =========================================================
   SHARING
   ========================================================= */
document.getElementById('btn-share').onclick = async () => {
  if (!wbId) return toast('Simpan dulu');
  await renderShareList();
  document.getElementById('modal-share').classList.remove('hidden');
};
document.getElementById('btn-share-close').onclick =
  () => document.getElementById('modal-share').classList.add('hidden');

async function renderShareList() {
  const res = await fetch(`/api/sheets/${wbId}/shares`,
    { headers: { 'Authorization': 'Bearer ' + token } });
  const list = await res.json();
  const box = document.getElementById('share-list');
  box.innerHTML = list.length ? '' :
    '<div class="muted" style="padding:12px;text-align:center">Belum dibagikan ke siapa pun</div>';
  for (const s of list) {
    const row = document.createElement('div');
    row.className = 'rev-row';
    row.innerHTML = `
      <div class="body">
        <div class="name">${escapeHtml(s.username)}
          <span class="share-badge badge-${s.permission}">${s.permission}</span>
        </div>
        <div class="meta">Dibagikan oleh ${escapeHtml(s.shared_by_username || '-')} • ${new Date(s.created_at).toLocaleDateString('id-ID')}</div>
      </div>
      <button class="preview" data-perm="view">👁</button>
      <button class="preview" data-perm="edit">✏️</button>
      <button class="restore" data-del="1" style="background:#dc2626">✕</button>
    `;
    row.querySelectorAll('[data-perm]').forEach(b => b.onclick = async () => {
      await fetch(`/api/sheets/${wbId}/shares/${s.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
        body: JSON.stringify({ permission: b.dataset.perm })
      });
      renderShareList();
    });
    row.querySelector('[data-del]').onclick = async () => {
      if (!confirm(`Hapus akses ${s.username}?`)) return;
      await fetch(`/api/sheets/${wbId}/shares/${s.id}`,
        { method: 'DELETE', headers: { 'Authorization': 'Bearer ' + token } });
      renderShareList();
    };
    box.appendChild(row);
  }
}
document.getElementById('btn-share-add').onclick = async () => {
  const username = document.getElementById('share-username').value.trim();
  const permission = document.getElementById('share-permission').value;
  if (!username) return;
  const res = await fetch(`/api/sheets/${wbId}/shares`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    body: JSON.stringify({ username, permission })
  });
  const data = await res.json();
  if (!res.ok) return toast(data.error || 'Gagal');
  document.getElementById('share-username').value = '';
  renderShareList();
  toast('✔ Dibagikan');
};

/* =========================================================
   REVISIONS
   ========================================================= */
document.getElementById('btn-history').onclick = async () => {
  if (!wbId) return toast('Simpan dulu');
  await renderRevisions();
  document.getElementById('modal-history').classList.remove('hidden');
};
document.getElementById('btn-history-close').onclick =
  () => document.getElementById('modal-history').classList.add('hidden');

async function renderRevisions() {
  const res = await fetch(`/api/sheets/${wbId}/revisions`,
    { headers: { 'Authorization': 'Bearer ' + token } });
  const list = await res.json();
  const box = document.getElementById('revision-list');
  box.innerHTML = list.length ? '' :
    '<div class="muted" style="padding:12px;text-align:center">Belum ada revisi</div>';
  for (const r of list) {
    const row = document.createElement('div');
    row.className = 'rev-row';
    const size = r.size_bytes
      ? (r.size_bytes < 1024 ? r.size_bytes + ' B' : (r.size_bytes / 1024).toFixed(1) + ' KB')
      : '-';
    row.innerHTML = `
      <div class="body">
        <div class="name">${escapeHtml(r.label || 'Snapshot')}</div>
        <div class="meta">${new Date(r.created_at).toLocaleString('id-ID')} • oleh ${escapeHtml(r.created_by || '-')} • ${size}</div>
      </div>
      <button class="restore">↺ Restore</button>
    `;
    row.querySelector('.restore').onclick = async () => {
      if (!confirm('Kembalikan ke revisi ini? Perubahan saat ini akan disimpan sebagai revisi baru.')) return;
      const res = await fetch(`/api/sheets/${wbId}/revisions/${r.id}/restore`, {
        method: 'POST', headers: { 'Authorization': 'Bearer ' + token }
      });
      if (!res.ok) return toast('Gagal restore');
      const fresh = await res.json();
      wbData = normalizeWB(fresh.data);
      activeSheetIdx = wbData.activeSheet || 0;
      clearDelta();
      resetHF(); renderSheetTabs(); renderSheet();
      document.getElementById('modal-history').classList.add('hidden');
      toast('✔ Direstore');
    };
    box.appendChild(row);
  }
}
document.getElementById('btn-snapshot').onclick = async () => {
  const label = prompt('Label snapshot:', 'Snapshot ' + new Date().toLocaleString('id-ID'));
  if (!label) return;
  await saveWB(true); // full save dulu biar konsisten
  await fetch(`/api/sheets/${wbId}/revisions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    body: JSON.stringify({ label })
  });
  renderRevisions();
  toast('✔ Snapshot dibuat');
};

/* =========================================================
   TEMPLATES
   ========================================================= */
document.getElementById('btn-templates').onclick = async () => {
  await renderTemplates();
  document.getElementById('modal-templates').classList.remove('hidden');
};
document.getElementById('btn-tpl-close').onclick =
  () => document.getElementById('modal-templates').classList.add('hidden');

async function renderTemplates() {
  const res = await fetch('/api/templates',
    { headers: { 'Authorization': 'Bearer ' + token } });
  const list = await res.json();
  const box = document.getElementById('template-list');
  box.innerHTML = '';
  const grouped = {};
  for (const t of list) (grouped[t.category] ||= []).push(t);
  for (const [cat, items] of Object.entries(grouped)) {
    const h = document.createElement('h4');
    h.style.cssText = 'margin:14px 0 6px;color:#64748b;font-size:11px;text-transform:uppercase;letter-spacing:.5px';
    h.textContent = cat;
    box.appendChild(h);
    for (const t of items) {
      const card = document.createElement('div');
      card.className = 'tpl-card';
      card.innerHTML = `
        <div class="icon"><i class="fa-solid fa-file-lines"></i></div>
        <div class="body">
          <div class="name">${escapeHtml(t.name)}</div>
          <div class="desc">${escapeHtml(t.description || '')}</div>
          <div class="meta">
            ${t.is_public ? '🌐 Publik' : '🔒 Pribadi'} • ${t.used_count}× dipakai
          </div>
        </div>
      `;
      card.onclick = async () => {
        const res = await fetch(`/api/templates/${t.id}/use`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
          body: JSON.stringify({})
        });
        const wb = await res.json();
        document.getElementById('modal-templates').classList.add('hidden');
        await openWB(wb.id);
        toast('✔ Template diterapkan');
      };
      box.appendChild(card);
    }
  }
}

document.getElementById('btn-save-tpl').onclick = async () => {
  if (!wbId) return toast('Simpan dulu');
  const name = prompt('Nama template:', wbName);
  if (!name) return;
  const description = prompt('Deskripsi singkat:', '') || '';
  const category = prompt('Kategori (Keuangan/Bisnis/Produktivitas/dll):', 'Custom') || 'Custom';
  const isPublic = confirm('Jadikan publik? (OK=publik, Cancel=pribadi)');
  await saveWB(true);
  const res = await fetch('/api/templates', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    body: JSON.stringify({ name, description, category, data: wbData, is_public: isPublic })
  });
  if (!res.ok) return toast('Gagal simpan template');
  renderTemplates();
  toast('✔ Template tersimpan');
};

document.getElementById('tab-prev')?.addEventListener('click', () => {
  if (activeSheetIdx > 0) switchSheet(activeSheetIdx - 1);
});
document.getElementById('tab-next')?.addEventListener('click', () => {
  if (activeSheetIdx < wbData.sheets.length - 1) switchSheet(activeSheetIdx + 1);
});
document.getElementById('btn-copy').onclick  = doCopy;
document.getElementById('btn-cut').onclick   = doCut;
document.getElementById('btn-paste').onclick = doPaste;

/* =========================================================
   INIT — dengan safety net biar UI tetap jalan
   ========================================================= */
function bootApp() {
  // Bind tombol-tombol penting DULUAN, sebelum apapun
  const safeBind = (id, fn) => {
    const el = document.getElementById(id);
    if (el) el.onclick = fn;
    else console.warn('Missing element #' + id);
  };

  // Bind titlebar
  safeBind('btn-sheets-list', () => showHome());
  safeBind('btn-templates', async () => {
    try { await renderTemplates(); } catch(e) { console.error(e); toast('Gagal buka template'); }
    document.getElementById('modal-templates')?.classList.remove('hidden');
  });
  safeBind('btn-share', async () => {
    if (!wbId) return toast('Simpan dulu');
    await renderShareList();
    document.getElementById('modal-share')?.classList.remove('hidden');
  });
  safeBind('btn-history', async () => {
    if (!wbId) return toast('Simpan dulu');
    await renderRevisions();
    document.getElementById('modal-history')?.classList.remove('hidden');
  });
  safeBind('btn-logout', async () => {
    if (dirty) { try { await saveWB(); } catch {} }
    localStorage.removeItem('wex-token');
    location.reload();
  });

  // Auth
  safeBind('btn-login', doLogin);
  safeBind('btn-register', doRegister);
  document.getElementById('auth-pass')?.addEventListener('keydown',
    e => e.key === 'Enter' && doLogin());

  // Init async
  if (!token) return;
  (async () => {
    try {
      const res = await fetch('/api/sheets', {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      if (!res.ok) {
        localStorage.removeItem('wex-token');
        token = '';
        return;
      }
      currentUser = { username: 'User' };
      await enterApp();
    } catch (err) {
      console.error('init failed:', err);
      localStorage.removeItem('wex-token');
    }
  })();
}

/* Jalankan setelah DOM siap */
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', bootApp);
} else {
  bootApp();
}

/* Save sebelum tab close */
window.addEventListener('beforeunload', () => {
  if (!dirty || !wbId) return;
  try {
    fetch(`/api/sheets/${wbId}/delta`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token
      },
      body: JSON.stringify({
        sheetIdx: activeSheetIdx,
        cells: dirtyCells,
        meta: {
          name: dirtyMeta.name ? wbName : undefined,
          activeSheet: dirtyMeta.activeSheet ? activeSheetIdx : undefined
        }
      }),
      keepalive: true
    });
  } catch {}
});

/* =========================================================
   PRINT — dengan dukungan MERGED cells (v2)
   ========================================================= */

/* Hitung nilai + merge map untuk 1 sheet */
function computeSheetValuesForPrint(sheetObj) {
  let maxR = 0, maxC = 0;

  // 1. Scan cells
  for (const a of Object.keys(sheetObj.cells || {})) {
    const p = parseAddr(a);
    if (!p) continue;
    if (p.r + 1 > maxR) maxR = p.r + 1;
    if (p.c + 1 > maxC) maxC = p.c + 1;
  }

  // 2. Scan merges — bisa memperluas area
  const merges = sheetObj.merged || [];
  const mergeInfos = [];
  for (const rangeStr of merges) {
    if (!rangeStr || typeof rangeStr !== 'string') continue;
    const [s, e] = rangeParts(rangeStr);
    const p1 = parseAddr(s), p2 = parseAddr(e);
    if (!p1 || !p2) continue;
    const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
    const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);
    if (r1 + 1 > maxR) maxR = r1 + 1;
    if (c1 + 1 > maxC) maxC = c1 + 1;
    mergeInfos.push({ range: rangeStr, r0, r1, c0, c1, tl: addrOf(r0, c0) });
  }

  if (maxR === 0 || maxC === 0) return { values: [], rows: 0, cols: 0, merges: [] };

  // 3. Build matrix HF
  const matrix = Array.from({ length: maxR }, () => new Array(maxC).fill(null));
  for (const [a, c] of Object.entries(sheetObj.cells || {})) {
    const p = parseAddr(a);
    if (!p) continue;
    if (c.f) matrix[p.r][p.c] = c.f;
    else if (c.v !== undefined && c.v !== null && c.v !== '') matrix[p.r][p.c] = c.v;
  }

  // 4. Evaluate
  let tempHF;
  try {
    tempHF = HyperFormulaClass.buildFromArray(matrix, { licenseKey: HF_LICENSE });
  } catch {
    return {
      values: matrix.map(row => row.map(v => {
        if (v === null || v === undefined) return '';
        if (typeof v === 'string' && v.startsWith('=')) return '';
        return v;
      })),
      rows: maxR, cols: maxC, merges: mergeInfos
    };
  }

  const values = [];
  for (let r = 0; r < maxR; r++) {
    const row = [];
    for (let c = 0; c < maxC; c++) {
      const v = tempHF.getCellValue({ sheet: 0, row: r, col: c });
      if (v === null || v === undefined) row.push('');
      else if (typeof v === 'object') row.push(v.value ?? '');
      else row.push(v);
    }
    values.push(row);
  }
  tempHF.destroy();
  return { values, rows: maxR, cols: maxC, merges: mergeInfos };
}

/* Bangun HTML print — dengan colSpan/rowSpan */
function buildPrintHTML(mode, orientation, gridlines) {
  const sheetsToPrint = mode === 'all' ? wbData.sheets : [sheet()];
  let body = '';

  for (const s of sheetsToPrint) {
    const { values, rows, cols, merges } = computeSheetValuesForPrint(s);
    body += '<div class="sheet-page">';
    body += `<h1>${escapeHtml(wbName)} — ${escapeHtml(s.name)}</h1>`;

    if (rows === 0 || cols === 0) {
      body += '<div style="color:#999;font-style:italic;padding:20px 0">(kosong)</div>';
      body += '</div>';
      continue;
    }

    // ---- Bangun peta merge ----
    const mergeLookup = {};      // "r_c" → merge info
    const covered = new Set();   // "r_c" yang jadi bagian merge (bukan TL)
    for (const m of merges) {
      for (let r = m.r0; r <= m.r1; r++) {
        for (let c = m.c0; c <= m.c1; c++) {
          mergeLookup[`${r}_${c}`] = m;
          if (r !== m.r0 || c !== m.c0) covered.add(`${r}_${c}`);
        }
      }
    }

    body += '<table>';
    for (let r = 0; r < rows; r++) {
      body += '<tr>';
      for (let c = 0; c < cols; c++) {
        const key = `${r}_${c}`;
        if (covered.has(key)) continue;    // cell ini bagian dari merge di atasnya

        const m = mergeLookup[key];
        const isTL = m && m.r0 === r && m.c0 === c;
        const cs = isTL ? (m.c1 - m.c0 + 1) : 1;
        const rs = isTL ? (m.r1 - m.r0 + 1) : 1;

        const v = String(values[r]?.[c] ?? '');
        const attrs = [];
        if (cs > 1) attrs.push(`colspan="${cs}"`);
        if (rs > 1) attrs.push(`rowspan="${rs}"`);
        const cls = isTL ? ' class="merged"' : '';
        const attrStr = attrs.length ? ' ' + attrs.join(' ') : '';
        body += `<td${attrStr}${cls}>${escapeHtml(v)}</td>`;
      }
      body += '</tr>';
    }
    body += '</table>';
    body += '</div>';
  }

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Print</title>
<style>
  @page { size: A4 ${orientation}; margin: 12mm; }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 10pt; color: #111; padding: 10px; }
  h1 { font-size: 13pt; margin: 0 0 10px; color: #333; font-weight: 600; }
  table { border-collapse: collapse; }
  td {
    padding: 3px 6px;
    ${gridlines ? 'border: 1px solid #cccccc;' : 'border: none;'}
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 220px;
    vertical-align: middle;
  }
  td.merged {
    text-align: center;
    font-weight: 600;
    background: #f0fdf4;
    border: 1.5px solid #10b981 !important;
  }
  tr { page-break-inside: avoid; }
  .sheet-page { page-break-after: always; }
  .sheet-page:last-child { page-break-after: auto; }
</style></head><body>${body}</body></html>`;
}

/* Refresh iframe preview */
function refreshPrintPreview() {
  const iframe = document.getElementById('print-iframe');
  if (!iframe) return;
  iframe.srcdoc = buildPrintHTML(
    printOptions.mode,
    printOptions.orientation,
    printOptions.gridlines
  );
}

function openPrintModal() {
  if (!wbId) return toast('Simpan dulu');
  // Reset ke default
  printOptions = { mode: 'current', orientation: 'portrait', gridlines: true };
  document.querySelector('input[name="print-mode"][value="current"]').checked = true;
  document.querySelector('input[name="print-orient"][value="portrait"]').checked = true;
  document.getElementById('print-gridlines').checked = true;

  refreshPrintPreview();
  document.getElementById('modal-print').classList.remove('hidden');
}

/* Bind tombol Print */
document.getElementById('btn-print')?.addEventListener('click', openPrintModal);
document.getElementById('btn-print-close')?.addEventListener('click', () =>
  document.getElementById('modal-print').classList.add('hidden'));

/* Bind opsi preview */
document.querySelectorAll('input[name="print-mode"]').forEach(el => {
  el.addEventListener('change', e => {
    printOptions.mode = e.target.value;
    refreshPrintPreview();
  });
});
document.querySelectorAll('input[name="print-orient"]').forEach(el => {
  el.addEventListener('change', e => {
    printOptions.orientation = e.target.value;
    refreshPrintPreview();
  });
});
document.getElementById('print-gridlines')?.addEventListener('change', e => {
  printOptions.gridlines = e.target.checked;
  refreshPrintPreview();
});

/* Eksekusi print */
document.getElementById('btn-do-print')?.addEventListener('click', () => {
  const iframe = document.getElementById('print-iframe');
  if (!iframe) return;
  try {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
  } catch (e) {
    console.error(e);
    toast('Gagal membuka dialog print');
  }
});

/* =========================================================
   API KEYS
   ========================================================= */
async function openApiModal() {
  await loadApiKeys();
  await populateApiSheetSelect();
  document.getElementById('modal-api').classList.remove('hidden');
}

async function populateApiSheetSelect() {
  try {
    const res = await fetch('/api/sheets', {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    const list = await res.json();
    const sel = document.getElementById('api-sheet');
    sel.innerHTML = '<option value="">Semua Workbook</option>';
    for (const s of list) {
      const o = document.createElement('option');
      o.value = s.id;
      o.textContent = s.name + (s.access !== 'owner' ? ` (${s.access})` : '');
      if (s.access !== 'owner') o.disabled = true;
      sel.appendChild(o);
    }
  } catch (e) { console.error(e); }
}

async function loadApiKeys() {
  const box = document.getElementById('api-list');
  box.innerHTML = '<div class="muted" style="padding:12px;text-align:center">Memuat…</div>';
  try {
    const res = await fetch('/api/keys', {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    const list = await res.json();
    box.innerHTML = '';
    if (!list.length) {
      box.innerHTML = '<div class="muted" style="padding:14px;text-align:center">Belum ada kunci API. Buat satu untuk mulai.</div>';
      return;
    }
    for (const k of list) {
      const row = document.createElement('div');
      row.className = 'api-key-row';
      const sheetBadge = k.sheet_name
        ? `<span class="badge-sheet">${escapeHtml(k.sheet_name)}</span>`
        : '<span class="badge-sheet">Semua Sheet</span>';
      const lastUsed = k.last_used_at
        ? new Date(k.last_used_at).toLocaleString('id-ID')
        : 'belum pernah';
      row.innerHTML = `
        <div class="api-key-info">
          <div class="api-key-label">
            ${escapeHtml(k.label || 'Default')} ${sheetBadge}
          </div>
          <div class="api-key-value">${escapeHtml(k.key)}</div>
          <div class="api-key-meta">
            Dibuat ${new Date(k.created_at).toLocaleDateString('id-ID')} •
            Terakhir dipakai: ${lastUsed} •
            ${k.request_count} request
          </div>
        </div>
        <div class="api-key-actions">
          <button data-copy="${escapeHtml(k.key)}" title="Copy">
            <i class="fa-solid fa-copy"></i>
          </button>
          <button data-rotate="${k.id}" title="Regenerate">
            <i class="fa-solid fa-rotate"></i>
          </button>
          <button class="danger" data-del="${k.id}" title="Hapus">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      `;
      row.querySelector('[data-copy]').onclick = () => {
        navigator.clipboard.writeText(k.key);
        toast('✔ Kunci dicopy');
      };
      row.querySelector('[data-rotate]').onclick = async () => {
        if (!confirm('Regenerate kunci ini? Kunci lama akan langsung tidak valid.')) return;
        await fetch(`/api/keys/${k.id}/rotate`, {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token }
        });
        loadApiKeys();
        toast('✔ Kunci diregenerasi');
      };
      row.querySelector('[data-del]').onclick = async () => {
        if (!confirm('Hapus kunci ini?')) return;
        await fetch(`/api/keys/${k.id}`, {
          method: 'DELETE',
          headers: { 'Authorization': 'Bearer ' + token }
        });
        loadApiKeys();
        toast('✔ Kunci dihapus');
      };
      box.appendChild(row);
    }
  } catch (e) {
    box.innerHTML = '<div class="muted" style="padding:14px;text-align:center;color:#dc2626">Gagal memuat kunci</div>';
  }
}

/* Bind tombol API */
document.getElementById('btn-api')?.addEventListener('click', openApiModal);
document.getElementById('btn-api-close')?.addEventListener('click', () =>
  document.getElementById('modal-api').classList.add('hidden'));

document.getElementById('btn-api-create')?.addEventListener('click', async () => {
  const label = document.getElementById('api-label').value.trim() || 'Default';
  const sheet_id = document.getElementById('api-sheet').value || null;
  const res = await fetch('/api/keys', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    body: JSON.stringify({ label, sheet_id: sheet_id ? +sheet_id : null })
  });
  if (!res.ok) return toast('Gagal buat kunci');
  const data = await res.json();
  document.getElementById('api-label').value = '';
  await loadApiKeys();
  navigator.clipboard.writeText(data.key).catch(() => {});
  toast('✔ Kunci dibuat & dicopy ke clipboard');
});

document.getElementById('btn-api-copy')?.addEventListener('click', () => {
  const txt = document.getElementById('api-sample')?.textContent || '';
  navigator.clipboard.writeText(txt);
  toast('✔ Contoh dicopy');
});

/* Update contoh curl dengan base URL asli + key pertama */
function updateApiSample() {
  const base = location.origin;
  const pre = document.getElementById('api-sample');
  if (pre) {
    pre.textContent = `# 1. Set header API key
curl -H "X-API-Key: wex_xxxxx" ${base}/api/v1/sheets

# 2. Ambil data (formula dievaluasi)
curl -H "X-API-Key: wex_xxxxx" ${base}/api/v1/sheets/1/data

# 3. Ambil sebagai array of object
curl -H "X-API-Key: wex_xxxxx" ${base}/api/v1/sheets/1/json

# 4. Download CSV
curl -H "X-API-Key: wex_xxxxx" ${base}/api/v1/sheets/1/csv -o data.csv`;
  }
}
updateApiSample();

/* =========================================================
   FORMULA AUTOCOMPLETE + KLIK SEL SAAT MERUMUS
   ========================================================= */

/* -------- Daftar fungsi (100+ fungsi umum) -------- */
const FX_FUNCTIONS = [
  // Math & Trig
  { name: 'SUM',        sig: '(number1, [number2], …)',                       desc: 'Jumlahkan angka dalam range' },
  { name: 'SUMIF',      sig: '(range, criteria, [sum_range])',                desc: 'Jumlahkan dengan kondisi' },
  { name: 'SUMIFS',     sig: '(sum_range, range1, criteria1, …)',             desc: 'Jumlah dengan banyak kondisi' },
  { name: 'PRODUCT',    sig: '(number1, [number2], …)',                       desc: 'Kalikan angka' },
  { name: 'ABS',        sig: '(number)',                                      desc: 'Nilai absolut' },
  { name: 'ROUND',      sig: '(number, num_digits)',                          desc: 'Pembulatan' },
  { name: 'ROUNDUP',    sig: '(number, num_digits)',                          desc: 'Bulatkan ke atas' },
  { name: 'ROUNDDOWN',  sig: '(number, num_digits)',                          desc: 'Bulatkan ke bawah' },
  { name: 'CEILING',    sig: '(number, significance)',                        desc: 'Bulatkan ke kelipatan' },
  { name: 'FLOOR',      sig: '(number, significance)',                        desc: 'Bulatkan ke bawah kelipatan' },
  { name: 'MOD',        sig: '(number, divisor)',                             desc: 'Sisa bagi' },
  { name: 'POWER',      sig: '(number, power)',                               desc: 'Perpangkatan' },
  { name: 'SQRT',       sig: '(number)',                                      desc: 'Akar kuadrat' },
  { name: 'INT',        sig: '(number)',                                      desc: 'Bulatkan ke bawah' },
  { name: 'RAND',       sig: '()',                                            desc: 'Random 0–1' },
  { name: 'RANDBETWEEN',sig: '(bottom, top)',                                 desc: 'Random antara 2 nilai' },
  // Statistik
  { name: 'AVERAGE',    sig: '(number1, [number2], …)',                       desc: 'Rata-rata' },
  { name: 'AVERAGEIF',  sig: '(range, criteria, [average_range])',            desc: 'Rata-rata dengan kondisi' },
  { name: 'AVERAGEIFS', sig: '(average_range, range1, criteria1, …)',         desc: 'Rata-rata banyak kondisi' },
  { name: 'COUNT',      sig: '(value1, [value2], …)',                         desc: 'Hitung sel berisi angka' },
  { name: 'COUNTA',     sig: '(value1, [value2], …)',                         desc: 'Hitung sel tidak kosong' },
  { name: 'COUNTIF',    sig: '(range, criteria)',                             desc: 'Hitung dengan kondisi' },
  { name: 'COUNTIFS',   sig: '(range1, criteria1, …)',                        desc: 'Hitung banyak kondisi' },
  { name: 'COUNTBLANK', sig: '(range)',                                       desc: 'Hitung sel kosong' },
  { name: 'MAX',        sig: '(number1, [number2], …)',                       desc: 'Nilai terbesar' },
  { name: 'MIN',        sig: '(number1, [number2], …)',                       desc: 'Nilai terkecil' },
  { name: 'MEDIAN',     sig: '(number1, [number2], …)',                       desc: 'Nilai tengah' },
  { name: 'MODE',       sig: '(number1, [number2], …)',                       desc: 'Nilai tersering' },
  { name: 'STDEV',      sig: '(number1, [number2], …)',                       desc: 'Standar deviasi' },
  { name: 'VAR',        sig: '(number1, [number2], …)',                       desc: 'Varians' },
  { name: 'LARGE',      sig: '(array, k)',                                    desc: 'Nilai terbesar ke-k' },
  { name: 'SMALL',      sig: '(array, k)',                                    desc: 'Nilai terkecil ke-k' },
  { name: 'RANK',       sig: '(number, ref, [order])',                        desc: 'Peringkat' },
  // Logika
  { name: 'IF',         sig: '(condition, value_if_true, [value_if_false])',  desc: 'Kondisi' },
  { name: 'IFS',        sig: '(cond1, val1, cond2, val2, …)',                 desc: 'Banyak kondisi' },
  { name: 'IFERROR',    sig: '(value, value_if_error)',                       desc: 'Pengganti jika error' },
  { name: 'IFNA',       sig: '(value, value_if_na)',                          desc: 'Pengganti jika #N/A' },
  { name: 'AND',        sig: '(logical1, [logical2], …)',                     desc: 'Semua TRUE' },
  { name: 'OR',         sig: '(logical1, [logical2], …)',                     desc: 'Salah satu TRUE' },
  { name: 'NOT',        sig: '(logical)',                                     desc: 'Kebalikan logika' },
  { name: 'XOR',        sig: '(logical1, [logical2], …)',                     desc: 'XOR logika' },
  { name: 'SWITCH',     sig: '(expr, val1, res1, …, [default])',              desc: 'Switch case' },
  { name: 'TRUE',       sig: '()',                                            desc: 'Nilai TRUE' },
  { name: 'FALSE',      sig: '()',                                            desc: 'Nilai FALSE' },
  // Lookup
  { name: 'VLOOKUP',    sig: '(lookup, table, col_index, [range_lookup])',    desc: 'Cari vertikal' },
  { name: 'HLOOKUP',    sig: '(lookup, table, row_index, [range_lookup])',    desc: 'Cari horizontal' },
  { name: 'XLOOKUP',    sig: '(lookup, lookup_arr, return_arr, [not_found])', desc: '⭐ Cari modern — lebih cepat' },
  { name: 'INDEX',      sig: '(array, row_num, [col_num])',                   desc: 'Ambil nilai di posisi' },
  { name: 'MATCH',      sig: '(lookup, lookup_array, [match_type])',          desc: 'Cari posisi' },
  { name: 'CHOOSE',     sig: '(index_num, value1, [value2], …)',              desc: 'Pilih dari daftar' },
  { name: 'ROW',        sig: '([reference])',                                 desc: 'Nomor baris' },
  { name: 'COLUMN',     sig: '([reference])',                                 desc: 'Nomor kolom' },
  { name: 'ROWS',       sig: '(array)',                                       desc: 'Jumlah baris' },
  { name: 'COLUMNS',    sig: '(array)',                                       desc: 'Jumlah kolom' },
  { name: 'OFFSET',     sig: '(reference, rows, cols, [h], [w])',             desc: 'Offset dari sel' },
  // Teks
  { name: 'CONCAT',     sig: '(text1, [text2], …)',                           desc: 'Gabung teks' },
  { name: 'CONCATENATE',sig: '(text1, [text2], …)',                           desc: 'Gabung teks' },
  { name: 'TEXTJOIN',   sig: '(delim, ignore_empty, text1, …)',               desc: 'Gabung dengan pemisah' },
  { name: 'LEFT',       sig: '(text, [num_chars])',                           desc: 'Ambil kiri' },
  { name: 'RIGHT',      sig: '(text, [num_chars])',                           desc: 'Ambil kanan' },
  { name: 'MID',        sig: '(text, start, num_chars)',                      desc: 'Ambil tengah' },
  { name: 'LEN',        sig: '(text)',                                        desc: 'Panjang teks' },
  { name: 'LOWER',      sig: '(text)',                                        desc: 'Huruf kecil' },
  { name: 'UPPER',      sig: '(text)',                                        desc: 'Huruf besar' },
  { name: 'PROPER',     sig: '(text)',                                        desc: 'Kapital di awal kata' },
  { name: 'TRIM',       sig: '(text)',                                        desc: 'Hapus spasi berlebih' },
  { name: 'SUBSTITUTE', sig: '(text, old, new, [instance])',                  desc: 'Ganti teks' },
  { name: 'REPLACE',    sig: '(old_text, start, n, new_text)',                desc: 'Ganti di posisi' },
  { name: 'FIND',       sig: '(find, within, [start])',                       desc: 'Cari posisi (case-sens)' },
  { name: 'SEARCH',     sig: '(find, within, [start])',                       desc: 'Cari posisi (case-insens)' },
  { name: 'TEXT',       sig: '(value, format)',                               desc: 'Format angka jadi teks' },
  { name: 'VALUE',      sig: '(text)',                                        desc: 'Teks jadi angka' },
  { name: 'REPT',       sig: '(text, number_times)',                          desc: 'Ulangi teks' },
  // Tanggal
  { name: 'TODAY',      sig: '()',                                            desc: 'Tanggal hari ini' },
  { name: 'NOW',        sig: '()',                                            desc: 'Tanggal & waktu sekarang' },
  { name: 'DATE',       sig: '(year, month, day)',                            desc: 'Buat tanggal' },
  { name: 'YEAR',       sig: '(serial)',                                      desc: 'Ambil tahun' },
  { name: 'MONTH',      sig: '(serial)',                                      desc: 'Ambil bulan' },
  { name: 'DAY',        sig: '(serial)',                                      desc: 'Ambil hari' },
  { name: 'HOUR',       sig: '(serial)',                                      desc: 'Ambil jam' },
  { name: 'MINUTE',     sig: '(serial)',                                      desc: 'Ambil menit' },
  { name: 'SECOND',     sig: '(serial)',                                      desc: 'Ambil detik' },
  { name: 'DAYS',       sig: '(end, start)',                                  desc: 'Selisih hari' },
  { name: 'WEEKDAY',    sig: '(serial, [return_type])',                       desc: 'Hari ke-berapa' },
  // Informasi
  { name: 'ISBLANK',    sig: '(value)',                                       desc: 'Cek kosong' },
  { name: 'ISNUMBER',   sig: '(value)',                                       desc: 'Cek angka' },
  { name: 'ISTEXT',     sig: '(value)',                                       desc: 'Cek teks' },
  { name: 'ISERROR',    sig: '(value)',                                       desc: 'Cek error' },
  { name: 'ISNA',       sig: '(value)',                                       desc: 'Cek #N/A' },
  { name: 'NA',         sig: '()',                                            desc: 'Nilai #N/A' },
];

/* -------- State -------- */
let fxSuggestEl = null;
let fxSuggestIndex = -1;
let fxSuggestMatches = [];
let fxSuggestContext = null;

/* -------- Helper -------- */
function ensureFxEl() {
  if (fxSuggestEl) return fxSuggestEl;
  fxSuggestEl = document.createElement('div');
  fxSuggestEl.id = 'fx-suggest';
  fxSuggestEl.className = 'fx-suggest hidden';
  fxSuggestEl.addEventListener('mousedown', e => e.preventDefault()); // jangan blur input
  document.body.appendChild(fxSuggestEl);
  return fxSuggestEl;
}

function hideFxSuggest() {
  if (fxSuggestEl) fxSuggestEl.classList.add('hidden');
  fxSuggestIndex = -1;
  fxSuggestMatches = [];
  fxSuggestContext = null;
}

function getFxContext(input) {
  const val = input.value;
  if (!val.startsWith('=')) return null;
  const pos = input.selectionStart ?? val.length;
  let i = pos;
  while (i > 0 && /[A-Za-z]/.test(val[i - 1])) i--;
  if (i === pos) return null;               // tidak ada huruf
  const token = val.slice(i, pos);
  if (!token) return null;
  const before = i > 0 ? val[i - 1] : '';
  // token harus berada di awal formula, atau setelah operator/pembuka
  if (i > 0 && !['=', '(', ',', '+', '-', '*', '/', '<', '>', '&', '^', ':'].includes(before)) {
    return null;
  }
  return { startPos: i, endPos: pos, token };
}

function renderFxSuggest(input) {
  const el = ensureFxEl();
  el.innerHTML = fxSuggestMatches.map((f, idx) => `
    <div class="fx-item ${idx === fxSuggestIndex ? 'active' : ''}" data-idx="${idx}">
      <div class="fx-item-name">${f.name}</div>
      <div class="fx-item-sig">${f.sig}</div>
      <div class="fx-item-desc">${f.desc}</div>
    </div>
  `).join('');
  el.classList.remove('hidden');

  // posisikan di bawah sel editor / formula bar
  const anchor = input.closest('td') || input.parentElement;
  if (anchor) {
    const rect = anchor.getBoundingClientRect();
    let top = rect.bottom + 4;
    let left = rect.left;
    const W = 400, H = 260;
    if (left + W > window.innerWidth - 10) left = window.innerWidth - W - 10;
    if (left < 10) left = 10;
    if (top + H > window.innerHeight - 10) top = rect.top - H - 4;
    el.style.left = left + 'px';
    el.style.top = top + 'px';
  }

  // klik item = pilih
  el.querySelectorAll('.fx-item').forEach(item => {
    item.addEventListener('click', e => {
      e.preventDefault();
      e.stopPropagation();
      applyFxSuggestion(+item.dataset.idx);
    });
  });
}

function applyFxSuggestion(idx) {
  if (!editor || !fxSuggestContext) {
    // kalau yang aktif formula-input (bukan editor cell)
    const fi = document.getElementById('formula-input');
    if (fi && fxSuggestContext && document.activeElement === fi) {
      const f = fxSuggestMatches[idx];
      if (!f) return;
      const val = fi.value;
      const { startPos, endPos } = fxSuggestContext;
      const insert = f.name + '(';
      fi.value = val.slice(0, startPos) + insert + val.slice(endPos);
      const newPos = startPos + insert.length;
      fi.setSelectionRange(newPos, newPos);
      hideFxSuggest();
      return;
    }
    return;
  }
  const f = fxSuggestMatches[idx];
  if (!f) return;
  const input = editor.input;
  const val = input.value;
  const { startPos, endPos } = fxSuggestContext;
  const insert = f.name + '(';
  input.value = val.slice(0, startPos) + insert + val.slice(endPos);
  const newPos = startPos + insert.length;
  input.setSelectionRange(newPos, newPos);
  hideFxSuggest();
  input.focus();
}

function updateFxSuggest(input) {
  const ctx = getFxContext(input);
  if (!ctx) { hideFxSuggest(); return; }
  const up = ctx.token.toUpperCase();
  const matches = FX_FUNCTIONS.filter(f => f.name.startsWith(up));
  // jangan tampilkan kalau token = nama fungsi lengkap (1 match exact)
  if (!matches.length || (matches.length === 1 && matches[0].name === up)) {
    hideFxSuggest(); return;
  }
  fxSuggestContext = ctx;
  fxSuggestMatches = matches.slice(0, 12);
  fxSuggestIndex = 0;
  renderFxSuggest(input);
}

/* handler input & keydown untuk editor */
function fxOnInput(e) {
  const el = e.target;
  if (el.classList.contains('cell-input') || el.id === 'formula-input') {
    if (el.value.startsWith('=')) updateFxSuggest(el);
    else hideFxSuggest();
  }
}
function fxOnKeydown(e) {
  const el = fxSuggestEl;
  if (!el || el.classList.contains('hidden')) return;
  if (e.key === 'ArrowDown') {
    e.preventDefault(); e.stopPropagation();
    fxSuggestIndex = (fxSuggestIndex + 1) % fxSuggestMatches.length;
    renderFxSuggest(e.target);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault(); e.stopPropagation();
    fxSuggestIndex = (fxSuggestIndex - 1 + fxSuggestMatches.length) % fxSuggestMatches.length;
    renderFxSuggest(e.target);
  } else if (e.key === 'Tab' || e.key === 'Enter') {
    if (fxSuggestIndex >= 0 && fxSuggestMatches[fxSuggestIndex]) {
      e.preventDefault(); e.stopPropagation();
      applyFxSuggestion(fxSuggestIndex);
    }
  } else if (e.key === 'Escape') {
    e.preventDefault(); e.stopPropagation();
    hideFxSuggest();
  }
}

/* =========================================================
   INSERT REFERENCE SAAT EDIT FORMULA
   ========================================================= */
function insertRefToEditor(addr) {
  if (!editor) return false;
  const input = editor.input;
  const val = input.value;
  if (!val.startsWith('=')) return false;
  const pos = input.selectionStart ?? val.length;
  const newVal = val.slice(0, pos) + addr + val.slice(pos);
  input.value = newVal;
  const newPos = pos + addr.length;
  input.setSelectionRange(newPos, newPos);
  input.focus();
  return true;
}

/* =========================================================
   PATCH openEditor — tambah listener fx + ref insert
   ========================================================= */
const _origOpenEditor = openEditor;
openEditor = function(addr, initial) {
  _origOpenEditor(addr, initial);
  if (!editor) return;
  const input = editor.input;

  // Autocomplete listeners
  input.addEventListener('input', fxOnInput);
  input.addEventListener('keydown', fxOnKeydown, true); // capture: jalan sebelum Enter commit

  // Re-position dropdown kalau user scroll
  const onScroll = () => { if (fxSuggestEl && !fxSuggestEl.classList.contains('hidden')) {
    renderFxSuggest(input);
  } };
  input.addEventListener('blur', () => hideFxSuggest(), true);
};

/* =========================================================
   PATCH grid mousedown — insert ref, bukan commit
   ========================================================= */
// Cari listener lama dan lepas — tapi karena anonymous, kita override pakai capturing phase
const gridEl2 = document.getElementById('grid');
gridEl2.addEventListener('mousedown', e => {
  if (!editor) return;
  const input = editor.input;
  if (!input.value.startsWith('=')) return;

  const td = e.target.closest('td[data-addr]');
  if (!td) return;
  // klik di dalam td yang sama (tempat editor) — biarkan
  if (td === input.closest('td')) return;

  // intercept! jangan blur, jangan selectCell — insert reference ke formula
  e.preventDefault();
  e.stopImmediatePropagation();
  insertRefToEditor(td.dataset.addr);
  hideFxSuggest();
}, true); // capture phase — jalan sebelum handler lama

/* =========================================================
   PATCH formula-input — autocomplete juga
   ========================================================= */
const fiEl = document.getElementById('formula-input');
if (fiEl) {
  fiEl.addEventListener('input', fxOnInput);
  fiEl.addEventListener('keydown', fxOnKeydown, true);
  fiEl.addEventListener('blur', hideFxSuggest, true);
}

/* =========================================================
   MERGE HANDLERS
   ========================================================= */
document.getElementById('btn-merge')?.addEventListener('click', () => {
  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;
  const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);
  if (r0 === r1 && c0 === c1) return toast('Pilih lebih dari 1 sel');

  const range = addrOf(r0, c0) + ':' + addrOf(r1, c1);
  const merges = sheet().merged || [];

  // Cek overlap
  for (const m of merges) {
    const [s, e] = rangeParts(m);
    const mp1 = parseAddr(s), mp2 = parseAddr(e);
    if (!mp1 || !mp2) continue;
    const overlap = !(r1 < Math.min(mp1.r, mp2.r) || r0 > Math.max(mp1.r, mp2.r)
                   || c1 < Math.min(mp1.c, mp2.c) || c0 > Math.max(mp1.c, mp2.c));
    if (overlap) return toast('Range beririsan dengan merge yang ada');
  }

  pushUndo();
  merges.push(range);

  // Kumpulkan nilai dari sel non-top-left → gabungkan dengan spasi
  const tlAddr = addrOf(r0, c0);
  const parts = [];
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const a = addrOf(r, c);
      const cell = sheet().cells[a];
      const v = cell?.v;
      if (v !== undefined && v !== null && v !== '') parts.push(String(v));
      if (a !== tlAddr) {
        delete sheet().cells[a];
        try { hf.setCellContents({ sheet: 0, row: r, col: c }, [[null]]); } catch {}
        markCellDirty(a);
      }
    }
  }
  if (parts.length) {
    ensureCell(tlAddr);
    sheet().cells[tlAddr].v = parts.join(' ');
    sheet().cells[tlAddr].s = { ...(sheet().cells[tlAddr].s || {}), align: 'center', bold: true };
    markCellDirty(tlAddr);
  }

  sheet().merged = merges;
  markMetaDirty('structure');
  selection = { start: addrOf(r0, c0), end: addrOf(r1, c1) };
  activeCell = addrOf(r0, c0);
  renderSheet();
  updateSelectionUI();
  toast('✔ Sel digabungkan');
});

document.getElementById('btn-unmerge')?.addEventListener('click', () => {
  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;
  const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  const before = (sheet().merged || []).length;
  sheet().merged = (sheet().merged || []).filter(m => {
    const [s, e] = rangeParts(m);
    const mp1 = parseAddr(s), mp2 = parseAddr(e);
    if (!mp1 || !mp2) return true;
    const mr0 = Math.min(mp1.r, mp2.r), mr1 = Math.max(mp1.r, mp2.r);
    const mc0 = Math.min(mp1.c, mp2.c), mc1 = Math.max(mp1.c, mp2.c);
    // Buang merge yang beririsan dengan selection (bukan hanya terkandung)
    const overlap = !(r1 < mr0 || r0 > mr1 || c1 < mc0 || c0 > mc1);
    return !overlap;
  });

  if (sheet().merged.length === before) return toast('Tidak ada merge di range ini');

  pushUndo();
  markMetaDirty('structure');
  selectCell(addrOf(r0, c0));
  renderSheet();
  toast('✔ Merge dibatalkan');
});

/* =========================================================
   FILTER HANDLERS
   ========================================================= */
document.getElementById('btn-filter-toggle')?.addEventListener('click', () => {
  filterMode = !filterMode;
  const btn = document.getElementById('btn-filter-toggle');
  if (btn) {
    btn.style.background = filterMode ? 'linear-gradient(135deg,#10b981,#059669)' : '';
    btn.style.color = filterMode ? '#fff' : '';
  }
  renderFilterHeaders();
  toast(filterMode ? 'Mode filter aktif — klik ikon di header kolom' : 'Mode filter nonaktif');
});

document.getElementById('btn-filter-clear')?.addEventListener('click', () => {
  sheet().filters = {};
  invalidateRowMap();
  pushUndo();
  markMetaDirty('structure');
  renderSheet();
  renderFilterHeaders();
  toast('✔ Filter dibersihkan');
});

function renderFilterHeaders() {
  const thead = document.querySelector('#grid thead tr');
  if (!thead) return;
  // Hapus tombol filter lama
  thead.querySelectorAll('.fx-filter-btn').forEach(b => b.remove());
  thead.querySelectorAll('th[data-col]').forEach(th => {
    th.classList.remove('filtered', 'filter-active');
  });

  if (!filterMode) return;

  const filters = sheet().filters || {};
  thead.querySelectorAll('th[data-col]').forEach(th => {
    const c = +th.dataset.col;
    const btn = document.createElement('button');
    btn.className = 'fx-filter-btn';
    btn.innerHTML = '<i class="fa-solid fa-caret-down"></i>';
    btn.title = 'Filter kolom ini';
    btn.addEventListener('mousedown', e => {
      e.stopPropagation();
      e.preventDefault();
      openFilterPopup(c, th);
    });
    th.appendChild(btn);
    if (filters[c]?.values?.length) {
      th.classList.add('filtered', 'filter-active');
    }
  });
}

function openFilterPopup(colIdx, thEl) {
  closeFilterPopup();
  const rect = thEl.getBoundingClientRect();

  // Ambil unique values dari kolom
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
  const activeSet = active ? new Set(active) : null;   // null = semua

  const popup = document.createElement('div');
  popup.className = 'fx-filter-popup';
  const PW = 300, PH = 420;
  let left = rect.left;
  let top = rect.bottom + 4;

  // Jangan keluar kanan
  if (left + PW > window.innerWidth - 10) left = window.innerWidth - PW - 10;
  if (left < 10) left = 10;

  // Jangan keluar bawah — kalau tidak cukup, tampilkan di atas
  if (top + PH > window.innerHeight - 10) {
    const above = rect.top - PH - 4;
    top = above >= 10 ? above : Math.max(10, window.innerHeight - PH - 10);
  }

  popup.style.left = left + 'px';
  popup.style.top = top + 'px';

  const selected = new Set(activeSet ? [...activeSet] : sortedKeys);

  popup.innerHTML = `
    <div class="fx-fp-header"><i class="fa-solid fa-filter"></i> Filter Kolom ${colName(colIdx)}</div>
    <div class="fx-fp-search"><input type="text" placeholder="Cari nilai…" id="fx-fp-search"></div>
    <div class="fx-fp-list" id="fx-fp-list"></div>
    <div class="fx-fp-actions">
      <button id="fx-fp-all">Pilih Semua</button>
      <button id="fx-fp-none">Kosongkan</button>
      <button id="fx-fp-clear">Clear Filter</button>
      <button class="primary" id="fx-fp-apply">Terapkan</button>
    </div>
  `;
  document.body.appendChild(popup);
  filterPopupEl = popup;

  const list = popup.querySelector('#fx-fp-list');
  const renderList = (query = '') => {
    const q = query.toLowerCase();
    const filtered = sortedKeys.filter(k => !q || k.toLowerCase().includes(q));
    if (!filtered.length) {
      list.innerHTML = '<div class="fx-fp-item empty">Tidak ada nilai</div>';
      return;
    }
    list.innerHTML = filtered.map(k => {
      const checked = selected.has(k);
      const esc = escapeHtml(k);
      return `<label class="fx-fp-item">
        <input type="checkbox" data-v="${esc}" ${checked ? 'checked' : ''}>
        <span>${esc}</span>
        <small style="color:#94a3b8">${values.get(k)}</small>
      </label>`;
    }).join('');
    list.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', () => {
        const v = cb.dataset.v;
        if (cb.checked) selected.add(v);
        else selected.delete(v);
      });
    });
  };
  renderList();

  popup.querySelector('#fx-fp-search').addEventListener('input', e => renderList(e.target.value));
  popup.querySelector('#fx-fp-all').onclick = () => {
    sortedKeys.forEach(k => selected.add(k));
    renderList(popup.querySelector('#fx-fp-search').value);
  };
  popup.querySelector('#fx-fp-none').onclick = () => {
    selected.clear();
    renderList(popup.querySelector('#fx-fp-search').value);
  };
  popup.querySelector('#fx-fp-clear').onclick = () => {
    if (sheet().filters) delete sheet().filters[colIdx];
    invalidateRowMap();
    pushUndo();
    markMetaDirty('structure');
    closeFilterPopup();
    renderSheet();
    renderFilterHeaders();
    toast('✔ Filter kolom dibersihkan');
  };
  popup.querySelector('#fx-fp-apply').onclick = () => {
    sheet().filters ||= {};
    const list_ = sortedKeys;
    if (selected.size === list_.length) {
      delete sheet().filters[colIdx];   // semua nilai = tidak memfilter
    } else {
      sheet().filters[colIdx] = { values: [...selected] };
    }
    invalidateRowMap();
    pushUndo();
    markMetaDirty('structure');
    closeFilterPopup();
    renderSheet();
    renderFilterHeaders();
    toast('✔ Filter diterapkan');
  };

  // Tutup kalau klik di luar
  setTimeout(() => {
    document.addEventListener('mousedown', outsideFilterHandler, true);
  }, 0);
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

/* =========================================================
   HOOK ke openWB & switchSheet — render filter headers juga
   ========================================================= */
const _origRenderSheet = renderSheet;
renderSheet = function() {
  _origRenderSheet();
  recalcTableWidth();
  renderFilterHeaders();
};

/* Kalau user klik di luar input / dropdown → tutup dropdown */
document.addEventListener('mousedown', e => {
  if (fxSuggestEl && !fxSuggestEl.classList.contains('hidden')) {
    if (!e.target.closest('#fx-suggest')
        && !e.target.classList.contains('cell-input')
        && e.target.id !== 'formula-input') {
      hideFxSuggest();
    }
  }
}, true);

/* =========================================================
   HOME SCREEN
   ========================================================= */
async function showHome() {
  // Simpan workbook aktif kalau ada perubahan
  if (wbId && dirty && hf && !saving) {
    try { await saveWB(); } catch (e) { console.warn('save before home failed:', e); }
  }

  // Bersihkan state dengan hati-hati
  try { if (hf) { hf.destroy(); hf = null; } } catch {}
  chartCache.forEach(m => m.forEach(c => { try { c.destroy(); } catch {} }));
  chartCache.clear();
  rowEls = {};

  // Sembunyikan app, tampilkan home
  document.getElementById('app').classList.add('hidden');
  document.getElementById('home-screen').classList.remove('hidden');

  // Reset state workbook aktif
  wbId = null;
  wbName = 'Untitled';
  document.getElementById('wb-name').value = wbName;

  // Render list
  await renderHomeList();
}

async function renderHomeList() {
  const box = document.getElementById('home-list');
  box.innerHTML = `
    <div class="home-loading">
      <div class="spinner-ring"></div>
      <div>Memuat workbook…</div>
    </div>`;

  let list = [];
  try {
    const res = await fetch('/api/sheets', {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    list = await res.json();
  } catch (e) {
    box.innerHTML = `
      <div class="home-empty">
        <i class="fa-solid fa-triangle-exclamation"></i>
        <h3>Gagal memuat</h3>
        <p>${escapeHtml(e.message)}</p>
      </div>`;
    return;
  }

  // Terapkan filter dari input
  const searchVal = (document.getElementById('home-search')?.value || '').trim().toLowerCase();
  const filterVal = document.getElementById('home-filter')?.value || 'all';

  let filtered = list;
  if (filterVal === 'mine') filtered = filtered.filter(s => s.access === 'owner');
  else if (filterVal === 'shared') filtered = filtered.filter(s => s.access !== 'owner');
  if (searchVal) {
    filtered = filtered.filter(s =>
      String(s.name || '').toLowerCase().includes(searchVal)
    );
  }

  box.innerHTML = '';
  if (!filtered.length) {
    const isEmpty = !list.length;
    box.innerHTML = `
      <div class="home-empty">
        <i class="fa-solid ${isEmpty ? 'fa-folder-open' : 'fa-magnifying-glass'}"></i>
        <h3>${isEmpty ? 'Belum ada workbook' : 'Tidak ada hasil'}</h3>
        <p>${isEmpty ? 'Buat workbook pertama untuk mulai bekerja' : 'Coba kata kunci lain atau ubah filter'}</p>
        ${isEmpty ? `<button class="btn primary" onclick="newWB()" style="display:inline-flex;gap:8px">
          <i class="fa-solid fa-plus"></i> Workbook Baru
        </button>` : ''}
      </div>`;
    return;
  }

  for (const s of filtered) {
    const card = document.createElement('div');
    card.className = 'home-card' + (s.access !== 'owner' ? ' shared' : '');

    const isOwner = s.access === 'owner';
    const badgeClass = `badge-${s.access || 'view'}`;
    const badgeText = s.access === 'owner' ? 'Milik Saya' : s.access;

    let dateStr = '';
    try {
      dateStr = new Date(s.updated_at).toLocaleString('id-ID', {
        day: 'numeric', month: 'short', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      });
    } catch { dateStr = '-'; }

    card.innerHTML = `
      <div class="home-card-icon">
        <i class="fa-solid ${isOwner ? 'fa-file-lines' : 'fa-share-nodes'}"></i>
      </div>
      <div class="home-card-name">${escapeHtml(s.name)}</div>
      <div class="home-card-meta">
        <span class="home-card-badge ${badgeClass}">${escapeHtml(badgeText)}</span>
        <span>${escapeHtml(dateStr)}</span>
      </div>
      ${isOwner ? `<button class="home-card-del" data-del="${s.id}" title="Hapus">
        <i class="fa-solid fa-trash"></i>
      </button>` : ''}
    `;

    card.addEventListener('click', async (e) => {
      if (e.target.closest('[data-del]')) return;   // jangan buka kalau klik delete
      await openWB(s.id);
    });

    const delBtn = card.querySelector('[data-del]');
    if (delBtn) {
      delBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        if (!confirm(`Hapus workbook "${s.name}"?`)) return;
        const r = await fetch(`/api/sheets/${s.id}`, {
          method: 'DELETE',
          headers: { 'Authorization': 'Bearer ' + token }
        });
        if (r.ok) {
          toast('✔ Workbook dihapus');
          renderHomeList();
        } else {
          toast('Gagal menghapus');
        }
      });
    }

    box.appendChild(card);
  }
}

/* =========================================================
   HOME — Event Bindings
   ========================================================= */
document.getElementById('btn-home-new')?.addEventListener('click', newWB);
document.getElementById('btn-home-templates')?.addEventListener('click', async () => {
  await renderTemplates();
  document.getElementById('modal-templates').classList.remove('hidden');
});
document.getElementById('btn-home-logout')?.addEventListener('click', async () => {
  if (dirty) { try { await saveWB(); } catch {} }
  localStorage.removeItem('wex-token');
  location.reload();
});

/* Search & filter */
let homeSearchTimer = null;
document.getElementById('home-search')?.addEventListener('input', () => {
  clearTimeout(homeSearchTimer);
  homeSearchTimer = setTimeout(renderHomeList, 150);
});
document.getElementById('home-filter')?.addEventListener('change', renderHomeList);

/* Tombol Home di titlebar app */
document.getElementById('btn-go-home')?.addEventListener('click', () => {
  showHome();
});

/* =========================================================
   PATCH listWBs — sembunyikan modal lama, langsung ke home
   ========================================================= */
const _origListWBs = listWBs;
listWBs = async function () {
  // Kalau wbId null (sudah di home), refresh home
  if (!wbId) {
    return renderHomeList();
  }
  // Kalau di dalam workbook, tampilkan modal seperti biasa
  return _origListWBs();
};

/* =========================================================
   READ-ONLY MODE
   ========================================================= */
function applyReadOnlyMode(isReadOnly) {
  document.body.classList.toggle('readonly-mode', isReadOnly);

  // Apply read-only ke editor kalau sedang aktif
  if (isReadOnly && editor) {
    closeEditor();
  }
}

document.getElementById('btn-readonly-home')?.addEventListener('click', () => {
  showHome();
});

/* =========================================================
   SHARE — USERNAME AUTOCOMPLETE
   ========================================================= */
(function initShareAutocomplete() {
  const input = document.getElementById('share-username');
  const box = document.getElementById('share-user-suggest');
  const clearBtn = document.getElementById('share-username-clear');
  if (!input || !box) return;

  let searchTimer = null;

  function hideSuggest() {
    box.classList.add('hidden');
    shareUserSuggestIdx = -1;
  }

  function showSuggest(list) {
    if (!list.length) {
      box.innerHTML = '<div class="no-result">Tidak ada user ditemukan</div>';
      box.classList.remove('hidden');
      return;
    }
    box.innerHTML = list.map((u, i) => `
      <div class="share-user-item ${i === shareUserSuggestIdx ? 'active' : ''}" data-idx="${i}" data-username="${escapeHtml(u.username)}">
        <div class="avatar-sm">${escapeHtml(u.username.charAt(0))}</div>
        <div class="user-name">${escapeHtml(u.username)}</div>
        <div class="user-hint">ID ${u.id}</div>
      </div>
    `).join('');
    box.classList.remove('hidden');

    box.querySelectorAll('.share-user-item').forEach(item => {
      item.addEventListener('mousedown', e => {
        e.preventDefault();   // jangan blur input
        e.stopPropagation();
        input.value = item.dataset.username;
        hideSuggest();
        updateClearBtn();
      });
    });
  }

  function updateClearBtn() {
    if (input.value.trim()) clearBtn?.classList.add('show');
    else clearBtn?.classList.remove('show');
  }

  input.addEventListener('input', () => {
    updateClearBtn();
    clearTimeout(searchTimer);
    const q = input.value.trim();
    if (q.length < 1) { hideSuggest(); return; }
    searchTimer = setTimeout(async () => {
      try {
        const res = await fetch(`/api/sheets/users/search?q=${encodeURIComponent(q)}`, {
          headers: { 'Authorization': 'Bearer ' + token }
        });
        if (!res.ok) { hideSuggest(); return; }
        shareUserSuggestList = await res.json();
        shareUserSuggestIdx = -1;
        showSuggest(shareUserSuggestList);
      } catch {
        hideSuggest();
      }
    }, 180);
  });

  input.addEventListener('keydown', e => {
    if (box.classList.contains('hidden')) return;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      shareUserSuggestIdx = Math.min(shareUserSuggestIdx + 1, shareUserSuggestList.length - 1);
      showSuggest(shareUserSuggestList);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      shareUserSuggestIdx = Math.max(shareUserSuggestIdx - 1, 0);
      showSuggest(shareUserSuggestList);
    } else if (e.key === 'Enter') {
      if (shareUserSuggestIdx >= 0) {
        e.preventDefault();
        e.stopPropagation();
        input.value = shareUserSuggestList[shareUserSuggestIdx].username;
        hideSuggest();
        updateClearBtn();
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      hideSuggest();
    }
  });

  input.addEventListener('blur', () => {
    // Delay biar klik item tetap bisa ter-capture
    setTimeout(hideSuggest, 150);
  });

  clearBtn?.addEventListener('click', () => {
    input.value = '';
    updateClearBtn();
    input.focus();
    hideSuggest();
  });

  // Reset saat modal dibuka
  const observer = new MutationObserver(() => {
    const modal = document.getElementById('modal-share');
    if (modal && !modal.classList.contains('hidden')) {
      input.value = '';
      updateClearBtn();
      hideSuggest();
    }
  });
  const modal = document.getElementById('modal-share');
  if (modal) observer.observe(modal, { attributes: true, attributeFilter: ['class'] });
})();