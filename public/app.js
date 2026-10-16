/* =========================================================
   WEB EXCEL — Frontend v4
   Virtual Scroll + Delta Save + Autosave
   ========================================================= */
let ROWS = 10000;
const COLS = 50;
const MIN_ROWS = 10000;
const ROWS_BUFFER = 100;
let ROW_H = 26;
const BUFFER = 3;
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
let insertingRefMode = false;
let crossSheetRef = null;
let chartInstances = new Map();
let wbUsesChunks = false;

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
let filterHeaderRow = null;
let currentAccessMode = 'owner';
/* Row/Col Selection */
let selectedRows = new Set();
let selectedCols = new Set();
let shareUserSuggestList = [];
let shareUserSuggestIdx = -1;

let _wbHasPassword = false;
/* =========================================================
   CUSTOM DIALOG (ganti alert/confirm/prompt native)
   ========================================================= */
const ICON_MAP = {
  info:    { icon: 'fa-circle-info',         cls: 'info' },
  question:{ icon: 'fa-circle-question',     cls: '' },
  warn:    { icon: 'fa-triangle-exclamation',cls: 'warn' },
  danger:  { icon: 'fa-triangle-exclamation',cls: 'danger' },
  success: { icon: 'fa-circle-check',        cls: '' }
};

function _buildDialog({ title, message, type = 'question', inputType, inputValue, okText, cancelText, showCancel = true, okClass = 'primary' }) {
  return new Promise(resolve => {
    const overlay = document.createElement('div');
    overlay.className = 'custom-dialog-overlay';

    const iconCfg = ICON_MAP[type] || ICON_MAP.question;
    const hasInput = inputType != null;

    overlay.innerHTML = `
      <div class="custom-dialog-card">
        <div class="custom-dialog-icon ${iconCfg.cls}">
          <i class="fa-solid ${iconCfg.icon}"></i>
        </div>
        ${title ? `<div style="font-weight:700;font-size:15px;color:#0f172a;margin-bottom:6px">${escapeHtml(title)}</div>` : ''}
        <div class="custom-dialog-message">${message}</div>
        ${hasInput ? `<input class="custom-dialog-input" type="${inputType}" value="${escapeHtml(inputValue || '')}" />` : ''}
        <div class="custom-dialog-actions">
          ${showCancel ? `<button class="cc-btn" data-cancel>${escapeHtml(cancelText || 'Batal')}</button>` : ''}
          <button class="cc-btn ${okClass}" data-ok>${escapeHtml(okText || 'OK')}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const input = overlay.querySelector('.custom-dialog-input');
    const okBtn = overlay.querySelector('[data-ok]');
    const cancelBtn = overlay.querySelector('[data-cancel]');

    const cleanup = (result) => {
      overlay.remove();
      document.removeEventListener('keydown', keyHandler, true);
      resolve(result);
    };

    const keyHandler = (e) => {
      if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') {
        e.preventDefault();
        e.stopPropagation();
        const val = hasInput ? input.value : true;
        cleanup(val);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        cleanup(hasInput ? null : false);
      }
    };
    document.addEventListener('keydown', keyHandler, true);

    okBtn.addEventListener('click', () => {
      const val = hasInput ? input.value : true;
      cleanup(val);
    });
    if (cancelBtn) {
      cancelBtn.addEventListener('click', () => cleanup(hasInput ? null : false));
    }

    // Klik backdrop = cancel
    overlay.addEventListener('click', e => {
      if (e.target === overlay) cleanup(hasInput ? null : false);
    });

    // Focus
    setTimeout(() => {
      if (hasInput && input) { input.focus(); input.select(); }
      else okBtn.focus();
    }, 50);
  });
}

/* Alert — ganti alert() native */
async function customAlert(message, options = {}) {
  return _buildDialog({
    title: options.title,
    message: escapeHtml(String(message)),
    type: options.type || 'info',
    okText: options.okText || 'OK',
    showCancel: false,
    okClass: options.okClass || 'primary'
  });
}

/* Confirm — ganti confirm() native */
async function customConfirm(message, options = {}) {
  return _buildDialog({
    title: options.title,
    message: escapeHtml(String(message)),
    type: options.type || 'question',
    okText: options.okText || 'Ya',
    cancelText: options.cancelText || 'Batal',
    showCancel: true,
    okClass: options.okClass || 'danger'  // default danger untuk konfirmasi hapus
  });
}

/* Prompt — ganti prompt() native */
async function customPrompt(message, defaultValue = '', options = {}) {
  return _buildDialog({
    title: options.title,
    message: escapeHtml(String(message)),
    type: options.type || 'question',
    inputType: options.inputType || 'text',
    inputValue: defaultValue,
    okText: options.okText || 'OK',
    cancelText: options.cancelText || 'Batal',
    showCancel: true,
    okClass: 'primary'
  });
}

/* =========================================================
   OVERRIDE NATIVE ALERT
   ========================================================= */
// alert() sync tidak mungkin jadi async — kita override dengan toast (fire-and-forget)
const _nativeAlert = window.alert;
window.alert = function (msg) {
  // Pakai toast kalau ada, kalau tidak fallback ke native
  if (typeof toast === 'function') {
    toast(String(msg));
  } else {
    _nativeAlert.call(window, msg);
  }
};

// confirm() & prompt() — kita tidak override native (banyak yang sinkron), 
// tapi kita ganti pemanggilan di kode kita dengan versi async custom

function blankWB() {
  return {
    sheets: [{
      name: 'Sheet1', cells: {},
      colWidths: {}, rowHeight: 26,
      cf: [], charts: [], merged: [], filters: {}, images: []
    }],
    activeSheet: 0
  };
}
function sheet() { return wbData.sheets[activeSheetIdx]; }
function ensureCell(addr) { return (sheet().cells[addr] ||= {}); }

/* =========================================================
   AUTO-SCALE ROWS — deteksi dari data + chunks + merges
   ========================================================= */
async function recalcRowsLimit(silent = false) {
  const sheetId = wbId;
  const sheetIdx = activeSheetIdx;
  let maxRow = MIN_ROWS;

  // 1️⃣ Cek dari cells (yang sudah di-load)
  const cells = sheet().cells || {};
  for (const addr of Object.keys(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!m) continue;
    const r = parseInt(m[2], 10);
    if (r > maxRow) maxRow = r;
  }

  // 2️⃣ Cek dari chunks metadata (kalau chunk mode)
  if (wbUsesChunks && window.Chunks && sheetId) {
    try {
      // Cek apakah Chunks punya method getMaxBlockRow
      if (typeof Chunks.getMaxBlockRow === 'function') {
        const maxBlockRow = Chunks.getMaxBlockRow(sheetIdx);
        if (maxBlockRow && maxBlockRow + Chunks.BLOCK_ROWS > maxRow) {
          maxRow = maxBlockRow + Chunks.BLOCK_ROWS;
        }
      } else {
        // Fallback: fetch dari server
        const res = await fetch(`/api/sheets/${sheetId}/blocks`, {
          headers: { 'Authorization': 'Bearer ' + token }
        });
        if (res.ok) {
          const data = await res.json();
          const blocks = data.blocks?.[sheetIdx] || [];
          for (const b of blocks) {
            const endRow = (b.row || 0) + 1000; // BLOCK_ROWS default
            if (endRow > maxRow) maxRow = endRow;
          }
        }
      }
    } catch (e) {
      console.warn('Chunk blocks check failed:', e);
    }
  }

  // 3️⃣ Cek dari merges
  for (const mergeStr of (sheet().merged || [])) {
    const [s, e] = rangeParts(mergeStr);
    const p2 = parseAddr(e || s);
    if (p2 && p2.r + 1 > maxRow) maxRow = p2.r + 1;
  }

  // 4️⃣ Tambah buffer
  const newRows = Math.max(MIN_ROWS, maxRow + ROWS_BUFFER);

  // Cek kalau sudah cukup
  if (newRows <= ROWS) {
    if (!silent) console.log(`📏 ROWS sudah cukup: ${ROWS} (max data: ${maxRow})`);
    return ROWS;
  }

  // 🎯 Naikkan ROWS
  const oldRows = ROWS;
  ROWS = newRows;
  console.log(`📏 ROWS auto-scale: ${oldRows} → ${ROWS} (max data: ${maxRow})`);

  // Update UI
  if (typeof recalcTableWidth === 'function') recalcTableWidth();
  window._lastRenderKey = null;
  if (typeof renderVirtual === 'function') renderVirtual(false);

  if (!silent && typeof toast === 'function') {
    toast(`📏 Grid diperluas ke ${ROWS.toLocaleString('id-ID')} baris`);
  }

  return ROWS;
}

/* Helper: pastikan row tertentu ada di dalam ROWS */
function ensureRowCapacity(r) {
  if (r + ROWS_BUFFER <= ROWS) return false;
  const oldRows = ROWS;
  ROWS = r + ROWS_BUFFER;
  console.log(`📏 ROWS grow: ${oldRows} → ${ROWS} (butuh row ${r + 1})`);
  return true;
}

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
function updateLoadingProgress(text, percent) {
  const txt = document.getElementById('loading-text');
  if (txt) txt.textContent = `${text} (${percent}%)`;
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

/* =========================================================
   HF BUILD — hanya untuk bagian yang dibutuhkan
   Hindari build matrix 500k sel yang bikin lag
   ========================================================= */
function computeUsedRangeForSheet(s) {
  let maxR = 50, maxC = 26;
  const cells = s?.cells || {};
  for (const a of Object.keys(cells)) {
    const p = parseAddr(a);
    if (!p) continue;
    if (p.r + 1 > maxR) maxR = p.r + 1;
    if (p.c + 1 > maxC) maxC = p.c + 1;
  }

  // ⚡ BATASI: HF hanya build sampai 2000 baris × 50 kolom (= 100k sel)
  // Sisanya di-render on-the-fly tanpa HF
  const HF_MAX_R = 500;
  const HF_MAX_C = COLS;

  maxR = Math.min(maxR + 20, HF_MAX_R);
  maxC = Math.min(maxC + 5, HF_MAX_C);
  return { maxR, maxC };
}

function buildSheetMatrix(s) {
  const { maxR, maxC } = computeUsedRangeForSheet(s);
  const cells = s?.cells || {};
  const R = Math.max(maxR, 1);
  const C = Math.max(maxC, 1);

  const matrix = Array.from({ length: R }, () => new Array(C).fill(null));
  for (const [a, c] of Object.entries(cells)) {
    const p = parseAddr(a);
    if (!p || p.r >= R || p.c >= C) continue;

    if (c.f) {
      // 🎯 Pakai helper yang sama
      matrix[p.r][p.c] = resolveNamedRanges(c.f);
    }
    else if (c.v !== undefined && c.v !== null && c.v !== '') {
      matrix[p.r][p.c] = c.v;
    }
  }
  return matrix;
}

/* =========================================================
   ✅ RESET HF — versi MULTI-SHEET + auto re-render listener
   ========================================================= */
function resetHF() {
  _sheetIdCache.clear();
  if (hf) { try { hf.destroy(); } catch {} hf = null; }

  // ⚠️ PENTING: sheetsData harus OBJECT, bukan array
  // Format: { "Sheet1": [[...], [...]], "Sheet2": [[...]] }
  const sheetsData = {};
  wbData.sheets.forEach((s, idx) => {
    let name = s.name || ('Sheet' + (idx + 1));
    // Pastikan nama unik
    let unique = name;
    let n = 1;
    while (sheetsData[unique] !== undefined) {
      n++;
      unique = name + '_' + n;
    }
    sheetsData[unique] = buildSheetMatrix(s);
  });

  try {
    hf = HyperFormulaClass.buildFromSheets(sheetsData, {
      licenseKey: HF_LICENSE
    });
    hf.__maxR = ROWS;
    hf.__maxC = COLS;
    console.log('✅ HF built. Sheets:', hf.getSheetNames());

    // 🎯 LISTENER: auto re-render saat HF selesai evaluate
    attachHFListener();

  } catch (e) {
    console.error('❌ HF multi-sheet build failed:', e.message);
    // Fallback: single-sheet dengan sheet aktif
    try {
      const cur = sheet();
      const matrix = buildSheetMatrix(cur);
      const safeMatrix = (matrix && matrix.length && matrix[0] && matrix[0].length)
        ? matrix
        : [['']];
      hf = HyperFormulaClass.buildFromArray(safeMatrix, { licenseKey: HF_LICENSE });
      hf.__maxR = safeMatrix.length || 50;
      hf.__maxC = (safeMatrix[0] || []).length || 26;
      console.log('⚠️ Fallback single-sheet:', hf.getSheetNames());

      // 🎯 LISTENER juga untuk fallback
      attachHFListener();

    } catch (e2) {
      console.error('❌ Fallback failed:', e2.message);
      hf = HyperFormulaClass.buildFromArray([['']], { licenseKey: HF_LICENSE });
      hf.__maxR = 50;
      hf.__maxC = 26;

      // 🎯 LISTENER juga untuk fallback minimal
      attachHFListener();
    }
  }
}

/* =========================================================
   🎯 Attach listener ke HF — auto re-render saat values updated
   ========================================================= */
function attachHFListener() {
  if (!hf) return;

  try {
    const subscribe = hf.on || hf.subscribe || hf.addListener;
    if (typeof subscribe === 'function') {
      subscribe.call(hf, 'valuesUpdated', () => {
        if (typeof clearHFCache === 'function') clearHFCache();
        window._lastRenderKey = null;
        if (!window._hfRenderPending) {
          window._hfRenderPending = true;
          requestAnimationFrame(() => {
            window._hfRenderPending = false;
            try { renderVirtual(false); } catch {}
          });
        }
      });
      console.log('🎧 HF listener attached (valuesUpdated)');
    } else {
      // 🎯 Fallback: polling ringan setiap 300ms
      console.warn('⚠️ HF tidak support events, pakai polling fallback');
      if (window._hfPollTimer) clearInterval(window._hfPollTimer);
      window._hfPollTimer = setInterval(() => {
        // Skip kalau sedang tidak ada formula aktif
        const hasFormulas = Object.values(sheet().cells || {})
          .some(c => c && c.f);
        if (!hasFormulas) return;

        // Cek apakah ada perubahan nilai di cache vs HF
        let changed = false;
        const checkCells = ['B1', 'C1', 'D1', 'A2', 'B2']; // sample cells
        for (const addr of checkCells) {
          const p = parseAddr(addr);
          if (!p) continue;
          const cached = _hfValueCache.get(activeSheetIdx + ':' + addr);
          const current = hfGet(p.r, p.c);
          if (cached !== current) { changed = true; break; }
        }

        if (changed) {
          clearHFCache();
          window._lastRenderKey = null;
          try { renderVirtual(false); } catch {}
        }
      }, 500);
    }
  } catch (e) {
    console.warn('Cannot attach HF listener:', e.message);
  }
}

/* =========================================================
   FILTER STATE PER-SHEET
   ========================================================= */
function saveFilterStateToSheet() {
  const sh = sheet();
  if (!sh) return;
  sh._filterMode = filterMode;
  sh._filterHeaderRow = filterHeaderRow;
}

function loadFilterStateFromSheet() {
  const sh = sheet();
  filterMode = !!sh?._filterMode;
  filterHeaderRow = sh?._filterHeaderRow ?? null;
  if (typeof invalidateRowMap === 'function') invalidateRowMap();
}

/* 🎯 Cache sheetId — hindari lookup HF tiap sel */
let _sheetIdCache = new Map();

function getSheetIdSafe(sheetIdx = activeSheetIdx) {
  if (!hf) return null;

  // 🎯 Cek cache dulu
  const cacheKey = sheetIdx + ':' + (wbData.sheets[sheetIdx]?.name || '');
  const cached = _sheetIdCache.get(cacheKey);
  if (cached !== undefined) return cached;

  const sh = wbData.sheets[sheetIdx];
  if (!sh) return null;
  const sheetName = sh.name;
  if (!sheetName) return null;

  let id = undefined;
  try { id = hf.getSheetId(sheetName); } catch { id = undefined; }

  if (id === undefined || id === null) {
    try {
      hf.addSheet(sheetName);
      id = hf.getSheetId(sheetName);
      console.log('➕ Auto-added sheet:', sheetName, '→ ID:', id);

      if (window.Chunks && hf) {
        const shData = wbData.sheets[sheetIdx];
        if (wbUsesChunks && shData.cells && Object.keys(shData.cells).length === 0) {
          Promise.all([
            Chunks.ensureBlockLoaded(sheetIdx, 0),
            Chunks.ensureBlockLoaded(sheetIdx, 256),
            Chunks.ensureBlockLoaded(sheetIdx, 512),
            Chunks.ensureBlockLoaded(sheetIdx, 768)
          ]).then(() => {
            const loadedCells = Chunks.getLoadedCells(sheetIdx);
            if (hf && Object.keys(loadedCells).length > 0) {
              for (const [addr, cell] of Object.entries(loadedCells)) {
                const p = parseAddr(addr);
                if (!p) continue;
                const val = cell.f || cell.v || null;
                try {
                  hf.setCellContents({ sheet: id, row: p.r, col: p.c }, [[val]]);
                } catch {}
              }
              window._lastRenderKey = null;
              setTimeout(() => renderVirtual(false), 100);
            }
          });
        }
      }
    } catch (e) {
      console.warn('Cannot add sheet:', e.message);
    }
  }

  if (id === undefined || id === null) {
    try {
      resetHF();
      id = hf.getSheetId(sheetName);
    } catch {}
  }

  const finalId = (id === undefined || id === null) ? null : id;
  _sheetIdCache.set(cacheKey, finalId);
  return finalId;
}

/* ✅ hfGet — pakai numeric sheet ID */
function hfGet(r, c, sheetIdx = activeSheetIdx) {
  if (!hf) return '';
  const sheetId = getSheetIdSafe(sheetIdx);
  if (sheetId === null) return '';
  try {
    return hf.getCellValue({ sheet: sheetId, row: r, col: c });
  } catch {
    return '';
  }
}

/* ✅ hfSet — pakai numeric sheet ID */
function hfSet(r, c, val, sheetIdx = activeSheetIdx) {
  if (!hf) return false;
  const sheetId = getSheetIdSafe(sheetIdx);
  if (sheetId === null) {
    console.warn('hfSet: sheet not found for index', sheetIdx);
    return false;
  }
  try {
    // 🎯 Resolve named ranges sebelum kirim ke HF
    const resolved = resolveNamedRanges(val);
    hf.setCellContents({ sheet: sheetId, row: r, col: c }, [[resolved]]);
    return true;
  } catch (e) {
    console.warn('hfSet failed (id=' + sheetId + '):', e.message);
    return false;
  }
}

async function resetHFLazy(reason = 'Menyiapkan formula engine…') {
  if (hf) { try { hf.destroy(); } catch {} hf = null; }

  const cellCount = Object.keys(sheet().cells || {}).length;
  console.log('🔧 Building HF for', cellCount, 'cells');

  // Data kecil — langsung build
  if (cellCount < 5000) {
    resetHF();
    return;
  }

  // Data besar — pakai spinner + yield
  showLoading(reason);
  await yieldToUI();
  updateLoading(`Membangun formula engine (${cellCount.toLocaleString('id-ID')} sel)…`);
  await yieldToUI();

  const t0 = performance.now();
  try {
    resetHF();
    const t1 = performance.now();
    console.log(`✅ HF built in ${(t1 - t0).toFixed(0)}ms`);
  } catch (e) {
    console.error('HF build failed:', e);
    hf = HyperFormulaClass.buildFromArray([['']], { licenseKey: HF_LICENSE });
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
  const colHeader = document.createElement('col');
  colHeader.style.width = '48px';
  colHeader.style.minWidth = '48px';
  colHeader.style.maxWidth = '48px';
  cg.appendChild(colHeader);
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
    try { v = hfGet(p.r, c); }
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
let scrollStopTimer = null;

function onGridScroll() {
  if (scrollRAF) return;
  scrollRAF = requestAnimationFrame(() => {
    scrollRAF = null;
    renderVirtual(true);   // fast mode saat scroll

    clearTimeout(scrollStopTimer);
    scrollStopTimer = setTimeout(() => {
      renderVirtual(false);   // full render setelah stop 400ms
    }, 400);   // ← naikkan dari 200ms ke 400ms
  });
}

/* =========================================================
   HF VALUE CACHE
   ========================================================= */
const _hfValueCache = new Map();
function clearHFCache() { _hfValueCache.clear(); }
function invalidateHF(addr) { _hfValueCache.delete(addr); }

/* =========================================================
   RENDER VIRTUAL v4 — pakai HTML string, super cepat
   ========================================================= */
function renderVirtual(fastMode = false) {
  const gridWrapEl = document.getElementById('grid-wrap');
  const scrollTop = gridWrapEl.scrollTop;
  const viewH = gridWrapEl.clientHeight;

  const neededRows = Math.ceil(scrollTop / ROW_H) + 10;
  if (neededRows > ROWS) {
    ROWS = neededRows;
    console.log(`📏 ROWS auto-grow dari scroll: ${ROWS}`);
  }

  const rowMap = getRowMap();
  const totalRows = rowMap ? rowMap.length : ROWS;

  const startRow = Math.max(0, Math.floor(scrollTop / ROW_H) - 3);
  const endRow = Math.min(totalRows, Math.ceil((scrollTop + viewH) / ROW_H) + 3);

  const rangeKey = startRow + ':' + endRow + ':' + (fastMode ? 'F' : 'S');
  if (rangeKey === window._lastRenderKey) return;
  window._lastRenderKey = rangeKey;

  visibleStart = startRow;
  visibleEnd = endRow;

  const tbody = document.getElementById('grid-body');
  if (!tbody) return;

  // =====================================================
  // 🎯 Load chunk untuk range visible
  // =====================================================
    if (window.Chunks) {
    const blockRowsNeeded = new Set();
    const BLOCK_ROWS = Chunks.BLOCK_ROWS;
    for (let di = startRow; di < endRow; di++) {
      const r = rowMap ? rowMap[di] : di;
      blockRowsNeeded.add(Math.floor(r / BLOCK_ROWS) * BLOCK_ROWS);
    }

    const needsLoad = [];
    for (const b of blockRowsNeeded) {
      needsLoad.push(Chunks.ensureBlockLoaded(activeSheetIdx, b));
    }

    // 🎯 Setelah block yang visible ke-load, load berikutnya di background
    Promise.all(needsLoad).then(() => {
      // Re-render cepat
      window._lastRenderKey = null;
      requestAnimationFrame(() => renderVirtual(fastMode));

      // 🎯 Background: load 3 block berikutnya untuk smooth scroll
      const nextBlocks = [];
      for (let i = 1; i <= 3; i++) {
        const nextBlock = Math.max(...blockRowsNeeded) + i * BLOCK_ROWS;
        nextBlocks.push(Chunks.ensureBlockLoaded(activeSheetIdx, nextBlock));
      }
      Promise.all(nextBlocks).catch(() => {});
    });
  }

  // =====================================================
  // Bangun HTML
  // =====================================================
  const parts = [];

  if (startRow > 0) {
    parts.push('<tr style="height:' + (startRow * ROW_H) + 'px"><td colspan="' + (COLS + 1) + '" style="padding:0;border:0"></td></tr>');
  }

  for (let di = startRow; di < endRow; di++) {
    const r = rowMap ? rowMap[di] : di;
    parts.push(buildRowHTML(r, fastMode));
  }

  if (endRow < totalRows) {
    parts.push('<tr style="height:' + ((totalRows - endRow) * ROW_H) + 'px"><td colspan="' + (COLS + 1) + '" style="padding:0;border:0"></td></tr>');
  }

  tbody.innerHTML = parts.join('');

  // Rebuild rowEls
  rowEls = {};
  const trs = tbody.children;
  for (let i = 0; i < trs.length; i++) {
    const tr = trs[i];
    const th = tr.firstElementChild;
    if (th && th.dataset && th.dataset.row !== undefined) {
      rowEls[+th.dataset.row] = tr;
    }
  }

  if (!fastMode) {
    try { renderCFVisible(); } catch(e) {}
    try { renderRowHeaderHighlight(); } catch(e) {}
    try { applyValidationIndicators(); } catch(e) {}
    try { applyCommentMarkers(); } catch(e) {}
  }

  applySelectionStylesFast();
}

function buildRowHTML(r, fastMode) {
  const colWidths = sheet().colWidths || {};
  const headerRow = getFilterHeaderRow();
  const isHeaderRow = filterMode && (r === headerRow);
  const filters = sheet().filters || {};

  let h = '<tr data-row="' + r + '" style="height:' + ROW_H + 'px">';

  // Row header
  const rowHeaderClass = isHeaderRow ? ' class="filter-header-row"' : '';
  const summaryGroup = getGroupSummaryAt(r);
  const insideGroup = getRowGroupAt(r);

  let thExtra = '';
  if (summaryGroup) {
    const label = summaryGroup.collapsed ? '+' : '−';
    const title = summaryGroup.collapsed ? 'Expand' : 'Collapse';
    thExtra = `<button class="row-group-btn${summaryGroup.collapsed ? '' : ' expanded'}" data-group-from="${summaryGroup.from}" title="${title}" type="button">${label}</button>`;
  }

  const thClass = (rowHeaderClass || '') + (insideGroup && !summaryGroup ? ' row-group-line' : '');
  h += '<th data-row="' + r + '" style="width:48px"' + thClass + '>' + (r + 1) + thExtra + '</th>';

  for (let c = 0; c < COLS; c++) {
    const addr = addrOf(r, c);

    // 🎯 Baca dari Chunks, fallback ke sheet().cells
    let cl = null;
    if (window.Chunks) {
      cl = Chunks.getCell(activeSheetIdx, addr);
    }
    if (cl === null && sheet().cells && sheet().cells[addr]) {
      cl = sheet().cells[addr];
    }

    let text = '';
    let style = '';
    let cls = '';

    if (cl) {
      let v;
      if (cl.f) {
        if (fastMode) {
          const cached = _hfValueCache.get(addr);
          v = cached !== undefined ? cached : '…';
        } else {
          try { v = hfGet(r, c); }
          catch { v = '#REF!'; }
          _hfValueCache.set(addr, v);
        }
      } else {
        v = cl.v ?? '';
      }

      if (_showFormulas && cl.f) {
        text = cl.f;
      } else {
        text = cl.s?.numFmt ? formatNumber(v, cl.s.numFmt) : displayVal(v);
      }
      // 🎯 Handle line break dalam HTML
      if (text && typeof text === 'string' && text.includes('\n')) {
        text = text.replace(/\n/g, '\n');   // biarkan \n, CSS white-space: pre-wrap yang render
      }

      if (cl.s) {
        const s = cl.s;
        if (s.bold) style += 'font-weight:700;';
        if (s.italic) style += 'font-style:italic;';
        if (s.underline) style += 'text-decoration:underline;';
        if (s.color) style += 'color:' + s.color + ';';
        if (s.bg) style += 'background:' + s.bg + ';';
        if (s.align) style += 'text-align:' + s.align + ';';
        if (s.fontFamily) style += 'font-family:' + s.fontFamily + ';';
        if (s.fontSize) style += 'font-size:' + s.fontSize + 'px;';
      }

      if (cl.link) cls += ' cell-link';
      if (cl.comment) cls += ' has-comment';
      // 🎯 Lock indicator
      const shRow = sheet();
      if (shRow.protected) {
        if (cl?.s?.locked === false) cls += ' cell-unlocked';
        else cls += ' cell-locked';
      }
      if (_highlightFormulas && cl.f) cls += ' has-formula';
      if (isErr(v) || (typeof v === 'string' && v.startsWith('#'))) cls += ' has-error';
    }

    // 🎯 Kalau baris header filter, tambahkan tombol filter & class
    let filterBtn = '';
    let headerCls = '';
    if (isHeaderRow) {
      headerCls = ' filter-header-cell';
      const isActive = filters[c]?.values?.length > 0;
      if (isActive) cls += ' filtered-cell';
      filterBtn = `<button class="fx-filter-btn-inline${isActive ? ' active' : ''}" data-filter-col="${c}" title="Filter ${colName(c)}">▼</button>`;
    }

    const width = colWidths[c] || 96;
    const esc = escapeHtml(text);

    h += '<td data-addr="' + addr + '" data-col="' + c + '" class="' + (cls + headerCls).trim() + '" style="width:' + width + 'px;' + style + '">'
       + '<div class="cell-inner' + cls + '">' + esc + '</div>'
       + filterBtn
       + '</td>';
  }

  h += '</tr>';
  return h;
}

/* =========================================================
   FILTER BUTTON CLICK — event delegation
   ========================================================= */
document.getElementById('grid')?.addEventListener('mousedown', e => {
  const btn = e.target.closest('.fx-filter-btn-inline');
  if (!btn) return;

  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();

  const colIdx = +btn.dataset.filterCol;
  const td = btn.closest('td');

  console.log(`🔽 Filter click: col=${colIdx} (${colName(colIdx)})`);
  openFilterPopupAtCell(colIdx, td);
}, true);   // capture phase — mencegah handler lain

/* =========================================================
   FAST SELECTION — pakai classList, tidak render ulang
   ========================================================= */
function applySelectionStylesFast() {
  // Clear selection lama
  const tbody = document.getElementById('grid-body');
  if (!tbody) return;

  const trs = tbody.children;
  for (let i = 0; i < trs.length; i++) {
    const tr = trs[i];
    const tds = tr.children;
    for (let j = 1; j < tds.length; j++) {
      const td = tds[j];
      if (td.classList.contains('active')) td.classList.remove('active');
      if (td.classList.contains('selected')) td.classList.remove('selected');
    }
  }

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  const selR0 = Math.min(p1.r, p2.r), selR1 = Math.max(p1.r, p2.r);
  const selC0 = Math.min(p1.c, p2.c), selC1 = Math.max(p1.c, p2.c);
  const activeP = parseAddr(activeCell);

  const rowMap = getRowMap();
  for (let di = visibleStart; di < visibleEnd; di++) {
    const r = rowMap ? rowMap[di] : di;
    if (r < selR0 || r > selR1) continue;

    const tr = rowEls[r];
    if (!tr) continue;

    const tds = tr.children;
    for (let j = 1; j < tds.length; j++) {
      const td = tds[j];
      const p = parseAddr(td.dataset.addr);
      if (!p) continue;
      if (p.c < selC0 || p.c > selC1) continue;

      if (activeP && p.r === activeP.r && p.c === activeP.c) {
        td.classList.add('active');
      } else {
        td.classList.add('selected');
      }
    }
  }
}

/* =========================================================
   SCROLL MODE — pakai raw value, tidak ada hfGet
   ========================================================= */
// function renderVisibleCellsScroll() {
//   const rowMap = getRowMap();

//   for (let di = visibleStart; di < visibleEnd; di++) {
//     const r = rowMap ? rowMap[di] : di;
//     const tr = rowEls[r];
//     if (!tr) continue;

//     const children = tr.children;
//     for (let i = 1; i < children.length; i++) {
//       const td = children[i];
//       if (!td.dataset || !td.dataset.addr) continue;

//       const p = parseAddr(td.dataset.addr);
//       if (!p) continue;

//       // Skip kalau sedang di-edit
//       const first = td.firstElementChild;
//       if (first && first.tagName === 'INPUT') continue;

//       const addr = td.dataset.addr;
//       const cl = sheet().cells[addr];

//       // 🎯 RAW value — NO HF CALL
//       let raw = '';
//       if (cl) {
//         // Kalau ada cache HF, pakai cache
//         const cached = _hfValueCache.get(addr);
//         if (cached !== undefined) {
//           raw = cached;
//         } else if (cl.f) {
//           // Ada formula tapi belum ada cache → tampilkan placeholder sementara
//           raw = '…';
//         } else {
//           raw = cl.v ?? '';
//         }
//       }

//       const displayText = cl?.s?.numFmt
//         ? formatNumber(raw, cl.s.numFmt)
//         : displayVal(raw);

//       let inner = first;
//       if (!inner || !inner.classList || !inner.classList.contains('cell-inner')) {
//         inner = td.querySelector('.cell-inner');
//         if (!inner) {
//           inner = document.createElement('div');
//           inner.className = 'cell-inner';
//           td.insertBefore(inner, td.firstChild);
//         }
//       }

//       if (inner.textContent !== displayText) {
//         inner.textContent = displayText;
//       }
//     }
//   }
// }

/* =========================================================
   FULL MODE — hitung HF, update cache
   ========================================================= */
// function renderVisibleCellsFull() {
//   const rowMap = getRowMap();

//   for (let di = visibleStart; di < visibleEnd; di++) {
//     const r = rowMap ? rowMap[di] : di;
//     const tr = rowEls[r];
//     if (!tr) continue;

//     const children = tr.children;
//     for (let i = 1; i < children.length; i++) {
//       const td = children[i];
//       if (!td.dataset || !td.dataset.addr) continue;

//       const p = parseAddr(td.dataset.addr);
//       if (!p) continue;

//       const first = td.firstElementChild;
//       if (first && first.tagName === 'INPUT') continue;

//       const addr = td.dataset.addr;
//       const cl = sheet().cells[addr];

//       let raw = '';
//       if (cl) {
//         if (cl.f) {
//           // 🎯 Panggil HF hanya untuk formula
//           try {
//             raw = hfGet(p.r, p.c);
//           } catch {
//             raw = '#REF!';
//           }
//           // Simpan ke cache
//           _hfValueCache.set(addr, raw);
//         } else {
//           raw = cl.v ?? '';
//         }
//       }

//       const displayText = cl?.s?.numFmt
//         ? formatNumber(raw, cl.s.numFmt)
//         : displayVal(raw);

//       let inner = first;
//       if (!inner || !inner.classList || !inner.classList.contains('cell-inner')) {
//         inner = td.querySelector('.cell-inner');
//         if (!inner) {
//           inner = document.createElement('div');
//           inner.className = 'cell-inner';
//           td.insertBefore(inner, td.firstChild);
//         }
//       }

//       if (inner.textContent !== displayText) {
//         inner.textContent = displayText;
//       }
//     }
//   }
// }

/* =========================================================
   BUILD ROW — dengan colSpan & rowSpan native + FILTER HEADER
   ========================================================= */
function buildRowWithMerges(r, visibleMerges) {
  const tr = document.createElement('tr');
  const sheetRowH = sheet().rowHeights?.[r] || sheet().rowHeight || ROW_H;
  tr.style.height = sheetRowH + 'px';
  tr.dataset.row = r;

  // 🎯 Cek apakah baris ini = header filter
  const headerRow = getFilterHeaderRow();
  const isHeaderRow = filterMode && (r === headerRow);
  const filters = sheet().filters || {};

  if (isHeaderRow) tr.classList.add('filter-header-row');

  /* ---------- ROW HEADER (th) ---------- */
  const th = document.createElement('th');
  th.textContent = r + 1;
  th.dataset.row = r;
  th.style.width = '48px';
  th.style.minWidth = '48px';
  th.style.maxWidth = '48px';
  th.style.boxSizing = 'border-box';
  if (isHeaderRow) th.classList.add('filter-header-row');

  const summaryGroup = getGroupSummaryAt(r);
  const insideGroup = getRowGroupAt(r);

  if (summaryGroup) {
    const btn = document.createElement('button');
    btn.className = 'row-group-btn' + (summaryGroup.collapsed ? '' : ' expanded');
    btn.textContent = summaryGroup.collapsed ? '+' : '−';
    btn.title = summaryGroup.collapsed ? 'Expand' : 'Collapse';
    btn.type = 'button';
    btn.dataset.groupFrom = summaryGroup.from;
    th.appendChild(btn);
  }

  if (insideGroup && !summaryGroup) {
    th.classList.add('row-group-line');
  }

  // Grip resize
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

  /* ---------- CELLS ---------- */
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
        // -------- Top-left merged cell --------
        const td = document.createElement('td');
        td.dataset.addr = addrOf(r, c);
        td.dataset.col = c;
        td.colSpan = found.c1 - found.c0 + 1;
        td.rowSpan = found.r1 - found.r0 + 1;
        td.classList.add('merged-tl');

        const inner = document.createElement('div');
        inner.className = 'cell-inner';
        td.appendChild(inner);
        tr.appendChild(td);
        c = found.c1 + 1;
      } else {
        // Covered — skip kolom ini
        c = found.c1 + 1;
      }
    } else {
      // -------- Cell normal --------
      const td = document.createElement('td');
      td.dataset.addr = addrOf(r, c);
      td.dataset.col = c;

      // 🎯 Kalau baris ini header filter → tambah tombol ▼
      if (isHeaderRow) {
        td.classList.add('filter-header-cell');
        const isActive = filters[c]?.values?.length > 0;
        if (isActive) td.classList.add('filtered-cell');

        const btn = document.createElement('button');
        btn.className = 'fx-filter-btn-inline' + (isActive ? ' active' : '');
        btn.dataset.filterCol = c;
        btn.title = 'Filter ' + colName(c);
        btn.type = 'button';
        btn.textContent = '▼';
        td.appendChild(btn);
      }

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

  // Frozen rows dulu
  const frozenRowCount = Math.min(sheet().freezeRow || 0, 20);
  for (let r = 0; r < frozenRowCount; r++) {
    const tr = rowEls[r];
    if (!tr) continue;
    const children = tr.children;
    for (let i = 1; i < children.length; i++) {
      const td = children[i];
      const addr = td.dataset?.addr;
      if (!addr) continue;
      const p = parseAddr(addr);
      if (p) updateCellDisplay(td, p.r, p.c);
    }
  }

  // Visible rows
  for (let di = visibleStart; di < visibleEnd; di++) {
    const r = rowMap ? rowMap[di] : di;
    const tr = rowEls[r];
    if (!tr) continue;

    const children = tr.children;
    for (let i = 1; i < children.length; i++) {
      const td = children[i];
      const addr = td.dataset?.addr;
      if (!addr) continue;
      const p = parseAddr(addr);
      if (p) updateCellDisplay(td, p.r, p.c);
    }
  }
}

// function renderVisibleCellsFast() {
//   const rowMap = getRowMap();

//   for (let di = visibleStart; di < visibleEnd; di++) {
//     const r = rowMap ? rowMap[di] : di;
//     const tr = rowEls[r];
//     if (!tr) continue;

//     // 🎯 Pakai children langsung (lebih cepat dari querySelectorAll)
//     const children = tr.children;
//     for (let i = 1; i < children.length; i++) {
//       const td = children[i];
//       if (!td.dataset || !td.dataset.addr) continue;
//       const p = parseAddr(td.dataset.addr);
//       if (p) updateCellDisplayFast(td, p.r, p.c);
//     }
//   }
// }

function updateCellDisplayFast(td, r, c) {
  if (!td) return;

  // Skip kalau sedang diedit
  const firstChild = td.firstElementChild;
  if (firstChild && firstChild.tagName === 'INPUT') return;

  const cacheKey = r + ':' + c;
  let displayText = _valueCache.get(cacheKey);

  if (displayText === undefined) {
    const addr = td.dataset.addr;
    const cl = sheet().cells[addr];
    let raw = '';

    if (cl && cl.f) {
      try { raw = hfGet(r, c); }
      catch { raw = '#REF!'; }
    } else {
      raw = cl?.v ?? '';
    }

    displayText = cl?.s?.numFmt
      ? formatNumber(raw, cl.s.numFmt)
      : displayVal(raw);

    _valueCache.set(cacheKey, displayText);
  }

  // Inner element
  let inner = firstChild;
  if (!inner || !inner.classList.contains('cell-inner')) {
    inner = td.querySelector('.cell-inner');
    if (!inner) {
      inner = document.createElement('div');
      inner.className = 'cell-inner';
      td.insertBefore(inner, td.firstChild);
    }
  }

  // Update hanya kalau text beda
  if (inner.textContent !== displayText) {
    inner.textContent = displayText;
  }
}

function updateCellDisplayFast(td, r, c) {
  if (!td) return;
  if (td.querySelector('input')) return;

  const addr = td.dataset.addr;
  const cl = sheet().cells[addr];

  let raw = '';
  // 🎯 FAST: pakai value langsung kalau bukan formula
  if (cl && cl.f) {
    try { raw = hfGet(r, c); }
    catch { raw = '#REF!'; }
  } else {
    raw = cl?.v ?? '';
  }

  let inner = td.querySelector('.cell-inner');
  if (!inner) {
    inner = document.createElement('div');
    inner.className = 'cell-inner';
    td.insertBefore(inner, td.firstChild);
  }

  const displayText = cl?.s?.numFmt
    ? formatNumber(raw, cl.s.numFmt)
    : displayVal(raw);

  // 🎯 OPTIMASI: skip update kalau text sama
  if (inner.textContent !== displayText) {
    inner.textContent = displayText;
  }
}


function forceRenderCell(td, r, c) {
  if (!td) return;
  const addr = td.dataset.addr;
  const cl = sheet().cells[addr];

  let raw = '';

  if (cl && cl.f) {
    // 🎯 Guard: skip HF kalau di luar range
    if (!hf || !hf.__maxR || r >= hf.__maxR || c >= hf.__maxC) {
      raw = cl.v ?? '';
    } else {
      try {
        raw = hfGet(r, c);
        if (raw === '' || raw === null || raw === undefined) raw = '#REF!';
      } catch {
        raw = '#REF!';
      }
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

  let newText;
  if (_showFormulas && cl?.f) {
    // 🎯 Tampilkan formula mentah, bukan hasil
    newText = cl.f;
  } else {
    newText = cl?.s?.numFmt
      ? formatNumber(raw, cl.s.numFmt)
      : displayVal(raw);
  }

  // 🎯 Hanya update kalau text beda
  if (inner.textContent !== newText) {
    inner.textContent = newText;
  }

  const isNum = typeof raw === 'number' || (!isNaN(Number(raw)) && raw !== '');
  const wantNumCell = isNum && newText.length > 8;
  if (td.classList.contains('num-cell') !== wantNumCell) {
    td.classList.toggle('num-cell', wantNumCell);
  }
  td.classList.toggle('cell-link', !!cl?.link);
  td.classList.toggle('error', !!isErr(raw));

  // 🎯 Style: apply hanya kalau berubah (via styleKey)
  const styleKey = JSON.stringify(cl?.s || {});
  if (td.dataset.styleKey !== styleKey) {
    td.dataset.styleKey = styleKey;
    applyFormat(td, cl?.s || {});
  }
  // 🎯 Highlight formula
  if (_highlightFormulas && cl?.f) {
    td.classList.add('has-formula');
  } else {
    td.classList.remove('has-formula');
  }

  // 🎯 Highlight error
  const isError = isErr(raw) || (typeof raw === 'string' && raw.startsWith('#'));
  if (isError) {
    td.classList.add('has-error');
  } else {
    td.classList.remove('has-error');
  }
  const text = String(newText || '');
  const hasNewline = text.includes('\n');
  td.classList.toggle('multiline', hasNewline);
  // 🎯 Lock indicator
  const sh = sheet();
  if (sh.protected) {
    if (isCellUnlocked(addr)) {
      td.classList.add('cell-unlocked');
      td.classList.remove('cell-locked');
    } else {
      td.classList.add('cell-locked');
      td.classList.remove('cell-unlocked');
    }
  } else {
    td.classList.remove('cell-locked', 'cell-unlocked');
  }
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
  renderImages();
  renderAllPivots();
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
  const advFilters = getAdvancedFilters();
  const collapsedSet = getCollapsedRowSet();

  const hasBasic = filters.length > 0;
  const hasAdv = advFilters.length > 0;
  const hasCollapse = collapsedSet.size > 0;

  // Fast path
  if (!hasBasic && !hasAdv && !hasCollapse) return null;
  if (cachedRowMap) return cachedRowMap;

  const headerRow = getFilterHeaderRow();
  const sh = sheet();

  let maxR = sh._cachedMaxR;
  if (maxR === undefined) {
    maxR = 100;
    for (const a of Object.keys(sh.cells || {})) {
      const p = parseAddr(a);
      if (p && p.r > maxR) maxR = p.r;
    }
    sh._cachedMaxR = maxR;
  }
  maxR += 20;

  const out = [];

  // Header filter row
  if (hasBasic) out.push(headerRow);

  for (let r = 0; r <= maxR; r++) {
    // Skip collapsed
    if (collapsedSet.has(r)) continue;

    // Skip header filter row (sudah di-push)
    if (hasBasic && r === headerRow) continue;
    if (hasBasic && r < headerRow) continue;

    // Basic filter
    if (hasBasic) {
      let pass = true;
      for (const [col, f] of filters) {
        const c = parseInt(col, 10);
        const cl = sh.cells[addrOf(r, c)];
        let v = '';
        if (cl) {
          if (cl.f) {
            try { v = String(hfGet(r, c) ?? ''); } catch { v = ''; }
          } else {
            v = String(cl.v ?? '');
          }
        }
        if (!f.values.includes(v)) { pass = false; break; }
      }
      if (!pass) continue;
    }

    // 🎯 Advanced filter
    if (hasAdv && r > headerRow) {
      if (!checkAdvancedFilterMatch(r, advFilters)) continue;
    }
    // Skip header row kalau cuma advanced filter
    if (hasAdv && !hasBasic && r === headerRow) continue;
    if (hasAdv && !hasBasic && r < headerRow) continue;

    out.push(r);
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

/* =========================================================
   BORDER STYLE MAP
   ========================================================= */
const BORDER_STYLE_MAP = {
  thin:    '1px',
  medium:  '2px',
  thick:   '3px',
  dashed:  '1px',
  dotted:  '1px',
  double:  '3px'
};

const BORDER_TYPE_MAP = {
  thin:   'solid',
  medium: 'solid',
  thick:  'solid',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double'
};

/* Format 1 sisi border → CSS value */
function formatBorderValue(styleKey, color) {
  if (!styleKey) return '';
  const width = BORDER_STYLE_MAP[styleKey] || '1px';
  const type = BORDER_TYPE_MAP[styleKey] || 'solid';
  const c = color || '#111111';
  return `${width} ${type} ${c}`;
}

/* Terapkan border dinamis ke td */
function applyBorder(td, s) {
  const defaultBorder = '1px solid #e9eef4';
  const b = s.borders || {};
  const style = b.style || 'thin';
  const color = b.color || '#111111';

  // Reset ke default dulu
  td.style.borderTop = defaultBorder;
  td.style.borderBottom = defaultBorder;
  td.style.borderLeft = defaultBorder;
  td.style.borderRight = defaultBorder;

  const borderVal = formatBorderValue(style, color);

  // Preset: all, outer, inner, top, bottom, left, right, none
  const preset = b.preset || (b.all ? 'all' : null);

  if (preset === 'none') {
    td.style.borderTop = '1px solid #e9eef4';
    td.style.borderBottom = '1px solid #e9eef4';
    td.style.borderLeft = '1px solid #e9eef4';
    td.style.borderRight = '1px solid #e9eef4';
    return;
  }

  if (preset === 'all' || b.all) {
    td.style.borderTop = borderVal;
    td.style.borderBottom = borderVal;
    td.style.borderLeft = borderVal;
    td.style.borderRight = borderVal;
  } else if (preset === 'outer') {
    // Cek posisi di selection — kalau di border, apply
    // Sederhananya: apply ke semua, tapi tandai outer
    // (Detail: nanti ada flag khusus dari handler)
    td.style.borderTop = b.top ? borderVal : '1px solid #e9eef4';
    td.style.borderBottom = b.bottom ? borderVal : '1px solid #e9eef4';
    td.style.borderLeft = b.left ? borderVal : '1px solid #e9eef4';
    td.style.borderRight = b.right ? borderVal : '1px solid #e9eef4';
  } else {
    if (b.top) td.style.borderTop = borderVal;
    if (b.bottom) td.style.borderBottom = borderVal;
    if (b.left) td.style.borderLeft = borderVal;
    if (b.right) td.style.borderRight = borderVal;
  }
}

function applyFormat(td, s) {
  s = s || {};   // guard

  td.style.fontWeight   = s.bold      ? '700' : '';
  td.style.fontStyle    = s.italic    ? 'italic' : '';
  td.style.textDecoration =
    [s.underline ? 'underline' : '', s.strike ? 'line-through' : '']
      .filter(Boolean).join(' ') || '';
  td.style.fontFamily = s.fontFamily || '';
  td.style.fontSize   = s.fontSize ? s.fontSize + 'px' : '';
  td.style.color      = s.color || '';
  // 🎯 JANGAN set background kalau tidak ada bg di style
  if (s.bg !== undefined) {
    td.style.background = s.bg || '';
  }
  td.style.textAlign = s.align || '';
  td.classList.toggle('wrap', !!s.wrap);

  applyBorder(td, s);
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
  // Clear dulu
  document.querySelectorAll('#grid tbody td.active, #grid tbody td.selected')
    .forEach(td => td.classList.remove('active', 'selected'));

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  const selR0 = Math.min(p1.r, p2.r), selR1 = Math.max(p1.r, p2.r);
  const selC0 = Math.min(p1.c, p2.c), selC1 = Math.max(p1.c, p2.c);
  const activeP = parseAddr(activeCell);

  // Iterate SEMUA td di DOM (hanya ~30 baris yang ter-render karena virtual)
  document.querySelectorAll('#grid tbody td[data-addr]').forEach(td => {
    const p = parseAddr(td.dataset.addr);
    if (!p) return;

    // Overlap check (akomodasi merge)
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
   if (typeof showFillHandle === 'function') {
    setTimeout(showFillHandle, 0);
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

/* =========================================================
   CEK FORMULA LENGKAP / BELUM
   ========================================================= */
function isFormulaComplete(formula) {
  if (!formula || typeof formula !== 'string') return true;
  if (!formula.startsWith('=')) return true;

  const body = formula.slice(1).trim();

  // Hanya "=" kosong → belum lengkap
  if (!body) return false;

  // Berakhir dengan karakter yang butuh operand lagi
  if (/[+\-*/^&<>=,(]$/.test(body)) return false;

  // Berakhir dengan referensi belum lengkap (mis. "Sheet2!" atau "A1:")
  if (body.endsWith('!') || body.endsWith(':')) return false;

  // Cek parentheses balance
  let open = 0;
  for (const ch of body) {
    if (ch === '(') open++;
    else if (ch === ')') open--;
    if (open < 0) return false;   // kelebihan ")"
  }
  if (open !== 0) return false;   // masih ada "(" yang belum ditutup

  // Cek string quotes (tanda petik)
  const dq = (body.match(/"/g) || []).length;
  if (dq % 2 !== 0) return false;

  return true;
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
      const v = (r < hfR && c < hfC) ? hfGet(r, c) : '';
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
  // 🎯 Invalidate semua cache
  if (typeof invalidateCommentCache === 'function') invalidateCommentCache();
  if (typeof invalidateValidationCache === 'function') invalidateValidationCache();

  const sh = sheet();
  if (sh) sh._cachedMaxR = undefined;   // reset cache maxR

  const cur = sh?.cells?.[addr];
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
  if (isCellLocked(addr)) {
    toast('🔒 Cell terkunci — klik kanan → Unlock Cell untuk membuka');
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

  let _resizeRAF = null;
  function autoResizeEditor() {
    if (typeof input.style.height !== 'undefined') {
      // Auto-height
      input.style.height = 'auto';
      const h = Math.max(td.offsetHeight, input.scrollHeight);
      input.style.height = h + 'px';
    }

    const lines = String(input.value).split('\n');
    const maxLineLen = Math.max(...lines.map(l => l.length), 1);
    const tdW = td.offsetWidth || 96;
    const spillW = (maxLineLen + 3) * 7.5 + 14;
    const targetW = Math.max(tdW, Math.min(spillW, 600));
    input.style.width = targetW + 'px';
  }

    autoResizeEditor();

    updateMultilineClass(input, td);

    input.addEventListener('input', () => {
      autoResizeEditor();
      updateMultilineClass(input, td);
    });

    input.addEventListener('keydown', e => {
      if (insertingRefMode) return;

      // 🎯 ALT+ENTER = line break (bukan commit)
      if (e.key === 'Enter' && e.altKey) {
        e.preventDefault();
        const pos = input.selectionStart ?? input.value.length;
        const before = input.value.slice(0, pos);
        const after = input.value.slice(pos);
        input.value = before + '\n' + after;
        input.setSelectionRange(pos + 1, pos + 1);
        updateMultilineClass(input, td);
        autoResizeEditor();
        return;
      }

      // 🎯 CTRL+ENTER = line break (alternatif Mac)
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        const pos = input.selectionStart ?? input.value.length;
        const before = input.value.slice(0, pos);
        const after = input.value.slice(pos);
        input.value = before + '\n' + after;
        input.setSelectionRange(pos + 1, pos + 1);
        updateMultilineClass(input, td);
        autoResizeEditor();
        return;
      }

      // ENTER biasa = commit & pindah bawah
      if (e.key === 'Enter') {
        e.preventDefault();
        const val = input.value;

        if (val.startsWith('=') && !isFormulaComplete(val)) {
          toast('⚠ Formula belum lengkap — tutup tanda kurung dulu');
          return;
        }

        commit(addr, val);
        closeEditor();
        if (!editor) moveSel(1, 0);
      }
      else if (e.key === 'Escape') {
        e.preventDefault();
        closeEditor();
      }
      else if (e.key === 'Tab') {
        e.preventDefault();
        const val = input.value;

        if (val.startsWith('=') && !isFormulaComplete(val)) {
          toast('⚠ Formula belum lengkap');
          return;
        }

        commit(addr, val);
        closeEditor();
        if (!editor) moveSel(0, e.shiftKey ? -1 : 1);
      }
    });
  input.addEventListener('blur', () => {
    if (insertingRefMode) return;
    if (!editor || editor.input !== input) return;

    const val = input.value;

    // 🎯 Kalau ini formula belum lengkap → DISCARD, jangan commit
    if (val.startsWith('=') && !isFormulaComplete(val)) {
      console.log('⏭ Formula belum lengkap, discard:', val);
      closeEditor();
      return;
    }

    commit(addr, val);
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

function setHFCellSafe(r, c, val) {
  if (!hf) return;
  // hfSet sudah handle auto-add & rebuild
  hfSet(r, c, val, activeSheetIdx);
}

function commit(addr, value) {
  const p = parseAddr(addr);
  if (!p) return;
  if (currentAccessMode !== 'view' && isCellLocked(addr)) {
    toast('🔒 Cell terkunci — tidak bisa edit');
    return;
  }
  if (ensureRowCapacity(p.r)) {
    if (typeof renderVirtual === 'function') {
      window._lastRenderKey = null;
      setTimeout(() => renderVirtual(false), 50);
    }
  }
  pushUndo();
  const raw = String(value);

  // =====================================================
  // 1. Build cell
  // =====================================================
  let newCell = null;
  if (raw === '') {
    newCell = null;
  } else if (raw.startsWith('=')) {
    newCell = { f: raw };
  } else {
    const n = Number(raw);
    const v = raw.trim() !== '' && !isNaN(n) ? n : raw;
    newCell = { v };
  }

  // Preserve style/link/comment
  const old = sheet().cells[addr];
  if (old && (old.s || old.link || old.comment)) {
    newCell = newCell || {};
    if (old.s) newCell.s = old.s;
    if (old.link) newCell.link = old.link;
    if (old.comment) newCell.comment = old.comment;
  }

  // =====================================================
  // 2. Update local state
  // =====================================================
  if (!sheet().cells) sheet().cells = {};
  if (newCell) sheet().cells[addr] = newCell;
  else delete sheet().cells[addr];

  // Update chunk (kalau chunk mode aktif)
  if (window.Chunks && Chunks.setCell && typeof wbUsesChunks !== 'undefined' && wbUsesChunks) {
    Chunks.setCell(activeSheetIdx, addr, newCell);
  }

  // =====================================================
  // 3. Update HF
  // =====================================================
  if (newCell && newCell.f) {
    setHFCellSafe(p.r, p.c, newCell.f);
  } else if (newCell && newCell.v !== undefined) {
    setHFCellSafe(p.r, p.c, newCell.v);
  } else {
    setHFCellSafe(p.r, p.c, null);
  }

  markCellDirty(addr);

  // =====================================================
  // 4. Invalidate cache & render key
  // =====================================================
  if (typeof clearHFCache === 'function') clearHFCache();
  if (typeof invalidateHF === 'function') invalidateHF(addr);
  window._lastRenderKey = null;
  window._lastFrozenKey = null;

  if (raw.startsWith('=') && /:\s*[A-Z]+\d{4,}/.test(raw)) warnBigRange();

  // =====================================================
  // 5. RE-RENDER — 1x sync, + 1x delay kalau formula
  // =====================================================
  const isFormula = newCell && newCell.f;

  // 5a. Force render cell target langsung (safe remove input)
  const tr = rowEls[p.r];
  if (tr) {
    const td = tr.children[p.c + 1];
    if (td) {
      // 🎯 Safe remove: cek parentNode dulu
      td.querySelectorAll('input').forEach(inp => {
        try { if (inp.parentNode) inp.parentNode.removeChild(inp); } catch {}
      });
      if (typeof forceRenderCell === 'function') {
        forceRenderCell(td, p.r, p.c);
      }
    }
  }

  // 5b. Full re-render — sync
  if (typeof renderVirtual === 'function') renderVirtual(false);

  // 5c. Hanya kalau formula — re-render lagi setelah HF evaluate
  if (isFormula) {
    setTimeout(() => {
      if (typeof clearHFCache === 'function') clearHFCache();
      window._lastRenderKey = null;
      window._lastFrozenKey = null;
      if (typeof renderVirtual === 'function') renderVirtual(false);
    }, 150);
  }

  scheduleSave();

  // Broadcast collab
  if (window.Collab && window.Collab.isConnected()) {
    const cur = sheet().cells[addr];
    window.Collab.sendEdit({ [addr]: cur || null }, activeSheetIdx);
  }
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
  // 🎯 Simpan header row LAMA (sebelum selection berubah)
  const oldHeaderRow = filterMode ? getFilterHeaderRow() : null;

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
  e_shiftDown = e.shiftKey;
  if (editor) return;
  const tag = (e.target.tagName || '').toUpperCase();
  const inputFocused = tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
  const ctrl = e.ctrlKey || e.metaKey;

  if (e.target && e.target.id === 'formula-input') return;

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
  if (ctrl && e.key.toLowerCase() === 'n') { e.preventDefault(); newWB(); return; }
  if (ctrl && e.key.toLowerCase() === 'p') { e.preventDefault(); openPrintModal?.(); return; }
  if (ctrl && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); return; }
  if (ctrl && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
  if (ctrl && e.key.toLowerCase() === 'b') { e.preventDefault(); toggleFmt('bold'); return; }
  if (ctrl && e.key.toLowerCase() === 'i') { e.preventDefault(); toggleFmt('italic'); return; }
  if (ctrl && e.key.toLowerCase() === 'u') { e.preventDefault(); toggleFmt('underline'); return; }
  if (ctrl && e.key.toLowerCase() === 'c') { if (!inputFocused) { e.preventDefault(); doCopy(); } return; }
  if (ctrl && e.key.toLowerCase() === 'x') { if (!inputFocused) { e.preventDefault(); doCut(); } return; }
  if (ctrl && e.key.toLowerCase() === 'v') { if (!inputFocused) { e.preventDefault(); doPaste(); } return; }
  if (ctrl && e.shiftKey && e.key.toLowerCase() === 'v') {
    e.preventDefault();
    doPasteSpecial('value');
    return;
  }

  if (inputFocused) return;

  const p = parseAddr(activeCell);
  if (!p) return;
  if (e.key === 'ArrowUp') {
  e.preventDefault();
  if (ctrl) jumpToEdge(p.r, p.c, -1, 0);
    else if (e.shiftKey) selectCell(addrOf(Math.max(0, p.r - 1), p.c), true);
    else moveSel(-1, 0);
  }
  else if (e.key === 'ArrowDown') {
    e.preventDefault();
    if (ctrl) jumpToEdge(p.r, p.c, 1, 0);
    else if (e.shiftKey) selectCell(addrOf(Math.min(ROWS - 1, p.r + 1), p.c), true);
    else moveSel(1, 0);
  }
  else if (e.key === 'ArrowLeft') {
    e.preventDefault();
    if (ctrl) jumpToEdge(p.r, p.c, 0, -1);
    else if (e.shiftKey) selectCell(addrOf(p.r, Math.max(0, p.c - 1)), true);
    else moveSel(0, -1);
  }
  else if (e.key === 'ArrowRight') {
    e.preventDefault();
    if (ctrl) jumpToEdge(p.r, p.c, 0, 1);
    else if (e.shiftKey) selectCell(addrOf(p.r, Math.min(COLS - 1, p.c + 1)), true);
    else moveSel(0, 1);
  }
  else if (e.key === 'Enter') { e.preventDefault(); openEditor(activeCell); }
  else if (e.key === 'F2') { e.preventDefault(); openEditor(activeCell); }
  else if (e.key === 'Tab') { e.preventDefault(); moveSel(0, e.shiftKey ? -1 : 1); }
  else if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    pushUndo();

    if (selectedRows.size || selectedCols.size) {
      deleteSelectedRowsAndCols();
      return;
    }

    const parts = expandRange(selection.start, selection.end);
    const sh = sheet();
    let deletedCount = 0;
    let skippedCount = 0;

    for (const a of parts) {
      // 🎯 Cek lock
      if (isCellLocked(a)) {
        skippedCount++;
        continue;
      }
      delete sheet().cells[a];
      const pp = parseAddr(a);
      if (pp) hfSet(pp.r, pp.c, null);
      markCellDirty(a);
      deletedCount++;
    }

    if (skippedCount > 0) {
      toast(`⚠ ${skippedCount} cell terkunci (skip)`);
    }

    window._lastRenderKey = null;
    window._lastFrozenKey = null;
    renderVirtual(false);
  } else if (e.key.length === 1 && !ctrl && !e.altKey) {
    e.preventDefault(); openEditor(activeCell, e.key);
  }
});

/* =========================================================
   CTRL+ARROW — jump ke data edge seperti Excel
   ========================================================= */
function jumpToEdge(r, c, dr, dc) {
  const hasData = (row, col) => {
    if (row < 0 || row >= ROWS || col < 0 || col >= COLS) return false;
    const addr = addrOf(row, col);
    const cell = sheet().cells[addr];
    if (!cell) return false;
    return (cell.v !== undefined && cell.v !== null && cell.v !== '') || !!cell.f;
  };

  let curR = r, curC = c;
  const startHasData = hasData(r, c);

  // Kalau sel sekarang kosong → cari data berikutnya
  // Kalau sel sekarang ada data → cari data terakhir di blok yang sama
  while (true) {
    const nextR = curR + dr;
    const nextC = curC + dc;
    if (nextR < 0 || nextR >= ROWS || nextC < 0 || nextC >= COLS) break;

    const nextHasData = hasData(nextR, nextC);

    if (!startHasData) {
      // Cari sampai ketemu data
      if (nextHasData) { curR = nextR; curC = nextC; break; }
      curR = nextR; curC = nextC;
    } else {
      // Cari sampai data habis (edge terakhir)
      if (!nextHasData) break;
      curR = nextR; curC = nextC;
    }
  }

  const targetAddr = addrOf(curR, curC);
  if (e_shiftDown) {
    selection.end = targetAddr;
    activeCell = selection.start;
  } else {
    selection = { start: targetAddr, end: targetAddr };
    activeCell = targetAddr;
  }
  scrollAddrIntoView(targetAddr);
  updateSelectionUI();
}

let e_shiftDown = false;  // simpan state shift

/* =========================================================
   FORMULA BAR — robust handler (capture phase)
   ========================================================= */
(function setupFormulaBarRobust() {
  const fi = document.getElementById('formula-input');
  if (!fi) return;

  // Pastikan editable
  fi.removeAttribute('readonly');
  fi.disabled = false;
  fi.readOnly = false;

  // 🎯 Handle Enter di CAPTURE phase — jalankan sebelum handler lain
  fi.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation();

      const val = e.target.value;
      const targetCell = activeCell;

      console.log('📝 Formula bar Enter:', JSON.stringify(val), '→', targetCell);

      // Commit ke cell aktif
      commit(targetCell, val);

      // 🎯 Force re-render langsung
      if (typeof renderVirtual === 'function') {
        window._lastRenderKey = null;
        renderVirtual(false);
      }

      // 🎯 Re-render setelah HF evaluate (delay untuk cross-sheet)
      setTimeout(() => {
        if (typeof clearHFCache === 'function') clearHFCache();
        window._lastRenderKey = null;
        if (typeof renderVirtual === 'function') renderVirtual(false);
      }, 100);

      // Pindah ke cell bawah (seperti Excel)
      if (typeof moveSel === 'function') {
        moveSel(1, 0);
      }

      // Sync formula bar dengan cell baru
      setTimeout(() => {
        if (fi) fi.value = getRaw(activeCell);
      }, 20);
    }
    else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      e.target.value = getRaw(activeCell);
      e.target.blur();
    }
    else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      const val = e.target.value;

      commit(activeCell, val);
      if (typeof renderVirtual === 'function') {
        window._lastRenderKey = null;
        renderVirtual(false);
      }

      // Pindah kanan
      if (typeof moveSel === 'function') moveSel(0, e.shiftKey ? -1 : 1);

      setTimeout(() => {
        if (fi) fi.value = getRaw(activeCell);
      }, 20);
    }
  }, true);   // ← capture phase

  // Blur → commit
  fi.addEventListener('blur', e => {
    const val = e.target.value;
    if (val !== getRaw(activeCell)) {
      commit(activeCell, val);
      if (typeof renderVirtual === 'function') {
        window._lastRenderKey = null;
        renderVirtual(false);
      }
    }
  });

  // Focus → select all
  fi.addEventListener('focus', e => {
    setTimeout(() => e.target.select(), 0);
  });

  console.log('✅ Formula bar handler installed');
})();

/* =========================================================
   CLIPBOARD
   ========================================================= */
function doCopy() {
  const parts = expandRange(selection.start, selection.end);
  clipboard = parts.map(a => ({
    addr: a,
    cell: sheet().cells[a] ? JSON.parse(JSON.stringify(sheet().cells[a])) : null
  }));

  // 🎯 Build TSV untuk Excel/Notepad
  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  const tsvRows = [];
  const htmlRows = [];

  for (let r = r0; r <= r1; r++) {
    const tsvCols = [];
    const htmlCols = [];

    for (let c = c0; c <= c1; c++) {
      const addr = addrOf(r, c);
      const cl = sheet().cells[addr];

      // 🎯 Nilai — pakai HF-evaluated kalau formula
      let val = '';
      if (cl) {
        if (cl.f && hf) {
          try { val = String(hfGet(r, c) ?? ''); } catch { val = ''; }
        } else {
          val = String(cl.v ?? '');
        }
      }

      // Escape TSV: buang tab & newline
      tsvCols.push(val.replace(/\t/g, ' ').replace(/\n/g, ' '));

      // HTML: preserve formatting
      htmlCols.push({
        text: val,
        style: cl?.s || {},
        numFmt: cl?.s?.numFmt
      });
    }

    tsvRows.push(tsvCols.join('\t'));
    htmlRows.push(htmlCols);
  }

  const tsv = tsvRows.join('\r\n');

  // 🎯 HTML table untuk paste ke Excel dengan format
  const html = buildClipboardHTML(htmlRows);

  // 🎯 Copy ke clipboard dengan multiple MIME types
  copyRichClipboard(tsv, html).then(ok => {
    if (ok) {
      toast('📋 ' + parts.length + ' sel dicopy');
    } else {
      toast('📋 Copy fallback — coba lagi');
    }
  });

  console.log('📋 Copied TSV:', tsv.slice(0, 200));
}
/* =========================================================
   BUILD HTML TABLE untuk clipboard
   ========================================================= */
function buildClipboardHTML(htmlRows) {
  if (!htmlRows.length) return '';

  const styleForCell = (s) => {
    if (!s) return '';
    const parts = [];
    if (s.bold) parts.push('font-weight:bold');
    if (s.italic) parts.push('font-style:italic');
    if (s.underline) parts.push('text-decoration:underline');
    if (s.color) parts.push(`color:${s.color}`);
    if (s.bg) parts.push(`background-color:${s.bg}`);
    if (s.align) parts.push(`text-align:${s.align}`);
    if (s.fontFamily) parts.push(`font-family:${s.fontFamily}`);
    if (s.fontSize) parts.push(`font-size:${s.fontSize}px`);
    return parts.join(';');
  };

  const escapeHtmlCell = (s) => String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  const rows = htmlRows.map(row => {
    const cells = row.map(c => {
      const style = styleForCell(c.style);
      const styleAttr = style ? ` style="${style}"` : '';
      return `<td${styleAttr}>${escapeHtmlCell(c.text)}</td>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>
  table { border-collapse: collapse; }
  td { padding: 2px 6px; border: 1px solid #ddd; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
</style></head><body>
<table>${rows}</table>
</body></html>`;
}
/* =========================================================
   COPY RICH CLIPBOARD — TSV + HTML
   Bekerja di HTTP & HTTPS
   ========================================================= */
async function copyRichClipboard(tsv, html) {
  // 1. Coba Clipboard API modern (butuh HTTPS/localhost)
  if (navigator.clipboard && window.ClipboardItem) {
    try {
      const items = {
        'text/plain': new Blob([tsv], { type: 'text/plain' }),
      };
      // Excel butuh text/html untuk preserve format
      if (html) {
        items['text/html'] = new Blob([html], { type: 'text/html' });
      }
      const item = new ClipboardItem(items);
      await navigator.clipboard.write([item]);
      return true;
    } catch (e) {
      console.warn('Clipboard API failed, fallback...', e);
    }
  }

  // 2. Fallback: event listener copy + execCommand
  try {
    const handler = (e) => {
      e.clipboardData.setData('text/plain', tsv);
      if (html) e.clipboardData.setData('text/html', html);
      e.preventDefault();
      document.removeEventListener('copy', handler);
    };
    document.addEventListener('copy', handler);

    // Trigger copy command
    const ok = document.execCommand('copy');

    // Cleanup
    document.removeEventListener('copy', handler);

    if (ok) return true;
  } catch (e) {
    console.warn('execCommand copy failed:', e);
  }

  // 3. Fallback terakhir: textarea + execCommand
  try {
    const ta = document.createElement('textarea');
    ta.value = tsv;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.setAttribute('readonly', '');
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    if (ok) return true;
  } catch (e) {
    console.warn('Textarea copy failed:', e);
  }

  return false;
}
function doCut() {
  doCopy();      // sudah handle clipboard rich
  pushUndo();
  for (const { addr } of clipboard) {
    delete sheet().cells[addr];
    const p = parseAddr(addr);
    if (p) hfSet(p.r, p.c, null);
    markCellDirty(addr);
  }
  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  renderVirtual(false);
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
    if (tp) hfSet(tp.r, tp.c, c?.f || c?.v || null);
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
/* =========================================================
   BORDER DROPDOWN HANDLER
   ========================================================= */
const borderMenu = document.getElementById('border-menu');
const borderMenuBtn = document.getElementById('btn-border-menu');

// State border saat ini
let currentBorderStyle = 'thin';
let currentBorderColor = '#000000';

// Toggle menu
borderMenuBtn?.addEventListener('click', e => {
  e.stopPropagation();
  borderMenu.classList.toggle('hidden');
});

// Close on click outside
document.addEventListener('mousedown', e => {
  if (borderMenu && !borderMenu.classList.contains('hidden')) {
    if (!borderMenu.contains(e.target) && !borderMenuBtn.contains(e.target)) {
      borderMenu.classList.add('hidden');
    }
  }
});

// Style picker
document.querySelectorAll('#border-style-list button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#border-style-list button').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentBorderStyle = btn.dataset.style;
  });
});

// Color picker
document.querySelectorAll('#border-color-list button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#border-color-list button').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentBorderColor = btn.dataset.color;
  });
});

/* Apply preset border ke selection */
function applyBorderPreset(preset) {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  pushUndo();

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const addr = addrOf(r, c);
      ensureCell(addr);
      const s = sheet().cells[addr].s || (sheet().cells[addr].s = {});

      if (preset === 'none') {
        s.borders = {};
      } else if (preset === 'all') {
        s.borders = {
          preset: 'all',
          style: currentBorderStyle,
          color: currentBorderColor
        };
      } else if (preset === 'outer') {
        // Apply outer border — cek apakah cell di pinggir selection
        s.borders = s.borders || {};
        s.borders.preset = 'outer';
        s.borders.style = currentBorderStyle;
        s.borders.color = currentBorderColor;
        s.borders.top = (r === r0);
        s.borders.bottom = (r === r1);
        s.borders.left = (c === c0);
        s.borders.right = (c === c1);
      } else if (preset === 'inner') {
        // Apply inner border — cek apakah cell di tengah
        s.borders = s.borders || {};
        s.borders.preset = 'inner';
        s.borders.style = currentBorderStyle;
        s.borders.color = currentBorderColor;
        // Inner = semua sisi KECUALI border luar selection
        s.borders.top = (r > r0) ? true : s.borders.top || false;
        s.borders.bottom = (r < r1) ? true : s.borders.bottom || false;
        s.borders.left = (c > c0) ? true : s.borders.left || false;
        s.borders.right = (c < c1) ? true : s.borders.right || false;
      } else {
        // top, bottom, left, right
        s.borders = s.borders || {};
        s.borders.style = currentBorderStyle;
        s.borders.color = currentBorderColor;
        s.borders.preset = null;
        s.borders[preset] = true;
      }

      markCellDirty(addr);
    }
  }

  renderVisibleCells();
  borderMenu.classList.add('hidden');
  toast(`✔ Border ${preset} diterapkan`);
}

// Bind preset buttons
document.querySelectorAll('[data-border]').forEach(btn => {
  btn.addEventListener('click', () => {
    applyBorderPreset(btn.dataset.border);
  });
});
document.getElementById('btn-clear-all').onclick = () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  pushUndo();
  for (const a of expandRange(selection.start, selection.end)) {
    delete sheet().cells[a];
    const p = parseAddr(a);
    if (p) hfSet(p.r, p.c, null);
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
  return JSON.stringify({ sheets: wbData.sheets, activeSheet: activeSheetIdx,wbName: wbName });
}
let _undoSkipWarned = false;
function pushUndo() {
  const snap = snapshot();
  if (snap.length > 20_000_000) {
    if (!_undoSkipWarned) {
      console.warn('⚠ Snapshot >20MB — undo dilewati untuk workbook besar');
      _undoSkipWarned = true;
    }
    return;
  }
  undoStack.push(snap);
  while (undoStack.length > 20) undoStack.shift();
  redoStack = [];
}
function undo() {
  if (!undoStack.length) return toast('Tidak ada undo');
  redoStack.push(snapshot());
  const prev = JSON.parse(undoStack.pop());
  wbData.sheets = prev.sheets;

  clearHFCache();

  // 🎯 RESET SEMUA CACHE RENDER
  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  window._lastFrozenRow = null;
  window._lastFrozenCol = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  activeSheetIdx = Math.min(prev.activeSheet, wbData.sheets.length - 1);

  // 🎯 Restore nama workbook (kalau ada di snapshot)
  if (prev.wbName !== undefined) {
    wbName = prev.wbName;
    const el = document.getElementById('wb-name');
    if (el) el.value = wbName;
  }

  // 🎯 Load filter state dari sheet yang di-restore
  if (typeof loadFilterStateFromSheet === 'function') {
    try { loadFilterStateFromSheet(); } catch {}
  }

  resetHF();
  renderSheetTabs();
  renderSheet();

  // 🎯 Paksa render ulang total
  window._lastFrozenKey = null;
  renderVirtual(false);

  markDirty();
  toast('↶ undo');
}

function redo() {
  if (!redoStack.length) return toast('Tidak ada redo');
  undoStack.push(snapshot());
  const next = JSON.parse(redoStack.pop());
  wbData.sheets = next.sheets;

  clearHFCache();

  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  window._lastFrozenRow = null;
  window._lastFrozenCol = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  activeSheetIdx = Math.min(next.activeSheet, wbData.sheets.length - 1);

  if (next.wbName !== undefined) {
    wbName = next.wbName;
    const el = document.getElementById('wb-name');
    if (el) el.value = wbName;
  }

  if (typeof loadFilterStateFromSheet === 'function') {
    try { loadFilterStateFromSheet(); } catch {}
  }

  resetHF();
  renderSheetTabs();
  renderSheet();

  window._lastFrozenKey = null;
  renderVirtual(false);

  markDirty();
  toast('↷ redo');
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
    // Skip hidden sheet (kecuali active — biar user tidak bingung)
    if (s.hidden && i !== activeSheetIdx) return;

    const btn = document.createElement('button');
    btn.className = 'sheet-tab' + (i === activeSheetIdx ? ' active' : '');
    btn.textContent = s.name;
    btn.dataset.idx = i;
    btn.draggable = true;

    // Tab color
    if (s.tabColor) btn.dataset.color = s.tabColor;

    // Prevent blur saat editor formula aktif
    btn.addEventListener('mousedown', e => {
      if (editor && editor.input && editor.input.value.startsWith('=')) {
        e.preventDefault();
      }
    }, true);

    // Click → switch
    btn.onclick = () => switchSheet(i);

    // Double-click → rename
    btn.ondblclick = async () => {
      const n = await customPrompt('Nama sheet:', s.name, { title: 'Rename Sheet' });
      if (n && n.trim() && n.trim() !== s.name) {
        const oldName = s.name;
        const newName = n.trim();

        // 🎯 Update named ranges yang refer ke sheet ini
        const ranges = wbData.namedRanges || {};
        const escapedOld = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(`^('?)${escapedOld}\\1!`);

        for (const key in ranges) {
          const ref = ranges[key];
          if (re.test(ref)) {
            const needsQuote = /[\s\-.]/.test(newName);
            const newSheetRef = needsQuote
              ? "'" + newName.replace(/'/g, "''") + "'!"
              : newName + '!';
            ranges[key] = ref.replace(re, newSheetRef);
            console.log(`🔧 Updated named range "${key}": ${ref} → ${ranges[key]}`);
          }
        }

        s.name = newName;
        markMetaDirty('structure');
        try { resetHF(); } catch (e) { console.error(e); }
        renderSheetTabs();
        renderSheet();
        scheduleSave();
        toast(`✔ Sheet di-rename: "${oldName}" → "${newName}"`);
      }
    };

    // Right-click → context menu
    btn.oncontextmenu = e => {
      e.preventDefault();
      e.stopPropagation();
      openTabContextMenu(e.clientX, e.clientY, i);
    };

    // Drag & drop reorder
    btn.addEventListener('dragstart', e => {
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', String(i));
      btn.classList.add('dragging');
      window._dragTabIdx = i;
    });

    btn.addEventListener('dragend', () => {
      btn.classList.remove('dragging');
      document.querySelectorAll('.sheet-tab').forEach(t => {
        t.classList.remove('drag-over-left', 'drag-over-right');
      });
      window._dragTabIdx = null;
    });

    btn.addEventListener('dragover', e => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      const rect = btn.getBoundingClientRect();
      const isLeft = e.clientX < rect.left + rect.width / 2;
      btn.classList.toggle('drag-over-left', isLeft);
      btn.classList.toggle('drag-over-right', !isLeft);
    });

    btn.addEventListener('dragleave', () => {
      btn.classList.remove('drag-over-left', 'drag-over-right');
    });

    btn.addEventListener('drop', e => {
      e.preventDefault();
      btn.classList.remove('drag-over-left', 'drag-over-right');
      const fromIdx = parseInt(e.dataTransfer.getData('text/plain'), 10);
      if (isNaN(fromIdx) || fromIdx === i) return;

      const rect = btn.getBoundingClientRect();
      const isLeft = e.clientX < rect.left + rect.width / 2;
      const toIdx = isLeft ? i : i + 1;

      moveSheet(fromIdx, toIdx);
    });

    cont.appendChild(btn);
  });
}

async function switchSheet(i) {
  invalidateCommentCache();
  invalidateValidationCache();
  if (i === activeSheetIdx) return;

  // 🎯 Simpan filter state sheet LAMA
  saveFilterStateToSheet();

  clearHFCache();
  window._lastRenderKey = null;

  // ==========================================
  // Handle editor (cross-sheet mode)
  // ==========================================
  if (editor) {
    const input = editor.input;
    const val = input.value;

    if (val.startsWith('=')) {
      console.log('🎯 Cross-sheet mode aktif. Formula:', val);

      crossSheetRef = {
        originalSheetIdx: activeSheetIdx,
        addr: editor.addr,
        formula: val
      };

      editor = null;
      try { input.remove(); } catch {}

      document.querySelectorAll('#grid td.active, #grid td.selected')
        .forEach(td => td.classList.remove('active', 'selected'));

      if (dirty) { try { await saveWB(); } catch {} }

      await lazyLoadSheet(i);

      activeSheetIdx = i;
      wbData.activeSheet = i;
      dirtyMeta.activeSheet = true;

      loadFilterStateFromSheet();

      await rebuildHFWithLoaded();

      // 🎯 RENDER FINAL — semua async sudah selesai
      forceRefreshGrid();
      renderSheetTabs();

      activeCell = 'A1';
      selection = { start: 'A1', end: 'A1' };
      updateSelectionUI();

      showCrossSheetBanner();
      return;
    } else {
      try { closeEditor(); } catch {}
    }
  }

  // ==========================================
  // Normal switch
  // ==========================================
  if (dirty) { try { await saveWB(); } catch {} }

  destroyChartsForSheet(activeSheetIdx);

  await lazyLoadSheet(i);

  activeSheetIdx = i;
  wbData.activeSheet = i;
  dirtyMeta.activeSheet = true;

  loadFilterStateFromSheet();

  await rebuildHFWithLoaded();
  await recalcRowsLimit(true);

  // 🎯 RENDER FINAL — setelah semua async selesai
  renderSheetTabs();
  forceRefreshGrid();

  activeCell = 'A1';
  selection = { start: 'A1', end: 'A1' };
  updateSelectionUI();
}

/* =========================================================
   FORCE REFRESH GRID — reset semua cache & render ulang
   ========================================================= */
function forceRefreshGrid() {
  // Reset scroll ke atas
  const gridWrapEl = document.getElementById('grid-wrap');
  if (gridWrapEl) gridWrapEl.scrollTop = 0;

  // Reset SEMUA cache
  window._lastFrozenKey = null;
  window._lastFrozenRow = null;
  window._lastFrozenCol = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  if (typeof invalidateRowMap === 'function') invalidateRowMap();
  if (typeof clearHFCache === 'function') clearHFCache();
  if (typeof invalidateCommentCache === 'function') invalidateCommentCache();
  if (typeof invalidateValidationCache === 'function') invalidateValidationCache();

  // Render full
  renderSheet();

  // Reset cache LAGI (renderSheet internal bisa set ulang)
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  // Final render
  renderVirtual(false);
}
/* =========================================================
   LAZY LOAD SHEET — fetch sheet dari server kalau belum loaded
   ========================================================= */
async function lazyLoadSheet(i) {
  if (!lazyLoadMode) return false;
  if (loadedSheetIdxs.has(i)) {
    console.log(`⚡ Sheet ${i} sudah loaded (skip fetch)`);
    return false;
  }

  const sheetName = wbData.sheets[i]?.name || `Sheet${i + 1}`;
  showLoading(`Memuat sheet "${sheetName}"…`);

  try {
    const t0 = performance.now();
    const res = await fetch(`/api/sheets/${wbId}/sheet/${i}`, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);

    const sheetData = await res.json();
    wbData.sheets[i] = {
      ...wbData.sheets[i],
      ...sheetData,
      cells: sheetData.cells || {}
    };
    loadedSheetIdxs.add(i);

    const ms = (performance.now() - t0).toFixed(0);
    const count = Object.keys(sheetData.cells || {}).length;
    console.log(`✅ Lazy loaded sheet ${i} "${sheetName}": ${count} cells in ${ms}ms`);

    hideLoading();
    return true;
  } catch (e) {
    console.warn('Lazy load sheet failed:', e);
    hideLoading();
    toast('⚠ Gagal memuat sheet ' + sheetName);
    return false;
  }
}

/* =========================================================
   REBUILD HF — hanya dengan sheet yang sudah di-load
   ========================================================= */
async function rebuildHFWithLoaded() {
  if (hf) { try { hf.destroy(); } catch {} hf = null; }
  if (typeof _sheetIdCache !== 'undefined') _sheetIdCache.clear();

  const sheetsData = {};
  let loadedCount = 0;

  wbData.sheets.forEach((s, idx) => {
    if (!loadedSheetIdxs.has(idx)) return;
    loadedCount++;

    let name = s.name || `Sheet${idx + 1}`;
    let unique = name;
    let n = 1;
    while (sheetsData[unique] !== undefined) {
      n++;
      unique = name + '_' + n;
    }
    sheetsData[unique] = buildSheetMatrix(s);
  });

  console.log(`🔧 Building HF with ${loadedCount}/${wbData.sheets.length} loaded sheets`);

  if (loadedCount === 0) {
    hf = HyperFormulaClass.buildFromArray([['']], { licenseKey: HF_LICENSE });
    hf.__maxR = 50;
    hf.__maxC = 26;
    return;
  }

  try {
    hf = HyperFormulaClass.buildFromSheets(sheetsData, { licenseKey: HF_LICENSE });
    hf.__maxR = ROWS;
    hf.__maxC = COLS;
    console.log('✅ HF rebuilt. Sheets:', hf.getSheetNames());
    attachHFListener();
  } catch (e) {
    console.error('HF multi-sheet build failed:', e);
    // Fallback single sheet
    const cur = sheet();
    const matrix = buildSheetMatrix(cur);
    const safeMatrix = (matrix && matrix.length && matrix[0] && matrix[0].length)
      ? matrix : [['']];
    hf = HyperFormulaClass.buildFromArray(safeMatrix, { licenseKey: HF_LICENSE });
    hf.__maxR = safeMatrix.length || 50;
    hf.__maxC = (safeMatrix[0] || []).length || 26;
    attachHFListener();
  }
}
document.getElementById('btn-add-sheet').onclick = async () => {
  const newIdx = wbData.sheets.length;

  // 🎯 Tambah dengan field lengkap
  wbData.sheets.push({
    name: 'Sheet' + (newIdx + 1),
    cells: {},
    colWidths: {},
    rowHeight: 26,
    cf: [],
    charts: [],
    merged: [],
    filters: {},
    images: [],
    validations: [],
    hidden: false,
    tabColor: null,
    rowGroups: [],
    protected: false,
    protectHash: null,
    freezeRow: 0,
    freezeCol: 0
  });

  // 🎯 Mark sheet baru as loaded
  loadedSheetIdxs.add(newIdx);

  // 🎯 Mark dirty SEBELUM switch (biar save tahu ada sheet baru)
  markMetaDirty('structure');

  // 🎯 Switch ke sheet baru
  await switchSheet(newIdx);

  // 🎯 Force save SEKARANG, jangan tunggu debounce
  try { await saveWB(); } catch (e) { console.warn('Save after add sheet failed:', e); }

  toast(`✔ Sheet "${wbData.sheets[newIdx].name}" dibuat`);
};

/* =========================================================
   RIBBON TABS
   ========================================================= */
document.querySelectorAll('#ribbon-tabs button').forEach(b => {
  if (!b.dataset.tab) return;

  b.onclick = () => {
    document.querySelectorAll('#ribbon-tabs button').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.ribbon-pane').forEach(x => x.classList.remove('active'));
    b.classList.add('active');

    const pane = document.querySelector(`.ribbon-pane[data-pane="${b.dataset.tab}"]`);
    if (pane) pane.classList.add('active');
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
    const v = hfGet(r, col);
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
  const rule = { range, type };
  const bg = document.getElementById('cf-bg').value;
  const color = document.getElementById('cf-color').value;

  if (type === 'colorScale') {
    rule.colorMin = document.getElementById('cf-color-min').value;
    rule.colorMid = document.getElementById('cf-color-mid').value;
    rule.colorMax = document.getElementById('cf-color-max').value;
  } else if (type === 'dataBar') {
    rule.barColor = document.getElementById('cf-bar-color').value;
  } else if (type === 'iconSet') {
    rule.iconType = document.getElementById('cf-icon-type').value;
  } else if (type === 'top' || type === 'bottom') {
    rule.value = document.getElementById('cf-value').value || '10';
    rule.style = { bg, color };
  } else if (type === 'aboveAverage' || type === 'belowAverage' || type === 'duplicate' || type === 'unique') {
    rule.style = { bg, color };
  } else {
    rule.value = document.getElementById('cf-value').value;
    rule.value2 = document.getElementById('cf-value2').value;
    rule.style = { bg, color };
  }

  sheet().cf ||= [];
  sheet().cf.push(rule);
  document.getElementById('modal-cf').classList.add('hidden');
  markDirty();
  renderCF();
};

/* Toggle UI berdasarkan tipe */
document.getElementById('cf-type').onchange = e => {
  const t = e.target.value;
  document.getElementById('cf-value2-row').style.display =
    (t === 'between') ? 'block' : 'none';
  document.getElementById('cf-colorscale-wrap').style.display =
    (t === 'colorScale') ? 'block' : 'none';
  document.getElementById('cf-databar-wrap').style.display =
    (t === 'dataBar') ? 'block' : 'none';
  document.getElementById('cf-iconset-wrap').style.display =
    (t === 'iconSet') ? 'block' : 'none';

  // Sembunyikan field value/bg/color untuk tipe visual
  const hideValue = ['colorScale', 'dataBar', 'iconSet'].includes(t);
  document.getElementById('cf-value').parentElement.style.display = hideValue ? 'none' : 'block';
  document.querySelector('.modal-row .modal-label:nth-child(1)').style.display = hideValue ? 'none' : 'block';
  document.querySelector('.modal-row .modal-label:nth-child(2)').style.display = hideValue ? 'none' : 'block';
};
document.getElementById('btn-cf-clear').onclick = () => {
  sheet().cf = []; markDirty(); renderCF(); toast('CF cleared');
};

function renderCF() { renderCFVisible(); }
function renderCFVisible() {
  // 🎯 Hanya reset CF override, JANGAN sentuh background biasa
  document.querySelectorAll('#grid td.cf, #grid td.cf-databar, #grid td.cf-iconset')
    .forEach(td => {
      td.classList.remove('cf', 'cf-databar', 'cf-iconset');
      td.style.removeProperty('background-image');
      td.querySelectorAll('.cf-icon-el').forEach(el => el.remove());
    });

  // 🎯 Re-apply format dasar dari style user (pakai data-addr, bukan index)
  document.querySelectorAll('#grid td[data-addr]').forEach(td => {
    const cl = sheet().cells[td.dataset.addr];
    applyFormat(td, cl?.s || {});
  });

  // Apply CF rules
  for (const rule of (sheet().cf || [])) {
    const [s, e] = rangeParts(rule.range);
    const p1 = parseAddr(s), p2 = parseAddr(e);
    if (!p1 || !p2) continue;
    const rMin = Math.max(visibleStart, Math.min(p1.r, p2.r));
    const rMax = Math.min(visibleEnd - 1, Math.max(p1.r, p2.r));
    const cMin = Math.min(p1.c, p2.c), cMax = Math.max(p1.c, p2.c);

    if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') {
      applyVisualCF(rule, rMin, rMax, cMin, cMax);
    } else if (rule.type === 'top' || rule.type === 'bottom' ||
               rule.type === 'aboveAverage' || rule.type === 'belowAverage' ||
               rule.type === 'duplicate' || rule.type === 'unique') {
      applyStatisticalCF(rule, p1, p2);
    } else {
      for (let r = rMin; r <= rMax; r++) {
        const tr = rowEls[r];
        if (!tr) continue;
        for (let c = cMin; c <= cMax; c++) {
          const td = tr.querySelector(`td[data-addr="${addrOf(r, c)}"]`);
          if (!td) continue;
          const v = hfGet(r, c);
          const n = Number(v);
          let match = false;
          switch (rule.type) {
            case 'greaterThan': match = !isNaN(n) && n > Number(rule.value); break;
            case 'lessThan':    match = !isNaN(n) && n < Number(rule.value); break;
            case 'equal':       match = String(v) === String(rule.value); break;
            case 'between':     match = !isNaN(n) && n >= Number(rule.value) && n <= Number(rule.value2); break;
            case 'contains':    match = String(v).toLowerCase().includes(String(rule.value).toLowerCase()); break;
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
}

/* Color Scale / Data Bar / Icon Set */
function applyVisualCF(rule, rMin, rMax, cMin, cMax) {
  // Kumpulkan nilai range
  const values = [];
  for (let r = rMin; r <= rMax; r++) {
    for (let c = cMin; c <= cMax; c++) {
      const v = Number(hfGet(r,c));
      if (!isNaN(v)) values.push(v);
    }
  }
  if (!values.length) return;

  const minV = Math.min(...values);
  const maxV = Math.max(...values);
  const midV = (minV + maxV) / 2;
  const range = maxV - minV || 1;

  for (let r = rMin; r <= rMax; r++) {
    const tr = rowEls[r];
    if (!tr) continue;
    for (let c = cMin; c <= cMax; c++) {
      const td = tr.children[c + 1];
      if (!td) continue;
      const v = Number(hfGet(r,c));
      if (isNaN(v)) continue;
      const t = (v - minV) / range; // 0..1

      if (rule.type === 'colorScale') {
        td.style.background = interpolateColor3(
          rule.colorMin || '#dc2626',
          rule.colorMid || '#fbbf24',
          rule.colorMax || '#10b981',
          t
        );
      } else if (rule.type === 'dataBar') {
        const barColor = rule.barColor || '#3b82f6';
        const pct = Math.max(2, t * 100);
        td.style.backgroundImage = `linear-gradient(to right, ${barColor} ${pct}%, transparent ${pct}%)`;
        td.classList.add('cf-databar');
      } else if (rule.type === 'iconSet') {
        td.classList.add('cf-iconset');
        let icon = '';
        const iconType = rule.iconType || 'arrows';
        if (iconType === 'arrows') {
          if (t > 0.66) icon = '▲';
          else if (t > 0.33) icon = '►';
          else icon = '▼';
        } else if (iconType === 'traffic') {
          if (t > 0.66) icon = '🟢';
          else if (t > 0.33) icon = '🟡';
          else icon = '🔴';
        } else if (iconType === 'stars') {
          if (t > 0.66) icon = '★★★';
          else if (t > 0.33) icon = '★★';
          else icon = '★';
        } else if (iconType === 'flags') {
          if (t > 0.66) icon = '🟢';
          else if (t > 0.33) icon = '🟡';
          else icon = '🔴';
        }

        const iconEl = document.createElement('span');
        iconEl.className = 'cf-icon-el';
        iconEl.textContent = icon;
        iconEl.style.cssText = `
          position: absolute; right: 3px; top: 50%; transform: translateY(-50%);
          font-size: 11px; pointer-events: none; color: #334155;
        `;
        td.appendChild(iconEl);
      }
    }
  }
}

function interpolateColor3(c1, c2, c3, t) {
  // t: 0..1
  if (t <= 0.5) {
    return interpolateColor2(c1, c2, t * 2);
  } else {
    return interpolateColor2(c2, c3, (t - 0.5) * 2);
  }
}
function interpolateColor2(a, b, t) {
  const ah = a.replace('#', '');
  const bh = b.replace('#', '');
  const ar = parseInt(ah.substr(0, 2), 16);
  const ag = parseInt(ah.substr(2, 2), 16);
  const ab = parseInt(ah.substr(4, 2), 16);
  const br = parseInt(bh.substr(0, 2), 16);
  const bg = parseInt(bh.substr(2, 2), 16);
  const bb = parseInt(bh.substr(4, 2), 16);
  const r = Math.round(ar + (br - ar) * t);
  const g = Math.round(ag + (bg - ag) * t);
  const bl = Math.round(ab + (bb - ab) * t);
  return `rgb(${r},${g},${bl})`;
}

function applyStatisticalCF(rule, p1, p2) {
  const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  // Kumpulkan nilai (single column atau seluruh range)
  const items = [];
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const v = hfGet(r,c);
      const n = Number(v);
      if (!isNaN(n)) items.push({ r, c, v: n });
    }
  }
  if (!items.length) return;

  // Tentukan highlight set
  const highlight = new Set();
  const N = rule.value ? parseInt(rule.value, 10) : 10;

  if (rule.type === 'top') {
    const sorted = [...items].sort((a, b) => b.v - a.v).slice(0, N);
    sorted.forEach(x => highlight.add(`${x.r}_${x.c}`));
  } else if (rule.type === 'bottom') {
    const sorted = [...items].sort((a, b) => a.v - b.v).slice(0, N);
    sorted.forEach(x => highlight.add(`${x.r}_${x.c}`));
  } else if (rule.type === 'aboveAverage') {
    const avg = items.reduce((s, x) => s + x.v, 0) / items.length;
    items.forEach(x => { if (x.v > avg) highlight.add(`${x.r}_${x.c}`); });
  } else if (rule.type === 'belowAverage') {
    const avg = items.reduce((s, x) => s + x.v, 0) / items.length;
    items.forEach(x => { if (x.v < avg) highlight.add(`${x.r}_${x.c}`); });
  } else if (rule.type === 'duplicate' || rule.type === 'unique') {
    const counter = new Map();
    items.forEach(x => counter.set(x.v, (counter.get(x.v) || 0) + 1));
    items.forEach(x => {
      const isDup = counter.get(x.v) > 1;
      if ((rule.type === 'duplicate' && isDup) || (rule.type === 'unique' && !isDup)) {
        highlight.add(`${x.r}_${x.c}`);
      }
    });
  }

  // Apply styles
  for (let r = Math.max(r0, visibleStart); r <= Math.min(r1, visibleEnd - 1); r++) {
    const tr = rowEls[r];
    if (!tr) continue;
    for (let c = c0; c <= c1; c++) {
      const td = tr.children[c + 1];
      if (!td) continue;
      if (highlight.has(`${r}_${c}`)) {
        td.classList.add('cf');
        td.style.background = rule.style?.bg || '#fef3c7';
        td.style.color = rule.style?.color || '#92400e';
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
    labels.push(String(hfGet(p.r,p.c) ?? ''));
  }
  for (let ci = 1; ci < colsArr.length; ci++) {
    const data = [];
    for (const r of rowsArr) {
      const p = parseAddr(addrOf(r, colsArr[ci]));
      const v = Number(hfGet(p.r,p.c));
      data.push(isNaN(v) ? 0 : v);
    }
    datasets.push({ label: 'Series ' + ci, data, backgroundColor: palette(ci - 1) });
  }
  if (!datasets.length) {
    const data = [];
    for (const r of rowsArr) {
      const p = parseAddr(addrOf(r, colsArr[0]));
      const v = Number(hfGet(p.r,p.c));
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
    const labels = rows.map(r => String(hfGet(r, cols[0]) ?? ''));
    const datasets = [];
    for (let ci = 1; ci < cols.length; ci++) {
      const data = rows.map(r => {
        const v = Number(hfGet(r, cols[ci]));
        return isNaN(v) ? 0 : v;
      });
      datasets.push({ label: 'Series ' + ci, data, backgroundColor: palette(ci - 1) });
    }
    if (!datasets.length) {
      const data = rows.map(r => {
        const v = Number(hfGet(r, cols[0]));
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
   SAVE WORKBOOK — dengan batching untuk hindari OOM
   + OFFLINE GUARD (stop retry saat server mati)
   ========================================================= */
async function saveWB(forceFull = false) {
  if (!wbId) return;
  if (saving) { saveAgain = true; return; }

  // 🎯 Skip total kalau offline
  if (window._isOnline === false) {
    const el = document.getElementById('save-status');
    if (el) el.textContent = '⚠ offline — perubahan tersimpan lokal';
    return;
  }

  const auth = { 'Authorization': 'Bearer ' + token };
  const statusEl = document.getElementById('save-status');

  // =====================================================
  // MODE CHUNK
  // =====================================================
  if (wbUsesChunks && window.Chunks) {
    const pending = Chunks.getPendingCount();
    const hasMeta = dirtyMeta.name || dirtyMeta.activeSheet || dirtyMeta.structure;

    if (!pending && !hasMeta && !forceFull) return;

    saving = true;
    if (statusEl) statusEl.textContent = '⟳ menyimpan…';

    try {
      if (pending > 0) {
        const r = await Chunks.flushChanges();
        console.log('💾 Chunk saved:', r.saved, 'blocks');
      }

      if (hasMeta) {
        await fetch(`/api/sheets/${wbId}/delta`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json', ...auth },
          body: JSON.stringify({
            sheetIdx: activeSheetIdx,
            cells: {},
            meta: {
              name: dirtyMeta.name ? wbName : undefined,
              activeSheet: dirtyMeta.activeSheet ? activeSheetIdx : undefined
            }
          })
        });
      }

      for (const k of Object.keys(dirtyMeta)) delete dirtyMeta[k];
      dirty = false;
      saveRetryCount = 0;

      if (statusEl) statusEl.textContent = '✔ tersimpan';
      setTimeout(() => {
        if (statusEl && statusEl.textContent === '✔ tersimpan') statusEl.textContent = '';
      }, 1800);
    } catch (e) {
      console.error('Chunk save failed:', e);
      if (statusEl) statusEl.textContent = '⚠ gagal save';
      dirty = true;
    } finally {
      saving = false;
      if (saveAgain) { saveAgain = false; saveWB(); }
    }
    return;
  }

  // =====================================================
  // MODE NORMAL — dengan BATCHING
  // =====================================================
  const cellsToSave = { ...dirtyCells };
  const metaToSave = { ...dirtyMeta };

  const hasCells = Object.keys(cellsToSave).length > 0;
  const hasMeta = metaToSave.name || metaToSave.activeSheet || metaToSave.structure;
  if (!hasCells && !hasMeta && !forceFull) return;

  saving = true;
  if (statusEl) statusEl.textContent = '⟳ menyimpan…';

  const getErrorText = async (res) => {
    try {
      const j = await res.json();
      return j.error || ('HTTP ' + res.status);
    } catch {
      return 'HTTP ' + res.status;
    }
  };

  try {
    const useFull = forceFull || metaToSave.structure;

    // =====================================================
    // FULL SAVE
    // =====================================================
    if (useFull) {
      // 🎯 Cek semua sheet loaded dulu
      const allLoaded = wbData.sheets.every((_, i) => loadedSheetIdxs.has(i));
      
      if (!allLoaded) {
        console.warn(`⚠️ Full save DIBATALKAN — ${loadedSheetIdxs.size}/${wbData.sheets.length} sheet loaded`);
        // Force delta save saja
        // metaToSave.structure = false;
        // Lanjut ke blok delta di bawah (kode tidak diubah)
      } else {
        // Aman, lanjut full save
        const res = await fetch(`/api/sheets/${wbId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', ...auth },
          body: JSON.stringify({ name: wbName, data: wbData })
        });

        if (!res.ok) throw new Error(await getErrorText(res));

        for (const k of Object.keys(cellsToSave)) delete dirtyCells[k];
        for (const k of Object.keys(metaToSave)) delete dirtyMeta[k];
        dirty = false;

        if (statusEl) statusEl.textContent = '✔ tersimpan';
        setTimeout(() => {
          if (statusEl && statusEl.textContent === '✔ tersimpan') statusEl.textContent = '';
        }, 1800);
        saveRetryCount = 0;
        return;
      }
    }

    // =====================================================
    // DELTA SAVE — batching per 500 cell
    // =====================================================
    const addrs = Object.keys(cellsToSave);
    const BATCH_SIZE = 500;
    const batches = [];

    for (let i = 0; i < addrs.length; i += BATCH_SIZE) {
      batches.push(addrs.slice(i, i + BATCH_SIZE));
    }

    console.log(`📤 Save: ${addrs.length} cells, ${batches.length} batches`);

    if (batches.length === 0) {
      const metaBody = {
        name: metaToSave.name ? wbName : undefined,
        activeSheet: metaToSave.activeSheet ? activeSheetIdx : undefined,
        namedRanges: wbData.namedRanges || {}
      };

      // 🎯 Kirim sheetMeta / sheetOrder kalau structure berubah
      if (metaToSave.structure) {
        metaBody.sheetOrder = wbData.sheets.map((_, idx) => idx);

        const sheetMeta = {};
        wbData.sheets.forEach((s, idx) => {
          sheetMeta[idx] = {
            hidden: !!s.hidden,
            tabColor: s.tabColor || null,
            name: s.name,
            rowGroups: s.rowGroups || [],
            protected: !!s.protected,
            protectHash: s.protectHash || null,
            advancedFilters: s.advancedFilters || [],
            pivots: s.pivots || []
          };
        });
        metaBody.sheetMeta = sheetMeta;

        console.log('📤 Sending meta (no cells):', metaBody);
      }

      const res = await fetch(`/api/sheets/${wbId}/delta`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json', ...auth },
        body: JSON.stringify({
          sheetIdx: activeSheetIdx,
          cells: {},
          meta: metaBody
        })
      });
      if (!res.ok && res.status !== 404) throw new Error(await getErrorText(res));
    } else {
      for (let i = 0; i < batches.length; i++) {
        const batchAddrs = batches[i];
        const batchCells = {};
        for (const addr of batchAddrs) {
          batchCells[addr] = cellsToSave[addr];
        }

        const body = {
          sheetIdx: activeSheetIdx,
          cells: batchCells,
          meta: {}
        };

        if (i === 0) {
          body.meta = {
            name: metaToSave.name ? wbName : undefined,
            activeSheet: metaToSave.activeSheet ? activeSheetIdx : undefined,
            namedRanges: wbData.namedRanges || {}
          };
          if (metaToSave.structure) {
            // Reorder array sesuai urutan saat ini
            body.meta.sheetOrder = wbData.sheets.map((_, i) => i);

            // Info per-sheet (hidden, tabColor, name)
            const sheetMeta = {};
            wbData.sheets.forEach((s, i) => {
              sheetMeta[i] = {
                hidden: !!s.hidden,
                tabColor: s.tabColor || null,
                name: s.name,
                protected: !!s.protected,
                protectHash: s.protectHash || null,
                advancedFilters: s.advancedFilters || [],
                pivots: s.pivots || []
              };
            });
            body.meta.sheetMeta = sheetMeta;
          }
        }

        const res = await fetch(`/api/sheets/${wbId}/delta`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json', ...auth },
          body: JSON.stringify(body)
        });

        if (res.status === 404 && i === 0) {
          const putRes = await fetch(`/api/sheets/${wbId}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json', ...auth },
            body: JSON.stringify({ name: wbName, data: wbData })
          });
          if (!putRes.ok) throw new Error(await getErrorText(putRes));

          for (const k of Object.keys(cellsToSave)) delete dirtyCells[k];
          for (const k of Object.keys(metaToSave)) delete dirtyMeta[k];
          dirty = false;
          if (statusEl) statusEl.textContent = '✔ tersimpan';
          setTimeout(() => {
            if (statusEl && statusEl.textContent === '✔ tersimpan') statusEl.textContent = '';
          }, 1800);
          saveRetryCount = 0;
          return;
        }

        if (!res.ok) {
          throw new Error(`Batch ${i + 1}/${batches.length}: ${await getErrorText(res)}`);
        }

        if (statusEl && batches.length > 1) {
          statusEl.textContent = `⟳ menyimpan ${i + 1}/${batches.length}…`;
        }
      }
    }

    // =====================================================
    // Sukses
    // =====================================================
    for (const k of Object.keys(cellsToSave)) delete dirtyCells[k];
    for (const k of Object.keys(metaToSave)) delete dirtyMeta[k];

    dirty = Object.keys(dirtyCells).length > 0 || Object.keys(dirtyMeta).length > 0;

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
    console.error('❌ Save failed:', err.message);

    // 🎯 DETEKSI: apakah ini network error?
    const isNetworkError = /Failed to fetch|ERR_CONNECTION|NetworkError|fetch/i.test(err.message);

    if (isNetworkError) {
      // ❌ JANGAN retry — tunggu event 'online'
      if (statusEl) statusEl.textContent = '⚠ offline — perubahan tersimpan lokal';
      dirty = true;
      saveRetryCount = 0;   // reset biar tidak numpuk
      return;
    }

    // Error non-network (mis. 500, validasi) → retry dengan backoff
    if (statusEl) statusEl.textContent = '⚠ gagal save — retry…';
    dirty = true;
    saveRetryCount = (saveRetryCount || 0) + 1;

    const delay = Math.min(30000, 2000 * saveRetryCount);
    setTimeout(() => {
      if (dirty && wbId && window._isOnline !== false) saveWB();
    }, delay);
  } finally {
    saving = false;
    if (saveAgain) {
      saveAgain = false;
      saveWB();
    }
  }
}

/* =========================================================
   ONLINE / OFFLINE DETECTION
   ========================================================= */
window._isOnline = navigator.onLine;

window.addEventListener('online', () => {
  console.log('🟢 Online — resume save');
  window._isOnline = true;
  if (dirty && wbId) saveWB();
});

window.addEventListener('offline', () => {
  console.log('🔴 Offline — save dijeda');
  window._isOnline = false;
  const el = document.getElementById('save-status');
  if (el) el.textContent = '⚠ offline — perubahan tersimpan lokal';
});

/* Debug helper */
window.clearDirty = () => {
  dirtyCells = {};
  dirtyMeta = {};
  dirty = false;
  saveRetryCount = 0;
  const el = document.getElementById('save-status');
  if (el) el.textContent = '';
  toast('✅ Dirty cleared');
};

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

/* =========================================================
   GLOBAL: track loaded sheets
   ========================================================= */
let loadedSheetIdxs = new Set();
let lazyLoadMode = false;

async function openWB(id) {
  document.getElementById('home-screen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');

  if (editor) closeEditor();
  if (wbId && dirty) { try { await saveWB(); } catch {} }

  // Reset state
  if (hf) { try { hf.destroy(); } catch {} hf = null; }
  chartCache.forEach((map) => {
    map.forEach(c => { try { c.destroy(); } catch {} });
  });
  chartCache.clear();
  rowEls = {};
  visibleStart = -1;
  visibleEnd = -1;
  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  window._lastFrozenRow = null;
  window._lastFrozenCol = null;
  if (typeof clearHFCache === 'function') clearHFCache();
  if (typeof _sheetIdCache !== 'undefined') _sheetIdCache.clear();
  if (typeof invalidateCommentCache === 'function') invalidateCommentCache();
  if (typeof invalidateValidationCache === 'function') invalidateValidationCache();
  undoStack = [];
  redoStack = [];
  if (autosaveTimer) { clearTimeout(autosaveTimer); autosaveTimer = null; }

  // Reset lazy load
  loadedSheetIdxs = new Set();
  lazyLoadMode = true;
  wbUsesChunks = false;   // 🎯 SELALU FALSE — jangan paksa chunk mode

  showLoading('Memuat workbook…');

  // =====================================================
  // STEP 1: Fetch metadata
  // =====================================================
  let meta;
  let wbMetaName = null;
  try {
    updateLoading('Memuat metadata…');
    const res = await fetch(`/api/sheets/${id}/sheets-meta`, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    meta = await res.json();

    // 🎯 CEK PASSWORD
    _wbHasPassword = !!meta.hasPassword;

    if (meta.hasPassword) {
      // Cek apakah user sudah pernah unlock di session ini
      const sessionPwd = sessionStorage.getItem(`wb-pass-${id}`);
      let isUnlocked = false;

      if (sessionPwd) {
        // Verify session password
        try {
          const vr = await fetch(`/api/sheets/${id}/verify-password`, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': 'Bearer ' + token
            },
            body: JSON.stringify({ password: sessionPwd })
          });
          if (vr.ok) isUnlocked = true;
          else sessionStorage.removeItem(`wb-pass-${id}`);
        } catch {}
      }

      if (!isUnlocked) {
        hideLoading();

        // Tutup app dulu
        document.getElementById('app').classList.add('hidden');
        document.getElementById('home-screen').classList.remove('hidden');

        // Prompt password
        const ok = await promptUnlockPassword(id, meta.name);
        if (!ok) {
          return;   // user batal
        }

        // User berhasil unlock — lanjut buka
        document.getElementById('home-screen').classList.add('hidden');
        document.getElementById('app').classList.remove('hidden');
        showLoading('Memuat workbook…');
      }
    }

    wbMetaName = meta.name || meta.workbookName || null;

    if (!wbMetaName) {
      try {
        // 🎯 light=1 → cells dikosongkan, hemat bandwidth
        const wbRes = await fetch(`/api/sheets/${id}?light=1`, {
          headers: { 'Authorization': 'Bearer ' + token }
        });
        if (wbRes.ok) {
          const wb = await wbRes.json();
          wbMetaName = wb.name || wb.data?.name || null;
          console.log('📥 wbName dari light fetch:', wbMetaName);
        }
      } catch (e) { console.warn('Gagal fetch wb name:', e); }
    }
  } catch (e) {
    hideLoading();
    return toast('Gagal memuat metadata: ' + e.message);
  }

  let activeIdx = Math.max(0,
    Math.min(meta.activeSheet || 0, (meta.sheets || []).length - 1));

  // 🎯 Kalau sheet aktif ternyata hidden, cari sheet visible pertama
  if (meta.sheets?.[activeIdx]?.hidden) {
    const visibleIdx = meta.sheets.findIndex(s => !s.hidden);
    if (visibleIdx >= 0) activeIdx = visibleIdx;
  }

  // =====================================================
  // STEP 2: Fetch full sheet (SELALU — jangan skip)
  // =====================================================
  let activeSheetData = { cells: {}, name: 'Sheet' };
  try {
    updateLoading(`Memuat sheet "${meta.sheets[activeIdx]?.name || activeIdx}"…`);
    const res = await fetch(`/api/sheets/${id}/sheet/${activeIdx}`, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    activeSheetData = await res.json();
    console.log(`✅ Sheet ${activeIdx} loaded: ${Object.keys(activeSheetData.cells || {}).length} cells`);
  } catch (e) {
    hideLoading();
    return toast('Gagal memuat sheet: ' + e.message);
  }

  // =====================================================
  // STEP 3: Build wbData
  // =====================================================
  wbData = {
    activeSheet: activeIdx,
    namedRanges: meta.namedRanges || {},
    sheets: (meta.sheets || []).map((s, idx) => ({
      name: s.name || `Sheet${idx + 1}`,
      colWidths: s.colWidths || {},
      rowHeight: s.rowHeight || 26,
      rowHeights: s.rowHeights || {},
      cf: s.cf || [],
      charts: s.charts || [],
      merged: s.merged || [],
      filters: s.filters || {},
      images: s.images || [],
      validations: s.validations || [],
      freezeRow: s.freezeRow || 0,
      freezeCol: s.freezeCol || 0,
      protected: s.protected || false,
      protectHash: s.protectHash || null,
      advancedFilters: s.advancedFilters || [],
      hidden: s.hidden || false,
      tabColor: s.tabColor || null,
      rowGroups: s.rowGroups || [],
      cells: idx === activeIdx ? (activeSheetData.cells || {}) : {}
    }))
  };

  loadedSheetIdxs.add(activeIdx);

  // =====================================================
  // STEP 4: Set state
  // =====================================================
  wbId = id;
  wbName = wbMetaName || 'Untitled';
  const wbNameEl = document.getElementById('wb-name');
  if (wbNameEl) wbNameEl.value = wbName;

  activeSheetIdx = activeIdx;
  clearDelta();
  loadFilterFromLocal(id);

  // =====================================================
  // STEP 5: Build HF
  // =====================================================
  const cellCount = Object.keys(sheet().cells || {}).length;
  const HF_CELL_LIMIT = 500_000;

  if (cellCount > HF_CELL_LIMIT) {
    console.warn(`⚠️ ${cellCount.toLocaleString('id-ID')} cells > ${HF_CELL_LIMIT.toLocaleString('id-ID')} — SKIP HF`);
    if (hf) { try { hf.destroy(); } catch {} hf = null; }
    hf = HyperFormulaClass.buildFromArray([['']], { licenseKey: HF_LICENSE });
    hf.__maxR = 50; hf.__maxC = 26;
    toast('⚠️ Data besar: formula tidak dievaluasi.');
  } else {
    if (cellCount > 5000) {
      updateLoading(`Membangun formula engine (${cellCount.toLocaleString('id-ID')} sel)…`);
    }
    await rebuildHFWithLoaded();
  }

  await recalcRowsLimit(true);

  // =====================================================
  // STEP 6: Render
  // =====================================================
  renderSheetTabs();
  const gridWrapEl = document.getElementById('grid-wrap');
  if (gridWrapEl) gridWrapEl.scrollTop = 0;
  forceRefreshGrid();
  selectCell('A1');
  document.getElementById('modal-sheets')?.classList.add('hidden');
  hideLoading();

  // =====================================================
  // STEP 7: Access mode & Collab
  // =====================================================
  currentAccessMode = meta.access || 'owner';
  applyReadOnlyMode(currentAccessMode === 'view');
  if (currentAccessMode === 'view') toast('👁 Mode lihat saja');

  if (window.Collab) window.Collab.connect(wbId, token);
  if (typeof applySheetProtection === 'function') applySheetProtection();
}

function normalizeWB(d) {
  if (d && Array.isArray(d.sheets)) {
    d.sheets.forEach(s => {
      if (typeof s.rowHeight !== 'number') s.rowHeight = 26;
      if (!s.colWidths) s.colWidths = {};
      if (!s.merged) s.merged = [];
      if (!s.filters) s.filters = {};
      if (!Array.isArray(s.images)) s.images = [];
      // 🎯 Hidden rows/cols
      if (!Array.isArray(s.hiddenCols)) s.hiddenCols = [];
      if (!Array.isArray(s.hiddenRows)) s.hiddenRows = [];
      // Existing
      if (!Array.isArray(s.validations)) s.validations = [];
      if (typeof s.freezeRow !== 'number') s.freezeRow = 0;
      if (typeof s.freezeCol !== 'number') s.freezeCol = 0;
      if (typeof s.protected !== 'boolean') s.protected = false;
    });
    return d;
  }
  return {
    sheets: [{
      name: 'Sheet1', cells: d || {},
      colWidths: {}, rowHeight: 26,
      cf: [], charts: [], merged: [], filters: {},
      hiddenCols: [], hiddenRows: []
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
            const ok = customConfirm(
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
                const ok = customConfirm(
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
        if (!(await customConfirm(`Hapus workbook "${s.name}"?`, { type: 'danger', okText: 'Hapus' }))) return;
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

  // Validasi
  const MAX_SIZE = 50 * 1024 * 1024;
  if (f.size > MAX_SIZE) {
    return toast(`⚠ File terlalu besar (${(f.size / 1024 / 1024).toFixed(1)} MB). Maks 50 MB.`);
  }

  if (!wbId) {
    await newWB();
    return;
  }

  const ok = await customConfirm(
    `Import file "${f.name}"?\n\nSemua sheet di file ini akan menggantikan data di workbook saat ini.`,
    { type: 'warn', okText: 'Import', cancelText: 'Batal' }
  );
  if (!ok) return;

  const fd = new FormData();
  fd.append('file', f);

  showLoading(`Mengimpor "${f.name}" (${(f.size / 1024 / 1024).toFixed(1)} MB)…`);

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 5 * 60 * 1000);

    const res = await fetch(`/api/sheets/${wbId}/import`, {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token },
      body: fd,
      signal: controller.signal
    });

    clearTimeout(timeout);

    if (!res.ok) {
      let errMsg = 'Import gagal';
      try {
        const j = await res.json();
        errMsg = j.error || errMsg;
      } catch {}
      hideLoading();
      return toast('⚠ ' + errMsg);
    }

    const result = await res.json();

    updateLoading(`Memuat ${result.sheets} sheet…`);
    await yieldToUI();

    // 🎯 Reload workbook
    await openWB(wbId);

    hideLoading();

    // Toast dengan info lengkap
    const sheetNames = (result.sheetNames || []).slice(0, 3).join(', ');
    const more = result.sheetNames?.length > 3 ? ` +${result.sheetNames.length - 3} lainnya` : '';
    toast(
      `✔ Import ${result.cells.toLocaleString('id-ID')} sel ` +
      `dari ${result.sheets} sheet: ${sheetNames}${more}`
    );
  } catch (err) {
    hideLoading();
    if (err.name === 'AbortError') {
      toast('⚠ Import timeout — file terlalu besar');
    } else {
      console.error('Import error:', err);
      toast('⚠ ' + err.message);
    }
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

    // Ikon & label berdasarkan permission
    const permMeta = {
      view:  { icon: 'fa-eye',  label: 'Hanya lihat' },
      edit:  { icon: 'fa-pen',  label: 'Bisa edit' },
      admin: { icon: 'fa-gear', label: 'Admin' }
    };
    const meta = permMeta[s.permission] || permMeta.view;

    row.innerHTML = `
      <div class="body">
        <div class="name">
          <i class="fa-solid fa-user" style="color:#10b981;font-size:11px;margin-right:6px"></i>
          ${escapeHtml(s.username)}
          <span class="share-badge badge-${s.permission}">
            <i class="fa-solid ${meta.icon}" style="font-size:9px;margin-right:3px"></i>${escapeHtml(meta.label)}
          </span>
        </div>
        <div class="meta">
          <i class="fa-solid fa-share" style="font-size:9px;margin-right:4px"></i>
          Dibagikan oleh ${escapeHtml(s.shared_by_username || '-')}
          <span style="opacity:.5;margin:0 4px">•</span>
          <i class="fa-regular fa-calendar" style="font-size:9px;margin-right:3px"></i>
          ${new Date(s.created_at).toLocaleDateString('id-ID')}
        </div>
      </div>
      <div class="perm-actions">
        <button class="perm-btn ${s.permission === 'view' ? 'active' : ''}"
                data-perm="view" title="Hanya lihat">
          <i class="fa-solid fa-eye"></i>
        </button>
        <button class="perm-btn ${s.permission === 'edit' ? 'active' : ''}"
                data-perm="edit" title="Bisa edit">
          <i class="fa-solid fa-pen"></i>
        </button>
        <button class="perm-btn ${s.permission === 'admin' ? 'active' : ''}"
                data-perm="admin" title="Admin">
          <i class="fa-solid fa-gear"></i>
        </button>
        <button class="perm-btn danger" data-del="1" title="Hapus akses">
          <i class="fa-solid fa-trash"></i>
        </button>
      </div>
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
      if (!(await customConfirm(`Hapus akses ${s.username}?`, { type: 'danger', okText: 'Hapus' }))) return;
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
            if (!(await customConfirm('Kembalikan ke revisi ini? Perubahan saat ini akan disimpan sebagai revisi baru.', { type: 'warn', okText: 'Restore' }))) return;
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
  const label = await customPrompt('Label snapshot:', 'Snapshot ' + new Date().toLocaleString('id-ID'), { title: 'Snapshot Baru', okText: 'Simpan' });
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
  const isPublic = await customConfirm('Jadikan template ini publik?', { type: 'question', okText: 'Publik', cancelText: 'Pribadi' });
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
   PARSE CLIPBOARD — handle TSV/CSV
   ========================================================= */
function parseClipboardTable(text) {
  if (!text) return [];

  // Normalize line ending
  let t = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  // Buang trailing newline
  t = t.replace(/\n+$/, '');

  // Deteksi delimiter: tab (Excel) atau comma (CSV)
  const tabCount = (t.match(/\t/g) || []).length;
  const commaCount = (t.match(/,/g) || []).length;
  const delimiter = tabCount > 0 ? '\t' : (commaCount > 0 ? ',' : '\t');

  // Split rows
  const lines = t.split('\n');
  const rows = [];

  for (const line of lines) {
    const cells = [];
    let cur = '';
    let inQuote = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuote && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuote = !inQuote;
        }
      } else if (ch === delimiter && !inQuote) {
        cells.push(cur);
        cur = '';
      } else {
        cur += ch;
      }
    }
    cells.push(cur);
    rows.push(cells);
  }

  return rows;
}

/* =========================================================
   PASTE 2D ARRAY — tulis ke cells mulai dari anchor
   ========================================================= */
function paste2DArray(rows, anchorAddr) {
  if (!rows.length) return false;
  if (currentAccessMode === 'view') {
    toast('👁 Mode lihat saja');
    return false;
  }

  const anchor = parseAddr(anchorAddr);
  if (!anchor) return false;

  const h = rows.length;
  const w = Math.max(...rows.map(r => r.length));

  if (h === 0 || w === 0) return false;

  pushUndo();

  let written = 0;

  for (let r = 0; r < h; r++) {
    for (let c = 0; c < w; c++) {
      const targetR = anchor.r + r;
      const targetC = anchor.c + c;
      const targetAddr = addrOf(targetR, targetC);
      const tp = parseAddr(targetAddr);
      if (!tp) continue;

      // Ensure row capacity
      ensureRowCapacity(targetR);

      const raw = rows[r][c] ?? '';
      const strVal = String(raw).trim();

      // Preserve style lama
      const old = sheet().cells[targetAddr];
      let newCell = null;

      if (strVal === '') {
        newCell = null;
      } else if (strVal.startsWith('=')) {
        newCell = { f: strVal };
      } else {
        const n = Number(strVal);
        const v = strVal !== '' && !isNaN(n) ? n : strVal;
        newCell = { v };
      }

      // Preserve style/link/comment
      if (old && newCell) {
        if (old.s) newCell.s = old.s;
        if (old.link) newCell.link = old.link;
        if (old.comment) newCell.comment = old.comment;
      }

      if (!sheet().cells) sheet().cells = {};
      if (newCell) {
        sheet().cells[targetAddr] = newCell;
        written++;
      } else {
        delete sheet().cells[targetAddr];
      }

      // Update HF
      if (newCell && newCell.f) {
        setHFCellSafe(targetR, targetC, newCell.f);
      } else if (newCell && newCell.v !== undefined) {
        setHFCellSafe(targetR, targetC, newCell.v);
      } else {
        setHFCellSafe(targetR, targetC, null);
      }

      markCellDirty(targetAddr);
    }
  }

  // Update selection ke range hasil paste
  const endAddr = addrOf(anchor.r + h - 1, anchor.c + w - 1);
  selection = { start: anchorAddr, end: endAddr };
  activeCell = anchorAddr;

  // Clear cache & re-render
  if (typeof clearHFCache === 'function') clearHFCache();
  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  // Full re-render
  renderVirtual(false);
  updateSelectionUI();

  // Save
  scheduleSave();

  toast(`📥 Paste ${h}×${w} = ${written} sel`);
  return true;
}

/* =========================================================
   GLOBAL PASTE HANDLER — dari Excel/CSV
   ========================================================= */
document.addEventListener('paste', async (e) => {
  // Skip kalau sedang edit cell
  if (editor) return;

  // Skip kalau target adalah input/textarea/modal
  const target = e.target;
  if (!target) return;
  const tag = (target.tagName || '').toUpperCase();
  if (tag === 'INPUT' || tag === 'TEXTAREA') return;
  if (target.closest && target.closest('.modal')) return;
  if (target.closest && target.closest('#formula-input')) return;
  if (target.closest && target.closest('#cs-formula-input')) return;
  if (target.closest && target.closest('.custom-dialog-overlay')) return;

  // Skip kalau view mode
  if (currentAccessMode === 'view') return;

  // Ambil data dari clipboard
  const cd = e.clipboardData;
  if (!cd) return;

  const text = cd.getData('text/plain');
  if (!text || !text.trim()) return;

  // Deteksi kalau cuma 1 cell (tanpa tab & newline) → biarkan default paste
  const hasTab = text.includes('\t');
  const hasNewline = text.includes('\n') || text.includes('\r');
  if (!hasTab && !hasNewline) return;   // biarkan handler lama

  // Parse
  const rows = parseClipboardTable(text);
  if (!rows.length || rows.length === 1 && rows[0].length === 1) return;

  e.preventDefault();
  e.stopPropagation();

  console.log(`📋 Paste detected: ${rows.length} rows × ${Math.max(...rows.map(r => r.length))} cols`);

  paste2DArray(rows, activeCell);
}, true);

/* =========================================================
   FORMAT PAINTER
   ========================================================= */
let _fpActive = false;       // sedang aktif?
let _fpLocked = false;       // double-click = locked (multi-paste)
let _fpSourceStyle = null;   // style yang mau diapply
let _fpSourceAddr = null;    // address sumber (untuk debug)

function startFormatPainter() {
  if (!editor && !crossSheetRef) {
    // Ambil style dari cell aktif
    const cur = sheet().cells[activeCell];
    if (!cur?.s || Object.keys(cur.s).length === 0) {
      toast('⚠ Cell sumber tidak punya format');
      return;
    }
    _fpSourceStyle = JSON.parse(JSON.stringify(cur.s));
    _fpSourceAddr = activeCell;
    _fpActive = true;
    _fpLocked = false;
  } else {
    toast('⚠ Selesaikan edit dulu');
    return;
  }

  // Visual feedback
  const btn = document.getElementById('btn-format-painter');
  if (btn) btn.classList.add('active');
  document.body.classList.add('format-painter-active');

  toast(`🖌 Format dari ${_fpSourceAddr} — klik cell target (Esc untuk batal)`);
  console.log('🖌 Format Painter START from', _fpSourceAddr, ':', _fpSourceStyle);
}

function stopFormatPainter() {
  _fpActive = false;
  _fpLocked = false;
  _fpSourceStyle = null;
  _fpSourceAddr = null;

  const btn = document.getElementById('btn-format-painter');
  if (btn) btn.classList.remove('active');
  document.body.classList.remove('format-painter-active');
}

function applyFormatPainterTo(addr) {
  if (!_fpSourceStyle) return;

  const parts = expandRange(selection.start, selection.end);
  // Kalau user klik 1 cell, langsung apply ke cell itu
  // Kalau ada selection lebih dari 1, apply ke seluruh selection
  const targets = parts.length > 1 ? parts : [addr];

  pushUndo();

  for (const a of targets) {
    const cur = sheet().cells[a];
    if (cur) {
      cur.s = JSON.parse(JSON.stringify(_fpSourceStyle));
      markCellDirty(a);
    } else {
      // Buat cell kosong dengan style
      ensureCell(a);
      sheet().cells[a].s = JSON.parse(JSON.stringify(_fpSourceStyle));
      markCellDirty(a);
    }
  }

  // Update HF tidak perlu (cuma style)

  // Force re-render
  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};
  renderVirtual(false);

  toast(`✅ Format ${targets.length} sel dari ${_fpSourceAddr}`);

  // Kalau bukan locked mode, matikan
  if (!_fpLocked) {
    stopFormatPainter();
  }
}

/* =========================================================
   BIND TOMBOL
   ========================================================= */
(function initFormatPainter() {
  const btn = document.getElementById('btn-format-painter');
  if (!btn) {
    console.warn('btn-format-painter tidak ditemukan');
    return;
  }

  // Single click = aktif 1x
  // Double click = aktif locked (multi-paste)
  let _clickTimer = null;

  btn.addEventListener('click', e => {
    e.preventDefault();
    e.stopPropagation();

    if (_fpActive && !_fpLocked) {
      // Klik lagi saat aktif → batal
      stopFormatPainter();
      toast('Format Painter dibatalkan');
      return;
    }

    // Klik pertama
    if (!_fpActive) {
      startFormatPainter();
    }

    // Tunggu untuk deteksi double-click
    if (_clickTimer) clearTimeout(_clickTimer);
    _clickTimer = setTimeout(() => {
      _clickTimer = null;
    }, 250);
  });

  btn.addEventListener('dblclick', e => {
    e.preventDefault();
    e.stopPropagation();
    if (_clickTimer) {
      clearTimeout(_clickTimer);
      _clickTimer = null;
    }
    // Kalau sudah aktif, jadikan locked
    if (_fpActive) {
      _fpLocked = true;
      toast('🖌 Format Painter LOCKED — multi-paste aktif');
    } else {
      startFormatPainter();
      _fpLocked = true;
      toast('🖌 Format Painter LOCKED — multi-paste aktif');
    }
  });

  console.log('✅ Format Painter initialized');
})();

/* =========================================================
   HANDLE CLICK DI GRID SAAT FORMAT PAINTER AKTIF
   ========================================================= */
document.getElementById('grid')?.addEventListener('mousedown', e => {
  if (!_fpActive) return;

  const td = e.target.closest('td[data-addr]');
  if (!td) return;

  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();

  const addr = td.dataset.addr;
  console.log('🖌 Format Painter apply to', addr);

  applyFormatPainterTo(addr);
}, true);   // capture phase — sebelum handler select cell

/* =========================================================
   ESCAPE UNTUK CANCEL
   ========================================================= */
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && _fpActive) {
    stopFormatPainter();
    toast('Format Painter dibatalkan');
  }
}, true);

/* =========================================================
   MULTILINE — update class berdasarkan isi editor
   ========================================================= */
function updateMultilineClass(input, td) {
  if (!input || !td) return;
  const hasNewline = String(input.value).includes('\n');
  input.classList.toggle('multiline-input', hasNewline);
  td.classList.toggle('multiline', hasNewline);
}

/* =========================================================
   SHEET TABS — Context Menu, Move, Duplicate, Delete
   ========================================================= */

let _tabCtxMenu = null;

function closeTabContextMenu() {
  if (_tabCtxMenu) {
    _tabCtxMenu.remove();
    _tabCtxMenu = null;
  }
}

function openTabContextMenu(x, y, sheetIdx) {
  closeTabContextMenu();

  const sh = wbData.sheets[sheetIdx];
  if (!sh) return;

  const isHidden = !!sh.hidden;

  const menu = document.createElement('div');
  menu.className = 'tab-context-menu';

  menu.innerHTML = `
    <button data-action="duplicate">
      <i class="fa-solid fa-copy"></i> Duplicate Sheet
    </button>
    <button data-action="rename">
      <i class="fa-solid fa-pen"></i> Rename Sheet
    </button>
    <button data-action="toggle-hidden">
      <i class="fa-solid ${isHidden ? 'fa-eye' : 'fa-eye-slash'}"></i>
      ${isHidden ? 'Unhide Sheet' : 'Hide Sheet'}
    </button>
    <div class="ct-divider"></div>
    <button data-action="delete" ${wbData.sheets.length <= 1 ? 'disabled' : ''}>
      <i class="fa-solid fa-trash" style="color:#dc2626"></i>
      <span style="color:#dc2626">Delete Sheet</span>
    </button>
    <div class="ct-divider"></div>
    <div style="padding: 4px 12px; font-size: 10.5px; color: #94a3b8; text-transform: uppercase; letter-spacing: .4px; font-weight: 700;">Warna Tab</div>
    <div class="tab-color-row">
      ${['none','blue','green','red','yellow','purple','pink','gray']
        .map(c => {
          const color = c === 'none' ? '#ffffff' : {
            blue: '#3b82f6', green: '#10b981', red: '#ef4444',
            yellow: '#f59e0b', purple: '#8b5cf6', pink: '#ec4899',
            gray: '#64748b'
          }[c];
          const active = (sh.tabColor || 'none') === c;
          return `<div class="tab-color-swatch ${active ? 'active' : ''}"
                       data-color="${c}"
                       title="${c}"
                       style="background: ${color}; ${c === 'none' ? 'border: 2px dashed #cbd5e1;' : ''}"></div>`;
        }).join('')}
    </div>
  `;

  document.body.appendChild(menu);
  _tabCtxMenu = menu;

  // Position
  const rect = menu.getBoundingClientRect();
  if (x + rect.width > window.innerWidth - 10) x = window.innerWidth - rect.width - 10;
  if (y + rect.height > window.innerHeight - 10) y = window.innerHeight - rect.height - 10;
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';

  // Button handlers
  menu.querySelectorAll('button[data-action]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const action = btn.dataset.action;
      closeTabContextMenu();

      if (action === 'duplicate') {
        await duplicateSheet(sheetIdx);
      } else if (action === 'rename') {
        const n = await customPrompt('Nama sheet baru:', sh.name, { title: 'Rename Sheet' });
        if (n && n.trim() && n.trim() !== sh.name) {
          const oldName = sh.name;
          const newName = n.trim();

          // 🎯 Update named ranges yang refer ke sheet ini
          const ranges = wbData.namedRanges || {};
          const escapedOld = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const re = new RegExp(`^('?)${escapedOld}\\1!`);

          for (const key in ranges) {
            const ref = ranges[key];
            if (re.test(ref)) {
              const needsQuote = /[\s\-.]/.test(newName);
              const newSheetRef = needsQuote
                ? "'" + newName.replace(/'/g, "''") + "'!"
                : newName + '!';
              ranges[key] = ref.replace(re, newSheetRef);
            }
          }

          sh.name = newName;
          markMetaDirty('structure');
          try { resetHF(); } catch {}
          renderSheetTabs();
          renderSheet();
          scheduleSave();
          toast('✔ Sheet di-rename');
        }
      } else if (action === 'toggle-hidden') {
        const visibleCount = wbData.sheets.filter(s => !s.hidden).length;
        
        if (!isHidden) {
          // SEDANG HIDE
          if (visibleCount <= 1) {
            return toast('⚠ Minimal 1 sheet terlihat');
          }
          
          sh.hidden = true;
          
          // 🎯 Kalau yang di-hide = active sheet, pindah ke sheet lain dulu
          if (activeSheetIdx === sheetIdx) {
            // Cari sheet lain yang tidak hidden
            const nextIdx = wbData.sheets.findIndex((s, i) => i !== sheetIdx && !s.hidden);
            if (nextIdx >= 0) {
              activeSheetIdx = nextIdx;
              wbData.activeSheet = nextIdx;
              await switchSheet(nextIdx);
            }
          }
        } else {
          // SEDANG UNHIDE
          sh.hidden = false;
        }
        
        markMetaDirty('structure');
        renderSheetTabs();
        scheduleSave();
        toast(isHidden ? '✔ Sheet ditampilkan' : '✔ Sheet disembunyikan');
      } else if (action === 'delete') {
        await deleteSheet(sheetIdx);
      }
    });
  });

  // Color swatches
  menu.querySelectorAll('.tab-color-swatch').forEach(sw => {
    sw.addEventListener('click', () => {
      const color = sw.dataset.color;
      sh.tabColor = color === 'none' ? null : color;
      markMetaDirty('structure');
      closeTabContextMenu();
      renderSheetTabs();
      scheduleSave();
      toast(`✔ Warna tab: ${color}`);
    });
  });

  // Close on outside
  setTimeout(() => {
    document.addEventListener('mousedown', function h(e) {
      if (_tabCtxMenu && !_tabCtxMenu.contains(e.target)) {
        closeTabContextMenu();
        document.removeEventListener('mousedown', h, true);
      }
    }, true);
  }, 0);
}

function moveSheet(fromIdx, toIdx) {
  if (fromIdx === toIdx) return;
  toIdx = Math.max(0, Math.min(toIdx, wbData.sheets.length));

  const moved = wbData.sheets.splice(fromIdx, 1)[0];
  if (fromIdx < toIdx) toIdx--;
  wbData.sheets.splice(toIdx, 0, moved);

  // Fix activeSheetIdx
  if (activeSheetIdx === fromIdx) {
    activeSheetIdx = toIdx;
    wbData.activeSheet = toIdx;
  } else if (fromIdx < activeSheetIdx && toIdx >= activeSheetIdx) {
    activeSheetIdx--;
  } else if (fromIdx > activeSheetIdx && toIdx <= activeSheetIdx) {
    activeSheetIdx++;
  }

  loadedSheetIdxs = new Set(wbData.sheets.map((_, i) => i));

  markMetaDirty('structure');
  renderSheetTabs();
  forceRefreshGrid();
  scheduleSave();
  toast('✔ Sheet dipindahkan');
}

async function duplicateSheet(idx) {
  const sh = wbData.sheets[idx];
  if (!sh) return;

  if (!loadedSheetIdxs.has(idx)) {
    showLoading('Memuat sheet…');
    try { await lazyLoadSheet(idx); } catch {}
    hideLoading();
  }

  const copy = JSON.parse(JSON.stringify(sh));
  copy.name = `Copy of ${sh.name}`.slice(0, 30);
  copy.hidden = false;
  copy.protected = false;
  delete copy.protectHash;

  wbData.sheets.splice(idx + 1, 0, copy);
  loadedSheetIdxs = new Set(wbData.sheets.map((_, i) => i));

  markMetaDirty('structure');
  try { resetHF(); } catch {}
  renderSheetTabs();
  scheduleSave();

  await switchSheet(idx + 1);
  toast(`✔ Sheet "${copy.name}" dibuat`);
}

async function deleteSheet(idx) {
  if (wbData.sheets.length <= 1) {
    return toast('Minimal harus ada 1 sheet');
  }

  const sh = wbData.sheets[idx];
  if (!(await customConfirm(`Hapus sheet "${sh.name}"?`, { type: 'danger', okText: 'Hapus' }))) return;

  showLoading('Menyiapkan…');
  try {
    for (let k = 0; k < wbData.sheets.length; k++) {
      if (!loadedSheetIdxs.has(k)) {
        updateLoading(`Memuat sheet "${wbData.sheets[k].name}"…`);
        await lazyLoadSheet(k);
      }
    }
  } catch (err) { console.warn('Gagal load semua:', err); }
  hideLoading();

  try { await saveWB(); } catch {}

  // 🎯 Hapus dari array
  wbData.sheets.splice(idx, 1);
  loadedSheetIdxs = new Set(wbData.sheets.map((_, k) => k));

  // 🎯 Fix activeSheetIdx — pastikan tidak out-of-range
  if (idx < activeSheetIdx) {
    // Hapus sheet sebelum active → active index geser -1
    activeSheetIdx--;
  } else if (idx === activeSheetIdx) {
    // Hapus active → pindah ke sheet tetangga
    activeSheetIdx = Math.min(idx, wbData.sheets.length - 1);
  }
  // Kalau idx > activeSheetIdx, tidak perlu ubah

  wbData.activeSheet = activeSheetIdx;

  saveFilterStateToSheet();
  if (typeof loadFilterStateFromSheet === 'function') loadFilterStateFromSheet();

  try { await rebuildHFWithLoaded(); } catch {}

  // 🎯 Force refresh TOTAL — reset semua cache
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  renderSheetTabs();
  if (typeof forceRefreshGrid === 'function') forceRefreshGrid();
  else renderSheet();

  markMetaDirty('structure');
  scheduleSave();
  toast(`✔ Sheet "${sh.name}" dihapus`);
}

/* =========================================================
   UNHIDE SHEET
   ========================================================= */
(function initUnhideSheet() {
  const btn = document.getElementById('btn-unhide-sheet');
  const badge = document.getElementById('hidden-sheet-count');
  if (!btn) return;

  // Update badge count
  function updateBadge() {
    if (!badge) return;
    const hiddenCount = wbData.sheets.filter(s => s.hidden).length;
    if (hiddenCount > 0) {
      badge.textContent = hiddenCount;
      badge.style.display = 'flex';
      btn.classList.add('has-hidden');
    } else {
      badge.style.display = 'none';
      btn.classList.remove('has-hidden');
    }
  }

  // Panggil setiap kali tab render
  const _origRenderTabs = renderSheetTabs;
  renderSheetTabs = function() {
    _origRenderTabs();
    updateBadge();
  };

  btn.addEventListener('click', e => {
    e.preventDefault();
    e.stopPropagation();

    // Tutup kalau sudah terbuka
    const existing = document.getElementById('btn-unhide-sheet-list');
    if (existing) {
      existing.remove();
      return;
    }

    const hidden = wbData.sheets
      .map((s, i) => ({ ...s, idx: i }))
      .filter(s => s.hidden);

    if (!hidden.length) {
      return toast('Tidak ada sheet tersembunyi');
    }

    const list = document.createElement('div');
    list.id = 'btn-unhide-sheet-list';
    list.innerHTML = hidden.map(s => `
      <button data-unhide-idx="${s.idx}">
        <i class="fa-solid fa-eye"></i>
        <span>${escapeHtml(s.name)}</span>
      </button>
    `).join('');

    btn.parentElement.style.position = 'relative';
    btn.parentElement.appendChild(list);

    list.querySelectorAll('button[data-unhide-idx]').forEach(b => {
      b.addEventListener('click', async () => {
        const idx = parseInt(b.dataset.unhideIdx, 10);
        const sh = wbData.sheets[idx];
        if (!sh) return;

        sh.hidden = false;
        list.remove();

        markMetaDirty('structure');
        renderSheetTabs();
        scheduleSave();

        // Switch ke sheet yang di-unhide
        await switchSheet(idx);
        toast(`✔ Sheet "${sh.name}" ditampilkan`);
      });
    });

    // Close on outside click
    setTimeout(() => {
      document.addEventListener('mousedown', function h(ev) {
        if (list && !list.contains(ev.target) && ev.target !== btn && !btn.contains(ev.target)) {
          list.remove();
          document.removeEventListener('mousedown', h, true);
        }
      }, true);
    }, 0);
  });

  console.log('✅ Unhide sheet initialized');
})();

/* =========================================================
   ROW GROUPING
   ========================================================= */
let _rowGroupCache = null;

function invalidateRowGroupCache() {
  _rowGroupCache = null;
  cachedRowMap = null;   // filter cache juga harus invalidate
}

function getRowGroups() {
  const sh = sheet();
  if (!sh) return [];
  if (!Array.isArray(sh.rowGroups)) sh.rowGroups = [];
  return sh.rowGroups;
}

/* Cek apakah row r ada di dalam group manapun */
function getRowGroupAt(r) {
  const groups = getRowGroups();
  for (const g of groups) {
    if (r >= g.from && r <= g.to) return g;
  }
  return null;
}

/* Cek apakah row r adalah "summary row" dari group */
function getGroupSummaryAt(r) {
  const groups = getRowGroups();
  for (const g of groups) {
    if (r === g.from - 1) return g;
  }
  return null;
}

/* Tambah group baru untuk range [from, to] */
function addRowGroup(from, to) {
  if (from > to) [from, to] = [to, from];
  if (from === to) {
    toast('⚠ Minimal 2 baris untuk group');
    return;
  }

  const sh = sheet();
  const groups = getRowGroups();

  // Cek overlap
  for (const g of groups) {
    const overlap = !(to < g.from || from > g.to);
    if (overlap) {
      toast('⚠ Range beririsan dengan group yang ada');
      return;
    }
  }

  pushUndo();
  groups.push({ from, to, collapsed: false });

  // Sort by from
  groups.sort((a, b) => a.from - b.from);

  invalidateRowGroupCache();
  markMetaDirty('structure');
  scheduleSave();

  window._lastFrozenKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual(false);

  toast(`✔ Group dibuat (${from + 1}–${to + 1})`);
}

/* Hapus group yang mencakup row r */
function removeRowGroup(r) {
  const sh = sheet();
  const groups = getRowGroups();
  const before = groups.length;

  const filtered = groups.filter(g => !(r >= g.from && r <= g.to));
  if (filtered.length === before) {
    toast('Tidak ada group di baris ini');
    return;
  }

  pushUndo();
  sh.rowGroups = filtered;
  invalidateRowGroupCache();
  markMetaDirty('structure');
  scheduleSave();

  window._lastFrozenKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual(false);

  toast('✔ Group dihapus');
}

/* Toggle collapse/expand group */
function toggleRowGroupCollapse(groupId) {
  const groups = getRowGroups();
  const g = groups.find(x => x === groupId || x.id === groupId);
  if (!g) return;

  g.collapsed = !g.collapsed;
  invalidateRowGroupCache();
  markMetaDirty('structure');
  scheduleSave();

  window._lastFrozenKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual(false);
}

/* Build set baris yang di-hide karena collapse */
function getCollapsedRowSet() {
  if (_rowGroupCache) return _rowGroupCache;

  const set = new Set();
  const groups = getRowGroups();
  for (const g of groups) {
    if (!g.collapsed) continue;
    for (let r = g.from; r <= g.to; r++) set.add(r);
  }

  _rowGroupCache = set;
  return set;
}

/* =========================================================
   KEYBOARD SHORTCUTS HELP
   ========================================================= */
const SHORTCUTS_DATA = [
  {
    category: 'File & Workbook',
    items: [
      { keys: ['Ctrl', 'S'], desc: 'Simpan workbook' },
      { keys: ['Ctrl', 'N'], desc: 'Workbook baru' },
      { keys: ['Ctrl', 'P'], desc: 'Print' },
    ]
  },
  {
    category: 'Edit',
    items: [
      { keys: ['Ctrl', 'Z'], desc: 'Undo' },
      { keys: ['Ctrl', 'Y'], desc: 'Redo' },
      { keys: ['Ctrl', 'C'], desc: 'Copy' },
      { keys: ['Ctrl', 'X'], desc: 'Cut' },
      { keys: ['Ctrl', 'V'], desc: 'Paste' },
      { keys: ['Ctrl', 'Shift', 'V'], desc: 'Paste value only' },
      { keys: ['Delete'], desc: 'Hapus sel' },
      { keys: ['Backspace'], desc: 'Hapus sel' },
      { keys: ['F2'], desc: 'Edit sel' },
      { keys: ['Enter'], desc: 'Commit & pindah bawah' },
      { keys: ['Shift', 'Enter'], desc: 'Commit & pindah atas' },
      { keys: ['Tab'], desc: 'Commit & pindah kanan' },
      { keys: ['Shift', 'Tab'], desc: 'Commit & pindah kiri' },
      { keys: ['Alt', 'Enter'], desc: 'Line break dalam cell' },
      { keys: ['Escape'], desc: 'Batal edit / tutup dialog' },
    ]
  },
  {
    category: 'Navigasi',
    items: [
      { keys: ['↑'], desc: 'Pindah atas' },
      { keys: ['↓'], desc: 'Pindah bawah' },
      { keys: ['←'], desc: 'Pindah kiri' },
      { keys: ['→'], desc: 'Pindah kanan' },
      { keys: ['Ctrl', '↑'], desc: 'Jump ke edge atas' },
      { keys: ['Ctrl', '↓'], desc: 'Jump ke edge bawah' },
      { keys: ['Ctrl', '←'], desc: 'Jump ke edge kiri' },
      { keys: ['Ctrl', '→'], desc: 'Jump ke edge kanan' },
      { keys: ['Shift', '↑'], desc: 'Extend selection atas' },
      { keys: ['Shift', '↓'], desc: 'Extend selection bawah' },
      { keys: ['Shift', '←'], desc: 'Extend selection kiri' },
      { keys: ['Shift', '→'], desc: 'Extend selection kanan' },
      { keys: ['Ctrl', 'G'], desc: 'Go To cell' },
      { keys: ['Page Up'], desc: 'Scroll halaman naik' },
      { keys: ['Page Down'], desc: 'Scroll halaman turun' },
      { keys: ['Home'], desc: 'Ke awal baris' },
    ]
  },
  {
    category: 'Format',
    items: [
      { keys: ['Ctrl', 'B'], desc: 'Bold' },
      { keys: ['Ctrl', 'I'], desc: 'Italic' },
      { keys: ['Ctrl', 'U'], desc: 'Underline' },
      { keys: ['F4'], desc: 'Cycle absolute reference ($A$1 → A$1 → $A1 → A1)' },
    ]
  },
  {
    category: 'Formula',
    items: [
      { keys: ['='], desc: 'Mulai formula' },
      { keys: ['F4'], desc: 'Toggle $ absolute ref di formula' },
      { keys: ['Ctrl', '`'], desc: 'Show formulas / values' },
      { keys: ['Ctrl', 'Shift', 'V'], desc: 'Paste value only' },
    ]
  },
  {
    category: 'Find & Replace',
    items: [
      { keys: ['Ctrl', 'F'], desc: 'Find' },
      { keys: ['Ctrl', 'H'], desc: 'Find & Replace' },
      { keys: ['Enter'], desc: 'Match berikutnya (dalam Find)' },
      { keys: ['Shift', 'Enter'], desc: 'Match sebelumnya' },
    ]
  },
  {
    category: 'Insert & Insert',
    items: [
      { keys: ['Ctrl', 'K'], desc: 'Insert hyperlink' },
    ]
  },
  {
    category: 'Bantuan',
    items: [
      { keys: ['Ctrl', '/'], desc: 'Tampilkan panel shortcut ini' },
      { keys: ['F1'], desc: 'Tampilkan panel shortcut ini' },
    ]
  }
];

function renderShortcuts(filter = '') {
  const body = document.getElementById('shortcuts-body');
  if (!body) return;

  const q = filter.trim().toLowerCase();

  const filtered = SHORTCUTS_DATA.map(cat => ({
    category: cat.category,
    items: cat.items.filter(it => {
      if (!q) return true;
      const keyStr = it.keys.join('+').toLowerCase();
      return it.desc.toLowerCase().includes(q) || keyStr.includes(q);
    })
  })).filter(cat => cat.items.length > 0);

  if (!filtered.length) {
    body.innerHTML = `<div class="sc-empty">
      <i class="fa-solid fa-magnifying-glass" style="font-size:32px;color:#cbd5e1;display:block;margin-bottom:12px"></i>
      Tidak ada shortcut yang cocok
    </div>`;
    return;
  }

  body.innerHTML = filtered.map(cat => `
    <div class="sc-category">
      <i class="fa-solid fa-circle-dot" style="color:#10b981;font-size:9px"></i>
      ${escapeHtml(cat.category)}
    </div>
    ${cat.items.map(it => `
      <div class="sc-item">
        <span class="sc-desc">${escapeHtml(it.desc)}</span>
        <span class="sc-keys">
          ${it.keys.map(k => `<span class="sc-key">${escapeHtml(k)}</span>`).join('<span style="color:#cbd5e1;font-size:10px">+</span>')}
        </span>
      </div>
    `).join('')}
  `).join('');
}

function openShortcutsModal() {
  const modal = document.getElementById('modal-shortcuts');
  if (!modal) return;
  modal.classList.remove('hidden');

  const input = document.getElementById('shortcuts-search');
  if (input) input.value = '';

  renderShortcuts();

  setTimeout(() => input?.focus(), 50);
}

function closeShortcutsModal() {
  document.getElementById('modal-shortcuts')?.classList.add('hidden');
}

/* Bind tombol */
document.getElementById('btn-shortcuts-close')?.addEventListener('click', closeShortcutsModal);
document.getElementById('btn-shortcuts-close2')?.addEventListener('click', closeShortcutsModal);

document.getElementById('shortcuts-search')?.addEventListener('input', e => {
  renderShortcuts(e.target.value);
});

document.getElementById('modal-shortcuts')?.addEventListener('click', e => {
  if (e.target.id === 'modal-shortcuts') closeShortcutsModal();
});

/* Shortcut: Ctrl+/ atau F1 */
document.addEventListener('keydown', e => {
  if (editor) return;
  const ctrl = e.ctrlKey || e.metaKey;

  // Ctrl+/ 
  if (ctrl && e.key === '/') {
    e.preventDefault();
    openShortcutsModal();
    return;
  }

  // F1
  if (e.key === 'F1') {
    e.preventDefault();
    openShortcutsModal();
    return;
  }
}, true);

document.getElementById('btn-shortcuts-help')?.addEventListener('click', openShortcutsModal);
document.getElementById('btn-shortcuts-help2')?.addEventListener('click', openShortcutsModal);
console.log('✅ Keyboard Shortcuts Help initialized');

/* =========================================================
   DARK MODE
   ========================================================= */
const DARK_MODE_KEY = 'wex-dark-mode';

function isDarkMode() {
  return document.body.classList.contains('dark-mode');
}

function applyDarkMode(enabled) {
  if (enabled) {
    document.body.classList.add('dark-mode');
    localStorage.setItem(DARK_MODE_KEY, '1');
  } else {
    document.body.classList.remove('dark-mode');
    localStorage.setItem(DARK_MODE_KEY, '0');
  }

  // Update toggle UI
  const toggle = document.getElementById('dark-mode-toggle');
  if (toggle) {
    toggle.checked = enabled;
  }
}

function toggleDarkMode() {
  applyDarkMode(!isDarkMode());
  toast(isDarkMode() ? '🌙 Dark mode aktif' : '☀️ Light mode aktif');
}

/* Load preference saat startup */
(function initDarkMode() {
  const saved = localStorage.getItem(DARK_MODE_KEY);
  
  // Kalau user belum pernah pilih, auto-detect dari OS
  if (saved === null) {
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    applyDarkMode(prefersDark);
  } else {
    applyDarkMode(saved === '1');
  }

  // Listen OS preference change (kalau user belum set manual)
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', e => {
    if (localStorage.getItem(DARK_MODE_KEY) === null) {
      applyDarkMode(e.matches);
    }
  });

  // Bind toggle checkbox
  const toggle = document.getElementById('dark-mode-toggle');
  if (toggle) {
    toggle.checked = isDarkMode();
    toggle.addEventListener('change', e => {
      applyDarkMode(e.target.checked);
    });
  }

  // Bind shortcut Ctrl+Shift+D
  document.addEventListener('keydown', e => {
    if (editor) return;
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && e.shiftKey && e.key.toLowerCase() === 'd') {
      e.preventDefault();
      toggleDarkMode();
    }
  }, true);

  console.log('✅ Dark Mode initialized. Current:', isDarkMode() ? 'dark' : 'light');
})();

/* =========================================================
   NAMED RANGES
   ========================================================= */
function getNamedRanges() {
  if (!wbData.namedRanges) wbData.namedRanges = {};
  return wbData.namedRanges;
}

/* =========================================================
   RESOLVE NAMED RANGES IN FORMULA
   ========================================================= */
function resolveNamedRanges(formula) {
  if (!formula || typeof formula !== 'string') return formula;
  if (!formula.startsWith('=')) return formula;

  const ranges = wbData.namedRanges || {};
  const keys = Object.keys(ranges);
  if (!keys.length) return formula;

  let result = formula;
  for (const name of keys) {
    const ref = ranges[name];
    const re = new RegExp(
      `\\b${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b(?!\\s*\\()`,
      'g'
    );
    result = result.replace(re, ref);
  }
  return result;
}

function addNamedRange(name, rangeRef) {
  if (!name || !rangeRef) {
    toast('⚠ Nama dan range wajib diisi');
    return false;
  }

  // Validasi nama: huruf, angka, underscore; tidak boleh mulai angka atau mirip cell ref
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    toast('⚠ Nama harus huruf/angka/underscore, tidak boleh mulai angka');
    return false;
  }

  if (/^[A-Z]+[0-9]+$/i.test(name)) {
    toast('⚠ Nama tidak boleh sama dengan alamat cell (A1, B2, ...)');
    return false;
  }

  const ranges = getNamedRanges();
  ranges[name] = rangeRef;
  markMetaDirty('structure');

  // 🎯 Rebuild HyperFormula biar named range kepakai di formula
  try { resetHF(); } catch (e) { console.warn('resetHF failed:', e); }

  scheduleSave();
  refreshNamedRangeCache();

  // 🎯 Force re-render
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  if (typeof renderVirtual === 'function') renderVirtual(false);

  return true;
}

function removeNamedRange(name) {
  const ranges = getNamedRanges();
  if (ranges[name] === undefined) return false;
  delete ranges[name];
  markMetaDirty('structure');

  // 🎯 Rebuild HyperFormula
  try { resetHF(); } catch (e) { console.warn('resetHF failed:', e); }

  scheduleSave();
  refreshNamedRangeCache();

  // 🎯 Force re-render
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  if (typeof renderVirtual === 'function') renderVirtual(false);

  return true;
}

/* Cache — refresh saat ada perubahan */
let _namedRangeCache = null;

function refreshNamedRangeCache() {
  _namedRangeCache = null;
}

/* Cek apakah sebuah string adalah named range */
function isNamedRange(name) {
  const ranges = getNamedRanges();
  return ranges[name] !== undefined;
}

/* Convert named range reference "Sheet1!A1:B10" menjadi absolut address */
function resolveNamedRange(ref) {
  if (!ref) return ref;

  // Kalau ref adalah named range
  const ranges = getNamedRanges();
  if (ranges[ref]) {
    return resolveNamedRange(ranges[ref]);
  }

  // Kalau ref mengandung named range di dalam formula, replace
  return ref.replace(/\b([A-Za-z_][A-Za-z0-9_]*)\b/g, (match) => {
    // Skip function names (diikuti `(`)
    if (ranges[match] !== undefined) {
      return ranges[match];
    }
    return match;
  });
}

/* Convert named range reference "Sheet1!A1:B10" menjadi {sheetIdx, addr} */
function parseNamedRangeRef(ref) {
  if (!ref) return null;
  // Format: SheetName!A1:B10 atau SheetName!A1
  const m = /^'?([^'!]+)'?!(.+)$/.exec(ref);
  if (!m) return null;

  const sheetName = m[1];
  const addrPart = m[2];

  // Cari sheet index
  const sheetIdx = wbData.sheets.findIndex(s => s.name === sheetName);
  if (sheetIdx < 0) return null;

  // Parse range addr
  const parts = addrPart.split(':');
  const startAddr = parts[0];
  const endAddr = parts[1] || parts[0];

  return { sheetIdx, startAddr, endAddr, addrPart };
}

/* =========================================================
   NAME MANAGER MODAL
   ========================================================= */
function openNameManager() {
  const modal = document.getElementById('modal-names');
  if (!modal) return;

  // Reset form
  document.getElementById('name-new-name').value = '';
  document.getElementById('name-new-range').value = '';

  renderNamesList();
  modal.classList.remove('hidden');

  setTimeout(() => document.getElementById('name-new-name')?.focus(), 50);
}

function closeNameManager() {
  document.getElementById('modal-names')?.classList.add('hidden');
}

function renderNamesList() {
  const box = document.getElementById('names-list');
  if (!box) return;

  const ranges = getNamedRanges();
  const entries = Object.entries(ranges);

  if (!entries.length) {
    box.innerHTML = `<div style="padding:30px 20px;text-align:center;color:#94a3b8">
      <i class="fa-solid fa-bookmark" style="font-size:32px;color:#cbd5e1;display:block;margin-bottom:12px"></i>
      Belum ada named range
    </div>`;
    return;
  }

  box.innerHTML = entries.map(([name, ref]) => `
    <div class="rev-row">
      <div class="body">
        <div class="name" style="font-weight:600;color:#059669">
          <i class="fa-solid fa-bookmark" style="font-size:11px;margin-right:6px"></i>
          ${escapeHtml(name)}
        </div>
        <div class="meta" style="font-family:monospace;font-size:11px;color:#64748b">
          ${escapeHtml(ref)}
        </div>
      </div>
      <button class="perm-btn danger" data-name-del="${escapeHtml(name)}" title="Hapus">
        <i class="fa-solid fa-trash"></i>
      </button>
    </div>
  `).join('');

  box.querySelectorAll('[data-name-del]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const name = btn.dataset.nameDel;
      if (!(await customConfirm(`Hapus named range "${name}"?`, { type: 'danger', okText: 'Hapus' }))) return;
      if (removeNamedRange(name)) {
        renderNamesList();
        toast('✔ Named range dihapus');
      }
    });
  });
}

/* Bind tombol */
document.getElementById('btn-name-manager')?.addEventListener('click', openNameManager);
document.getElementById('btn-names-close')?.addEventListener('click', closeNameManager);
document.getElementById('btn-names-cancel')?.addEventListener('click', closeNameManager);
document.getElementById('modal-names')?.addEventListener('click', e => {
  if (e.target.id === 'modal-names') closeNameManager();
});

/* Define Name — dari selection */
document.getElementById('btn-name-from-selection')?.addEventListener('click', async () => {
  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return toast('Pilih range dulu');

  const name = await customPrompt('Nama untuk range ini:', '', {
    title: 'Define Name',
    okText: 'Buat'
  });
  if (!name || !name.trim()) return;

  const sheetName = sheet().name;
  const needsQuote = /[\s\-.]/.test(sheetName);
  const safeSheetName = needsQuote
    ? "'" + sheetName.replace(/'/g, "''") + "'"
    : sheetName;

  const rangeRef = `${safeSheetName}!${selection.start}:${selection.end}`;

  if (addNamedRange(name.trim(), rangeRef)) {
    toast(`✔ Named range "${name}" dibuat`);
  }
});

/* Tambah dari modal */
document.getElementById('btn-name-add')?.addEventListener('click', () => {
  const name = document.getElementById('name-new-name').value.trim();
  const ref = document.getElementById('name-new-range').value.trim();

  if (!name || !ref) {
    return toast('⚠ Isi nama dan range');
  }

  // Auto-format ref — kalau user ketik "A1:B10" saja, tambahkan sheet aktif
  let finalRef = ref;
  if (!ref.includes('!')) {
    const sheetName = sheet().name;
    const needsQuote = /[\s\-.]/.test(sheetName);
    const safeSheetName = needsQuote
      ? "'" + sheetName.replace(/'/g, "''") + "'"
      : sheetName;
    finalRef = `${safeSheetName}!${ref}`;
  }

  if (addNamedRange(name, finalRef)) {
    document.getElementById('name-new-name').value = '';
    document.getElementById('name-new-range').value = '';
    renderNamesList();
    toast(`✔ Named range "${name}" dibuat`);
  }
});

/* Auto-fill range dari selection */
document.getElementById('name-new-range')?.addEventListener('focus', e => {
  if (e.target.value) return;
  if (!selection.start || !selection.end) return;
  const sheetName = sheet().name;
  const needsQuote = /[\s\-.]/.test(sheetName);
  const safeSheetName = needsQuote
    ? "'" + sheetName.replace(/'/g, "''") + "'"
    : sheetName;
  e.target.value = `${safeSheetName}!${selection.start}:${selection.end}`;
});

/* Enter di input */
document.getElementById('name-new-name')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    document.getElementById('name-new-range')?.focus();
  }
});

document.getElementById('name-new-range')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    document.getElementById('btn-name-add')?.click();
  }
});

console.log('✅ Named Ranges initialized');

/* =========================================================
   FORMULA AUDITING
   ========================================================= */
let _showFormulas = false;
let _highlightFormulas = false;

function toggleShowFormulas() {
  _showFormulas = !_showFormulas;
  document.body.classList.toggle('show-formulas', _showFormulas);

  // Update button state
  const btn = document.getElementById('btn-show-formulas');
  if (btn) {
    btn.classList.toggle('active', _showFormulas);
    btn.style.background = _showFormulas
      ? 'linear-gradient(135deg,#0369a1,#0284c7)' : '';
    btn.style.color = _showFormulas ? '#fff' : '';
  }

  // Force re-render — bypass cache
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  // Clear HF cache supaya re-evaluate
  if (typeof clearHFCache === 'function') clearHFCache();

  renderVirtual(false);
  toast(_showFormulas ? '👁 Tampil formula' : '👁 Tampil nilai');
}

/* =========================================================
   HIGHLIGHT FORMULA CELLS
   ========================================================= */
function toggleHighlightFormulas() {
  _highlightFormulas = !_highlightFormulas;

  const btn = document.getElementById('btn-highlight-formulas');
  if (btn) {
    btn.classList.toggle('active', _highlightFormulas);
    btn.style.background = _highlightFormulas
      ? 'linear-gradient(135deg,#f59e0b,#d97706)' : '';
    btn.style.color = _highlightFormulas ? '#fff' : '';
  }

  applyFormulaHighlight();
  toast(_highlightFormulas ? '✨ Formula di-highlight' : 'Highlight formula off');
}

function applyFormulaHighlight() {
  // Clear dulu
  document.querySelectorAll('#grid td.has-formula')
    .forEach(td => td.classList.remove('has-formula'));

  if (!_highlightFormulas) return;

  // Apply ke visible cells
  document.querySelectorAll('#grid tbody td[data-addr]').forEach(td => {
    const addr = td.dataset.addr;
    const cl = sheet().cells[addr];
    if (cl && cl.f) {
      td.classList.add('has-formula');
    }
  });
}

/* =========================================================
   ERROR CHECKING
   ========================================================= */
function runErrorCheck() {
  const modal = document.getElementById('modal-errors');
  const list = document.getElementById('errors-list');
  const summary = document.getElementById('errors-summary');
  if (!modal || !list) return;

  modal.classList.remove('hidden');

  // Scan sheet aktif
  const cells = sheet().cells || {};
  const errors = [];

  for (const [addr, cl] of Object.entries(cells)) {
    if (!cl.f) continue;   // hanya cell formula yang bisa error
    const p = parseAddr(addr);
    if (!p) continue;

    let value;
    try {
      value = hfGet(p.r, p.c);
    } catch {
      value = '#ERROR';
    }

    let errType = null;
    if (isErr(value)) {
      errType = value.type || 'ERROR';
    } else if (typeof value === 'string' && value.startsWith('#')) {
      errType = value;
    }

    if (errType) {
      errors.push({
        addr,
        formula: cl.f,
        type: errType,
        r: p.r,
        c: p.c
      });
    }
  }

  // Render summary
  if (summary) {
    if (errors.length === 0) {
      summary.innerHTML = `✅ Tidak ada error di sheet <strong>${escapeHtml(sheet().name)}</strong>`;
    } else {
      summary.innerHTML = `⚠️ Ditemukan <strong>${errors.length}</strong> error di sheet <strong>${escapeHtml(sheet().name)}</strong>`;
    }
  }

  // Render list
  if (errors.length === 0) {
    list.innerHTML = `<div class="error-empty">
      <i class="fa-solid fa-circle-check"></i>
      Sheet bersih! Tidak ada error.
    </div>`;
    return;
  }

  // Sort by addr
  errors.sort((a, b) => a.r - b.r || a.c - b.c);

  list.innerHTML = errors.map(e => `
    <div class="error-item" data-addr="${escapeHtml(e.addr)}">
      <span class="addr">${escapeHtml(e.addr)}</span>
      <span class="err-type">${escapeHtml(e.type)}</span>
      <span class="formula">${escapeHtml(e.formula)}</span>
      <i class="fa-solid fa-arrow-right" style="color:#94a3b8;font-size:11px"></i>
    </div>
  `).join('');

  // Click → jump to cell
  list.querySelectorAll('.error-item').forEach(item => {
    item.addEventListener('click', () => {
      const addr = item.dataset.addr;
      selectCell(addr);
      scrollAddrIntoView(addr);
      modal.classList.add('hidden');
      toast(`🎯 Ke ${addr}`);
    });
  });
}

function closeErrorModal() {
  document.getElementById('modal-errors')?.classList.add('hidden');
}

/* =========================================================
   ADVANCED FILTER
   ========================================================= */
let _advFilterDraft = [];   // kondisi sementara (belum di-apply)

function getAdvancedFilters() {
  const sh = sheet();
  if (!sh) return [];
  if (!Array.isArray(sh.advancedFilters)) sh.advancedFilters = [];
  return sh.advancedFilters;
}

/* Ambil header text dari kolom untuk label dropdown */
function getColumnHeaders() {
  const headers = [];
  const cells = sheet().cells || {};
  const headerRow = getFilterHeaderRow();

  // Scan kolom yang punya data di header row atau 20 baris pertama
  let maxC = 30;
  for (const addr of Object.keys(cells)) {
    const p = parseAddr(addr);
    if (p && p.c > maxC) maxC = p.c;
  }

  for (let c = 0; c <= maxC; c++) {
    let label = colName(c);
    // Cari header text
    for (let r = 0; r < 20; r++) {
      const cl = cells[addrOf(r, c)];
      if (cl && cl.v !== undefined && cl.v !== null && cl.v !== '') {
        const text = String(cl.v).trim();
        if (text) {
          label = `${colName(c)} — ${text.slice(0, 25)}`;
          break;
        }
      }
    }
    headers.push({ col: c, label });
  }

  return headers;
}

function openAdvancedFilterModal() {
  const modal = document.getElementById('modal-adv-filter');
  if (!modal) return;

  // Load draft dari existing
  _advFilterDraft = JSON.parse(JSON.stringify(getAdvancedFilters()));
  if (_advFilterDraft.length === 0) {
    // Tambah 1 baris kosong default
    _advFilterDraft.push({ col: 0, op: 'eq', value: '', logic: 'AND' });
  }

  renderAdvConditions();
  modal.classList.remove('hidden');
  updateAdvPreview();
}

function closeAdvancedFilterModal() {
  document.getElementById('modal-adv-filter')?.classList.add('hidden');
}

function renderAdvConditions() {
  const cont = document.getElementById('advf-conditions');
  if (!cont) return;

  const headers = getColumnHeaders();

  cont.innerHTML = _advFilterDraft.map((cond, i) => {
    // Kolom options
    const colOpts = headers.map(h =>
      `<option value="${h.col}" ${h.col === cond.col ? 'selected' : ''}>${escapeHtml(h.label)}</option>`
    ).join('');

    // Operator options
    const opList = [
      { v: 'eq', l: '= sama dengan' },
      { v: 'neq', l: '≠ tidak sama' },
      { v: 'gt', l: '> lebih besar' },
      { v: 'lt', l: '< lebih kecil' },
      { v: 'gte', l: '≥ besar sama' },
      { v: 'lte', l: '≤ kecil sama' },
      { v: 'contains', l: 'mengandung' },
      { v: 'startswith', l: 'diawali' },
      { v: 'endswith', l: 'diakhiri' },
      { v: 'between', l: 'di antara' }
    ];
    const opOpts = opList.map(o =>
      `<option value="${o.v}" ${o.v === cond.op ? 'selected' : ''}>${escapeHtml(o.l)}</option>`
    ).join('');

    // Logic (hanya baris ke-2 ke atas)
    const logicHtml = i === 0
      ? `<div class="advf-logic">WHERE</div>`
      : `<select class="advf-logic-select" data-idx="${i}" style="background:#ecfdf5;color:#059669;font-weight:700;border-color:#d1fae5">
           <option value="AND" ${cond.logic === 'AND' ? 'selected' : ''}>AND</option>
           <option value="OR" ${cond.logic === 'OR' ? 'selected' : ''}>OR</option>
         </select>`;

    // Value — 1 atau 2 input (kalau between)
    let valueHtml;
    if (cond.op === 'between') {
      valueHtml = `<div style="display:flex;gap:4px;align-items:center">
        <input class="advf-value" data-idx="${i}" value="${escapeHtml(cond.value || '')}" placeholder="min" style="flex:1" />
        <span style="color:#94a3b8;font-size:11px">–</span>
        <input class="advf-value2" data-idx="${i}" value="${escapeHtml(cond.value2 || '')}" placeholder="max" style="flex:1" />
      </div>`;
    } else {
      valueHtml = `<input class="advf-value" data-idx="${i}" value="${escapeHtml(cond.value || '')}" placeholder="nilai" />`;
    }

    return `
      <div class="advf-row" data-idx="${i}">
        ${logicHtml}
        <select class="advf-col" data-idx="${i}">${colOpts}</select>
        <select class="advf-op" data-idx="${i}">${opOpts}</select>
        ${valueHtml}
        <button class="advf-del" data-idx="${i}" title="Hapus">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
    `;
  }).join('');

  // Bind events
  cont.querySelectorAll('.advf-col').forEach(el => {
    el.addEventListener('change', e => {
      const i = +e.target.dataset.idx;
      _advFilterDraft[i].col = +e.target.value;
    });
  });

  cont.querySelectorAll('.advf-op').forEach(el => {
    el.addEventListener('change', e => {
      const i = +e.target.dataset.idx;
      _advFilterDraft[i].op = e.target.value;
      renderAdvConditions();   // re-render karena between butuh 2 input
      updateAdvPreview();
    });
  });

  cont.querySelectorAll('.advf-value').forEach(el => {
    el.addEventListener('input', e => {
      const i = +e.target.dataset.idx;
      _advFilterDraft[i].value = e.target.value;
      updateAdvPreview();
    });
  });

  cont.querySelectorAll('.advf-value2').forEach(el => {
    el.addEventListener('input', e => {
      const i = +e.target.dataset.idx;
      _advFilterDraft[i].value2 = e.target.value;
      updateAdvPreview();
    });
  });

  cont.querySelectorAll('.advf-logic-select').forEach(el => {
    el.addEventListener('change', e => {
      const i = +e.target.dataset.idx;
      _advFilterDraft[i].logic = e.target.value;
      updateAdvPreview();
    });
  });

  cont.querySelectorAll('.advf-del').forEach(btn => {
    btn.addEventListener('click', e => {
      const i = +e.currentTarget.dataset.idx;
      _advFilterDraft.splice(i, 1);
      if (_advFilterDraft.length === 0) {
        _advFilterDraft.push({ col: 0, op: 'eq', value: '', logic: 'AND' });
      }
      renderAdvConditions();
      updateAdvPreview();
    });
  });
}

/* =========================================================
   MATCH CHECK — apakah row pass kondisi
   ========================================================= */
function checkAdvancedFilterMatch(row, conds) {
  if (!conds.length) return true;

  const sh = sheet();
  const results = [];

  for (const cond of conds) {
    const addr = addrOf(row, cond.col);
    const cl = sh.cells[addr];
    let val = '';

    if (cl) {
      if (cl.f && hf) {
        try { val = hfGet(row, cond.col); } catch { val = ''; }
      } else {
        val = cl.v ?? '';
      }
    }

    const strVal = String(val).trim();
    const numVal = Number(strVal);
    const isNum = !isNaN(numVal) && strVal !== '';
    const compareVal = String(cond.value ?? '').trim();
    const numCompare = Number(compareVal);

    let match = false;

    switch (cond.op) {
      case 'eq':
        if (isNum && !isNaN(numCompare)) match = numVal === numCompare;
        else match = strVal === compareVal;
        break;
      case 'neq':
        if (isNum && !isNaN(numCompare)) match = numVal !== numCompare;
        else match = strVal !== compareVal;
        break;
      case 'gt':
        match = isNum && !isNaN(numCompare) && numVal > numCompare;
        break;
      case 'lt':
        match = isNum && !isNaN(numCompare) && numVal < numCompare;
        break;
      case 'gte':
        match = isNum && !isNaN(numCompare) && numVal >= numCompare;
        break;
      case 'lte':
        match = isNum && !isNaN(numCompare) && numVal <= numCompare;
        break;
      case 'contains':
        match = strVal.toLowerCase().includes(compareVal.toLowerCase());
        break;
      case 'startswith':
        match = strVal.toLowerCase().startsWith(compareVal.toLowerCase());
        break;
      case 'endswith':
        match = strVal.toLowerCase().endsWith(compareVal.toLowerCase());
        break;
      case 'between': {
        const c2 = String(cond.value2 ?? '').trim();
        const n2 = Number(c2);
        match = isNum && !isNaN(n2) && numVal >= numCompare && numVal <= n2;
        break;
      }
    }

    results.push(match);
  }

  // Combine AND/OR
  // Baris pertama = dasar
  // Baris berikutnya pakai logic dari kondisi itu sendiri
  let result = results[0];
  for (let i = 1; i < conds.length; i++) {
    const logic = conds[i].logic || 'AND';
    if (logic === 'AND') result = result && results[i];
    else result = result || results[i];
  }
  return result;
}

/* =========================================================
   UPDATE PREVIEW
   ========================================================= */
function updateAdvPreview() {
  const preview = document.getElementById('advf-preview');
  if (!preview) return;

  // Cek draft valid
  const valid = _advFilterDraft.filter(c => c.value !== '' || c.op === 'between');
  if (!valid.length) {
    preview.classList.remove('show', 'error');
    return;
  }

  // Hitung jumlah baris yang match (scan up to 1000 baris untuk preview cepat)
  const sh = sheet();
  let maxR = 0;
  for (const addr of Object.keys(sh.cells || {})) {
    const p = parseAddr(addr);
    if (p && p.r > maxR) maxR = p.r;
  }

  const previewLimit = Math.min(maxR, 2000);
  let matched = 0;
  let totalChecked = 0;
  const headerRow = getFilterHeaderRow();

  for (let r = headerRow + 1; r <= previewLimit; r++) {
    totalChecked++;
    if (checkAdvancedFilterMatch(r, _advFilterDraft)) matched++;
  }

  preview.classList.add('show');
  preview.classList.remove('error');

  preview.innerHTML = `
    <i class="fa-solid fa-circle-info" style="font-size:14px"></i>
    <div>
      <strong>${matched.toLocaleString('id-ID')}</strong> dari
      <strong>${totalChecked.toLocaleString('id-ID')}</strong> baris match
      ${previewLimit < maxR ? `<span style="color:#64748b"> (preview 2.000 baris pertama)</span>` : ''}
    </div>
  `;
}

/* =========================================================
   LOCK CELL (per-cell)
   ========================================================= */
function isCellLocked(addr) {
  const cl = sheet().cells[addr];
  if (!cl || !cl.s) return false;
  return cl.s.locked === true;
}

function isCellUnlocked(addr) {
  const cl = sheet().cells[addr];
  if (!cl || !cl.s) return false;
  return cl.s.locked === false;
}

function setCellLocked(addr, locked) {
  const sh = sheet();
  if (!sh) return;
  if (!sh.cells) sh.cells = {};

  const existing = sh.cells[addr];

  if (locked === true) {
    // === LOCK ===
    // Buat cell kalau belum ada
    if (!existing) {
      sh.cells[addr] = { s: { locked: true } };
    } else {
      existing.s = existing.s || {};
      existing.s.locked = true;
    }
    markCellDirty(addr);
  } else {
    // === UNLOCK ===
    if (!existing) return;   // tidak ada cell, tidak perlu unlock

    if (existing.s) {
      delete existing.s.locked;

      // Kalau style kosong dan tidak ada value/formula, hapus cell
      const sKeys = Object.keys(existing.s);
      const hasValue = existing.v !== undefined && existing.v !== null && existing.v !== '';
      const hasFormula = !!existing.f;

      if (sKeys.length === 0 && !hasValue && !hasFormula) {
        delete sh.cells[addr];
        markCellDirty(addr);
        return;
      }
    }

    markCellDirty(addr);
  }
}

function lockSelection() {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  pushUndo();
  const parts = expandRange(selection.start, selection.end);
  for (const addr of parts) {
    setCellLocked(addr, true);
  }

  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  markMetaDirty('structure');
  scheduleSave();
  renderVirtual(false);

  toast(`🔒 ${parts.length} cell di-lock`);
}

function unlockSelection() {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  pushUndo();
  const parts = expandRange(selection.start, selection.end);
  for (const addr of parts) {
    setCellLocked(addr, false);
  }

  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  markMetaDirty('structure');
  scheduleSave();
  renderVirtual(false);

  toast(`🔓 ${parts.length} cell di-unlock`);
}

/* =========================================================
   APPLY VISUAL INDICATORS
   ========================================================= */
function applyLockIndicators() {
  // Clear
  document.querySelectorAll('#grid td.cell-locked, #grid td.cell-unlocked')
    .forEach(td => td.classList.remove('cell-locked', 'cell-unlocked'));

  document.querySelectorAll('#grid tbody td[data-addr]').forEach(td => {
    const addr = td.dataset.addr;
    const cl = sheet().cells[addr];
    if (cl?.s?.locked === true) {
      td.classList.add('cell-locked');
    }
  });
}

/* =========================================================
   WORKBOOK PASSWORD
   ========================================================= */

/* Simpan wbId yang menunggu unlock */
let _pendingUnlock = null;

/* =========================================================
   OPEN PASSWORD PROMPT
   ========================================================= */
function promptUnlockPassword(wbIdToUnlock, wbNameToShow) {
  return new Promise((resolve) => {
    const modal = document.getElementById('modal-unlock');
    if (!modal) return resolve(false);

    modal.classList.remove('hidden');
    const input = document.getElementById('unlock-password');
    const errEl = document.getElementById('unlock-error');

    input.value = '';
    errEl.textContent = '';

    setTimeout(() => input.focus(), 50);

    // Update judul dengan nama workbook
    const titleEl = modal.querySelector('.modal-header h3');
    if (titleEl && wbNameToShow) {
      titleEl.innerHTML = `<i class="fa-solid fa-shield-halved"></i> "${escapeHtml(wbNameToShow)}" Terkunci`;
    }

    _pendingUnlock = { wbId: wbIdToUnlock, resolve };

    // Handle Enter
    const onKeydown = (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        attemptUnlock();
      }
    };

    const onCancel = () => {
      cleanup();
      resolve(false);
    };

    const cleanup = () => {
      modal.classList.add('hidden');
      input.removeEventListener('keydown', onKeydown);
      document.getElementById('btn-unlock-apply').removeEventListener('click', attemptUnlock);
      document.getElementById('btn-unlock-cancel').removeEventListener('click', onCancel);
      _pendingUnlock = null;
    };

    const attemptUnlock = async () => {
      const pwd = input.value;
      if (!pwd) {
        errEl.textContent = '⚠ Password wajib diisi';
        return;
      }

      errEl.textContent = '';

      try {
        const res = await fetch(`/api/sheets/${wbIdToUnlock}/verify-password`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + token
          },
          body: JSON.stringify({ password: pwd })
        });

        const data = await res.json();

        if (res.ok && data.ok) {
          // Simpan password di session untuk re-check jika perlu
          sessionStorage.setItem(`wb-pass-${wbIdToUnlock}`, pwd);
          cleanup();
          resolve(true);
        } else {
          errEl.textContent = '⚠ ' + (data.error || 'Password salah');
          input.select();
        }
      } catch (e) {
        errEl.textContent = '⚠ Gagal verifikasi: ' + e.message;
      }
    };

    input.addEventListener('keydown', onKeydown);
    document.getElementById('btn-unlock-apply').addEventListener('click', attemptUnlock);
    document.getElementById('btn-unlock-cancel').addEventListener('click', onCancel);
  });
}

/* =========================================================
   SET / CHANGE PASSWORD
   ========================================================= */
function openSetPasswordModal() {
  const modal = document.getElementById('modal-set-password');
  if (!modal) return;

  // Cek apakah sudah ada password
  const hasPass = !!_wbHasPassword;

  const titleEl = document.getElementById('set-pass-title');
  const currentWrap = document.getElementById('current-pass-wrap');
  const removeBtn = document.getElementById('btn-set-pass-remove');

  titleEl.textContent = hasPass ? 'Ganti Password' : 'Set Password';
  currentWrap.style.display = hasPass ? '' : 'none';
  removeBtn.style.display = hasPass ? '' : 'none';

  // Reset form
  document.getElementById('current-password').value = '';
  document.getElementById('new-password').value = '';
  document.getElementById('confirm-password').value = '';
  document.getElementById('set-pass-error').textContent = '';

  modal.classList.remove('hidden');

  setTimeout(() => {
    const focusEl = hasPass
      ? document.getElementById('current-password')
      : document.getElementById('new-password');
    focusEl?.focus();
  }, 50);
}

function closeSetPasswordModal() {
  document.getElementById('modal-set-password')?.classList.add('hidden');
}

async function submitSetPassword(removeMode = false) {
  const errEl = document.getElementById('set-pass-error');
  errEl.textContent = '';

  const hasPass = !!_wbHasPassword;
  const currentPwd = document.getElementById('current-password').value;
  const newPwd = document.getElementById('new-password').value;
  const confirmPwd = document.getElementById('confirm-password').value;

  if (hasPass && !currentPwd) {
    errEl.textContent = '⚠ Password lama wajib diisi';
    return;
  }

  if (!removeMode) {
    if (!newPwd) {
      errEl.textContent = '⚠ Password baru wajib diisi';
      return;
    }
    if (newPwd.length < 4) {
      errEl.textContent = '⚠ Password minimal 4 karakter';
      return;
    }
    if (newPwd !== confirmPwd) {
      errEl.textContent = '⚠ Konfirmasi password tidak cocok';
      return;
    }
  }

  try {
    const body = {
      currentPassword: hasPass ? currentPwd : undefined,
      newPassword: removeMode ? null : newPwd
    };

    const res = await fetch(`/api/sheets/${wbId}/password`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token
      },
      body: JSON.stringify(body)
    });

    const data = await res.json();
    if (!res.ok) {
      errEl.textContent = '⚠ ' + (data.error || 'Gagal');
      return;
    }

    _wbHasPassword = !removeMode;
    closeSetPasswordModal();

    if (removeMode) {
      sessionStorage.removeItem(`wb-pass-${wbId}`);
      toast('✔ Password dihapus');
    } else {
      sessionStorage.setItem(`wb-pass-${wbId}`, newPwd);
      toast('✔ Password ' + (hasPass ? 'diganti' : 'diset'));
    }
  } catch (e) {
    errEl.textContent = '⚠ ' + e.message;
  }
}

/* =========================================================
   BIND
   ========================================================= */
document.getElementById('btn-set-pass-cancel')?.addEventListener('click', closeSetPasswordModal);
document.getElementById('btn-set-pass-apply')?.addEventListener('click', () => submitSetPassword(false));
document.getElementById('btn-set-pass-remove')?.addEventListener('click', async () => {
  if (!(await customConfirm('Hapus password workbook ini?', { type: 'danger', okText: 'Hapus' }))) return;
  submitSetPassword(true);
});

/* Enter di input */
['current-password', 'new-password', 'confirm-password'].forEach(id => {
  document.getElementById(id)?.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      document.getElementById('btn-set-pass-apply')?.click();
    }
    if (e.key === 'Escape') {
      closeSetPasswordModal();
    }
  });
});

console.log('✅ Workbook Password initialized');

/* =========================================================
   AUDIT LOG
   ========================================================= */
let _auditEntries = [];

function openAuditLog() {
  if (!wbId) return toast('Buka workbook dulu');

  const modal = document.getElementById('modal-audit');
  if (!modal) return;

  modal.classList.remove('hidden');

  // Reset filters
  document.getElementById('audit-filter-user').value = '';
  document.getElementById('audit-filter-addr').value = '';
  document.getElementById('audit-filter-action').value = '';

  loadAuditLog();
}

function closeAuditLog() {
  document.getElementById('modal-audit')?.classList.add('hidden');
}

async function loadAuditLog() {
  const list = document.getElementById('audit-list');
  const summary = document.getElementById('audit-summary');
  if (!list) return;

  list.innerHTML = '<div class="audit-empty"><i class="fa-solid fa-spinner fa-spin"></i>Memuat…</div>';

  const user = document.getElementById('audit-filter-user').value.trim();
  const addr = document.getElementById('audit-filter-addr').value.trim();
  const action = document.getElementById('audit-filter-action').value;

  const params = new URLSearchParams({ limit: '200' });
  if (user) params.set('user', user);
  if (addr) params.set('addr', addr);
  if (action) params.set('action', action);

  try {
    const res = await fetch(`/api/sheets/${wbId}/audit-log?${params}`, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);

    const data = await res.json();
    _auditEntries = data.entries || [];

    if (summary) {
      summary.textContent = `Menampilkan ${_auditEntries.length} dari ${data.total} entri`;
    }

    if (!_auditEntries.length) {
      list.innerHTML = `<div class="audit-empty">
        <i class="fa-solid fa-clipboard-list"></i>
        Belum ada riwayat perubahan
      </div>`;
      return;
    }

    renderAuditList();
  } catch (e) {
    list.innerHTML = `<div class="audit-empty" style="color:#dc2626">
      <i class="fa-solid fa-triangle-exclamation"></i>
      Gagal memuat: ${escapeHtml(e.message)}
    </div>`;
  }
}

function renderAuditList() {
  const list = document.getElementById('audit-list');
  if (!list) return;

  list.innerHTML = _auditEntries.map((e, idx) => {
    // Icon & label action
    const iconMap = {
      create: { icon: 'fa-plus', label: 'Dibuat' },
      edit:   { icon: 'fa-pen',  label: 'Diubah' },
      delete: { icon: 'fa-trash', label: 'Dihapus' }
    };
    const meta = iconMap[e.action] || iconMap.edit;

    // Value display
    const oldText = formatAuditValue(e.old_value);
    const newText = formatAuditValue(e.new_value);

    // Time display
    const time = new Date(e.created_at).toLocaleString('id-ID', {
      day: '2-digit', month: 'short', year: 'numeric',
      hour: '2-digit', minute: '2-digit'
    });

    // Sheet name
    const sheetName = wbData.sheets[e.sheet_idx]?.name || `Sheet${e.sheet_idx + 1}`;

    // Value diff
    let valuesHtml = '';
    if (e.action === 'create') {
      valuesHtml = `<div class="ae-values"><span class="ae-new">${escapeHtml(newText)}</span></div>`;
    } else if (e.action === 'delete') {
      valuesHtml = `<div class="ae-values"><span class="ae-old">${escapeHtml(oldText)}</span></div>`;
    } else {
      valuesHtml = `<div class="ae-values">
        <span class="ae-old">${escapeHtml(oldText)}</span>
        <span class="ae-arrow">→</span>
        <span class="ae-new">${escapeHtml(newText)}</span>
      </div>`;
    }

    return `
      <div class="audit-entry" data-idx="${idx}" data-addr="${escapeHtml(e.addr)}" data-sheet="${e.sheet_idx}">
        <div class="ae-icon ${e.action}">
          <i class="fa-solid ${meta.icon}"></i>
        </div>
        <div class="ae-body">
          <div class="ae-header">
            <span class="ae-addr">${escapeHtml(e.addr)}</span>
            <span style="font-size:11px;color:#94a3b8">·</span>
            <span class="ae-user">
              <i class="fa-solid fa-user" style="font-size:10px;margin-right:3px"></i>
              ${escapeHtml(e.username || 'Unknown')}
            </span>
            <span style="font-size:11px;color:#94a3b8">·</span>
            <span style="font-size:11px;color:#64748b">
              <i class="fa-solid fa-table" style="font-size:10px;margin-right:3px"></i>
              ${escapeHtml(sheetName)}
            </span>
            <span class="ae-time">${escapeHtml(time)}</span>
          </div>
          <div style="font-size:11px;color:#64748b;margin-bottom:2px">
            ${escapeHtml(meta.label)}
          </div>
          ${valuesHtml}
        </div>
      </div>
    `;
  }).join('');

  // Click handler — jump to cell
  list.querySelectorAll('.audit-entry').forEach(item => {
    item.addEventListener('click', async () => {
      const addr = item.dataset.addr;
      const sheetIdx = parseInt(item.dataset.sheet, 10);

      closeAuditLog();

      // Switch sheet kalau perlu
      if (sheetIdx !== activeSheetIdx) {
        await switchSheet(sheetIdx);
      }

      // Jump ke cell
      selectCell(addr);
      scrollAddrIntoView(addr);

      // Highlight sementara
      const td = findTd(addr);
      if (td) {
        td.style.transition = 'all .3s';
        td.style.boxShadow = '0 0 0 4px rgba(16,185,129,.5)';
        setTimeout(() => {
          td.style.boxShadow = '';
        }, 1500);
      }

      toast(`🎯 Ke ${addr}`);
    });
  });
}

/* Helper: format value untuk ditampilkan */
function formatAuditValue(val) {
  if (val === null || val === undefined) return '(kosong)';
  if (typeof val === 'object') {
    if (val.f) return val.f;
    if (val.v !== undefined) return String(val.v);
    return JSON.stringify(val);
  }
  return String(val);
}

/* =========================================================
   BIND
   ========================================================= */
document.getElementById('btn-audit-close')?.addEventListener('click', closeAuditLog);
document.getElementById('btn-audit-cancel')?.addEventListener('click', closeAuditLog);
document.getElementById('btn-audit-refresh')?.addEventListener('click', loadAuditLog);
document.getElementById('modal-audit')?.addEventListener('click', e => {
  if (e.target.id === 'modal-audit') closeAuditLog();
});

// Filter — debounce input
let _auditTimer = null;
['audit-filter-user', 'audit-filter-addr', 'audit-filter-action'].forEach(id => {
  document.getElementById(id)?.addEventListener('input', () => {
    clearTimeout(_auditTimer);
    _auditTimer = setTimeout(loadAuditLog, 300);
  });
  document.getElementById(id)?.addEventListener('change', () => {
    clearTimeout(_auditTimer);
    _auditTimer = setTimeout(loadAuditLog, 100);
  });
});

console.log('✅ Audit Log initialized');
document.getElementById('btn-audit')?.addEventListener('click', openAuditLog);

/* =========================================================
   PIVOT TABLE — Interactive
   ========================================================= */
let _pivotCounter = 0;

function getPivots() {
  const sh = sheet();
  if (!sh) return [];
  if (!Array.isArray(sh.pivots)) sh.pivots = [];
  return sh.pivots;
}

/* =========================================================
   CREATE PIVOT dari selection
   ========================================================= */
function createPivotFromSelection() {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return toast('Pilih range dulu');

  const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  if (r1 - r0 < 1) return toast('⚠ Minimal 2 baris (header + data)');

  const rangeStr = addrOf(r0, c0) + ':' + addrOf(r1, c1);

  const pivot = {
    id: 'pv' + Date.now() + '_' + Math.random().toString(36).slice(2, 5),
    name: `Pivot ${getPivots().length + 1}`,
    sourceSheetIdx: activeSheetIdx,
    sourceRange: rangeStr,
    rows: [c0],       // kolom pertama jadi baris default
    cols: [],
    values: [{ col: c1, agg: 'SUM' }],   // kolom terakhir jadi value default
    filters: [],
    x: 200 + (getPivots().length * 30),
    y: 150 + (getPivots().length * 30),
    w: 700,
    h: 450
  };

  getPivots().push(pivot);
  markMetaDirty('structure');
  scheduleSave();

  renderPivotWindow(pivot);
  toast(`✔ Pivot "${pivot.name}" dibuat`);
}

/* =========================================================
   RENDER SEMUA PIVOT untuk sheet aktif
   ========================================================= */
function renderAllPivots() {
  // Hapus DOM pivot yang lama
  document.querySelectorAll('.pivot-window').forEach(el => el.remove());

  const pivots = getPivots();
  for (const p of pivots) {
    // Skip kalau bukan untuk sheet ini (pivot yang source-nya beda sheet)
    if (p.sourceSheetIdx !== activeSheetIdx) continue;
    renderPivotWindow(p);
  }
}

/* =========================================================
   RENDER SATU PIVOT WINDOW
   ========================================================= */
function renderPivotWindow(pivot) {
  // Hapus window lama dengan id sama
  document.querySelector(`.pivot-window[data-pivot-id="${pivot.id}"]`)?.remove();

  const el = document.createElement('div');
  el.className = 'pivot-window';
  el.dataset.pivotId = pivot.id;
  el.style.left = pivot.x + 'px';
  el.style.top = pivot.y + 'px';
  el.style.width = pivot.w + 'px';
  el.style.height = pivot.h + 'px';

  el.innerHTML = `
    <div class="pivot-header">
      <span class="pivot-title">${escapeHtml(pivot.name)}</span>
      <div class="pivot-actions">
        <button class="pivot-btn pivot-refresh" title="Refresh"><i class="fa-solid fa-rotate"></i></button>
        <button class="pivot-btn pivot-source" title="Edit Source Range"><i class="fa-solid fa-crosshairs"></i></button>
        <button class="pivot-btn pivot-close" title="Tutup"><i class="fa-solid fa-xmark"></i></button>
      </div>
    </div>
    <div class="pivot-body">
      <div class="pivot-config">
        <div class="pivot-section">
          <div class="pivot-section-label"><i class="fa-solid fa-table-cells"></i> Fields</div>
          <div id="pivot-fields-${pivot.id}" class="pivot-fields-list"></div>
        </div>
      </div>
      <div class="pivot-result-wrap">
        <div id="pivot-result-${pivot.id}"></div>
      </div>
    </div>
    <div class="pivot-resize"></div>
  `;

  document.body.appendChild(el);

  // Render fields list & result
  renderPivotFields(pivot);
  renderPivotResult(pivot);

  // Drag
  makePivotDraggable(el, el.querySelector('.pivot-header'), pivot);
  makePivotResizable(el, el.querySelector('.pivot-resize'), pivot);

  // Button handlers
  el.querySelector('.pivot-close').addEventListener('click', () => {
    if (!customConfirm(`Hapus pivot "${pivot.name}"?`)) return;
    const pivots = getPivots();
    const idx = pivots.findIndex(x => x.id === pivot.id);
    if (idx >= 0) {
      pivots.splice(idx, 1);
      markMetaDirty('structure');
      scheduleSave();
    }
    el.remove();
    toast('✔ Pivot dihapus');
  });

  el.querySelector('.pivot-refresh').addEventListener('click', () => {
    renderPivotResult(pivot);
    toast('🔄 Pivot refresh');
  });

  el.querySelector('.pivot-source').addEventListener('click', () => {
    editPivotSource(pivot);
  });
}

/* =========================================================
   RENDER FIELDS — drop zones + field list
   ========================================================= */
function renderPivotFields(pivot) {
  const cont = document.getElementById(`pivot-fields-${pivot.id}`);
  if (!cont) return;

  // Ambil header fields dari source range
  const srcCells = wbData.sheets[pivot.sourceSheetIdx]?.cells || {};
  const [s, e] = rangeParts(pivot.sourceRange);
  const p1 = parseAddr(s), p2 = parseAddr(e);
  if (!p1 || !p2) return;

  const r0 = Math.min(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  // Header row = row pertama source
  const fields = [];
  for (let c = c0; c <= c1; c++) {
    const addr = addrOf(r0, c);
    const cl = srcCells[addr];
    const headerText = cl?.v ? String(cl.v) : colName(c);
    fields.push({ col: c, label: headerText, addr });
  }

  // Cek apakah ada di drop zone
  const isAssigned = (col) =>
    pivot.rows.includes(col) ||
    pivot.cols.includes(col) ||
    pivot.values.some(v => v.col === col);

  const availableFields = fields.filter(f => !isAssigned(f.col));

  // Build UI
  let html = '';

  // === ROWS ===
  html += `
    <div class="pivot-section">
      <div class="pivot-section-label"><i class="fa-solid fa-bars"></i> Rows</div>
      <div class="pivot-drop-zone" data-zone="rows" data-pivot="${pivot.id}">
        ${pivot.rows.map(c => {
          const f = fields.find(x => x.col === c);
          return `<div class="pivot-field-chip" draggable="true" data-col="${c}" data-zone="rows" data-pivot="${pivot.id}">
            ${escapeHtml(f?.label || colName(c))}
            <span class="chip-remove" data-col="${c}" data-zone="rows" data-pivot="${pivot.id}">×</span>
          </div>`;
        }).join('')}
      </div>
    </div>
  `;

  // === COLUMNS ===
  html += `
    <div class="pivot-section">
      <div class="pivot-section-label"><i class="fa-solid fa-table-columns"></i> Columns</div>
      <div class="pivot-drop-zone" data-zone="cols" data-pivot="${pivot.id}">
        ${pivot.cols.map(c => {
          const f = fields.find(x => x.col === c);
          return `<div class="pivot-field-chip" draggable="true" data-col="${c}" data-zone="cols" data-pivot="${pivot.id}">
            ${escapeHtml(f?.label || colName(c))}
            <span class="chip-remove" data-col="${c}" data-zone="cols" data-pivot="${pivot.id}">×</span>
          </div>`;
        }).join('')}
      </div>
    </div>
  `;

  // === VALUES ===
  html += `
    <div class="pivot-section">
      <div class="pivot-section-label"><i class="fa-solid fa-sigma"></i> Values</div>
      <div class="pivot-drop-zone" data-zone="values" data-pivot="${pivot.id}">
        ${pivot.values.map(v => {
          const f = fields.find(x => x.col === v.col);
          return `<div class="pivot-field-chip agg" draggable="true" data-col="${v.col}" data-zone="values" data-pivot="${pivot.id}">
            <select data-col="${v.col}" data-pivot="${pivot.id}">
              <option value="SUM" ${v.agg === 'SUM' ? 'selected' : ''}>SUM</option>
              <option value="COUNT" ${v.agg === 'COUNT' ? 'selected' : ''}>COUNT</option>
              <option value="AVERAGE" ${v.agg === 'AVERAGE' ? 'selected' : ''}>AVG</option>
              <option value="MIN" ${v.agg === 'MIN' ? 'selected' : ''}>MIN</option>
              <option value="MAX" ${v.agg === 'MAX' ? 'selected' : ''}>MAX</option>
            </select>
            ${escapeHtml(f?.label || colName(v.col))}
            <span class="chip-remove" data-col="${v.col}" data-zone="values" data-pivot="${pivot.id}">×</span>
          </div>`;
        }).join('')}
      </div>
    </div>
  `;

  // === AVAILABLE FIELDS ===
  html += `
    <div class="pivot-section">
      <div class="pivot-section-label"><i class="fa-solid fa-cube"></i> Available Fields</div>
      <div class="pivot-fields-list">
        ${availableFields.map(f => `
          <div class="pivot-field-item" draggable="true" data-col="${f.col}" data-label="${escapeHtml(f.label)}" data-pivot="${pivot.id}">
            ${escapeHtml(f.label)}
          </div>
        `).join('')}
      </div>
    </div>
  `;

  cont.innerHTML = html;

  // === BIND DRAG & DROP ===
  bindPivotDragDrop(pivot, cont);
}

/* =========================================================
   BIND DRAG DROP
   ========================================================= */
function bindPivotDragDrop(pivot, cont) {
  // Field chips — draggable
  cont.querySelectorAll('.pivot-field-chip[draggable]').forEach(chip => {
    chip.addEventListener('dragstart', e => {
      e.dataTransfer.setData('text/plain', JSON.stringify({
        col: parseInt(chip.dataset.col, 10),
        fromZone: chip.dataset.zone,
        pivotId: chip.dataset.pivot
      }));
      e.dataTransfer.effectAllowed = 'move';
    });
  });

  // Available fields — draggable
  cont.querySelectorAll('.pivot-field-item[draggable]').forEach(item => {
    item.addEventListener('dragstart', e => {
      e.dataTransfer.setData('text/plain', JSON.stringify({
        col: parseInt(item.dataset.col, 10),
        label: item.dataset.label,
        fromZone: null,
        pivotId: item.dataset.pivot
      }));
      e.dataTransfer.effectAllowed = 'move';
    });
  });

  // Drop zones
  cont.querySelectorAll('.pivot-drop-zone').forEach(zone => {
    zone.addEventListener('dragover', e => {
      e.preventDefault();
      zone.classList.add('drag-over');
    });
    zone.addEventListener('dragleave', () => {
      zone.classList.remove('drag-over');
    });
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('drag-over');
      try {
        const data = JSON.parse(e.dataTransfer.getData('text/plain'));
        if (data.pivotId !== pivot.id) return;
        assignPivotField(pivot, data.col, data.fromZone, zone.dataset.zone);
      } catch (err) {
        console.warn('Drop failed:', err);
      }
    });
  });

  // Chip remove buttons
  cont.querySelectorAll('.chip-remove').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const col = parseInt(btn.dataset.col, 10);
      const zone = btn.dataset.zone;
      removePivotField(pivot, col, zone);
    });
  });

  // Aggregate select
  cont.querySelectorAll('.pivot-field-chip.agg select').forEach(sel => {
    sel.addEventListener('change', e => {
      const col = parseInt(e.target.dataset.col, 10);
      const v = pivot.values.find(x => x.col === col);
      if (v) {
        v.agg = e.target.value;
        markMetaDirty('structure');
        scheduleSave();
        renderPivotResult(pivot);
        renderPivotFields(pivot);
      }
    });
  });
}

/* =========================================================
   ASSIGN / REMOVE FIELD
   ========================================================= */
function assignPivotField(pivot, col, fromZone, toZone) {
  // Hapus dari zone lama
  if (fromZone === 'rows') pivot.rows = pivot.rows.filter(c => c !== col);
  else if (fromZone === 'cols') pivot.cols = pivot.cols.filter(c => c !== col);
  else if (fromZone === 'values') pivot.values = pivot.values.filter(v => v.col !== col);

  // Tambah ke zone baru
  if (toZone === 'rows') {
    if (!pivot.rows.includes(col)) pivot.rows.push(col);
  } else if (toZone === 'cols') {
    if (!pivot.cols.includes(col)) pivot.cols.push(col);
  } else if (toZone === 'values') {
    if (!pivot.values.some(v => v.col === col)) {
      pivot.values.push({ col, agg: 'SUM' });
    }
  }

  markMetaDirty('structure');
  scheduleSave();
  renderPivotFields(pivot);
  renderPivotResult(pivot);
}

function removePivotField(pivot, col, zone) {
  if (zone === 'rows') pivot.rows = pivot.rows.filter(c => c !== col);
  else if (zone === 'cols') pivot.cols = pivot.cols.filter(c => c !== col);
  else if (zone === 'values') pivot.values = pivot.values.filter(v => v.col !== col);

  markMetaDirty('structure');
  scheduleSave();
  renderPivotFields(pivot);
  renderPivotResult(pivot);
}

/* =========================================================
   COMPUTE PIVOT
   ========================================================= */
function computePivot(pivot) {
  const srcSheet = wbData.sheets[pivot.sourceSheetIdx];
  if (!srcSheet) return { rows: [], cols: [], data: [] };

  const [s, e] = rangeParts(pivot.sourceRange);
  const p1 = parseAddr(s), p2 = parseAddr(e);
  if (!p1 || !p2) return { rows: [], cols: [], data: [] };

  const r0 = Math.min(p1.r, p2.r);
  const r1 = Math.max(p1.r, p2.r);
  const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);

  // Get cell value helper
  const getVal = (r, c) => {
    const addr = addrOf(r, c);
    const cl = srcSheet.cells[addr];
    if (!cl) return '';
    if (cl.f && pivot.sourceSheetIdx === activeSheetIdx) {
      try { return hfGet(r, c); } catch { return ''; }
    }
    return cl.v ?? '';
  };

  // Header row = r0
  const headers = {};
  for (let c = c0; c <= c1; c++) {
    headers[c] = String(getVal(r0, c) || colName(c));
  }

  // Collect data
  const groups = new Map();  // key → { rowKeys, colKeys, values }

  for (let r = r0 + 1; r <= r1; r++) {
    const rowKeyParts = pivot.rows.map(c => String(getVal(r, c) || '(blank)'));
    const colKeyParts = pivot.cols.map(c => String(getVal(r, c) || '(blank)'));
    const rowKey = rowKeyParts.join('\u0001');
    const colKey = colKeyParts.join('\u0001');
    const combinedKey = rowKey + '\u0002' + colKey;

    if (!groups.has(combinedKey)) {
      groups.set(combinedKey, {
        rowKeys: rowKeyParts,
        colKeys: colKeyParts,
        vals: pivot.values.map(() => [])
      });
    }

    const g = groups.get(combinedKey);
    pivot.values.forEach((v, vi) => {
      const raw = getVal(r, v.col);
      const num = Number(raw);
      if (v.agg === 'COUNT') {
        if (raw !== '' && raw !== null && raw !== undefined) g.vals[vi].push(1);
      } else {
        if (!isNaN(num)) g.vals[vi].push(num);
      }
    });
  }

  // Aggregate
  const aggFn = (arr, agg) => {
    if (!arr.length) return 0;
    switch (agg) {
      case 'SUM': return arr.reduce((a, b) => a + b, 0);
      case 'COUNT': return arr.length;
      case 'AVERAGE': return arr.reduce((a, b) => a + b, 0) / arr.length;
      case 'MIN': return Math.min(...arr);
      case 'MAX': return Math.max(...arr);
      default: return 0;
    }
  };

  // Extract unique row keys & col keys
  const uniqueRowKeys = new Map();  // rowKeyStr → rowKeyParts
  const uniqueColKeys = new Map();

  for (const [key, g] of groups) {
    const rk = g.rowKeys.join('\u0001');
    const ck = g.colKeys.join('\u0001');
    if (!uniqueRowKeys.has(rk)) uniqueRowKeys.set(rk, g.rowKeys);
    if (!uniqueColKeys.has(ck)) uniqueColKeys.set(ck, g.colKeys);
  }

  // Sort keys
  const sortFn = (a, b) => {
    const na = Number(a), nb = Number(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return String(a).localeCompare(String(b));
  };

  const rowKeysArr = [...uniqueRowKeys.values()].sort((a, b) => sortFn(a[0], b[0]));
  const colKeysArr = [...uniqueColKeys.values()].sort((a, b) => sortFn(a[0], b[0]));

  // Build result matrix
  const rowLabels = rowKeysArr.map(rk => rk.join(' / '));
  const colLabels = colKeysArr.map(ck => ck.join(' / '));

  const data = [];
  for (const rk of rowKeysArr) {
    const rowData = [];
    for (const ck of colKeysArr) {
      const combinedKey = rk.join('\u0001') + '\u0002' + ck.join('\u0001');
      const g = groups.get(combinedKey);
      if (g) {
        rowData.push(pivot.values.map((v, vi) => aggFn(g.vals[vi], v.agg)));
      } else {
        rowData.push(pivot.values.map(() => 0));
      }
    }
    data.push({ rowLabel: rk.join(' / '), values: rowData });
  }

  return {
    headers,
    rowKeys: rowKeysArr,
    colKeys: colKeysArr,
    rowLabels,
    colLabels,
    data,
    values: pivot.values
  };
}

/* =========================================================
   RENDER PIVOT RESULT
   ========================================================= */
function renderPivotResult(pivot) {
  const cont = document.getElementById(`pivot-result-${pivot.id}`);
  if (!cont) return;

  if (!pivot.rows.length && !pivot.cols.length) {
    cont.innerHTML = `<div class="pivot-empty">
      <i class="fa-solid fa-table-columns"></i>
      Tambahkan field ke <b>Rows</b> atau <b>Columns</b> dari panel kiri
    </div>`;
    return;
  }

  const result = computePivot(pivot);

  if (!result.data.length) {
    cont.innerHTML = `<div class="pivot-empty">
      <i class="fa-solid fa-table-columns"></i>
      Tidak ada data untuk ditampilkan
    </div>`;
    return;
  }

  // Build table
  const hasCols = pivot.cols.length > 0;

  let html = '<table class="pivot-result-table">';

  // Header
  html += '<thead><tr>';
  html += '<th>' + (pivot.rows.length ? result.headers[pivot.rows[0]] || 'Row' : 'Row') + '</th>';

  if (hasCols) {
    // Multiple columns headers
    for (const ck of result.colKeys) {
      html += `<th colspan="${pivot.values.length}">${escapeHtml(ck.join(' / ') || '(blank)')}</th>`;
    }
  } else {
    for (const v of pivot.values) {
      const h = result.headers[v.col] || colName(v.col);
      html += `<th>${escapeHtml(v.agg)} of ${escapeHtml(h)}</th>`;
    }
  }
  html += '</tr>';

  // Sub-header kalau ada cols
  if (hasCols && pivot.values.length > 1) {
    html += '<tr>';
    html += '<th></th>';
    for (let ci = 0; ci < result.colKeys.length; ci++) {
      for (const v of pivot.values) {
        const h = result.headers[v.col] || colName(v.col);
        html += `<th style="font-weight:500;font-size:11px">${escapeHtml(v.agg)} ${escapeHtml(h)}</th>`;
      }
    }
    html += '</tr>';
  } else if (hasCols && pivot.values.length === 1) {
    html += '<tr>';
    html += '<th></th>';
    for (let ci = 0; ci < result.colKeys.length; ci++) {
      const h = result.headers[pivot.values[0].col] || colName(pivot.values[0].col);
      html += `<th style="font-weight:500;font-size:11px">${escapeHtml(pivot.values[0].agg)} ${escapeHtml(h)}</th>`;
    }
    html += '</tr>';
  }
  html += '</thead>';

  // Body
  html += '<tbody>';

  for (const row of result.data) {
    html += '<tr>';
    html += `<td><b>${escapeHtml(row.rowLabel || '(blank)')}</b></td>`;

    for (const cellVals of row.values) {
      if (Array.isArray(cellVals)) {
        for (const v of cellVals) {
          html += `<td class="num">${formatPivotNumber(v)}</td>`;
        }
      } else {
        html += `<td class="num">${formatPivotNumber(cellVals)}</td>`;
      }
    }
    html += '</tr>';
  }

  // Grand total
  if (result.data.length > 1) {
    html += '<tr class="total-row">';
    html += '<td>Grand Total</td>';

    const totalCols = result.colKeys.length * pivot.values.length;
    for (let ci = 0; ci < totalCols; ci++) {
      let total = 0;
      for (const row of result.data) {
        const flat = row.values.flat();
        total += Number(flat[ci]) || 0;
      }
      html += `<td class="num">${formatPivotNumber(total)}</td>`;
    }
    html += '</tr>';
  }

  html += '</tbody></table>';
  cont.innerHTML = html;
}

function formatPivotNumber(n) {
  if (n === null || n === undefined) return '—';
  if (typeof n === 'number') {
    if (Number.isInteger(n)) return n.toLocaleString('id-ID');
    return n.toLocaleString('id-ID', { maximumFractionDigits: 2 });
  }
  return String(n);
}

/* =========================================================
   DRAG & RESIZE
   ========================================================= */
function makePivotDraggable(el, handle, model) {
  handle.addEventListener('mousedown', e => {
    if (e.target.closest('.pivot-btn')) return;
    const sx = e.clientX, sy = e.clientY, ox = model.x, oy = model.y;
    const mv = ev => {
      model.x = Math.max(0, ox + ev.clientX - sx);
      model.y = Math.max(0, oy + ev.clientY - sy);
      el.style.left = model.x + 'px';
      el.style.top = model.y + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      markMetaDirty('structure');
      scheduleSave();
    };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });
}

function makePivotResizable(el, handle, model) {
  handle.addEventListener('mousedown', e => {
    e.preventDefault();
    const sx = e.clientX, sy = e.clientY, ow = model.w, oh = model.h;
    const mv = ev => {
      model.w = Math.max(400, ow + ev.clientX - sx);
      model.h = Math.max(250, oh + ev.clientY - sy);
      el.style.width = model.w + 'px';
      el.style.height = model.h + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      markMetaDirty('structure');
      scheduleSave();
    };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });
}

/* =========================================================
   EDIT SOURCE
   ========================================================= */
async function editPivotSource(pivot) {
  const newRange = await customPrompt(
    'Ubah source range (contoh: A1:F100):',
    pivot.sourceRange,
    { title: 'Edit Source Range' }
  );
  if (!newRange || newRange === pivot.sourceRange) return;

  // Validasi
  const [s, e] = rangeParts(newRange);
  if (!parseAddr(s) || !parseAddr(e)) {
    return toast('⚠ Range tidak valid');
  }

  pivot.sourceRange = newRange.toUpperCase();
  markMetaDirty('structure');
  scheduleSave();
  renderPivotFields(pivot);
  renderPivotResult(pivot);
  toast('✔ Source diubah');
}

/* =========================================================
   MODAL: LIST PIVOTS
   ========================================================= */
function openPivotList() {
  if (!wbId) return toast('Buka workbook dulu');

  const modal = document.getElementById('modal-pivots');
  if (!modal) return;

  const list = document.getElementById('pivots-list');
  const pivots = getPivots();

  if (!pivots.length) {
    list.innerHTML = `<div style="padding:30px 20px;text-align:center;color:#94a3b8">
      <i class="fa-solid fa-table-columns" style="font-size:32px;color:#cbd5e1;display:block;margin-bottom:12px"></i>
      Belum ada pivot. Pilih range data lalu klik <b>Pivot Table</b>.
    </div>`;
  } else {
    list.innerHTML = pivots.map(p => `
      <div class="pivot-list-item" data-id="${p.id}">
        <div class="pl-icon"><i class="fa-solid fa-table-columns"></i></div>
        <div class="pl-body">
          <div class="pl-name">${escapeHtml(p.name)}</div>
          <div class="pl-meta">
            ${wbData.sheets[p.sourceSheetIdx]?.name || 'Sheet'}!${p.sourceRange} ·
            ${p.rows.length} rows · ${p.values.length} values
          </div>
        </div>
        <div class="pl-del" data-id="${p.id}" title="Hapus">
          <i class="fa-solid fa-trash"></i>
        </div>
      </div>
    `).join('');

    // Click to open
    list.querySelectorAll('.pivot-list-item').forEach(item => {
      item.addEventListener('click', e => {
        if (e.target.closest('.pl-del')) return;
        const id = item.dataset.id;
        const p = getPivots().find(x => x.id === id);
        if (p) {
          // Switch ke source sheet kalau perlu
          if (p.sourceSheetIdx !== activeSheetIdx) {
            switchSheet(p.sourceSheetIdx).then(() => renderPivotWindow(p));
          } else {
            renderPivotWindow(p);
          }
          modal.classList.add('hidden');
        }
      });
    });

    // Delete buttons
    list.querySelectorAll('.pl-del').forEach(btn => {
      btn.addEventListener('click', async e => {
        e.stopPropagation();
        const id = btn.dataset.id;
        if (!(await customConfirm('Hapus pivot ini?', { type: 'danger', okText: 'Hapus' }))) return;

        const pivots = getPivots();
        const idx = pivots.findIndex(x => x.id === id);
        if (idx >= 0) {
          pivots.splice(idx, 1);
          markMetaDirty('structure');
          scheduleSave();
        }
        document.querySelector(`.pivot-window[data-pivot-id="${id}"]`)?.remove();
        openPivotList();
      });
    });
  }

  modal.classList.remove('hidden');
}

function closePivotList() {
  document.getElementById('modal-pivots')?.classList.add('hidden');
}

/* =========================================================
   BIND
   ========================================================= */
document.getElementById('btn-pivot-new')?.addEventListener('click', createPivotFromSelection);
document.getElementById('btn-pivot-list')?.addEventListener('click', openPivotList);
document.getElementById('btn-pivots-close')?.addEventListener('click', closePivotList);
document.getElementById('btn-pivots-cancel')?.addEventListener('click', closePivotList);
document.getElementById('modal-pivots')?.addEventListener('click', e => {
  if (e.target.id === 'modal-pivots') closePivotList();
});

console.log('✅ Pivot Table initialized');

/* =========================================================
   IMPORT FROM URL
   ========================================================= */
let _iuPreviewData = null;   // data terakhir yang di-fetch

function openImportURLModal() {
  if (!wbId) return toast('Buka workbook dulu');

  const modal = document.getElementById('modal-import-url');
  if (!modal) return;

  // Reset
  document.getElementById('iu-url').value = '';
  document.getElementById('iu-format').value = 'auto';
  document.getElementById('iu-target').value = 'new';
  document.getElementById('iu-sheet-name').value = 'Data Import';
  document.getElementById('iu-header-row').value = '1';
  document.getElementById('iu-has-header').checked = true;
  document.getElementById('iu-error').textContent = '';
  document.getElementById('iu-preview-meta').textContent = '';
  document.getElementById('iu-preview').innerHTML = '<div class="iu-preview-empty">Klik <b>Fetch & Preview</b> untuk melihat data</div>';
  document.getElementById('btn-iu-import').disabled = true;
  _iuPreviewData = null;

  modal.classList.remove('hidden');
  setTimeout(() => document.getElementById('iu-url')?.focus(), 50);
}

function closeImportURLModal() {
  document.getElementById('modal-import-url')?.classList.add('hidden');
}

/* =========================================================
   FETCH URL
   ========================================================= */
async function fetchURLData() {
  const url = document.getElementById('iu-url').value.trim();
  const errEl = document.getElementById('iu-error');
  errEl.textContent = '';

  if (!url) {
    errEl.textContent = '⚠ URL wajib diisi';
    return;
  }

  const btn = document.getElementById('btn-iu-fetch');
  const origHTML = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Fetching…';

  try {
    const res = await fetch('/api/sheets/fetch-url', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token
      },
      body: JSON.stringify({ url })
    });

    const data = await res.json();
    if (!res.ok) {
      errEl.textContent = '⚠ ' + (data.error || 'Gagal fetch');
      return;
    }

    // Detect format
    const format = document.getElementById('iu-format').value;
    const detected = detectFormat(url, data.contentType, data.content, format);

    // Parse
    const parsed = parseImportData(data.content, detected);

    if (!parsed || !parsed.rows || !parsed.rows.length) {
      errEl.textContent = '⚠ Tidak ada data yang bisa di-parse';
      return;
    }

    _iuPreviewData = {
      url,
      format: detected,
      rows: parsed.rows,
      headers: parsed.headers,
      size: data.size
    };

    renderImportPreview();
    document.getElementById('btn-iu-import').disabled = false;
    document.getElementById('iu-preview-meta').textContent =
      `${parsed.rows.length} baris · ${parsed.headers.length} kolom · ${(data.size / 1024).toFixed(1)} KB`;

  } catch (e) {
    errEl.textContent = '⚠ ' + e.message;
  } finally {
    btn.disabled = false;
    btn.innerHTML = origHTML;
  }
}

function detectFormat(url, contentType, content, userChoice) {
  if (userChoice !== 'auto') return userChoice;

  const lowerUrl = url.toLowerCase();
  const lowerCT = (contentType || '').toLowerCase();

  if (lowerUrl.endsWith('.json') || lowerCT.includes('json')) return 'json';
  if (lowerUrl.endsWith('.tsv') || lowerUrl.endsWith('.tab')) return 'tsv';
  if (lowerUrl.endsWith('.csv')) return 'csv';

  // Content-based
  const trimmed = content.trim();
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) return 'json';

  // Cek tab vs comma
  const firstLine = trimmed.split('\n')[0] || '';
  if (firstLine.split('\t').length > firstLine.split(',').length) return 'tsv';

  return 'csv';
}

/* =========================================================
   PARSE DATA
   ========================================================= */
function parseImportData(content, format) {
  const hasHeader = document.getElementById('iu-has-header').checked;

  if (format === 'json') {
    return parseJSONData(content, hasHeader);
  }
  return parseDelimitedData(content, format === 'tsv' ? '\t' : ',', hasHeader);
}

function parseJSONData(content, hasHeader) {
  try {
    const obj = JSON.parse(content);
    let arr = null;

    // Cari array pertama
    if (Array.isArray(obj)) {
      arr = obj;
    } else if (obj && typeof obj === 'object') {
      // Cari key yang isinya array
      for (const key of Object.keys(obj)) {
        if (Array.isArray(obj[key]) && obj[key].length > 0) {
          arr = obj[key];
          break;
        }
      }
    }

    if (!arr || !arr.length) return null;

    // Array of primitives
    if (typeof arr[0] !== 'object') {
      return {
        headers: ['Value'],
        rows: arr.map(v => [String(v)])
      };
    }

    // Array of objects
    const allKeys = new Set();
    for (const item of arr) {
      if (item && typeof item === 'object') {
        Object.keys(item).forEach(k => allKeys.add(k));
      }
    }

    const keys = [...allKeys];
    const headers = hasHeader ? keys : keys;

    const rows = arr.map(item => keys.map(k => {
      const v = item?.[k];
      if (v === null || v === undefined) return '';
      if (typeof v === 'object') return JSON.stringify(v);
      return String(v);
    }));

    return { headers, rows };
  } catch (e) {
    console.warn('JSON parse failed:', e);
    return null;
  }
}

function parseDelimitedData(content, delimiter, hasHeader) {
  // Normalize line endings
  const text = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n+$/, '');
  const lines = text.split('\n');

  const rows = [];
  for (const line of lines) {
    const cells = [];
    let cur = '';
    let inQuote = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuote && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuote = !inQuote;
        }
      } else if (ch === delimiter && !inQuote) {
        cells.push(cur);
        cur = '';
      } else {
        cur += ch;
      }
    }
    cells.push(cur);
    rows.push(cells);
  }

  if (!rows.length) return null;

  let headers, dataRows;
  if (hasHeader) {
    headers = rows[0];
    dataRows = rows.slice(1);
  } else {
    const maxCols = Math.max(...rows.map(r => r.length));
    headers = Array.from({ length: maxCols }, (_, i) => `Col ${i + 1}`);
    dataRows = rows;
  }

  return { headers, rows: dataRows };
}

/* =========================================================
   RENDER PREVIEW
   ========================================================= */
function renderImportPreview() {
  const cont = document.getElementById('iu-preview');
  if (!cont || !_iuPreviewData) return;

  const { headers, rows } = _iuPreviewData;

  // Ambil 20 baris pertama untuk preview
  const preview = rows.slice(0, 20);

  let html = '<table>';
  html += '<thead><tr>';
  html += '<th style="width:40px;text-align:center">#</th>';
  headers.forEach(h => html += `<th>${escapeHtml(String(h))}</th>`);
  html += '</tr></thead><tbody>';

  preview.forEach((row, i) => {
    html += '<tr>';
    html += `<td style="text-align:center;color:#94a3b8;font-size:10.5px">${i + 1}</td>`;
    headers.forEach((_, ci) => {
      const v = row[ci] ?? '';
      const s = String(v);
      const isNum = s !== '' && !isNaN(Number(s));
      html += `<td class="${isNum ? 'num' : ''}">${escapeHtml(s)}</td>`;
    });
    html += '</tr>';
  });

  if (rows.length > 20) {
    html += `<tr><td colspan="${headers.length + 1}" style="text-align:center;color:#94a3b8;font-size:11px;padding:8px">
      … dan ${rows.length - 20} baris lainnya
    </td></tr>`;
  }

  html += '</tbody></table>';
  cont.innerHTML = html;
}

/* =========================================================
   IMPORT DATA
   ========================================================= */
async function importURLData() {
  console.log('🔵 importURLData called');

  if (!_iuPreviewData) {
    console.warn('❌ _iuPreviewData null');
    return toast('⚠ Fetch data dulu');
  }

  const { headers, rows } = _iuPreviewData;
  const target = document.getElementById('iu-target').value;
  const sheetName = document.getElementById('iu-sheet-name').value.trim() || 'Data Import';
  const hasHeader = document.getElementById('iu-has-header').checked;

  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  pushUndo();

  let targetSheetIdx;
  let startRow = 0;

  // =====================================================
  // 🎯 STEP 1: Buat sheet baru & SET ACTIVE LANGSUNG
  // (JANGAN pakai switchSheet — biar tidak trigger save internal)
  // =====================================================
  if (target === 'new') {
    // Cek duplikat nama
    let finalName = sheetName;
    let n = 1;
    while (wbData.sheets.some(s => s.name === finalName)) {
      n++;
      finalName = sheetName + ' ' + n;
    }

    const newSheet = {
      name: finalName,
      cells: {},
      colWidths: {}, rowHeight: 26,
      rowHeights: {},
      cf: [], charts: [], merged: [], filters: {},
      images: [], validations: [],
      hidden: false, tabColor: null, rowGroups: [],
      protected: false, protectHash: null,
      freezeRow: 0, freezeCol: 0,
      pivots: [], advancedFilters: []
    };

    wbData.sheets.push(newSheet);
    targetSheetIdx = wbData.sheets.length - 1;
    loadedSheetIdxs.add(targetSheetIdx);

    // 🎯 SET ACTIVE LANGSUNG — tanpa switchSheet
    activeSheetIdx = targetSheetIdx;
    wbData.activeSheet = targetSheetIdx;
    dirtyMeta.activeSheet = true;

    startRow = 0;
  } else {
    targetSheetIdx = activeSheetIdx;
    const sh = wbData.sheets[targetSheetIdx];
    let maxR = -1;
    for (const addr of Object.keys(sh.cells || {})) {
      const p = parseAddr(addr);
      if (p && p.r > maxR) maxR = p.r;
    }
    startRow = maxR >= 0 ? maxR + 2 : 0;
  }

  // =====================================================
  // 🎯 STEP 2: Tulis data ke sheet target
  // =====================================================
  const sh = wbData.sheets[targetSheetIdx];
  if (!sh.cells) sh.cells = {};

  let totalCells = 0;
  let currentRow = startRow;

  // Header
  if (hasHeader) {
    for (let c = 0; c < headers.length; c++) {
      const addr = addrOf(currentRow, c);
      sh.cells[addr] = {
        v: String(headers[c]),
        s: { bold: true, bg: '#d1fae5', color: '#065f46', align: 'center' }
      };
      markCellDirty(addr);
      totalCells++;
    }
    currentRow++;
  }

  // Data rows
  for (const row of rows) {
    for (let c = 0; c < headers.length; c++) {
      const addr = addrOf(currentRow, c);
      const raw = row[c] ?? '';
      const str = String(raw).trim();

      let cell = null;
      if (str !== '') {
        const num = Number(str);
        const v = !isNaN(num) && str !== '' ? num : str;
        cell = { v };
      }

      if (cell) {
        sh.cells[addr] = cell;
        markCellDirty(addr);
        totalCells++;
      }
    }
    currentRow++;
  }

  console.log(`📊 Imported ${totalCells} cells, activeSheetIdx=${activeSheetIdx}, targetSheetIdx=${targetSheetIdx}`);

  // =====================================================
  // 🎯 STEP 3: Rebuild HF dengan sheet baru
  // =====================================================
  try { await rebuildHFWithLoaded(); } catch (e) { console.warn(e); }
  await recalcRowsLimit(true);

  // =====================================================
  // 🎯 STEP 4: Reset state & render
  // =====================================================
  saveFilterStateToSheet();
  if (typeof loadFilterStateFromSheet === 'function') loadFilterStateFromSheet();

  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  renderSheetTabs();

  const gridWrapEl = document.getElementById('grid-wrap');
  if (gridWrapEl) gridWrapEl.scrollTop = 0;

  forceRefreshGrid();

  activeCell = 'A1';
  selection = { start: 'A1', end: 'A1' };
  updateSelectionUI();

  // =====================================================
  // 🎯 STEP 5: Save — paksa tunggu sampai selesai
  // =====================================================
  markMetaDirty('structure');

  // Tunggu save sebelumnya selesai (kalau ada)
  let tries = 0;
  while (saving && tries < 50) {
    await new Promise(r => setTimeout(r, 100));
    tries++;
  }

  try {
    await saveWB();
    console.log('✅ Save after import completed');
  } catch (e) {
    console.error('Save failed:', e);
    toast('⚠ Coba Ctrl+S manual');
  }

  // =====================================================
  // STEP 6: Auto-refresh config
  // =====================================================
  if (target === 'new') {
    try {
      addAutoRefreshConfig(targetSheetIdx, _iuPreviewData.url, {
        format: _iuPreviewData.format,
        hasHeader
      });
    } catch (e) { console.warn(e); }
  }

  closeImportURLModal();
  toast(`✔ Import ${totalCells.toLocaleString('id-ID')} sel dari URL`);
}

/* =========================================================
   AUTO-REFRESH CONFIG (simpel — pakai localStorage)
   ========================================================= */
const AUTO_REFRESH_KEY = 'wex-auto-refresh';

function getAutoRefreshConfigs() {
  try {
    const raw = localStorage.getItem(AUTO_REFRESH_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch { return {}; }
}

function saveAutoRefreshConfigs(configs) {
  try {
    localStorage.setItem(AUTO_REFRESH_KEY, JSON.stringify(configs));
  } catch (e) { console.warn(e); }
}

function addAutoRefreshConfig(sheetIdx, url, opts) {
  if (!wbId) return;
  const configs = getAutoRefreshConfigs();
  const key = `${wbId}:${sheetIdx}`;
  configs[key] = {
    url,
    format: opts.format,
    hasHeader: opts.hasHeader,
    addedAt: Date.now()
  };
  saveAutoRefreshConfigs(configs);
}

/* =========================================================
   BIND
   ========================================================= */
document.getElementById('btn-iu-close')?.addEventListener('click', closeImportURLModal);
document.getElementById('btn-iu-cancel')?.addEventListener('click', closeImportURLModal);
document.getElementById('btn-iu-fetch')?.addEventListener('click', fetchURLData);
document.addEventListener('click', e => {
  if (e.target.closest('#btn-iu-import')) {
    e.preventDefault();
    console.log('🔵 Import button clicked');
    importURLData();
  }
});

document.getElementById('modal-import-url')?.addEventListener('click', e => {
  if (e.target.id === 'modal-import-url') closeImportURLModal();
});

// Enter di URL → fetch
document.getElementById('iu-url')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    fetchURLData();
  }
});

console.log('✅ Import from URL initialized');
document.getElementById('btn-import-url')?.addEventListener('click', openImportURLModal);

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
      const v = temphfGet(r,c);
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
        copyToClipboard(k.key);
        toast('✔ Kunci dicopy');
      };
      row.querySelector('[data-rotate]').onclick = async () => {
        if (!(await customConfirm('Regenerate kunci ini? Kunci lama akan langsung tidak valid.', { type: 'warn', okText: 'Regenerate' }))) return;
        await fetch(`/api/keys/${k.id}/rotate`, {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token }
        });
        loadApiKeys();
        toast('✔ Kunci diregenerasi');
      };
      row.querySelector('[data-del]').onclick = async () => {
        if (!(await customConfirm('Hapus kunci ini?', { type: 'danger', okText: 'Hapus' }))) return;
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
  copyToClipboard(data.key).catch(() => {});
  toast('✔ Kunci dibuat & dicopy ke clipboard');
});

document.getElementById('btn-api-copy')?.addEventListener('click', () => {
  const txt = document.getElementById('api-sample')?.textContent || '';
  copyToClipboard(txt);
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

  try { input.focus(); } catch {}

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
   GLOBAL GUARD — cegah blur saat formula aktif
   Capture phase paling atas, cegah focus shift
   ========================================================= */
document.addEventListener('mousedown', e => {
  if (!editor) return;
  const input = editor.input;
  if (!input || !input.value.startsWith('=')) return;

  // Target adalah input editor sendiri → biarkan
  if (e.target === input) return;

  // 🎯 PENGECUALIAN: formula bar, input, dropdown, dialog
  if (e.target.closest('#formula-input')) return;    // ← TAMBAHKAN
  if (e.target.tagName === 'INPUT') return;           // ← TAMBAHKAN (semua input)
  if (e.target.tagName === 'TEXTAREA') return;        // ← TAMBAHKAN
  if (e.target.closest('#fx-suggest')) return;
  if (e.target.closest('.custom-dialog-overlay')) return;
  if (e.target.closest('.val-dropdown')) return;
  if (e.target.closest('.modal')) return;             // ← TAMBAHKAN (modal)
  if (e.target.closest('.sheet-tab')) return;         // ← sheet tab biarkan
  if (e.target.closest('button')) return;             // ← tombol biarkan

  // Sisanya (grid cell, dll) → preventDefault
  e.preventDefault();
}, true);

/* =========================================================
   INSERT REFERENCE + DRAG RANGE
   ========================================================= */
/* =========================================================
   REF DRAG — support editor normal + cross-sheet mode
   ========================================================= */
let refDrag = null;

document.getElementById('grid').addEventListener('mousedown', e => {
  // 🎯 Cek 2 kondisi: editor normal ATAU crossSheetRef aktif
  const isEditorMode = !!editor && editor.input?.value?.startsWith('=');
  const isCrossMode = !!crossSheetRef;

  if (!isEditorMode && !isCrossMode) return;

  const td = e.target.closest('td[data-addr]');
  if (!td) return;

  // Skip kalau klik di cell yang sama dengan editor
  if (isEditorMode && td === editor.input.closest('td')) return;

  e.preventDefault();
  e.stopPropagation();

  insertingRefMode = true;
  if (typeof hideFxSuggest === 'function') hideFxSuggest();

  const startAddr = td.dataset.addr;
  let curAddr = startAddr;
  const mode = isCrossMode ? 'CROSS' : 'EDITOR';

  console.log(`[RefDrag] START [${mode}]:`, startAddr);

  highlightRefRange(startAddr, startAddr);

  refDrag = { startAddr, curAddr, mode };

  // RAF loop untuk track cursor
  let lastX = e.clientX, lastY = e.clientY;
  let rafId = null, stopped = false;

  const trackCursor = () => {
    if (stopped) return;
    const el = document.elementFromPoint(lastX, lastY);
    const moveTd = el?.closest('td[data-addr]');
    if (moveTd) {
      const newAddr = moveTd.dataset.addr;
      if (newAddr !== curAddr) {
        curAddr = newAddr;
        refDrag.curAddr = curAddr;
        console.log(`[RefDrag] DRAG [${mode}]:`, startAddr, '→', curAddr);
        highlightRefRange(startAddr, curAddr);
      }
    }
    rafId = requestAnimationFrame(trackCursor);
  };
  rafId = requestAnimationFrame(trackCursor);

  const onMove = ev => { lastX = ev.clientX; lastY = ev.clientY; };

  const onUp = () => {
    stopped = true;
    if (rafId) cancelAnimationFrame(rafId);
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);

    const endAddr = refDrag?.curAddr || startAddr;
    const refText = (startAddr === endAddr) ? startAddr : `${startAddr}:${endAddr}`;

    console.log(`[RefDrag] END [${mode}]:`, refText);

        if (isCrossMode) {
      const rawName = wbData.sheets[activeSheetIdx].name;

      // 🎯 Auto-quote kalau perlu
      const needsQuote = /[\s\-\.\(\)]/.test(rawName) || /^[0-9]/.test(rawName);
      const sheetName = needsQuote
        ? "'" + rawName.replace(/'/g, "''") + "'"
        : rawName;

      const fullRef = (startAddr === endAddr)
        ? `${sheetName}!${startAddr}`
        : `${sheetName}!${startAddr}:${endAddr}`;

      insertCrossSheetRef2(fullRef);
    } else {
      insertRefToEditor(refText);
    }

    refDrag = null;
    setTimeout(clearRefHighlight, 800);
    setTimeout(() => { insertingRefMode = false; }, 200);
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}, true);

/* Clear highlight saat klik biasa (bukan insert ref) */
document.addEventListener('mousedown', e => {
  if (editor) return;   // jangan clear kalau sedang edit
  if (e.target.closest('#grid')) {
    clearRefHighlight();
  }
}, true);

function highlightRefRange(start, end) {
  document.querySelectorAll('#grid td.ref-highlight')
    .forEach(td => td.classList.remove('ref-highlight'));

  const addrs = expandRange(start, end);
  for (const addr of addrs) {
    const td = document.querySelector(`#grid td[data-addr="${addr}"]`);
    if (td) td.classList.add('ref-highlight');
  }
}

function clearRefHighlight() {
  document.querySelectorAll('#grid td.ref-highlight')
    .forEach(td => td.classList.remove('ref-highlight'));
}

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
        try { hfSet(r, c, raw); } catch {}
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

  // 🎯 Snapshot baris selection jadi header filter
  if (filterMode) {
    const p = parseAddr(selection.start);
    filterHeaderRow = p ? p.r : null;
  } else {
    filterHeaderRow = null;   // reset saat nonaktif
  }

  invalidateRowMap();
  window._lastRenderKey = null;
  renderVirtual(false);

  if (filterMode) {
    const hr = getFilterHeaderRow();
    toast(`🔽 Filter aktif di baris ${hr + 1}. Klik ▼ di cell kolom untuk filter.`);
  } else {
    toast('Filter dinonaktifkan');
  }
}, true);

document.getElementById('btn-filter-clear')?.addEventListener('click', () => {
  sheet().filters = {};
  sheet().advancedFilters = [];
  invalidateRowMap();
  saveFilterToLocal();
  window._lastRenderKey = null;
  window._lastFrozenKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  markMetaDirty('structure');
  scheduleSave();
  renderVirtual(false);
  renderFilterHeaders();
  toast('✔ Semua filter dibersihkan');
});

/* =========================================================
   RENDER FILTER HEADERS — single source of truth
   ========================================================= */
function renderFilterHeaders() {
  // const thead = document.querySelector('#grid thead tr');
  // if (!thead) return;

  // // 🎯 Hapus SEMUA tombol filter lama (cleanup)
  // thead.querySelectorAll('.fx-filter-btn').forEach(b => b.remove());
  // thead.querySelectorAll('th[data-col]').forEach(th => {
  //   th.classList.remove('filtered', 'filter-active');
  // });

  // // Ambil filter state
  // const filters = sheet().filters || {};

  // // 🎯 Buat tombol filter untuk SETIAP kolom dengan closure benar
  // thead.querySelectorAll('th[data-col]').forEach(th => {
  //   const c = +th.dataset.col;   // ← closure benar per kolom

  //   const btn = document.createElement('button');
  //   btn.className = 'fx-filter-btn';
  //   btn.innerHTML = '<i class="fa-solid fa-caret-down"></i>';
  //   btn.title = `Filter kolom ${colName(c)}`;

  //   // 🎯 Handler dengan kolom YANG BENAR
  //   btn.addEventListener('mousedown', ev => {
  //     ev.stopPropagation();
  //     ev.preventDefault();
  //     console.log(`🔽 Open filter popup for column ${colName(c)} (idx=${c})`);
  //     openFilterPopup(c, th);
  //   });

  //   th.appendChild(btn);

  //   // Highlight filter aktif
  //   if (filters[c]?.values?.length) {
  //     th.classList.add('filtered', 'filter-active');
  //   }
  // });
}

function openFilterPopup(colIdx, thEl) {
  console.log(`📋 openFilterPopup: colIdx=${colIdx} (${colName(colIdx)}), th=${thEl.dataset.col}`);
  closeFilterPopup();
  const rect = thEl.getBoundingClientRect();

  const values = new Map();
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (!p || p.c !== colIdx) continue;
    if (p.r === 0) continue;
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
    saveFilterToLocal();
    closeFilterPopup();
    window._lastRenderKey = null;
    renderVirtual(false);
    renderFilterHeaders();
    toast('✔ Filter kolom dibersihkan');
  };
  popup.querySelector('#fx-fp-apply').onclick = () => {
    sheet().filters ||= {};
    const list_ = sortedKeys;
    if (selected.size === list_.length) {
      delete sheet().filters[colIdx];
    } else {
      sheet().filters[colIdx] = { values: [...selected] };
    }
    invalidateRowMap();
    saveFilterToLocal();
    closeFilterPopup();
    window._lastRenderKey = null;
    renderVirtual(false);
    renderFilterHeaders();
    toast('✔ Filter diterapkan');
  };

  // Tutup kalau klik di luar
  setTimeout(() => {
    document.addEventListener('mousedown', outsideFilterHandler, true);
  }, 0);
}

/* =========================================================
   FILTER STATE — localStorage (bukan DB)
   ========================================================= */
function saveFilterToLocal() {
  if (!wbId) return;
  try {
    const key = `filter-${wbId}`;
    const filtersAll = {};
    wbData.sheets.forEach((sh, idx) => {
      filtersAll[idx] = sh.filters || {};
    });
    localStorage.setItem(key, JSON.stringify(filtersAll));
  } catch (e) {
    console.warn('saveFilterToLocal failed:', e);
  }
}

function loadFilterFromLocal(sheetId) {
  try {
    const raw = localStorage.getItem(`filter-${sheetId}`);
    if (!raw) return;
    const filtersAll = JSON.parse(raw);
    wbData.sheets.forEach((sh, idx) => {
      sh.filters = filtersAll[idx] || {};
    });
    console.log('📁 Filter loaded from localStorage');
  } catch (e) {
    console.warn('loadFilterFromLocal failed:', e);
  }
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
  if (window.Collab) window.Collab.disconnect();

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

  // 🎯 Reset state — JANGAN reset wbName
  wbId = null;
  // wbName tetap seperti terakhir kali, tidak di-reset ke 'Untitled'
  // karena openWB() nanti akan set wbName dari server

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
        if (!(await customConfirm(`Hapus workbook "${s.name}"?`, { type: 'danger', okText: 'Hapus' }))) return;
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
/* =========================================================
   ROW / COL HEADER SELECTION
   ========================================================= */
document.querySelector('#grid')?.addEventListener('mousedown', e => {
  // Header kolom
  const colTh = e.target.closest('thead th[data-col]');
  if (colTh) {
    if (e.target.closest('.col-resize-grip')) return;   // biarkan resize
    e.preventDefault();
    const c = +colTh.dataset.col;
    if (e.shiftKey && selectedCols.size) {
      const arr = [...selectedCols].sort((a,b)=>a-b);
      const from = Math.min(...arr, c), to = Math.max(...arr, c);
      selectedCols = new Set();
      for (let i = from; i <= to; i++) selectedCols.add(i);
    } else {
      selectedCols = new Set([c]);
    }
    selectedRows.clear();
    applyHeaderSelection();
    return;
  }

  // Header baris
  const rowTh = e.target.closest('tbody th[data-row]');
  if (rowTh) {
    if (e.target.closest('.row-resize-grip')) return;
    e.preventDefault();
    const r = +rowTh.dataset.row;
    if (e.shiftKey && selectedRows.size) {
      const arr = [...selectedRows].sort((a,b)=>a-b);
      const from = Math.min(...arr, r), to = Math.max(...arr, r);
      selectedRows = new Set();
      for (let i = from; i <= to; i++) selectedRows.add(i);
    } else {
      selectedRows = new Set([r]);
    }
    selectedCols.clear();
    applyHeaderSelection();
    return;
  }
});

function applyHeaderSelection() {
  document.querySelectorAll('#grid th.sel').forEach(th => th.classList.remove('sel'));
  selectedRows.forEach(r => {
    const tr = rowEls[r];
    if (tr?.firstElementChild) tr.firstElementChild.classList.add('sel');
  });
  selectedCols.forEach(c => {
    const th = document.querySelector(`#grid thead th[data-col="${c}"]`);
    if (th) th.classList.add('sel');
  });
}

function clearHeaderSelection() {
  if (!selectedRows.size && !selectedCols.size) return;
  selectedRows.clear();
  selectedCols.clear();
  applyHeaderSelection();
}
function deleteSelectedRowsAndCols() {
  const rows = [...selectedRows].sort((a,b)=>a-b);
  const cols = [...selectedCols].sort((a,b)=>a-b);

  // Hapus semua cell di baris terpilih
  for (const a of Object.keys(sheet().cells)) {
    const p = parseAddr(a);
    if (!p) continue;
    if (rows.includes(p.r) || cols.includes(p.c)) {
      delete sheet().cells[a];
      markCellDirty(a);
      try { hfSet(p.r, p.c, null); } catch {}
    }
  }

  // Hapus merge yang overlap
  if (rows.length || cols.length) {
    sheet().merged = (sheet().merged || []).filter(m => {
      const [s, e] = rangeParts(m);
      const p1 = parseAddr(s), p2 = parseAddr(e);
      if (!p1 || !p2) return true;
      const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
      const c0 = Math.min(p1.c, p2.c), c1 = Math.max(p1.c, p2.c);
      // buang kalau overlap dengan selection
      const overlapR = rows.some(r => r >= r0 && r <= r1);
      const overlapC = cols.some(c => c >= c0 && c <= c1);
      return !(overlapR || overlapC);
    });
  }

  selectedRows.clear();
  selectedCols.clear();
  markMetaDirty('structure');
  renderSheet();
  if (window.Collab && window.Collab.isConnected()) {
    const changes = {};
    // Cells yang dihapus sudah di-delete, jadi kirim null
    for (const r of rows) {
      for (let c = 0; c < COLS; c++) {
        changes[addrOf(r, c)] = null;
      }
    }
    for (const c of cols) {
      for (let r = 0; r < ROWS; r++) {
        changes[addrOf(r, c)] = null;
      }
    }
    window.Collab.sendEdit(changes, activeSheetIdx);
  }
  toast('✔ ' + (rows.length ? rows.length + ' baris' : '') +
        (cols.length ? (rows.length ? ' & ' : '') + cols.length + ' kolom' : '') + ' dihapus');
}
/* =========================================================
   FILE MENU
   ========================================================= */
const fileTabBtn = document.getElementById('tab-file');
const fileDropdown = document.getElementById('file-dropdown');

fileTabBtn?.addEventListener('click', e => {
  e.stopPropagation();
  const hidden = fileDropdown.classList.contains('hidden');
  if (hidden) {
    openFileMenu();
  } else {
    fileDropdown.classList.add('hidden');
  }
});

document.addEventListener('mousedown', e => {
  if (!fileDropdown || fileDropdown.classList.contains('hidden')) return;
  if (e.target.closest('#file-dropdown') || e.target.closest('#tab-file')) return;
  fileDropdown.classList.add('hidden');
});

function openFileMenu() {
  const nameEl = document.getElementById('fd-wb-name');
  const metaEl = document.getElementById('fd-wb-meta');
  nameEl.textContent = wbName || 'Untitled';

  if (wbData?.sheets) {
    const totalCells = wbData.sheets.reduce((s, sh) => s + Object.keys(sh.cells || {}).length, 0);
    metaEl.textContent = `${wbData.sheets.length} sheet · ${totalCells.toLocaleString('id-ID')} sel`;
  }

  fileDropdown.classList.remove('hidden');
}

document.querySelectorAll('#file-dropdown .fd-item').forEach(btn => {
  btn.addEventListener('click', async () => {
    fileDropdown.classList.add('hidden');
    const action = btn.dataset.action;
    switch (action) {
      case 'new-wb': newWB(); break;
      case 'open': listWBs(); break;
      case 'home': showHome(); break;
      case 'save': await saveWB(true); toast('✔ Tersimpan'); break;
      case 'save-as': saveAsCopy(); break;
      case 'info': showInfoModal(); break;
      case 'print': openPrintModal(); break;
      case 'export-xlsx': exportFile('xlsx'); break;
      case 'export-csv': exportFile('csv'); break;
      case 'export-pdf': exportFile('pdf'); break;
      case 'logout': document.getElementById('btn-logout')?.click(); break;
      case 'optimize': optimizeWorkbook(); break;
      case 'set-password': openSetPasswordModal(); break;
    }
  });
});

function exportFile(kind) {
  if (!wbId) return toast('Simpan dulu');
  saveWB(true).then(() => {
    fetch(`/api/sheets/${wbId}/export/${kind}`, {
      headers: { 'Authorization': 'Bearer ' + token }
    })
      .then(r => r.blob())
      .then(blob => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = wbName + '.' + kind; a.click();
        URL.revokeObjectURL(url);
      });
  });
}

async function saveAsCopy() {
  if (!wbId) return toast('Simpan dulu');
  const newName = prompt('Nama salinan:', wbName + ' - Copy');
  if (!newName) return;
  await saveWB(true);
  const res = await fetch('/api/sheets', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    body: JSON.stringify({ name: newName, data: wbData })
  });
  if (!res.ok) return toast('Gagal menyimpan salinan');
  const wb = await res.json();
  await openWB(wb.id);
  toast('✔ Salinan dibuat');
}

function showInfoModal() {
  const body = document.getElementById('info-body');
  if (!body) return;

  const totalCells = (wbData.sheets || []).reduce((s, sh) => s + Object.keys(sh.cells || {}).length, 0);
  const totalMerges = (wbData.sheets || []).reduce((s, sh) => s + (sh.merged?.length || 0), 0);
  const totalCharts = (wbData.sheets || []).reduce((s, sh) => s + (sh.charts?.length || 0), 0);
  const dataSize = (JSON.stringify(wbData).length / 1024).toFixed(1);

  body.innerHTML = `
    <div class="info-card"><div class="label">Nama</div><div class="value">${escapeHtml(wbName)}</div></div>
    <div class="info-card"><div class="label">Jumlah Sheet</div><div class="value">${wbData.sheets.length}</div></div>
    <div class="info-card"><div class="label">Total Sel</div><div class="value">${totalCells.toLocaleString('id-ID')}</div></div>
    <div class="info-card"><div class="label">Ukuran Data</div><div class="value">${dataSize} KB</div></div>
    <div class="info-card"><div class="label">Merge</div><div class="value">${totalMerges}</div></div>
    <div class="info-card"><div class="label">Chart</div><div class="value">${totalCharts}</div></div>
  `;
  document.getElementById('modal-info').classList.remove('hidden');
}
document.getElementById('btn-info-close')?.addEventListener('click', () =>
  document.getElementById('modal-info').classList.add('hidden'));
/* =========================================================
   HYPERLINK
   ========================================================= */
document.getElementById('btn-insert-link')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  const cur = sheet().cells[activeCell];
  document.getElementById('link-text').value = cur?.link?.text || displayVal(cur?.v || '');
  document.getElementById('link-url').value = cur?.link?.url || '';
  document.getElementById('modal-link').classList.remove('hidden');
  setTimeout(() => document.getElementById('link-url')?.focus(), 50);
});

document.getElementById('btn-link-cancel')?.addEventListener('click', () =>
  document.getElementById('modal-link').classList.add('hidden'));

function normalizeUrl(url) {
  url = url.trim();
  if (!url) return '';
  if (url.startsWith('#')) return url;
  if (/^[a-z][a-z0-9+.-]*:/i.test(url)) return url;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(url)) return 'mailto:' + url;
  return 'https://' + url;
}

document.getElementById('btn-link-apply')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  const text = document.getElementById('link-text').value.trim();
  const rawUrl = document.getElementById('link-url').value.trim();
  if (!rawUrl) return toast('URL wajib diisi');
  const url = normalizeUrl(rawUrl);

  pushUndo();
  ensureCell(activeCell);
  const cell = sheet().cells[activeCell];
  if (text) cell.v = text;
  cell.link = { url, text: text || url };

  // Styling underline biru
  cell.s = { ...(cell.s || {}), color: '#2563eb', underline: true };

  markCellDirty(activeCell);
  document.getElementById('modal-link').classList.add('hidden');
  renderRow(parseAddr(activeCell).r);
  renderVisibleCells();
  toast('✔ Hyperlink ditambahkan');
});

document.getElementById('btn-link-remove')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  pushUndo();
  const cell = sheet().cells[activeCell];
  if (cell) {
    delete cell.link;
    if (cell.s) {
      delete cell.s.color;
      delete cell.s.underline;
    }
  }
  markCellDirty(activeCell);
  document.getElementById('modal-link').classList.add('hidden');
  renderVisibleCells();
  toast('✔ Link dihapus');
});

/* Klik link untuk buka */
document.getElementById('grid')?.addEventListener('click', e => {
  const inner = e.target.closest('.cell-inner');
  if (!inner) return;
  const td = inner.closest('td[data-addr]');
  if (!td) return;
  const cell = sheet().cells[td.dataset.addr];
  if (!cell?.link) return;
  // Buka dengan Ctrl+Click atau selalu buka kalau mode view
  if (e.ctrlKey || e.metaKey || currentAccessMode === 'view') {
    e.preventDefault();
    e.stopPropagation();
    const url = normalizeUrl(cell.link.url);
    if (url.startsWith('#')) {
      // Internal link: #Sheet2!A1
      const m = /^#([^!]+)!(.+)$/.exec(url);
      if (m) {
        const sheetName = m[1], addr = m[2];
        const idx = wbData.sheets.findIndex(s => s.name === sheetName);
        if (idx >= 0) {
          switchSheet(idx).then(() => selectCell(addr));
        }
      }
    } else {
      window.open(url, '_blank', 'noopener');
    }
  }
});

/* Ctrl+K shortcut */
document.addEventListener('keydown', e => {
  if (editor) return;
  const ctrl = e.ctrlKey || e.metaKey;
  if (ctrl && e.key.toLowerCase() === 'k') {
    e.preventDefault();
    document.getElementById('btn-insert-link')?.click();
  }
});
/* =========================================================
   INSERT GAMBAR
   ========================================================= */
document.getElementById('btn-insert-image')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  document.getElementById('image-upload').click();
});

document.getElementById('image-upload')?.addEventListener('change', async e => {
  const f = e.target.files[0];
  if (!f) return;
  e.target.value = '';

  if (f.size > 5 * 1024 * 1024) {
    return toast('⚠ Gambar maksimal 5 MB');
  }

  const reader = new FileReader();
  reader.onload = () => {
    const dataUrl = reader.result;
    // Ukuran default: 240 x 160
    const id = 'img_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
    const pos = parseAddr(activeCell) || { r: 0, c: 0 };

    // Hitung posisi pixel relatif terhadap grid
    const tr = rowEls[pos.r];
    const td = tr?.children[pos.c + 1];
    const rect = td?.getBoundingClientRect();
    const wrapRect = document.getElementById('grid-wrap').getBoundingClientRect();

    const x = rect ? (rect.left - wrapRect.left + gridWrap.scrollLeft) : 100;
    const y = rect ? (rect.top - wrapRect.top + gridWrap.scrollTop) : 100;

    sheet().images ||= [];
    sheet().images.push({
      id, data: dataUrl,
      x, y, w: 240, h: 160
    });
    markMetaDirty('structure');
    renderImages();
    toast('✔ Gambar ditambahkan');
  };
  reader.readAsDataURL(f);
});

function renderImages() {
  document.querySelectorAll('.sheet-image').forEach(el => el.remove());
  const wrap = document.getElementById('grid-wrap');
  if (!wrap) return;

  const images = sheet().images || [];
  const zBase = 200;

  for (const img of images) {
    const el = document.createElement('div');
    el.className = 'sheet-image';
    el.dataset.imgId = img.id;
    el.style.left = img.x + 'px';
    el.style.top = img.y + 'px';
    el.style.width = img.w + 'px';
    el.style.height = img.h + 'px';
    el.style.zIndex = zBase;
    el.innerHTML = `
      <img src="${img.data}" alt="image" draggable="false" />
      <div class="img-toolbar">
        <button data-act="front" title="Ke depan"><i class="fa-solid fa-layer-group"></i></button>
        <button data-act="down" title="Unduh"><i class="fa-solid fa-download"></i></button>
        <button class="danger" data-act="del" title="Hapus"><i class="fa-solid fa-trash"></i></button>
      </div>
      <div class="img-resize"></div>
    `;
    wrap.appendChild(el);

    makeImageDraggable(el, img);
    makeImageResizable(el.querySelector('.img-resize'), el, img);
    bindImageToolbar(el, img);
  }
}

function makeImageDraggable(el, img) {
  el.addEventListener('mousedown', e => {
    if (e.target.closest('.img-toolbar') || e.target.closest('.img-resize')) return;
    e.preventDefault();
    e.stopPropagation();

    // Select
    document.querySelectorAll('.sheet-image.selected').forEach(x => x.classList.remove('selected'));
    el.classList.add('selected');

    const sx = e.clientX, sy = e.clientY, ox = img.x, oy = img.y;
    const move = ev => {
      img.x = ox + ev.clientX - sx;
      img.y = oy + ev.clientY - sy;
      el.style.left = img.x + 'px';
      el.style.top = img.y + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      markMetaDirty('structure');
      scheduleSave();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  });
}

function makeImageResizable(handle, el, img) {
  handle.addEventListener('mousedown', e => {
    e.preventDefault();
    e.stopPropagation();
    const sx = e.clientX, sy = e.clientY, ow = img.w, oh = img.h;
    const move = ev => {
      img.w = Math.max(60, ow + ev.clientX - sx);
      img.h = Math.max(40, oh + ev.clientY - sy);
      el.style.width = img.w + 'px';
      el.style.height = img.h + 'px';
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      markMetaDirty('structure');
      scheduleSave();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  });
}

function bindImageToolbar(el, img) {
  el.querySelector('[data-act="front"]')?.addEventListener('click', e => {
    e.stopPropagation();
    // pindahkan ke akhir array (paling depan)
    const arr = sheet().images || [];
    const idx = arr.findIndex(x => x.id === img.id);
    if (idx >= 0) {
      arr.splice(idx, 1);
      arr.push(img);
      markMetaDirty('structure');
      renderImages();
    }
  });
  el.querySelector('[data-act="down"]')?.addEventListener('click', e => {
    e.stopPropagation();
    const a = document.createElement('a');
    a.href = img.data;
    a.download = 'image.png';
    a.click();
  });
  el.querySelector('[data-act="del"]')?.addEventListener('click', async e => {
    e.stopPropagation();
    if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
    if (!(await customConfirm('Hapus gambar ini?', { type: 'danger', okText: 'Hapus' }))) return;
    sheet().images = (sheet().images || []).filter(x => x.id !== img.id);
    markMetaDirty('structure');
    renderImages();
    toast('✔ Gambar dihapus');
  });
}

/* Klik di luar gambar → deselect */
document.addEventListener('mousedown', e => {
  if (!e.target.closest('.sheet-image')) {
    document.querySelectorAll('.sheet-image.selected').forEach(x => x.classList.remove('selected'));
  }
});
/* =========================================================
   VIEW LINKS (Copy Link)
   ========================================================= */
async function renderViewLinks() {
  const box = document.getElementById('view-link-list');
  if (!box) return;
  box.innerHTML = '<div class="muted" style="padding:8px;text-align:center">Memuat…</div>';

  try {
    const res = await fetch(`/api/sheets/${wbId}/view-links`, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    const list = await res.json();
    box.innerHTML = '';
    if (!list.length) {
      box.innerHTML = '<div class="muted" style="padding:8px;text-align:center">Belum ada link</div>';
      return;
    }
    for (const l of list) {
      const row = document.createElement('div');
      row.className = 'view-link-row';
      const url = `${location.origin}/view.html?token=${l.token}`;
      row.innerHTML = `
        <span class="token" title="${escapeHtml(url)}">${escapeHtml(url)}</span>
        <button data-copy>Copy</button>
        <button data-open>Buka</button>
        <button class="danger" data-del>Hapus</button>
      `;
      row.querySelector('[data-copy]').onclick = () => {
        copyToClipboard(url);
        toast('✔ Link dicopy');
      };
      row.querySelector('[data-open]').onclick = () => {
        window.open(url, '_blank');
      };
      row.querySelector('[data-del]').onclick = async () => {
        if (!(await customConfirm('Hapus link ini?', { type: 'danger', okText: 'Hapus' }))) return;
        await fetch(`/api/sheets/${wbId}/view-links/${l.id}`, {
          method: 'DELETE',
          headers: { 'Authorization': 'Bearer ' + token }
        });
        renderViewLinks();
      };
      box.appendChild(row);
    }
  } catch {
    box.innerHTML = '<div class="muted" style="padding:8px">Gagal memuat</div>';
  }
}

document.getElementById('btn-create-view-link')?.addEventListener('click', async () => {
  if (!wbId) return;

  const res = await fetch(`/api/sheets/${wbId}/view-links`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    body: JSON.stringify({ label: 'View Link' })
  });

  if (!res.ok) return toast('Gagal buat link');

  const data = await res.json();
  const url = `${location.origin}/view.html?token=${data.token}`;

  const ok = await copyToClipboard(url);
  toast(ok ? '✔ Link dibuat & dicopy' : '✔ Link dibuat — copy dari list di bawah');

  renderViewLinks();
});

/* =========================================================
   CLIPBOARD HELPER — bekerja di HTTP & HTTPS
   ========================================================= */
async function copyToClipboard(text) {
  if (!text) return false;

  // 1. Modern API (butuh HTTPS / localhost)
  if (navigator.clipboard && copyToClipboard) {
    try {
      await copyToClipboard(text);
      return true;
    } catch (e) {
      console.warn('Clipboard API failed, fallback...', e);
    }
  }

  // 2. Fallback: textarea + execCommand
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.style.top = '0';
    ta.style.opacity = '0';
    ta.setAttribute('readonly', '');
    document.body.appendChild(ta);

    ta.select();
    ta.setSelectionRange(0, ta.value.length);
    const ok = document.execCommand('copy');

    document.body.removeChild(ta);
    if (ok) return true;
  } catch (e) {
    console.warn('execCommand failed:', e);
  }

  // 3. Fallback terakhir: prompt manual
  try {
    window.prompt('Salin manual (Ctrl+C):', text);
    return true;
  } catch {
    return false;
  }
}

/* Panggil saat share modal dibuka */
const _origRenderShareList = renderShareList;
renderShareList = async function() {
  await _origRenderShareList();
  await renderViewLinks();
};
/* =========================================================
   PERMISSION CUSTOM SELECT
   ========================================================= */
(function initPermissionSelect() {
  const wrap = document.getElementById('share-perm-select');
  const trigger = document.getElementById('share-perm-trigger');
  const menu = document.getElementById('share-perm-menu');
  const hidden = document.getElementById('share-permission');
  if (!wrap || !trigger || !menu || !hidden) return;

  const iconMap = {
    view:  'fa-solid fa-eye',
    edit:  'fa-solid fa-pen',
    admin: 'fa-solid fa-gear'
  };
  const labelMap = {
    view:  'Hanya lihat',
    edit:  'Bisa edit',
    admin: 'Admin'
  };

  function openMenu() {
    wrap.classList.add('open');
    menu.classList.remove('hidden');
    document.addEventListener('mousedown', outsideHandler, true);
  }

  function closeMenu() {
    wrap.classList.remove('open');
    menu.classList.add('hidden');
    document.removeEventListener('mousedown', outsideHandler, true);
  }

  function outsideHandler(e) {
    if (!wrap.contains(e.target)) closeMenu();
  }

  function setValue(val) {
    hidden.value = val;
    // Update trigger
    trigger.querySelector('i:first-child').className = iconMap[val];
    trigger.querySelector('.perm-label').textContent = labelMap[val];
    // Update active
    menu.querySelectorAll('.perm-option').forEach(o => {
      o.classList.toggle('active', o.dataset.value === val);
    });
  }

  trigger.addEventListener('click', e => {
    e.stopPropagation();
    if (menu.classList.contains('hidden')) openMenu();
    else closeMenu();
  });

  menu.querySelectorAll('.perm-option').forEach(opt => {
    opt.addEventListener('click', () => {
      setValue(opt.dataset.value);
      closeMenu();
    });
  });

  // Reset saat modal share dibuka
  const modal = document.getElementById('modal-share');
  if (modal) {
    const obs = new MutationObserver(() => {
      if (!modal.classList.contains('hidden')) {
        setValue('view');
        closeMenu();
      }
    });
    obs.observe(modal, { attributes: true, attributeFilter: ['class'] });
  }
})();
/* =========================================================
   FIND & REPLACE
   ========================================================= */
let findState = {
  open: false,
  matches: [],           // [{sheetIdx, addr, preview}]
  currentIdx: -1,
  query: '',
  replaceWith: '',
  options: {
    matchCase: false,
    wholeCell: false,
    inFormulas: false,
    allSheets: false,
    scope: 'sheet'
  }
};

function openFindModal(mode = 'find') {
  const modal = document.getElementById('modal-find');
  if (!modal) return;
  modal.classList.remove('hidden');
  findState.open = true;
  document.getElementById('replace-row').style.display = mode === 'replace' ? 'flex' : 'none';
  setTimeout(() => document.getElementById('find-input').focus(), 50);
  // Preload query dari selection kalau ada
  const selAddr = activeCell;
  const selCell = sheet().cells[selAddr];
  if (selCell?.v && !document.getElementById('find-input').value) {
    document.getElementById('find-input').value = String(selCell.v).slice(0, 100);
  }
  doFind();
}

function closeFindModal() {
  document.getElementById('modal-find').classList.add('hidden');
  findState.open = false;
  clearFindHighlights();
}

function clearFindHighlights() {
  document.querySelectorAll('#grid td.find-match, #grid td.find-current')
    .forEach(td => td.classList.remove('find-match', 'find-current'));
}

/* ============ SEARCH CORE ============ */
function matchesQuery(text, query, opts) {
  if (text === null || text === undefined) text = '';
  let t = String(text);
  let q = query;
  if (!opts.matchCase) { t = t.toLowerCase(); q = q.toLowerCase(); }
  if (opts.wholeCell) return t === q;
  return t.indexOf(q) !== -1;
}

function doFind() {
  const query = document.getElementById('find-input').value.trim();
  findState.query = query;
  findState.options = {
    matchCase: document.getElementById('find-match-case').checked,
    wholeCell: document.getElementById('find-whole-cell').checked,
    inFormulas: document.getElementById('find-in-formulas').checked,
    allSheets: document.getElementById('find-all-sheets').checked,
    scope: document.querySelector('input[name="find-scope"]:checked')?.value || 'sheet'
  };

  findState.matches = [];
  findState.currentIdx = -1;
  clearFindHighlights();

  if (!query) {
    renderFindResults();
    return;
  }

  // Tentukan sheet-scope
  let sheetIdxs = [activeSheetIdx];
  if (findState.options.allSheets || findState.options.scope === 'workbook') {
    sheetIdxs = wbData.sheets.map((_, i) => i);
  }

  // Range untuk scope 'selection'
  let rangeFilter = null;
  if (findState.options.scope === 'selection' && selection.start && selection.end) {
    const p1 = parseAddr(selection.start), p2 = parseAddr(selection.end);
    if (p1 && p2) {
      rangeFilter = {
        r0: Math.min(p1.r, p2.r), r1: Math.max(p1.r, p2.r),
        c0: Math.min(p1.c, p2.c), c1: Math.max(p1.c, p2.c)
      };
    }
  }

  // Loop setiap sheet
  for (const sIdx of sheetIdxs) {
    const sh = wbData.sheets[sIdx];
    if (!sh?.cells) continue;
    for (const [addr, cell] of Object.entries(sh.cells)) {
      const p = parseAddr(addr);
      if (!p) continue;

      // Filter range
      if (rangeFilter && sIdx === activeSheetIdx) {
        if (p.r < rangeFilter.r0 || p.r > rangeFilter.r1 ||
            p.c < rangeFilter.c0 || p.c > rangeFilter.c1) continue;
      }

      // Cek value atau formula
      let text;
      if (findState.options.inFormulas && cell.f) {
        text = cell.f;
      } else {
        // Nilai terhitung kalau ada formula, kalau tidak pakai raw
        if (sIdx === activeSheetIdx && cell.f && hf) {
          try { text = hfGet(p.r,p.c); }
          catch { text = cell.v; }
        } else {
          text = cell.v;
        }
      }

      if (matchesQuery(text, query, findState.options)) {
        findState.matches.push({
          sheetIdx: sIdx,
          addr,
          preview: String(text ?? ''),
          raw: text
        });
      }
    }
  }

  if (findState.matches.length) {
    findState.currentIdx = 0;
    gotoMatch(0);
  }
  renderFindResults();
}

function gotoMatch(idx) {
  if (idx < 0 || idx >= findState.matches.length) return;
  const m = findState.matches[idx];

  // Pindah sheet kalau perlu
  if (m.sheetIdx !== activeSheetIdx) {
    switchSheet(m.sheetIdx).then(() => {
      selectCell(m.addr);
      highlightFindMatch(m.addr, true);
    });
  } else {
    selectCell(m.addr);
    highlightFindMatch(m.addr, true);
  }
}

function highlightFindMatch(addr, isCurrent = false) {
  clearFindHighlights();
  // Highlight semua match di sheet aktif
  for (const m of findState.matches) {
    if (m.sheetIdx !== activeSheetIdx) continue;
    const td = findAddrEl(m.addr);
    if (td) td.classList.add('find-match');
  }
  // Current dalam warna beda
  if (isCurrent) {
    const td = findAddrEl(addr);
    if (td) {
      td.classList.remove('find-match');
      td.classList.add('find-current');
    }
  }
}

function findAddrEl(addr) {
  // Cari langsung di DOM
  let td = document.querySelector(`#grid td[data-addr="${addr}"]`);
  if (td) return td;
  // Coba scroll dulu (untuk virtual scroll)
  scrollAddrIntoView(addr);
  return document.querySelector(`#grid td[data-addr="${addr}"]`);
}

function renderFindResults() {
  const box = document.getElementById('find-results');
  if (!box) return;
  if (!findState.query) {
    box.innerHTML = '';
    return;
  }
  if (!findState.matches.length) {
    box.innerHTML = '<div style="padding:12px;text-align:center;color:#94a3b8">Tidak ditemukan</div>';
    return;
  }

  const cnt = findState.matches.length;
  const cur = findState.currentIdx + 1;
  const query = findState.query;
  const opts = findState.options;

  let html = `<div class="find-result-count">
    <i class="fa-solid fa-circle-check"></i>
    Ditemukan ${cnt} hasil — ${cur} dari ${cnt}
  </div>`;

  findState.matches.slice(0, 100).forEach((m, idx) => {
    // Highlight query dalam preview
    let preview = escapeHtml(m.preview.slice(0, 80));
    let q = query;
    if (!opts.matchCase) {
      const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      preview = escapeHtml(m.preview.slice(0, 80)).replace(
        re, match => `<mark>${match}</mark>`
      );
    } else {
      preview = preview.split(escapeHtml(q)).join(`<mark>${escapeHtml(q)}</mark>`);
    }

    const sheetBadge = m.sheetIdx !== activeSheetIdx
      ? `<span class="sheet-badge">${escapeHtml(wbData.sheets[m.sheetIdx].name)}</span>`
      : '';

    html += `
      <div class="find-result-item" data-idx="${idx}">
        <span class="addr">${escapeHtml(m.addr)}</span>
        ${sheetBadge}
        <span class="preview">${preview}</span>
      </div>
    `;
  });

  if (cnt > 100) {
    html += `<div style="padding:8px;text-align:center;color:#94a3b8;font-size:11px">
      … dan ${cnt - 100} hasil lainnya
    </div>`;
  }

  box.innerHTML = html;
  box.querySelectorAll('.find-result-item').forEach(item => {
    item.addEventListener('click', () => {
      const idx = +item.dataset.idx;
      findState.currentIdx = idx;
      gotoMatch(idx);
      renderFindResults();
    });
  });
}

/* ============ NAVIGASI ============ */
function findNext() {
  if (!findState.matches.length) return;
  findState.currentIdx = (findState.currentIdx + 1) % findState.matches.length;
  gotoMatch(findState.currentIdx);
  renderFindResults();
}
function findPrev() {
  if (!findState.matches.length) return;
  findState.currentIdx = (findState.currentIdx - 1 + findState.matches.length) % findState.matches.length;
  gotoMatch(findState.currentIdx);
  renderFindResults();
}

/* ============ REPLACE ============ */
function replaceOne() {
  if (!findState.matches.length) return;
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  const m = findState.matches[findState.currentIdx];
  if (!m) return;

  const replaceWith = document.getElementById('replace-input').value;
  const sIdx = m.sheetIdx;
  const sh = wbData.sheets[sIdx];

  const cell = sh.cells[m.addr];
  if (!cell) return;

  pushUndo();

  // Replace di value atau formula
  if (findState.options.inFormulas && cell.f) {
    const regex = buildReplaceRegex(findState.query, findState.options);
    cell.f = cell.f.replace(regex, replaceWith);
  } else {
    const v = String(cell.v ?? '');
    const regex = buildReplaceRegex(findState.query, findState.options);
    const newV = v.replace(regex, replaceWith);
    // Auto convert ke number kalau numeric
    const n = Number(newV);
    cell.v = (!isNaN(n) && newV.trim() !== '') ? n : newV;
  }

  // Update HF
  if (sIdx === activeSheetIdx) {
    const p = parseAddr(m.addr);
    try {
      hfSet(p.r, p.c, cell.f || cell.v);
    } catch {}
  }

  markCellDirty(m.addr);
  renderVisibleCells();
  toast('✔ Diganti');

  // Refresh matches
  doFind();
}

function replaceAll() {
  if (!findState.matches.length) return;
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  const replaceWith = document.getElementById('replace-input').value;
  const count = findState.matches.length;
  if (!customConfirm(`Ganti ${count} hasil dengan "${replaceWith}"?`)) return;

  pushUndo();
  const regex = buildReplaceRegex(findState.query, findState.options);

  // Group by sheet untuk efisiensi
  const bySheet = new Map();
  for (const m of findState.matches) {
    if (!bySheet.has(m.sheetIdx)) bySheet.set(m.sheetIdx, []);
    bySheet.get(m.sheetIdx).push(m);
  }

  let replaced = 0;
  for (const [sIdx, matches] of bySheet.entries()) {
    const sh = wbData.sheets[sIdx];
    for (const m of matches) {
      const cell = sh.cells[m.addr];
      if (!cell) continue;

      if (findState.options.inFormulas && cell.f) {
        cell.f = cell.f.replace(regex, replaceWith);
      } else {
        const v = String(cell.v ?? '');
        const newV = v.replace(regex, replaceWith);
        const n = Number(newV);
        cell.v = (!isNaN(n) && newV.trim() !== '') ? n : newV;
      }
      if (sIdx === activeSheetIdx) {
        const p = parseAddr(m.addr);
        try { hfSet(p.r, p.c, cell.f || cell.v); } catch {}
      }
      markCellDirty(m.addr);
      replaced++;
    }
  }

  renderVisibleCells();
  toast(`✔ ${replaced} sel diganti`);
  doFind();
}

function buildReplaceRegex(query, opts) {
  const flags = opts.matchCase ? 'g' : 'gi';
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (opts.wholeCell) {
    return new RegExp(`^${escaped}$`, flags);
  }
  return new RegExp(escaped, flags);
}

/* ============ BINDINGS ============ */
document.getElementById('btn-find')?.addEventListener('click', () => openFindModal('find'));
document.getElementById('btn-replace')?.addEventListener('click', () => openFindModal('replace'));
document.getElementById('btn-find-close')?.addEventListener('click', closeFindModal);
document.getElementById('find-next')?.addEventListener('click', findNext);
document.getElementById('find-prev')?.addEventListener('click', findPrev);
document.getElementById('btn-replace-one')?.addEventListener('click', replaceOne);
document.getElementById('btn-replace-all')?.addEventListener('click', replaceAll);

/* Enter untuk next, Shift+Enter untuk prev */
document.getElementById('find-input')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    e.preventDefault();
    if (e.shiftKey) findPrev(); else findNext();
  } else if (e.key === 'Escape') {
    closeFindModal();
  }
});

/* Live search saat ketik (debounce 250ms) */
let findTimer = null;
document.getElementById('find-input')?.addEventListener('input', () => {
  clearTimeout(findTimer);
  findTimer = setTimeout(doFind, 250);
});

/* Re-run kalau opsi berubah */
['find-match-case', 'find-whole-cell', 'find-in-formulas', 'find-all-sheets'].forEach(id => {
  document.getElementById(id)?.addEventListener('change', doFind);
});
document.querySelectorAll('input[name="find-scope"]').forEach(r => {
  r.addEventListener('change', doFind);
});

/* Klik backdrop → close */
document.getElementById('modal-find')?.addEventListener('click', e => {
  if (e.target.id === 'modal-find') closeFindModal();
});

/* ============ KEYBOARD SHORTCUTS ============ */
document.addEventListener('keydown', e => {
  if (editor) return;
  const ctrl = e.ctrlKey || e.metaKey;
  if (!ctrl) return;

  if (e.key.toLowerCase() === 'f') {
    e.preventDefault();
    openFindModal('find');
  } else if (e.key.toLowerCase() === 'h') {
    e.preventDefault();
    openFindModal('replace');
  }
}, true);
/* =========================================================
   GO TO CELL (Ctrl+G)
   ========================================================= */
let gotoCellsHistory = [];

function openGoToModal() {
  const modal = document.getElementById('modal-goto');
  if (!modal) return;
  modal.classList.remove('hidden');
  const input = document.getElementById('goto-input');
  input.value = '';
  setTimeout(() => input.focus(), 50);
  renderGotoHistory();
}

function closeGoToModal() {
  document.getElementById('modal-goto')?.classList.add('hidden');
}

function renderGotoHistory() {
  const box = document.getElementById('goto-history');
  if (!box) return;
  box.innerHTML = '';
  gotoCellsHistory.slice(-8).reverse().forEach(addr => {
    const chip = document.createElement('span');
    chip.className = 'goto-history-item';
    chip.textContent = addr;
    chip.addEventListener('click', () => {
      document.getElementById('goto-input').value = addr;
      applyGoTo();
    });
    box.appendChild(chip);
  });
}

function applyGoTo() {
  const raw = document.getElementById('goto-input').value.trim();
  if (!raw) return;

  // Format: A1, A1:C10, Sheet2!A1, Sheet2!A1:C10
  let sheetName = null;
  let refPart = raw;

  if (raw.includes('!')) {
    const [sn, ref] = raw.split('!');
    sheetName = sn.trim();
    refPart = ref.trim();
  }

  // Cari sheet
  let targetSheetIdx = activeSheetIdx;
  if (sheetName) {
    const clean = sheetName.replace(/^'|'$/g, '');
    const idx = wbData.sheets.findIndex(s =>
      s.name.toLowerCase() === clean.toLowerCase()
    );
    if (idx < 0) {
      toast('⚠ Sheet "' + clean + '" tidak ditemukan');
      return;
    }
    targetSheetIdx = idx;
  }

  // Parse range
  const parts = refPart.split(':');
  const startAddr = parts[0].toUpperCase();
  const endAddr = (parts[1] || parts[0]).toUpperCase();

  if (!parseAddr(startAddr) || !parseAddr(endAddr)) {
    toast('⚠ Alamat tidak valid');
    return;
  }

  // Simpan history
  gotoCellsHistory.push(raw);
  if (gotoCellsHistory.length > 20) gotoCellsHistory.shift();

  closeGoToModal();

  // Pindah sheet dulu kalau perlu
  const jump = () => {
    selection = { start: startAddr, end: endAddr };
    activeCell = startAddr;
    scrollAddrIntoView(startAddr);
    updateSelectionUI();
    toast('→ ' + raw);
  };

  if (targetSheetIdx !== activeSheetIdx) {
    switchSheet(targetSheetIdx).then(jump);
  } else {
    jump();
  }
}

document.getElementById('btn-goto')?.addEventListener('click', openGoToModal);
document.getElementById('btn-goto-cancel')?.addEventListener('click', closeGoToModal);
document.getElementById('btn-goto-apply')?.addEventListener('click', applyGoTo);
document.getElementById('goto-input')?.addEventListener('keydown', e => {
  if (e.key === 'Enter') { e.preventDefault(); applyGoTo(); }
  else if (e.key === 'Escape') closeGoToModal();
});
document.getElementById('modal-goto')?.addEventListener('click', e => {
  if (e.target.id === 'modal-goto') closeGoToModal();
});

/* Shortcut Ctrl+G */
document.addEventListener('keydown', e => {
  if (editor) return;
  const ctrl = e.ctrlKey || e.metaKey;
  if (ctrl && e.key.toLowerCase() === 'g') {
    e.preventDefault();
    openGoToModal();
  }
}, true);

/* =========================================================
   SORT MULTI-KOLOM
   ========================================================= */
function openSortModal() {
  const modal = document.getElementById('modal-sort');
  if (!modal) return;
  modal.classList.remove('hidden');
  document.getElementById('sort-keys').innerHTML = '';
  addSortKey();
  document.getElementById('sort-has-header').checked = true;
}

function closeSortModal() {
  document.getElementById('modal-sort')?.classList.add('hidden');
}

function getUsedColumnRange() {
  let maxC = 0;
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (p && p.c > maxC) maxC = p.c;
  }
  return Math.min(maxC + 1, COLS);
}

function addSortKey() {
  const box = document.getElementById('sort-keys');
  if (!box) return;
  if (box.children.length >= 5) return toast('Maksimal 5 kolom');

  const usedCols = getUsedColumnRange();
  const row = document.createElement('div');
  row.className = 'sort-key-row';

  let opts = '';
  for (let c = 0; c < usedCols; c++) {
    // Cari header name
    const headerCell = sheet().cells[addrOf(0, c)];
    const headerText = headerCell?.v ? String(headerCell.v) : '';
    const label = headerText
      ? `${colName(c)} — ${headerText.slice(0, 20)}`
      : colName(c);
    opts += `<option value="${c}">${escapeHtml(label)}</option>`;
  }

  row.innerHTML = `
    <select class="sort-col">${opts}</select>
    <select class="sort-dir">
      <option value="asc">A → Z / 1 → 9</option>
      <option value="desc">Z → A / 9 → 1</option>
    </select>
    <button class="sort-key-del" title="Hapus">
      <i class="fa-solid fa-xmark"></i>
    </button>
  `;
  row.querySelector('.sort-key-del').addEventListener('click', () => {
    row.remove();
    if (!box.children.length) addSortKey();
  });
  box.appendChild(row);
}

document.getElementById('btn-sort-multi')?.addEventListener('click', openSortModal);
document.getElementById('btn-sort-add')?.addEventListener('click', addSortKey);
document.getElementById('btn-sort-cancel')?.addEventListener('click', closeSortModal);

document.getElementById('btn-sort-apply')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  // Kumpulkan sort keys
  const keys = [];
  document.querySelectorAll('#sort-keys .sort-key-row').forEach(row => {
    const c = +row.querySelector('.sort-col').value;
    const dir = row.querySelector('.sort-dir').value;
    keys.push({ c, dir });
  });
  if (!keys.length) return toast('Tambahkan minimal 1 kolom');

  const hasHeader = document.getElementById('sort-has-header').checked;
  applyMultiSort(keys, hasHeader);
  closeSortModal();
});

function applyMultiSort(keys, hasHeader) {
  // Tentukan range data
  let maxR = -1, maxC = -1;
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (!p) continue;
    if (p.r > maxR) maxR = p.r;
    if (p.c > maxC) maxC = p.c;
  }
  if (maxR < 0) return toast('Tidak ada data');

  const startRow = hasHeader ? 1 : 0;

  // Kumpulkan baris
  const rows = [];
  for (let r = startRow; r <= maxR; r++) {
    const cells = {};
    let hasData = false;
    for (let c = 0; c <= maxC; c++) {
      const a = addrOf(r, c);
      if (sheet().cells[a]) {
        cells[c] = JSON.parse(JSON.stringify(sheet().cells[a]));
        hasData = true;
      }
    }
    if (hasData) rows.push({ r, cells });
  }

  if (rows.length < 2) return toast('Data tidak cukup untuk disortir');

  // Ambil nilai sort untuk tiap baris
  const sorted = rows.map(item => {
    const vals = keys.map(k => {
      let v;
      const cell = item.cells[k.c];
      if (cell?.f) {
        try { v = hfGet(item.r, k.c); }
        catch { v = cell.v ?? ''; }
      } else {
        v = cell?.v ?? '';
      }
      return v;
    });
    return { item, vals };
  });

  // Sort stable by priority
  sorted.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const dir = keys[i].dir === 'desc' ? -1 : 1;
      const va = a.vals[i], vb = b.vals[i];
      const na = Number(va), nb = Number(vb);
      let cmp = 0;
      if (!isNaN(na) && !isNaN(nb) && va !== '' && vb !== '') {
        cmp = na - nb;
      } else {
        cmp = String(va ?? '').localeCompare(String(vb ?? ''), 'id');
      }
      if (cmp !== 0) return dir * cmp;
    }
    return 0;
  });

  pushUndo();

  // Hapus baris lama (dari startRow)
  for (let r = startRow; r <= maxR; r++) {
    for (let c = 0; c <= maxC; c++) {
      delete sheet().cells[addrOf(r, c)];
    }
  }

  // Tulis ulang
  sorted.forEach((s, i) => {
    const targetR = startRow + i;
    for (const [c, cell] of Object.entries(s.item.cells)) {
      sheet().cells[addrOf(targetR, +c)] = cell;
    }
  });

  resetHF();
  renderSheet();
  markDirty();
  toast(`✔ Sortir ${sorted.length} baris (${keys.length} kolom)`);
}

/* =========================================================
   REMOVE DUPLICATES
   ========================================================= */
function openDupModal() {
  const modal = document.getElementById('modal-dup');
  if (!modal) return;
  modal.classList.remove('hidden');

  const box = document.getElementById('dup-columns');
  box.innerHTML = '';
  const usedCols = getUsedColumnRange();

  for (let c = 0; c < usedCols; c++) {
    const headerCell = sheet().cells[addrOf(0, c)];
    const headerText = headerCell?.v ? String(headerCell.v).slice(0, 25) : '';
    const item = document.createElement('label');
    item.className = 'dup-col-item';
    item.innerHTML = `
      <input type="checkbox" data-col="${c}" checked />
      <span class="col-letter">${colName(c)}</span>
      <span>${escapeHtml(headerText || colName(c))}</span>
    `;
    box.appendChild(item);
  }

  document.getElementById('dup-has-header').checked = true;
}

function closeDupModal() {
  document.getElementById('modal-dup')?.classList.add('hidden');
}

document.getElementById('btn-remove-dup')?.addEventListener('click', openDupModal);
document.getElementById('btn-dup-cancel')?.addEventListener('click', closeDupModal);

document.getElementById('btn-dup-apply')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  const cols = [];
  document.querySelectorAll('#dup-columns input[type=checkbox]:checked').forEach(cb => {
    cols.push(+cb.dataset.col);
  });
  if (!cols.length) return toast('Pilih minimal 1 kolom');

  const hasHeader = document.getElementById('dup-has-header').checked;
  const startRow = hasHeader ? 1 : 0;

  // Cari max
  let maxR = -1, maxC = -1;
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (!p) continue;
    if (p.r > maxR) maxR = p.r;
    if (p.c > maxC) maxC = p.c;
  }
  if (maxR < startRow) return toast('Tidak ada data');

  // Detect duplicates
  const seen = new Map();
  const toRemove = [];

  for (let r = startRow; r <= maxR; r++) {
    const key = cols.map(c => {
      const cell = sheet().cells[addrOf(r, c)];
      let v;
      if (cell?.f) {
        try { v = hfGet(r,c); }
        catch { v = cell.v ?? ''; }
      } else {
        v = cell?.v ?? '';
      }
      return String(v ?? '').trim().toLowerCase();
    }).join('\u0001');

    // Skip baris kosong
    if (key.replace(/\u0001/g, '') === '') continue;

    if (seen.has(key)) {
      toRemove.push(r);
    } else {
      seen.set(key, r);
    }
  }

  if (!toRemove.length) return toast('Tidak ada duplikat');

  if (!customConfirm(`Hapus ${toRemove.length} baris duplikat?`)) return;

  pushUndo();

  // Geser baris
  const removeSet = new Set(toRemove);
  let writeRow = startRow;
  for (let r = startRow; r <= maxR; r++) {
    if (removeSet.has(r)) continue;
    if (writeRow !== r) {
      // Copy cells dari r → writeRow
      for (let c = 0; c <= maxC; c++) {
        delete sheet().cells[addrOf(writeRow, c)];
        const src = sheet().cells[addrOf(r, c)];
        if (src) sheet().cells[addrOf(writeRow, c)] = src;
        delete sheet().cells[addrOf(r, c)];
      }
    }
    writeRow++;
  }

  // Hapus baris sisa
  for (let r = writeRow; r <= maxR; r++) {
    for (let c = 0; c <= maxC; c++) {
      delete sheet().cells[addrOf(r, c)];
    }
  }

  resetHF();
  renderSheet();
  markDirty();
  closeDupModal();
  toast(`✔ ${toRemove.length} baris duplikat dihapus`);
});

/* =========================================================
   AUTO FILTER DROPDOWN DI HEADER (selalu tampil)
   ========================================================= */
function enableAutoFilterHeaders() {
  // const thead = document.querySelector('#grid thead tr');
  // if (!thead) return;

  // // Selalu tampilkan ikon filter di header (bukan hanya mode filter)
  // thead.querySelectorAll('th[data-col]').forEach(th => {
  //   if (th.querySelector('.fx-filter-btn')) return;

  //   const c = +th.dataset.col;
  //   const btn = document.createElement('button');
  //   btn.className = 'fx-filter-btn';
  //   btn.innerHTML = '<i class="fa-solid fa-caret-down"></i>';
  //   btn.title = 'Filter kolom ini';
  //   btn.addEventListener('mousedown', e => {
  //     e.stopPropagation();
  //     e.preventDefault();
  //     openFilterPopup(c, th);
  //   });
  //   th.appendChild(btn);
  // });

  // // Highlight filter aktif
  // const filters = sheet().filters || {};
  // thead.querySelectorAll('th[data-col]').forEach(th => {
  //   const c = +th.dataset.col;
  //   th.classList.toggle('filtered', !!filters[c]?.values?.length);
  //   th.classList.toggle('filter-active', !!filters[c]?.values?.length);
  // });
}

// // Override renderFilterHeaders — panggil auto filter selalu
// const _origRenderFilterHeaders = renderFilterHeaders;
// renderFilterHeaders = function () {
//   // Selalu tampilkan (sekarang filter bukan mode lagi)
//   const thead = document.querySelector('#grid thead tr');
//   if (!thead) return;
//   thead.querySelectorAll('.fx-filter-btn').forEach(b => b.remove());
//   thead.querySelectorAll('th[data-col]').forEach(th => th.classList.remove('filtered', 'filter-active'));
//   enableAutoFilterHeaders();
// };

// Panggil saat build grid
const _origBuildGrid = buildGrid;
buildGrid = function () {
  _origBuildGrid();
  // enableAutoFilterHeaders();
};

/* Tombol Filter toggle jadi Clear All Filters */
document.getElementById('btn-filter-toggle')?.addEventListener('click', async () => {
  filterMode = !filterMode;
  const btn = document.getElementById('btn-filter-toggle');

  if (btn) {
    btn.style.background = filterMode ? 'linear-gradient(135deg,#10b981,#059669)' : '';
    btn.style.color = filterMode ? '#fff' : '';
  }

  window._lastRenderKey = null;
  renderVirtual(false);

  if (filterMode) {
    const hr = getFilterHeaderRow();
    toast(`🔽 Filter aktif di baris ${hr + 1}. Klik ▼ di header kolom untuk filter.`);
  } else {
    toast('Filter dinonaktifkan');
  }
}, true);

// /* =========================================================
//    Right-click row header → "Set sebagai Filter Header"
//    ========================================================= */
// document.getElementById('grid')?.addEventListener('contextmenu', e => {
//   const rowTh = e.target.closest('tbody th[data-row]');
//   if (!rowTh) return;   // hanya row header

//   e.preventDefault();
//   const rowNum = +rowTh.dataset.row;

//   // Konfirmasi
//   if (customConfirm(`Set baris ${rowNum + 1} sebagai header filter?`)) {
//     filterHeaderRow = rowNum;
//     filterMode = true;

//     const btn = document.getElementById('btn-filter-toggle');
//     if (btn) {
//       btn.style.background = 'linear-gradient(135deg,#10b981,#059669)';
//       btn.style.color = '#fff';
//     }

//     invalidateRowMap();
//     window._lastRenderKey = null;
//     renderVirtual(false);
//     toast(`✅ Baris ${rowNum + 1} jadi header filter`);
//   }
// });

/* =========================================================
   SHEET PROTECTION
   ========================================================= */
function hashPassword(pwd) {
  // Simple hash (bukan kriptografi kuat, cukup untuk proteksi UI)
  let h = 0;
  for (let i = 0; i < pwd.length; i++) {
    h = ((h << 5) - h) + pwd.charCodeAt(i);
    h |= 0;
  }
  return String(h);
}

function openProtectModal() {
  const cur = sheet();
  if (cur.protected) {
    return toast('Sheet sudah diproteksi');
  }
  document.getElementById('protect-password').value = '';
  document.getElementById('modal-protect').classList.remove('hidden');
  setTimeout(() => document.getElementById('protect-password').focus(), 50);
}

function closeProtectModal() {
  document.getElementById('modal-protect')?.classList.add('hidden');
}

function openUnprotectModal() {
  const cur = sheet();
  if (!cur.protected) return toast('Sheet tidak diproteksi');
  document.getElementById('unprotect-password').value = '';
  document.getElementById('modal-unprotect').classList.remove('hidden');
  setTimeout(() => document.getElementById('unprotect-password').focus(), 50);
}

function closeUnprotectModal() {
  document.getElementById('modal-unprotect')?.classList.add('hidden');
}

document.getElementById('btn-protect-sheet')?.addEventListener('click', openProtectModal);
document.getElementById('btn-protect-cancel')?.addEventListener('click', closeProtectModal);
document.getElementById('btn-unprotect-sheet')?.addEventListener('click', openUnprotectModal);
document.getElementById('btn-unprotect-cancel')?.addEventListener('click', closeUnprotectModal);

document.getElementById('btn-protect-apply')?.addEventListener('click', () => {
  const pwd = document.getElementById('protect-password').value;
  const cur = sheet();
  cur.protected = true;
  cur.protectHash = pwd ? hashPassword(pwd) : '';
  markMetaDirty('structure');
  applySheetProtection();
  closeProtectModal();
  toast('Sheet diproteksi');
});

document.getElementById('btn-unprotect-apply')?.addEventListener('click', () => {
  const pwd = document.getElementById('unprotect-password').value;
  const cur = sheet();
  if (cur.protectHash && hashPassword(pwd) !== cur.protectHash) {
    toast('⚠ Password salah');
    return;
  }
  cur.protected = false;
  delete cur.protectHash;
  markMetaDirty('structure');
  applySheetProtection();
  closeUnprotectModal();
  toast('Proteksi dibuka');
});

function applySheetProtection() {
  const cur = sheet();
  const isProtected = !!cur.protected;
  const isViewMode = currentAccessMode === 'view';

  document.body.classList.toggle('sheet-protected', isProtected);
  document.body.classList.toggle('readonly-mode', isViewMode);

  const btnProtect = document.getElementById('btn-protect-sheet');
  const btnUnprotect = document.getElementById('btn-unprotect-sheet');
  if (btnProtect && btnUnprotect) {
    btnProtect.style.display = isProtected ? 'none' : '';
    btnUnprotect.style.display = isProtected ? '' : 'none';
  }

  // 🎯 Apply lock indicators
  applyLockIndicators();

  renderSheetTabs();
}

/* Bind banner unprotect */
document.getElementById('btn-banner-unprotect')?.addEventListener('click', () => {
  openUnprotectModal();
});

/* Patch renderSheetTabs agar tampilkan indikator proteksi */
const _origRenderSheetTabs = renderSheetTabs;
renderSheetTabs = function () {
  _origRenderSheetTabs();
  // Tandai tab yang protected
  const tabs = document.querySelectorAll('#sheet-tabs .sheet-tab');
  wbData.sheets.forEach((s, i) => {
    if (tabs[i] && s.protected) tabs[i].classList.add('protected');
  });
};

/* Patch openEditor untuk cek proteksi */
const _origOpenEditorProtect = openEditor;
openEditor = function (addr, initial) {
  const cur = sheet();
  if (cur.protected && currentAccessMode !== 'view') {
    toast('Sheet diproteksi — buka proteksi dulu');
    return;
  }
  return _origOpenEditorProtect(addr, initial);
};

/* Update saat ganti sheet */
const _origSwitchSheet = switchSheet;
switchSheet = async function (i) {
  await _origSwitchSheet(i);
  applySheetProtection();
};
/* =========================================================
   FREEZE PANES
   ========================================================= */
function freezeTopRow() {
  sheet().freezeRow = 1;
  sheet().freezeCol = 0;
  markMetaDirty('structure');
  visibleStart = -1;
  visibleEnd = -1;
  renderSheet();
  toast('Baris 1 di-freeze');
}

function freezeFirstCol() {
  sheet().freezeRow = 0;
  sheet().freezeCol = 1;
  markMetaDirty('structure');
  visibleStart = -1;
  visibleEnd = -1;
  renderSheet();
  toast('Kolom A di-freeze');
}

function freezeHere() {
  const p = parseAddr(activeCell);
  if (!p) return;
  if (p.r === 0 && p.c === 0) {
    return toast('Pilih sel di luar A1');
  }

  const MAX_FREEZE_ROW = 4;
  const MAX_FREEZE_COL = 3;
  const fr = Math.min(p.r, MAX_FREEZE_ROW);
  const fc = Math.min(p.c, MAX_FREEZE_COL);

  if (fr === 0 && fc === 0) {
    return toast('Pilih sel yang lebih jauh dari A1');
  }

  sheet().freezeRow = fr;
  sheet().freezeCol = fc;
  markMetaDirty('structure');
  visibleStart = -1;
  visibleEnd = -1;
  renderSheet();
  toast(`Freeze ${fr} baris + ${fc} kolom`);
}

function unfreezeAll() {
  sheet().freezeRow = 0;
  sheet().freezeCol = 0;
  markMetaDirty('structure');
  visibleStart = -1;
  visibleEnd = -1;
  renderSheet();
  toast('Freeze dihapus');
}

document.getElementById('btn-freeze-top')?.addEventListener('click', freezeTopRow);
document.getElementById('btn-freeze-first-col')?.addEventListener('click', freezeFirstCol);
document.getElementById('btn-freeze-here')?.addEventListener('click', freezeHere);
document.getElementById('btn-unfreeze')?.addEventListener('click', unfreezeAll);

/* =========================================================
   FREEZE RENDERING PATCH — modify renderVirtual
   ========================================================= */
const _origRenderVirtual = renderVirtual;
renderVirtual = function (fastMode = false) {
  const gridWrapEl = document.getElementById('grid-wrap');
  const scrollTop = gridWrapEl.scrollTop;
  const viewH = gridWrapEl.clientHeight;

  const rowMap = getRowMap();
  const totalRows = rowMap ? rowMap.length : ROWS;

  const frozenRowCount = Math.min(sheet().freezeRow || 0, 20);
  const frozenColCount = Math.min(sheet().freezeCol || 0, 5);

  const effectiveScrollTop = Math.max(0, scrollTop);
  const scrollStart = Math.max(0, Math.floor(effectiveScrollTop / ROW_H) - BUFFER);
  const scrollEnd = Math.min(totalRows, Math.ceil((effectiveScrollTop + viewH) / ROW_H) + BUFFER);

  // 🎯 stateKey WAJIB include activeSheetIdx + nama sheet
  // biar ganti sheet selalu trigger re-render walau posisi scroll sama
  const sheetName = sheet()?.name || '';
  const stateKey =
    activeSheetIdx + '|' + sheetName +
    '|' + scrollStart + ':' + scrollEnd +
    '|' + frozenRowCount + ':' + frozenColCount +
    '|' + (filterMode ? '1' : '0') +
    '|' + (filterHeaderRow == null ? 'x' : filterHeaderRow);

  if (stateKey === window._lastFrozenKey) {
    // console.log('⏭ renderVirtual SKIP:', stateKey);
    return;
  }
  // console.log('🎨 renderVirtual RUN:', stateKey);
  window._lastFrozenKey = stateKey;

  visibleStart = scrollStart;
  visibleEnd = scrollEnd;
  window._lastFrozenRow = frozenRowCount;
  window._lastFrozenCol = frozenColCount;

  const tbody = document.getElementById('grid-body');
  if (!tbody) return;
  const frag = document.createDocumentFragment();

  // Merges visible
  const visibleMerges = [];
  const allVisibleStart = Math.min(scrollStart, frozenRowCount);
  const allVisibleEnd = Math.max(scrollEnd, frozenRowCount);
  for (const rangeStr of (sheet().merged || [])) {
    const [s, e] = rangeParts(rangeStr);
    const p1 = parseAddr(s), p2 = parseAddr(e);
    if (!p1 || !p2) continue;
    const r0 = Math.min(p1.r, p2.r), r1 = Math.max(p1.r, p2.r);
    if (r0 < allVisibleEnd && r1 >= allVisibleStart) {
      visibleMerges.push({
        r0, r1,
        c0: Math.min(p1.c, p2.c),
        c1: Math.max(p1.c, p2.c)
      });
    }
  }

  /* ---- 1. Frozen rows ---- */
  for (let r = 0; r < frozenRowCount; r++) {
    const actualR = rowMap ? rowMap[r] : r;
    if (actualR === undefined) continue;
    const tr = buildRowWithMerges(actualR, visibleMerges);
    tr.classList.add('frozen-row');
    tr.style.position = 'sticky';
    tr.style.top = (26 + r * ROW_H) + 'px';
    tr.style.zIndex = '8';

    tr.querySelectorAll('td, th').forEach(cell => {
      cell.style.position = 'sticky';
      cell.style.top = (26 + r * ROW_H) + 'px';
      cell.style.background = cell.tagName === 'TH' ? '#ecfdf5' : '#fff';
      cell.style.zIndex = '8';
    });
    frag.appendChild(tr);
  }

  /* ---- 2. Top spacer ---- */
  if (scrollStart > frozenRowCount) {
    const spacerH = (scrollStart - frozenRowCount) * ROW_H;
    const tr = document.createElement('tr');
    tr.style.height = spacerH + 'px';
    const td = document.createElement('td');
    td.colSpan = COLS + 1;
    td.style.cssText = 'padding:0;border:0';
    tr.appendChild(td);
    frag.appendChild(tr);
  }

  /* ---- 3. Visible rows ---- */
  const visStart = Math.max(frozenRowCount, Math.floor(scrollTop / ROW_H) + frozenRowCount - BUFFER);
  const visEnd = Math.min(totalRows, Math.ceil((scrollTop + viewH) / ROW_H) + frozenRowCount + BUFFER);
  for (let di = visStart; di < visEnd; di++) {
    const r = rowMap ? rowMap[di] : di;
    const tr = buildRowWithMerges(r, visibleMerges);
    frag.appendChild(tr);
  }

  /* ---- 4. Bottom spacer ---- */
  if (scrollEnd < totalRows) {
    const spacerH = (totalRows - scrollEnd) * ROW_H;
    const tr = document.createElement('tr');
    tr.style.height = spacerH + 'px';
    const td = document.createElement('td');
    td.colSpan = COLS + 1;
    td.style.cssText = 'padding:0;border:0';
    tr.appendChild(td);
    frag.appendChild(tr);
  }

  tbody.replaceChildren(frag);

  // Update rowEls
  rowEls = {};
  Array.from(tbody.children).forEach(tr => {
    const th = tr.firstElementChild;
    if (th && th.dataset.row !== undefined) {
      rowEls[+th.dataset.row] = tr;
    }
  });

  if (frozenColCount > 0) {
    applyFreezeCols(frozenColCount);
  }

  if (fastMode) {
    // 🎯 FAST: cuma render cells + selection
    renderVisibleCells();
    applySelectionStyles();
  } else {
    // 🎯 FULL: semua fitur
    renderVisibleCells();
    renderCFVisible();
    applySelectionStyles();
    renderRowHeaderHighlight();
    applyValidationIndicators();
    applyCommentMarkers();
  }

  if (typeof applyHiddenRowsCols === 'function') {
    applyHiddenRowsCols();
  }
};

/* Patch visibleStart untuk renderVisibleCells dll */
// const _origRenderVisibleCells = renderVisibleCells;
// renderVisibleCells = function () {
//   const frozenRowCount = Math.min(sheet().freezeRow || 0, 20);
//   for (let r = 0; r < frozenRowCount; r++) {
//     const tr = rowEls[r];
//     if (!tr) continue;
//     tr.querySelectorAll('td[data-addr]').forEach(td => {
//       const p = parseAddr(td.dataset.addr);
//       if (p) updateCellDisplay(td, p.r, p.c);
//     });
//   }
//   _origRenderVisibleCells();
// };

/* Terapkan sticky left untuk freeze cols */
function applyFreezeCols(frozenColCount) {
  const rowHeaderWidth = 48;
  const colOffsets = [rowHeaderWidth];
  for (let c = 0; c < COLS; c++) {
    colOffsets.push(colOffsets[c] + (sheet().colWidths?.[c] || 96));
  }

  document.querySelectorAll('#grid tbody tr').forEach(tr => {
    const isFrozenRow = tr.classList.contains('frozen-row');

    tr.querySelectorAll('td[data-addr]').forEach(td => {
      const p = parseAddr(td.dataset.addr);
      if (!p) return;
      if (p.c < frozenColCount) {
        td.classList.add('freeze-col');
        td.style.position = 'sticky';
        td.style.left = colOffsets[p.c] + 'px';
        td.style.zIndex = isFrozenRow ? '9' : '6';
      } else {
        td.classList.remove('freeze-col');
        td.style.position = '';
        td.style.left = '';
      }
    });
  });
}

/* =========================================================
   DATA VALIDATION
   ========================================================= */
function parseRange(range) {
  const parts = range.split(':');
  const p1 = parseAddr(parts[0]?.trim().toUpperCase());
  const p2 = parseAddr((parts[1] || parts[0])?.trim().toUpperCase());
  if (!p1 || !p2) return null;
  return {
    r0: Math.min(p1.r, p2.r), r1: Math.max(p1.r, p2.r),
    c0: Math.min(p1.c, p2.c), c1: Math.max(p1.c, p2.c)
  };
}

function findValidationForCell(addr) {
  const p = parseAddr(addr);
  if (!p) return null;
  const vals = sheet().validations || [];
  for (const v of vals) {
    const r = parseRange(v.range);
    if (!r) continue;
    if (p.r >= r.r0 && p.r <= r.r1 && p.c >= r.c0 && p.c <= r.c1) return v;
  }
  return null;
}

function validateValue(v, rule) {
  const s = String(v ?? '').trim();
  if (s === '') return true; // kosong selalu valid

  switch (rule.type) {
    case 'list': {
      const opts = (rule.options || []).map(o => String(o).trim());
      return opts.includes(s);
    }
    case 'number': {
      const n = Number(s);
      if (isNaN(n)) return false;
      return checkRange(n, rule.condition, Number(rule.min), Number(rule.max));
    }
    case 'date': {
      const d = new Date(s);
      if (isNaN(d.getTime())) return false;
      const t = d.getTime();
      const min = rule.min ? new Date(rule.min).getTime() : null;
      const max = rule.max ? new Date(rule.max).getTime() : null;
      return checkRange(t, rule.condition, min, max);
    }
    case 'text': {
      const len = s.length;
      const min = rule.minLength != null ? Number(rule.minLength) : null;
      const max = rule.maxLength != null ? Number(rule.maxLength) : null;
      if (min != null && len < min) return false;
      if (max != null && len > max) return false;
      return true;
    }
    case 'email': {
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
    }
    case 'url': {
      try { new URL(s); return true; } catch { return false; }
    }
    default: return true;
  }
}

function checkRange(n, cond, min, max) {
  switch (cond) {
    case 'between': return n >= min && n <= max;
    case 'notBetween': return n < min || n > max;
    case 'gt': return n > min;
    case 'lt': return n < min;
    case 'gte': return n >= min;
    case 'lte': return n <= min;
    case 'equal': return n === min;
    default: return true;
  }
}

function openValidationModal() {
  const modal = document.getElementById('modal-validation');
  if (!modal) return;
  modal.classList.remove('hidden');

  // Preload range dari selection
  const range = selection.start === selection.end
    ? activeCell
    : `${selection.start}:${selection.end}`;
  document.getElementById('val-range').value = range;

  // Reset form
  document.getElementById('val-type').value = 'list';
  document.getElementById('val-condition').value = 'between';
  document.getElementById('val-min').value = '';
  document.getElementById('val-max').value = '';
  document.getElementById('val-options').value = '';
  document.getElementById('val-text-min').value = '';
  document.getElementById('val-text-max').value = '';
  document.getElementById('val-error-msg').value = '';

  updateValidationUI();
  renderExistingValidations();
  setTimeout(() => document.getElementById('val-range').focus(), 50);
}

function updateValidationUI() {
  const type = document.getElementById('val-type').value;
  document.getElementById('val-condition-wrap').style.display =
    (type === 'number' || type === 'date') ? '' : 'none';
  document.getElementById('val-list-wrap').style.display =
    type === 'list' ? '' : 'none';
  document.getElementById('val-text-wrap').style.display =
    type === 'text' ? '' : 'none';

  if (type === 'number' || type === 'date') {
    const cond = document.getElementById('val-condition').value;
    const showMax = (cond === 'between' || cond === 'notBetween');
    document.getElementById('val-max-wrap').style.display = showMax ? '' : 'none';
    document.getElementById('val-min-label').textContent =
      cond === 'gt' ? 'Nilai minimum' :
      cond === 'lt' ? 'Nilai maksimum' :
      cond === 'gte' ? 'Nilai minimum' :
      cond === 'lte' ? 'Nilai maksimum' :
      cond === 'equal' ? 'Nilai' : 'Minimal';
  }
}

document.getElementById('val-type')?.addEventListener('change', updateValidationUI);
document.getElementById('val-condition')?.addEventListener('change', updateValidationUI);
document.getElementById('btn-validation')?.addEventListener('click', openValidationModal);
document.getElementById('btn-val-cancel')?.addEventListener('click', () =>
  document.getElementById('modal-validation').classList.add('hidden'));

document.getElementById('btn-val-apply')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  const range = document.getElementById('val-range').value.trim();
  const parsed = parseRange(range);
  if (!parsed) return toast('⚠ Range tidak valid');

  const type = document.getElementById('val-type').value;
  const rule = {
    range: range.toUpperCase(),
    type,
    errorMsg: document.getElementById('val-error-msg').value.trim() || 'Nilai tidak valid'
  };

  if (type === 'list') {
    const raw = document.getElementById('val-options').value;
    const opts = raw.split(/[,\n]/).map(s => s.trim()).filter(Boolean);
    if (!opts.length) return toast('⚠ Isi minimal 1 opsi');
    rule.options = opts;
  } else if (type === 'number' || type === 'date') {
    rule.condition = document.getElementById('val-condition').value;
    rule.min = document.getElementById('val-min').value;
    rule.max = document.getElementById('val-max').value;
    if (!rule.min) return toast('⚠ Isi nilai');
  } else if (type === 'text') {
    rule.minLength = document.getElementById('val-text-min').value;
    rule.maxLength = document.getElementById('val-text-max').value;
  }

  pushUndo();
  sheet().validations ||= [];
  // Hapus validation lama di range yang sama
  sheet().validations = sheet().validations.filter(v => v.range !== rule.range);
  sheet().validations.push(rule);
  markMetaDirty('structure');

  document.getElementById('modal-validation').classList.add('hidden');
  applyValidationIndicators();
  toast('✔ Validasi diterapkan');
});

document.getElementById('btn-validation-clear')?.addEventListener('click', () => {
  if (!customConfirm('Hapus semua validasi di sheet ini?')) return;
  pushUndo();
  sheet().validations = [];
  markMetaDirty('structure');
  applyValidationIndicators();
  toast('✔ Validasi dibersihkan');
});

function renderExistingValidations() {
  const box = document.getElementById('val-existing');
  const vals = sheet().validations || [];
  if (!vals.length) {
    box.innerHTML = '';
    return;
  }
  box.innerHTML = `<div style="font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:.4px;font-weight:700;margin-bottom:6px">Validasi aktif (${vals.length})</div>` +
    vals.map(v => `
      <div class="view-link-row" style="justify-content:space-between">
        <span style="font-weight:600;color:#334155">${escapeHtml(v.range)}</span>
        <span style="font-size:11px;color:#64748b;background:#f1f5f9;padding:2px 8px;border-radius:5px">${v.type}</span>
        <button data-val-del="${escapeHtml(v.range)}" style="background:#fee2e2;color:#dc2626">
          <i class="fa-solid fa-trash"></i>
        </button>
      </div>
    `).join('');

  box.querySelectorAll('[data-val-del]').forEach(btn => {
    btn.addEventListener('click', () => {
      const r = btn.dataset.valDel;
      sheet().validations = (sheet().validations || []).filter(v => v.range !== r);
      markMetaDirty('structure');
      renderExistingValidations();
      applyValidationIndicators();
    });
  });
}

function getValidationAddrs() {
  if (_validationAddrsCache) return _validationAddrsCache;
  const vals = sheet().validations || [];
  if (!vals.length) { _validationAddrsCache = []; return []; }

  const addrs = [];
  const cells = sheet().cells || {};
  for (const addr in cells) {
    const p = parseAddr(addr);
    if (!p) continue;
    for (const v of vals) {
      const r = parseRange(v.range);
      if (!r) continue;
      if (p.r >= r.r0 && p.r <= r.r1 && p.c >= r.c0 && p.c <= r.c1) {
        addrs.push(addr);
        break;
      }
    }
  }
  _validationAddrsCache = addrs;
  console.log('✅ Validation cache built:', addrs.length);
  return addrs;
}

function applyValidationIndicators() {
  document.querySelectorAll('#grid td.validated, #grid td.invalid')
    .forEach(td => td.classList.remove('validated', 'invalid'));

  const addrs = getValidationAddrs();
  if (!addrs.length) return;

  for (const addr of addrs) {
    const td = document.querySelector(`#grid td[data-addr="${addr}"]`);
    if (!td) continue;
    td.classList.add('validated');

    const cl = sheet().cells[addr];
    const v = cl?.v;
    if (v === undefined || v === null || v === '') continue;
    const rule = findValidationForCell(addr);
    if (rule && !validateValue(v, rule)) td.classList.add('invalid');
  }
}

/* Cek saat commit — reject kalau tidak valid */
const _origCommitValidate = commit;
commit = function (addr, value) {
  const rule = findValidationForCell(addr);
  if (rule) {
    const raw = String(value);
    if (raw !== '' && raw.startsWith('=')) {
      // Formula boleh
    } else if (!validateValue(raw, rule)) {
      toast('⚠ ' + (rule.errorMsg || 'Nilai tidak valid'));
      return;
    }
  }
  return _origCommitValidate(addr, value);
};

/* Dropdown saat edit sel dengan list validation */
const _origOpenEditorList = openEditor;
openEditor = function (addr, initial) {
  const result = _origOpenEditorList(addr, initial);
  if (!editor) return;

  const rule = findValidationForCell(addr);
  if (rule?.type !== 'list') return;

  const input = editor.input;
  const rect = input.getBoundingClientRect();

  // Tampilkan dropdown setelah 200ms
  setTimeout(() => {
    if (!editor || editor.input !== input) return;

    const dd = document.createElement('div');
    dd.className = 'val-dropdown';
    dd.style.left = rect.left + 'px';
    dd.style.top = (rect.bottom + 4) + 'px';
    dd.innerHTML = (rule.options || []).map(o =>
      `<div class="val-option" data-val="${escapeHtml(o)}">${escapeHtml(o)}</div>`
    ).join('');
    document.body.appendChild(dd);

    dd.querySelectorAll('.val-option').forEach(opt => {
      opt.addEventListener('mousedown', e => {
        e.preventDefault();
        input.value = opt.dataset.val;
        dd.remove();
        if (editor && editor.input === input) {
          commit(addr, input.value);
          closeEditor();
        }
      });
    });

    // Close on blur
    const closeDD = () => { if (dd.parentNode) dd.remove(); };
    input.addEventListener('blur', closeDD, { once: true });
    setTimeout(() => {
      document.addEventListener('mousedown', function h(e) {
        if (!dd.contains(e.target) && e.target !== input) {
          closeDD();
          document.removeEventListener('mousedown', h, true);
        }
      }, true);
    }, 100);
  }, 200);

  return result;
};

/* =========================================================
   COMMENTS / NOTES
   ========================================================= */
function openCommentModal(addr) {
  const modal = document.getElementById('modal-comment');
  if (!modal) return;
  modal.classList.remove('hidden');

  const cell = sheet().cells[addr] || {};
  const cmt = cell.comment;

  document.getElementById('comment-target').innerHTML = `
    <i class="fa-solid fa-crosshairs"></i>
    <span class="cmt-addr">${escapeHtml(addr)}</span>
    <span>${cmt ? 'Edit Komentar' : 'Tambah Komentar'}</span>
  `;

  document.getElementById('comment-text').value = cmt?.text || '';

  const meta = document.getElementById('comment-meta');
  if (cmt) {
    meta.innerHTML = `<i class="fa-solid fa-user"></i> ${escapeHtml(cmt.author || 'User')} • <i class="fa-regular fa-clock"></i> ${new Date(cmt.createdAt).toLocaleString('id-ID')}`;
    document.getElementById('btn-comment-delete').style.display = '';
  } else {
    meta.innerHTML = '';
    document.getElementById('btn-comment-delete').style.display = 'none';
  }

  modal.dataset.addr = addr;
  setTimeout(() => document.getElementById('comment-text').focus(), 50);
}

function closeCommentModal() {
  document.getElementById('modal-comment')?.classList.add('hidden');
}

document.getElementById('btn-comment-cancel')?.addEventListener('click', closeCommentModal);

document.getElementById('btn-comment-save')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  const addr = document.getElementById('modal-comment').dataset.addr;
  const text = document.getElementById('comment-text').value.trim();
  if (!text) return toast('⚠ Tulis komentar dulu');

  pushUndo();
  ensureCell(addr);
  const existing = sheet().cells[addr].comment;
  sheet().cells[addr].comment = {
    text,
    author: currentUser?.username || 'User',
    createdAt: existing?.createdAt || Date.now(),
    updatedAt: Date.now()
  };
  markCellDirty(addr);

  closeCommentModal();
  applyCommentMarkers();
  toast('✔ Komentar disimpan');
});

document.getElementById('btn-comment-delete')?.addEventListener('click', async () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  const addr = document.getElementById('modal-comment').dataset.addr;
  if (!(await customConfirm('Hapus komentar ini?', { type: 'danger', okText: 'Hapus' }))) return;
  pushUndo();
  const cell = sheet().cells[addr];
  if (cell) delete cell.comment;
  markCellDirty(addr);
  closeCommentModal();
  applyCommentMarkers();
  toast('✔ Komentar dihapus');
});

document.getElementById('btn-add-comment')?.addEventListener('click', () => {
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
  openCommentModal(activeCell);
});

document.getElementById('btn-delete-comment')?.addEventListener('click', () => {
  const cell = sheet().cells[activeCell];
  if (!cell?.comment) return toast('Sel ini tidak punya komentar');
  document.getElementById('modal-comment').dataset.addr = activeCell;
  document.getElementById('btn-comment-delete').click();
});

/* =========================================================
   COMMENT MARKERS — versi cache (jangan iterasi 2M cells)
   ========================================================= */
let _commentAddrsCache = null;
let _validationAddrsCache = null;

function invalidateCommentCache() { _commentAddrsCache = null; }
function invalidateValidationCache() { _validationAddrsCache = null; }

function getCommentAddrs() {
  if (_commentAddrsCache) return _commentAddrsCache;
  const addrs = [];
  const cells = sheet().cells || {};
  for (const addr in cells) {
    if (cells[addr]?.comment) addrs.push(addr);
  }
  _commentAddrsCache = addrs;
  console.log('💬 Comment cache built:', addrs.length);
  return addrs;
}

function applyCommentMarkers() {
  // Clear dulu
  document.querySelectorAll('#grid td.has-comment')
    .forEach(td => td.classList.remove('has-comment'));

  const addrs = getCommentAddrs();
  if (!addrs.length) return;

  // Hanya apply ke td yang ada di DOM (bukan iterasi semua cells)
  for (const addr of addrs) {
    const td = document.querySelector(`#grid td[data-addr="${addr}"]`);
    if (td) td.classList.add('has-comment');
  }
}

/* Hover tooltip */
let commentTooltipEl = null;
document.getElementById('grid')?.addEventListener('mouseover', e => {
  const td = e.target.closest('td.has-comment');
  if (!td) {
    if (commentTooltipEl) { commentTooltipEl.remove(); commentTooltipEl = null; }
    return;
  }
  const cell = sheet().cells[td.dataset.addr];
  if (!cell?.comment) return;

  if (commentTooltipEl) commentTooltipEl.remove();
  const tt = document.createElement('div');
  tt.className = 'comment-tooltip';
  const rect = td.getBoundingClientRect();
  tt.style.left = Math.min(rect.right + 8, window.innerWidth - 320) + 'px';
  tt.style.top = rect.top + 'px';
  tt.innerHTML = `
    <div class="cmt-author">
      <i class="fa-solid fa-user"></i> ${escapeHtml(cell.comment.author || 'User')}
    </div>
    <div>${escapeHtml(cell.comment.text).replace(/\n/g, '<br>')}</div>
    <div class="cmt-time">
      <i class="fa-regular fa-clock"></i>
      ${new Date(cell.comment.createdAt).toLocaleString('id-ID')}
    </div>
  `;
  document.body.appendChild(tt);
  commentTooltipEl = tt;
});

document.getElementById('grid')?.addEventListener('mouseout', e => {
  const related = e.relatedTarget;
  if (commentTooltipEl && (!related || !related.closest('td.has-comment'))) {
    commentTooltipEl.remove();
    commentTooltipEl = null;
  }
});

/* Show all comments modal */
document.getElementById('btn-show-comments')?.addEventListener('click', () => {
  const modal = document.getElementById('modal-comments-list');
  const body = document.getElementById('comments-list-body');
  const comments = [];
  for (const [addr, cell] of Object.entries(sheet().cells || {})) {
    if (cell.comment) {
      comments.push({ addr, ...cell.comment });
    }
  }
  if (!comments.length) {
    body.innerHTML = '<div class="comments-empty"><i class="fa-solid fa-comment-slash" style="font-size:32px;color:#cbd5e1;display:block;margin-bottom:12px"></i>Belum ada komentar di sheet ini</div>';
  } else {
    body.innerHTML = comments.map(c => `
      <div class="comment-item" data-addr="${escapeHtml(c.addr)}">
        <div class="cmt-avatar">${escapeHtml((c.author || 'U').charAt(0).toUpperCase())}</div>
        <div class="cmt-body">
          <div class="cmt-header">
            <span class="cmt-addr">${escapeHtml(c.addr)}</span>
            <span class="cmt-author">${escapeHtml(c.author || 'User')}</span>
            <span class="cmt-time">${new Date(c.createdAt).toLocaleString('id-ID')}</span>
          </div>
          <div class="cmt-text">${escapeHtml(c.text).replace(/\n/g, '<br>')}</div>
        </div>
      </div>
    `).join('');

    body.querySelectorAll('.comment-item').forEach(item => {
      item.addEventListener('click', () => {
        selectCell(item.dataset.addr);
        modal.classList.add('hidden');
      });
    });
  }
  modal.classList.remove('hidden');
});

document.getElementById('btn-comments-close')?.addEventListener('click', () =>
  document.getElementById('modal-comments-list').classList.add('hidden'));

/* =========================================================
   CONTEXT MENU
   ========================================================= */
const ctxMenu = document.getElementById('cell-context-menu');

document.getElementById('grid')?.addEventListener('contextmenu', e => {
  if (!ctxMenu) return;   // ← guard

  const td = e.target.closest('td[data-addr]');
  if (!td) return;
  e.preventDefault();

  // Select cell
  selectCell(td.dataset.addr);

  // Show/hide comment-view (aman dengan optional chaining)
  const cell = sheet().cells[td.dataset.addr];
  const hasComment = !!cell?.comment;
  const commentBtn = ctxMenu.querySelector('[data-action="comment"]');
  const commentViewBtn = ctxMenu.querySelector('[data-action="comment-view"]');
  if (commentBtn) commentBtn.style.display = hasComment ? 'none' : '';
  if (commentViewBtn) commentViewBtn.style.display = hasComment ? '' : 'none';

  // Position
  let x = e.clientX, y = e.clientY;
  ctxMenu.classList.remove('hidden');
  const rect = ctxMenu.getBoundingClientRect();
  if (x + rect.width > window.innerWidth - 10) x = window.innerWidth - rect.width - 10;
  if (y + rect.height > window.innerHeight - 10) y = window.innerHeight - rect.height - 10;
  ctxMenu.style.left = x + 'px';
  ctxMenu.style.top = y + 'px';
});

document.addEventListener('mousedown', e => {
  if (ctxMenu && !ctxMenu.classList.contains('hidden') && !ctxMenu.contains(e.target)) {
    ctxMenu.classList.add('hidden');
  }
});

ctxMenu?.querySelectorAll('button').forEach(btn => {
  btn.addEventListener('click', async () => {
    ctxMenu.classList.add('hidden');
    const action = btn.dataset.action;
    switch (action) {
      case 'comment': openCommentModal(activeCell); break;
      case 'comment-view': openCommentModal(activeCell); break;
      case 'validation': openValidationModal(); break;
      case 'copy': doCopy(); break;
      case 'paste': doPaste(); break;
      case 'clear': document.getElementById('btn-clear-all')?.click(); break;
      case 'insert-row': document.getElementById('btn-insert-row')?.click(); break;
      case 'delete-row': document.getElementById('btn-del-row')?.click(); break;
      case 'paste-value':   doPasteSpecial('value'); break;
      case 'paste-formula': doPasteSpecial('formula'); break;
      case 'paste-format':  doPasteSpecial('format'); break;
      case 'lock-cell': lockSelection(); break;
      case 'unlock-cell': unlockSelection(); break;
    }
  });
});

/* =========================================================
   UPDATE RENDER CHAIN — Apply new indicators
   ========================================================= */
const _origRenderSheetBatch2 = renderSheet;
renderSheet = function () {
  _origRenderSheetBatch2();
  applyValidationIndicators();
  applyCommentMarkers();
  applyFreezeCols(Math.min(sheet().freezeCol || 0, 5));
};

/* Normalize validations + comments */
const _origNormalizeWB2 = normalizeWB;
normalizeWB = function (d) {
  const wb = _origNormalizeWB2(d);
  wb.sheets.forEach(s => {
    if (!Array.isArray(s.validations)) s.validations = [];
    if (typeof s.freezeRow !== 'number') s.freezeRow = 0;
    if (typeof s.freezeCol !== 'number') s.freezeCol = 0;
    if (typeof s.protected !== 'boolean') s.protected = false;
  });
  return wb;
};

/* Blank WB: tambah field baru */
const _origBlankWB2 = blankWB;
blankWB = function () {
  const wb = _origBlankWB2();
  wb.sheets[0].validations = [];
  wb.sheets[0].freezeRow = 0;
  wb.sheets[0].freezeCol = 0;
  wb.sheets[0].protected = false;
  return wb;
};
/* =========================================================
   GLOBAL ENTER HANDLER UNTUK SEMUA MODAL
   ========================================================= */
document.addEventListener('keydown', e => {
  // Hanya untuk modal aktif
  const openModal = document.querySelector('.modal:not(.hidden)');
  if (!openModal) return;

  // Skip kalau di textarea (biarkan Enter untuk new line)
  if (e.target.tagName === 'TEXTAREA') return;

  // Skip kalau custom dialog sudah handle sendiri
  if (document.querySelector('.custom-dialog-overlay')) return;

  if (e.key === 'Enter') {
    // Cari tombol primary
    const primaryBtn = openModal.querySelector('.btn.primary:not([disabled])') ||
                       openModal.querySelector('button[type=submit]:not([disabled])');
    if (primaryBtn) {
      e.preventDefault();
      e.stopPropagation();
      primaryBtn.click();
    }
  } else if (e.key === 'Escape') {
    // Cari tombol close/cancel
    const cancelBtn = openModal.querySelector('.btn.ghost') ||
                      openModal.querySelector('[id$="-cancel"]') ||
                      openModal.querySelector('[id$="-close"]') ||
                      openModal.querySelector('.chart-close') ||
                      openModal.querySelector('.find-close-btn');
    if (cancelBtn) {
      e.preventDefault();
      e.stopPropagation();
      cancelBtn.click();
    }
  }
}, true);
/* =========================================================
   CROSS-SHEET REFERENCE (Excel-like)
   ========================================================= */
function showCrossSheetBanner() {
  if (!crossSheetRef) return;
  const banner = document.getElementById('cross-sheet-banner');
  const target = document.getElementById('cs-target');
  const input = document.getElementById('cs-formula-input');
  if (!banner || !input) return;

  const originalSheet = wbData.sheets[crossSheetRef.originalSheetIdx];
  target.textContent = `Menyunting rumus untuk ${originalSheet.name}!${crossSheetRef.addr}`;
  input.value = crossSheetRef.formula;
  banner.classList.remove('hidden');

  // 🎯 Auto-focus & cursor di akhir
  setTimeout(() => {
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
  }, 50);

  console.log('🎯 Cross-sheet mode:', crossSheetRef);
}

function hideCrossSheetBanner() {
  const banner = document.getElementById('cross-sheet-banner');
  if (banner) banner.classList.add('hidden');
}

// function insertCrossSheetRef(sheetName, addr) {
//   if (!crossSheetRef) return;
//   const ref = `${sheetName}!${addr}`;
//   crossSheetRef.formula += ref;

//   const formula = document.getElementById('cs-formula');
//   if (formula) formula.textContent = crossSheetRef.formula;
// }

/* Insert cross-sheet ref langsung (dengan prefix sheet name) */
/* =========================================================
   INSERT CROSS-SHEET REF — dengan auto-quote
   ========================================================= */
function insertCrossSheetRef2(fullRef) {
  if (!crossSheetRef) return;

  const input = document.getElementById('cs-formula-input');
  if (!input) return;

  // 🎯 Insert di cursor position, bukan append ke akhir
  const val = input.value;
  const pos = input.selectionStart ?? val.length;
  const newVal = val.slice(0, pos) + fullRef + val.slice(pos);

  input.value = newVal;

  // Cursor pindah setelah ref
  const newPos = pos + fullRef.length;
  input.focus();
  input.setSelectionRange(newPos, newPos);

  // Sync ke crossSheetRef
  crossSheetRef.formula = newVal;

  console.log('✍ Cross-sheet ref ditambahkan:', fullRef);
  console.log('   Formula sekarang:', newVal);
  if (typeof toast === 'function') toast('✔ ' + fullRef);
}

function insertCrossSheetRange(sheetName, startAddr, endAddr) {
  if (!crossSheetRef) return;
  const ref = `${sheetName}!${startAddr}:${endAddr}`;
  crossSheetRef.formula += ref;

  const formula = document.getElementById('cs-formula');
  if (formula) formula.textContent = crossSheetRef.formula;
}

async function finishCrossSheetEdit() {
  if (!crossSheetRef) return;

  const input = document.getElementById('cs-formula-input');
  const finalFormula = input?.value || crossSheetRef.formula;

  const { originalSheetIdx, addr } = crossSheetRef;
  const refBackup = { ...crossSheetRef, formula: finalFormula };

  crossSheetRef = null;
  hideCrossSheetBanner();

  if (!isFormulaComplete(finalFormula)) {
    toast('⚠ Formula belum lengkap');
    crossSheetRef = refBackup;
    showCrossSheetBanner();
    return;
  }

  // Reset cache SEBELUM switch
  window._lastFrozenKey = null;
  window._lastFrozenRow = null;
  window._lastFrozenCol = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  rowEls = {};

  await switchSheet(originalSheetIdx);

  const p = parseAddr(addr);
  if (p) {
    ensureCell(addr);
    sheet().cells[addr].f = finalFormula;
    delete sheet().cells[addr].v;

    try { await rebuildHFWithLoaded(); } catch (e) { console.warn(e); }

    setHFCellSafe(p.r, p.c, finalFormula);
    markCellDirty(addr);

    clearHFCache();
    window._lastFrozenKey = null;
    window._lastFrozenRow = null;
    window._lastFrozenCol = null;
    window._lastRenderKey = null;
    visibleStart = -1;
    visibleEnd = -1;
    rowEls = {};

    renderSheet();

    setTimeout(() => {
      clearHFCache();
      window._lastFrozenKey = null;
      window._lastRenderKey = null;
      renderSheet();
    }, 150);

    toast('✔ Formula disimpan');
  }
}

function cancelCrossSheetEdit() {
  if (!crossSheetRef) return;
  const originalIdx = crossSheetRef.originalSheetIdx;
  crossSheetRef = null;
  hideCrossSheetBanner();

  const input = document.getElementById('cs-formula-input');
  if (input) input.value = '';

  switchSheet(originalIdx);
  toast('Dibatalkan');
}

// Bind tombol
document.getElementById('cs-finish')?.addEventListener('click', finishCrossSheetEdit);
document.getElementById('cs-cancel')?.addEventListener('click', cancelCrossSheetEdit);

// Enter di banner → finish
document.addEventListener('keydown', e => {
  if (!crossSheetRef) return;
  if (e.key === 'Enter') {
    e.preventDefault();
    finishCrossSheetEdit();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    cancelCrossSheetEdit();
  }
}, true);

/* =========================================================
   INDEXEDDB CACHE
   ========================================================= */
const IDB = {
  db: null,
  async open() {
    if (this.db) return this.db;
    return new Promise((resolve, reject) => {
      const req = indexedDB.open('tabelin-cache', 1);
      req.onupgradeneeded = e => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('workbooks')) {
          db.createObjectStore('workbooks', { keyPath: 'id' });
        }
      };
      req.onsuccess = () => { this.db = req.result; resolve(this.db); };
      req.onerror = () => reject(req.error);
    });
  },
  async set(id, data) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('workbooks', 'readwrite');
      tx.objectStore('workbooks').put({ id, data, ts: Date.now() });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  },
  async get(id) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('workbooks', 'readonly');
      const req = tx.objectStore('workbooks').get(id);
      req.onsuccess = () => resolve(req.result?.data || null);
      req.onerror = () => reject(req.error);
    });
  },
  async clear() {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('workbooks', 'readwrite');
      tx.objectStore('workbooks').clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }
};
async function optimizeWorkbook() {
  if (!wbId) return toast('Buka workbook dulu');
  if (wbUsesChunks) return toast('Sudah dalam mode chunk');

  const ok = await customConfirm(
    'Optimize workbook? Data akan dipindah ke format chunk untuk loading lebih cepat.\n\nData asli tetap tersimpan.',
    { type: 'question', okText: 'Optimize' }
  );
  if (!ok) return;

  showLoading('Mengoptimasi…');
  try {
    const res = await fetch(`/api/sheets/${wbId}/chunk-migrate`, {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token }
    });
    const result = await res.json();
    if (!res.ok) throw new Error(result.error || 'Gagal');

    hideLoading();
    toast(`✔ ${result.blocks} block dibuat`);
    await openWB(wbId);   // reload dengan chunk mode
  } catch (e) {
    hideLoading();
    toast('⚠ ' + e.message);
  }
}

/* =========================================================
   DRAG SELECT — blok cell dengan drag mouse
   ========================================================= */
let isDragging = false;
let dragStartAddr = null;

document.getElementById('grid').addEventListener('mousedown', e => {
  // Skip kalau sedang edit, klik header, grip resize, atau ref-mode
  if (editor || insertingRefMode || crossSheetRef) return;
  if (e.target.closest('th')) return;
  if (e.target.closest('.col-resize-grip') || e.target.closest('.row-resize-grip')) return;

  const td = e.target.closest('td[data-addr]');
  if (!td) return;

  if (e.button !== 0) return;  // hanya left click

  isDragging = true;
  dragStartAddr = td.dataset.addr;

  if (!e.shiftKey) {
    selection = { start: dragStartAddr, end: dragStartAddr };
    activeCell = dragStartAddr;
    updateSelectionUI();
  }
});

document.addEventListener('mousemove', e => {
  if (!isDragging || !dragStartAddr) return;

  const el = document.elementFromPoint(e.clientX, e.clientY);
  const td = el?.closest('td[data-addr]');
  if (!td) return;

  const currentAddr = td.dataset.addr;
  if (currentAddr === selection.end) return;

  selection.end = currentAddr;
  activeCell = selection.start;
  updateSelectionUI();
  updateStatusStats();
});

document.addEventListener('mouseup', () => {
  isDragging = false;
  dragStartAddr = null;
});

/* =========================================================
   HEADER SELECT — klik header row/col untuk select semua
   ========================================================= */
document.getElementById('grid').addEventListener('mousedown', e => {
  // Skip kalau sedang edit atau ref-mode
  if (editor || crossSheetRef) return;

  const target = e.target;

  // Header kolom
  if (target.matches('#grid thead th[data-col]')) {
    e.preventDefault();
    const c = +target.dataset.col;
    const lastRow = Math.max(1, ROWS);
    selection = { start: addrOf(0, c), end: addrOf(ROWS - 1, c) };
    activeCell = addrOf(0, c);
    updateSelectionUI();
    updateStatusStats();
    // Highlight header
    document.querySelectorAll('#grid thead th').forEach(th => th.classList.remove('active'));
    target.classList.add('active');
    return;
  }

  // Header baris
  if (target.matches('#grid tbody th[data-row]')) {
    e.preventDefault();
    const r = +target.dataset.row;
    selection = { start: addrOf(r, 0), end: addrOf(r, COLS - 1) };
    activeCell = addrOf(r, 0);
    updateSelectionUI();
    updateStatusStats();
    document.querySelectorAll('#grid tbody th').forEach(th => th.classList.remove('active'));
    target.classList.add('active');
    return;
  }

  // Klik di sudut (pojok kiri atas) → select all
  if (target.matches('#grid thead th:first-child')) {
    e.preventDefault();
    selection = { start: 'A1', end: addrOf(ROWS - 1, COLS - 1) };
    activeCell = 'A1';
    updateSelectionUI();
    updateStatusStats();
    return;
  }
}, true);   // capture phase
/* =========================================================
   FILTER HEADER ROW — baris yang jadi header filter
   Auto-detect: baris dengan data terbanyak (dalam 10 baris pertama)
   ========================================================= */
function getFilterHeaderRow() {
  // 🎯 Kalau sudah di-set (dari toggle / right-click), pakai itu — JANGAN baca selection lagi
  if (filterHeaderRow !== null && filterHeaderRow >= 0) return filterHeaderRow;

  // Fallback: auto-detect baris dengan data terbanyak
  const cells = sheet().cells || {};
  let bestRow = 0, bestCount = 0;
  for (let r = 0; r < 10; r++) {
    let count = 0;
    for (let c = 0; c < COLS; c++) {
      const cell = cells[addrOf(r, c)];
      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') count++;
    }
    if (count > bestCount) { bestCount = count; bestRow = r; }
  }
  return bestRow;
}
/* =========================================================
   OPEN FILTER POPUP — anchor di cell header
   ========================================================= */
function openFilterPopupAtCell(colIdx, anchorEl) {
  closeFilterPopup();
  if (!anchorEl) return;

  const rect = anchorEl.getBoundingClientRect();
  const headerRow = getFilterHeaderRow();

  // Kumpulkan nilai UNIQUE dari kolom (skip header row)
  const values = new Map();
  for (const a of Object.keys(sheet().cells || {})) {
    const p = parseAddr(a);
    if (!p || p.c !== colIdx) continue;
    if (p.r === headerRow) continue;

    const cl = sheet().cells[a];
    if (!cl) continue;

    // 🎯 Baca nilai HF kalau formula
    let v;
    if (cl.f) {
      try { v = hfGet(p.r, p.c); } catch { v = ''; }
    } else {
      v = cl.v;
    }
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

  // Buat popup
  const popup = document.createElement('div');
  popup.className = 'fx-filter-popup';

  const PW = 320, PH = 420;
  let left = rect.left;
  let top = rect.bottom + 4;

  if (left + PW > window.innerWidth - 10) left = window.innerWidth - PW - 10;
  if (left < 10) left = 10;
  if (top + PH > window.innerHeight - 10) {
    const above = rect.top - PH - 4;
    top = above >= 10 ? above : Math.max(10, window.innerHeight - PH - 10);
  }

  popup.style.left = left + 'px';
  popup.style.top = top + 'px';

  const headerName = sheet().cells[addrOf(headerRow, colIdx)]?.v || '';
  const title = headerName
    ? `${colName(colIdx)} — ${String(headerName).slice(0, 30)}`
    : `Filter Kolom ${colName(colIdx)}`;

  popup.innerHTML = `
    <div class="fx-fp-header">
      <i class="fa-solid fa-filter"></i> ${escapeHtml(title)}
      <span class="fx-fp-count">${sortedKeys.length} nilai</span>
    </div>
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
    saveFilterToLocal();
    closeFilterPopup();
    window._lastRenderKey = null;
    renderVirtual(false);
    toast(`✔ Filter ${colName(colIdx)} dibersihkan`);
  };

  popup.querySelector('#fx-fp-apply').onclick = () => {
    sheet().filters ||= {};
    if (selected.size === sortedKeys.length) {
      delete sheet().filters[colIdx];
    } else {
      sheet().filters[colIdx] = { values: [...selected] };
    }
    invalidateRowMap();
    saveFilterToLocal();
    closeFilterPopup();
    window._lastRenderKey = null;
    renderVirtual(false);
    toast(`✔ Filter ${colName(colIdx)} diterapkan`);
  };

  setTimeout(() => {
    document.addEventListener('mousedown', outsideFilterHandler, true);
  }, 0);
}
/* =========================================================
   FIX FILTER TOGGLE — single source of truth
   ========================================================= */
(function initFilterToggleFix() {
  const oldBtn = document.getElementById('btn-filter-toggle');
  if (!oldBtn) {
    console.warn('btn-filter-toggle tidak ditemukan');
    return;
  }

  // 🔑 Clone tombol → semua addEventListener lama otomatis hilang
  const btn = oldBtn.cloneNode(true);
  oldBtn.parentNode.replaceChild(btn, oldBtn);

  btn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();

    filterMode = !filterMode;

    // Style visual
    btn.style.background = filterMode
      ? 'linear-gradient(135deg,#10b981,#059669)' : '';
    btn.style.color = filterMode ? '#fff' : '';

    // 🎯 Kalau aktif: baris header = baris cell yang sedang dipilih
    if (filterMode) {
      const p = parseAddr(selection.start);
      filterHeaderRow = p ? p.r : 0;
      console.log('🟢 Filter ON, header row =', filterHeaderRow);
    } else {
      filterHeaderRow = null;
      console.log('🔴 Filter OFF');
    }

    // Reset cache & re-render
    invalidateRowMap();
    window._lastRenderKey = null;
    renderVirtual(false);

    // Debug: cek tombol ▼ benar-benar ke-render
    setTimeout(() => {
      const btns = document.querySelectorAll('.fx-filter-btn-inline');
      console.log('🔽 Jumlah tombol ▼ yang ke-render:', btns.length);
      if (filterMode && btns.length === 0) {
        toast('⚠ Filter aktif tapi 0 tombol render — cek console');
      }
    }, 100);

    toast(filterMode
      ? `🔽 Filter aktif di baris ${filterHeaderRow + 1}`
      : 'Filter dinonaktifkan');
  });

  console.log('✅ Filter toggle di-rebind (single listener)');
})();

(function rebindFilterToggle() {
  const oldBtn = document.getElementById('btn-filter-toggle');
  if (!oldBtn) return;

  const btn = oldBtn.cloneNode(true);
  oldBtn.parentNode.replaceChild(btn, oldBtn);

  btn.addEventListener('click', e => {
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();

    filterMode = !filterMode;

    btn.style.background = filterMode
      ? 'linear-gradient(135deg,#10b981,#059669)' : '';
    btn.style.color = filterMode ? '#fff' : '';

    if (filterMode) {
      const p = parseAddr(selection.start);
      filterHeaderRow = p ? p.r : 0;
    } else {
      filterHeaderRow = null;
    }

    saveFilterStateToSheet()
    invalidateRowMap();
    window._lastRenderKey = null;   // untuk renderVirtual versi lama (jaga-jaga)
    renderVirtual(false);

    setTimeout(() => {
      const n = document.querySelectorAll('.fx-filter-btn-inline').length;
      console.log(`🔽 Filter ${filterMode ? 'ON' : 'OFF'} | row=${filterHeaderRow} | tombol=${n}`);
    }, 100);

    toast(filterMode
      ? `🔽 Filter aktif — baris ${filterHeaderRow + 1}`
      : 'Filter dinonaktifkan');
  }, true);

  console.log('✅ Filter toggle di-rebind');
})();
/* =========================================================
   CROSS-SHEET INPUT BINDINGS
   ========================================================= */
(function initCrossSheetInput() {
  const input = document.getElementById('cs-formula-input');
  if (!input) return;

  // 🎯 Cegah keydown bubble ke document (biar gak buka editor di A1)
  input.addEventListener('keydown', e => {
    e.stopPropagation();

    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      finishCrossSheetEdit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelCrossSheetEdit();
    }
  }, true);   // capture phase

  // Sync ke crossSheetRef saat user ketik
  input.addEventListener('input', () => {
    if (crossSheetRef) {
      crossSheetRef.formula = input.value;
    }
  });

  // 🎯 Cegah mousedown di input "bocor" ke grid (yang bikin cell A1 ter-edit)
  input.addEventListener('mousedown', e => {
    e.stopPropagation();
  }, true);

  console.log('✅ Cross-sheet input bindings installed');
})();

/* =========================================================
   F4 — CYCLE ABSOLUTE REFERENCE ($)
   ========================================================= */
function cycleAbsoluteRef(input) {
  const val = input.value;
  const pos = input.selectionStart ?? val.length;

  // Cari ref di sekitar cursor: A1, $A$1, A$1, $A1
  // Range bisa 1-3 huruf kolom, angka row
  const before = val.slice(0, pos);
  const after = val.slice(pos);

  // Cari ref yang berakhir tepat di cursor
  const m = /(\$?)([A-Z]{1,3})(\$?)(\d+)$/.exec(before);
  if (!m) return false;

  const [full, colDollar, col, rowDollar, row] = m;
  const startPos = pos - full.length;

  // Cycle: A1 → $A$1 → A$1 → $A1 → A1
  let newRef;
  if (!colDollar && !rowDollar)      newRef = `$${col}$${row}`;
  else if (colDollar && rowDollar)   newRef = `${col}$${row}`;
  else if (!colDollar && rowDollar)  newRef = `$${col}${row}`;
  else                                newRef = `${col}${row}`;

  input.value = val.slice(0, startPos) + newRef + after;
  const newPos = startPos + newRef.length;
  input.setSelectionRange(newPos, newPos);
  return true;
}

// Bind F4 ke editor & formula bar
document.addEventListener('keydown', e => {
  if (e.key !== 'F4') return;

  const fi = document.getElementById('formula-input');
  const target = document.activeElement;

  if (editor && editor.input === target) {
    e.preventDefault();
    cycleAbsoluteRef(editor.input);
  } else if (fi && fi === target) {
    e.preventDefault();
    cycleAbsoluteRef(fi);
  }
}, true);
/* =========================================================
   ADJUST FORMULA REFS SAAT FILL/COPY
   ========================================================= */
function adjustFormula(formula, dRow, dCol) {
  if (!formula || (dRow === 0 && dCol === 0)) return formula;

  // Match: prefix(non-alnum) + optional$ + col(1-3) + optional$ + row(digits)
  // Not followed by letter/digit (avoid LOG10 false positive)
  return formula.replace(
    /(^|[^A-Za-z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/g,
    (match, prefix, colDollar, col, rowDollar, row) => {
      let newCol = col;
      let newRow = parseInt(row, 10);

      if (colDollar !== '$' && dCol !== 0) {
        let c = 0;
        for (const ch of col) c = c * 26 + (ch.charCodeAt(0) - 64);
        c = c - 1 + dCol;
        if (c < 0) c = 0;
        newCol = colName(c);
      }
      if (rowDollar !== '$' && dRow !== 0) {
        newRow = newRow + dRow;
        if (newRow < 1) newRow = 1;
      }
      return prefix + colDollar + newCol + rowDollar + newRow;
    }
  );
}
/* =========================================================
   FILL HANDLE — drag untuk copy formula seperti Excel
   ========================================================= */
let fillHandleEl = null;
let fillDrag = null;

function ensureFillHandle() {
  if (fillHandleEl) return fillHandleEl;
  fillHandleEl = document.createElement('div');
  fillHandleEl.id = 'fill-handle';
  fillHandleEl.className = 'fill-handle';
  fillHandleEl.addEventListener('mousedown', startFillDrag, true);
  document.body.appendChild(fillHandleEl);
  return fillHandleEl;
}

function showFillHandle() {
  const h = ensureFillHandle();

  if (editor || currentAccessMode === 'view') {
    h.style.display = 'none';
    return;
  }

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) { h.style.display = 'none'; return; }

  const r1 = Math.max(p1.r, p2.r);
  const c1 = Math.max(p1.c, p2.c);
  const td = document.querySelector(`#grid td[data-addr="${addrOf(r1, c1)}"]`);

  if (!td) { h.style.display = 'none'; return; }

  const rect = td.getBoundingClientRect();
  const gridRect = document.getElementById('grid-wrap').getBoundingClientRect();

  // Sembunyikan kalau td di luar viewport grid
  if (rect.bottom < gridRect.top || rect.top > gridRect.bottom ||
      rect.right < gridRect.left || rect.left > gridRect.right) {
    h.style.display = 'none';
    return;
  }

  h.style.display = 'block';
  h.style.left = (rect.right - 5) + 'px';
  h.style.top  = (rect.bottom - 5) + 'px';
}

function hideFillHandle() {
  if (fillHandleEl) fillHandleEl.style.display = 'none';
}

function startFillDrag(e) {
  e.preventDefault();
  e.stopPropagation();

  const p1 = parseAddr(selection.start);
  const p2 = parseAddr(selection.end);
  if (!p1 || !p2) return;

  fillDrag = {
    r0: Math.min(p1.r, p2.r), r1: Math.max(p1.r, p2.r),
    c0: Math.min(p1.c, p2.c), c1: Math.max(p1.c, p2.c),
    targetR: Math.max(p1.r, p2.r),
    targetC: Math.max(p1.c, p2.c)
  };

  const onMove = ev => {
    const el = document.elementFromPoint(ev.clientX, ev.clientY);
    const td = el?.closest('td[data-addr]');
    if (!td) return;
    const p = parseAddr(td.dataset.addr);
    if (!p) return;
    fillDrag.targetR = p.r;
    fillDrag.targetC = p.c;
    highlightFillRange(fillDrag);
  };

  const onUp = () => {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    clearFillHighlight();
    if (fillDrag) {
      const copy = fillDrag;
      fillDrag = null;
      applyFill(copy);
    }
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function highlightFillRange(f) {
  document.querySelectorAll('#grid td.fill-preview').forEach(td => td.classList.remove('fill-preview'));
  const tr = Math.max(f.targetR, f.r1);
  const tc = Math.max(f.targetC, f.c1);
  for (let r = f.r0; r <= tr; r++) {
    for (let c = f.c0; c <= tc; c++) {
      const td = document.querySelector(`#grid td[data-addr="${addrOf(r, c)}"]`);
      if (td) td.classList.add('fill-preview');
    }
  }
}

function clearFillHighlight() {
  document.querySelectorAll('#grid td.fill-preview').forEach(td => td.classList.remove('fill-preview'));
}

function applyFill(f) {
  const tR = f.targetR;
  const tC = f.targetC;

  // Cek apakah ada perubahan
  if (tR >= f.r0 && tR <= f.r1 && tC >= f.c0 && tC <= f.c1) return;

  pushUndo();

  const srcH = f.r1 - f.r0 + 1;
  const srcW = f.c1 - f.c0 + 1;

  // Range akhir = union dari source + target
  const endR0 = Math.min(tR, f.r0);
  const endR1 = Math.max(tR, f.r1);
  const endC0 = Math.min(tC, f.c0);
  const endC1 = Math.max(tC, f.c1);

  for (let r = endR0; r <= endR1; r++) {
    for (let c = endC0; c <= endC1; c++) {
      // Skip cell source
      if (r >= f.r0 && r <= f.r1 && c >= f.c0 && c <= f.c1) continue;

      // 🎯 Modulo positif — handle arah negatif
      const offR = ((r - f.r0) % srcH + srcH) % srcH;
      const offC = ((c - f.c0) % srcW + srcW) % srcW;
      const srcR = f.r0 + offR;
      const srcC = f.c0 + offC;

      const srcAddr = addrOf(srcR, srcC);
      const srcCell = sheet().cells[srcAddr];
      const dstAddr = addrOf(r, c);
      const dRow = r - srcR;
      const dCol = c - srcC;
      const tp = parseAddr(dstAddr);

      if (!srcCell) {
        delete sheet().cells[dstAddr];
        if (tp) hfSet(tp.r, tp.c, null);
        markCellDirty(dstAddr);
        continue;
      }

      const newCell = JSON.parse(JSON.stringify(srcCell));

      if (newCell.f) {
        newCell.f = adjustFormula(newCell.f, dRow, dCol);
        delete newCell.v;
        if (tp) setHFCellSafe(tp.r, tp.c, newCell.f);
      } else if (newCell.v !== undefined) {
        if (tp) setHFCellSafe(tp.r, tp.c, newCell.v);
      }

      sheet().cells[dstAddr] = newCell;
      markCellDirty(dstAddr);
    }
  }

  selection = { start: addrOf(endR0, endC0), end: addrOf(endR1, endC1) };
  activeCell = addrOf(f.r0, f.c0);

  clearHFCache();
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual(false);
  updateSelectionUI();

  toast('📋 Fill selesai');
}
function doPasteSpecial(mode) {
  if (!clipboard?.length) return toast('Clipboard kosong');
  if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');

  const p = parseAddr(activeCell);
  if (!p) return;
  const origin = parseAddr(clipboard[0].addr);
  if (!origin) return;

  pushUndo();

  for (const { addr, cell } of clipboard) {
    const sp = parseAddr(addr);
    const target = addrOf(p.r + (sp.r - origin.r), p.c + (sp.c - origin.c));
    const tp = parseAddr(target);
    if (!tp) continue;

    if (mode === 'value') {
      // 🎯 Paste nilai HF-evaluated (rumus → value)
      let v = null;
      if (cell) {
        if (cell.f) {
          try { v = hfGet(sp.r, sp.c); } catch { v = null; }
        } else {
          v = cell.v;
        }
      }
      if (v === null || v === undefined || v === '') {
        delete sheet().cells[target];
      } else {
        sheet().cells[target] = { v };
      }
      hfSet(tp.r, tp.c, v ?? null);
    } else if (mode === 'formula') {
      if (!cell) { delete sheet().cells[target]; hfSet(tp.r, tp.c, null); continue; }
      const dst = sheet().cells[target] || {};
      if (cell.f) {
        dst.f = cell.f;
        delete dst.v;
        hfSet(tp.r, tp.c, cell.f);
      } else {
        dst.v = cell.v;
        delete dst.f;
        hfSet(tp.r, tp.c, cell.v);
      }
      sheet().cells[target] = dst;
    } else if (mode === 'format') {
      if (cell?.s) {
        const dst = sheet().cells[target] || {};
        dst.s = JSON.parse(JSON.stringify(cell.s));
        sheet().cells[target] = dst;
      }
    }
    markCellDirty(target);
  }

  clearHFCache();
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  renderVirtual(false);
  toast('📥 Paste ' + mode);
}
const wbNameInput = document.getElementById('wb-name');
if (wbNameInput) {
  wbNameInput.addEventListener('input', e => {
    wbName = e.target.value;
    markMetaDirty('name');
  });

  wbNameInput.addEventListener('blur', () => {
    if (wbName && dirtyMeta.name) scheduleSave();
  });
}
/* =========================================================
   WB NAME INPUT BINDINGS
   ========================================================= */
(function bindWbNameInput() {
  const el = document.getElementById('wb-name');
  if (!el) return;

  // Real-time update
  el.addEventListener('input', e => {
    wbName = e.target.value;
    markMetaDirty('name');
  });

  // Blur → save
  el.addEventListener('blur', e => {
    wbName = e.target.value;
    markMetaDirty('name');
    scheduleSave();
    console.log('💾 wbName saved:', wbName);
  });

  // Enter → blur
  el.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.target.blur();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.target.value = wbName;   // restore
      e.target.blur();
    }
  });

  console.log('✅ wb-name bindings installed');
})();

/* =========================================================
   ANALYZE — Excel-like Dashboard + Pivot + Charts
   ========================================================= */
(function initAnalyze() {
  const btn = document.getElementById('btn-analyze');
  const modal = document.getElementById('modal-analyze');
  if (!btn || !modal) return;

  const nameInput = document.getElementById('an-sheet-name');

  btn.addEventListener('click', () => {
    nameInput.value = 'Dashboard';
    modal.classList.remove('hidden');
  });

  document.getElementById('btn-an-cancel').addEventListener('click', () => {
    modal.classList.add('hidden');
  });

  document.getElementById('btn-an-run').addEventListener('click', async () => {
    if (currentAccessMode === 'view') return toast('👁 Mode lihat saja');
    const dashName = nameInput.value.trim() || 'Dashboard';
    modal.classList.add('hidden');
    await analyzeWorkbook(dashName);
  });
})();

/* =========================================================
   MAIN: analyzeWorkbook
   ========================================================= */
async function analyzeWorkbook(dashName = 'Dashboard') {
  if (!wbData?.sheets?.length) return toast('Workbook kosong');

  showLoading('Menganalisa…');
  await yieldToUI();

  try {
    // 1. Load semua sheet
    updateLoading('Memuat semua sheet…');
    await yieldToUI();
    for (let k = 0; k < wbData.sheets.length; k++) {
      if (!loadedSheetIdxs.has(k)) {
        updateLoading(`Memuat "${wbData.sheets[k].name}"…`);
        await lazyLoadSheet(k);
      }
    }

    // 2. Kumpulkan records
    updateLoading('Menggabungkan data…');
    await yieldToUI();
    const records = collectAllRecords();
    if (!records.length) {
      hideLoading();
      return toast('⚠ Tidak ada data untuk dianalisa');
    }
    console.log(`📊 Records: ${records.length}`);

    // 3. Analisa kolom
    const cols = analyzeColumns(records);
    console.log('📋 Columns:', cols);

    const pivotCols = cols.filter(c => c.type === 'category' && c.unique <= 30 && c.unique > 1);
    const valueCols = cols.filter(c => c.type === 'number');
    const dateCols = cols.filter(c => c.type === 'date');

    // 4. Buat Dashboard
    updateLoading('Membuat dashboard…');
    await yieldToUI();
    const dashIdx = createDashboardSheet(records, cols, pivotCols, valueCols, dashName);

    // 5. Buat Pivot sheets (max 5)
    const maxPivots = 5;
    const pivotResults = [];
    for (let i = 0; i < Math.min(pivotCols.length, maxPivots); i++) {
      const pcol = pivotCols[i];
      updateLoading(`Membuat pivot: ${pcol.name}…`);
      await yieldToUI();
      const vcol = valueCols[0]?.name || null;
      const pivotIdx = createPivotSheet(records, pcol.name, vcol, 'Dashboard', dashName);
      if (pivotIdx >= 0) pivotResults.push({ name: pcol.name, idx: pivotIdx });
    }

    // 6. Buat Trend sheet
    if (dateCols.length && valueCols.length) {
      updateLoading('Membuat trend…');
      await yieldToUI();
      createTrendSheet(records, dateCols[0].name, valueCols[0].name, dashName);
    }

    // 7. Selesai
    markMetaDirty('structure');
    scheduleSave();
    hideLoading();

    await switchSheet(dashIdx);
    toast(`✔ Dashboard + ${pivotResults.length} pivot dibuat`);

  } catch (e) {
    hideLoading();
    console.error(e);
    toast('⚠ ' + e.message);
  }
}

/* =========================================================
   COLLECT RECORDS dari semua sheet
   ========================================================= */
function collectAllRecords() {
  const records = [];

  wbData.sheets.forEach((sh, shIdx) => {
    const cells = sh.cells || {};
    if (!Object.keys(cells).length) return;

    let maxR = -1, maxC = -1;
    for (const a of Object.keys(cells)) {
      const p = parseAddr(a);
      if (!p) continue;
      if (p.r > maxR) maxR = p.r;
      if (p.c > maxC) maxC = p.c;
    }
    if (maxR < 1) return;

    // 🎯 Skip sheet tanpa header
    let headerCount = 0;
    for (let c = 0; c <= maxC; c++) {
      const cl = cells[addrOf(0, c)];
      if (cl && cl.v !== undefined && cl.v !== null && cl.v !== '') headerCount++;
    }
    if (headerCount < 2) {
      console.log(`⏭ Skip sheet "${sh.name}" — tidak ada header`);
      return;
    }

    // Ambil headers
    const headers = [];
    for (let c = 0; c <= maxC; c++) {
      const cl = cells[addrOf(0, c)];
      let h = '';
      if (cl) {
        if (cl.f && shIdx === activeSheetIdx) {
          try { h = displayVal(hfGet(0, c)); } catch { h = ''; }
        } else {
          h = displayVal(cl.v);
        }
      }
      headers.push(h || colName(c));
    }

    // Data rows
    for (let r = 1; r <= maxR; r++) {
      const row = { __sheet: sh.name };
      let hasData = false;
      for (let c = 0; c <= maxC; c++) {
        const addr = addrOf(r, c);
        const cl = cells[addr];
        if (!cl) continue;

        let v;
        if (cl.f && shIdx === activeSheetIdx) {
          try { v = hfGet(r, c); } catch { v = null; }
        } else {
          v = cl.v;
        }

        if (v !== null && v !== undefined && v !== '') {
          row[headers[c]] = v;
          hasData = true;
        }
      }
      if (hasData) records.push(row);
    }
  });

  return records;
}

/* =========================================================
   ANALYZE COLUMNS
   Deteksi tipe tiap kolom: number / date / category / text
   ========================================================= */
function analyzeColumns(records) {
  const keys = new Set();
  records.forEach(r => {
    Object.keys(r).forEach(k => {
      if (!k.startsWith('__')) keys.add(k);
    });
  });

  const cols = [];
  for (const k of keys) {
    const values = records.map(r => r[k]).filter(v => v !== null && v !== undefined && v !== '');
    if (!values.length) continue;

    const sample = values.slice(0, 200);
    const numCount = sample.filter(v => !isNaN(Number(v)) && String(v).trim() !== '').length;
    const dateCount = sample.filter(v => isDateLike(v)).length;
    const unique = new Set(values.map(v => String(v))).size;

    let type = 'text';
    if (dateCount / sample.length > 0.7) type = 'date';
    else if (numCount / sample.length > 0.7) type = 'number';
    else if (unique <= Math.max(50, records.length * 0.5)) type = 'category';

    cols.push({ name: k, type, unique, sample: values.slice(0, 3) });
  }
  return cols;
}

function isDateLike(v) {
  if (v == null) return false;
  const s = String(v).trim();
  if (!s) return false;
  // ISO date
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  // Excel serial (50000-50000)
  if (/^\d{5,6}(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (n > 20000 && n < 60000) return true;
  }
  return false;
}

/* =========================================================
   CREATE DASHBOARD SHEET
   ========================================================= */
function createDashboardSheet(records, cols, pivotCols, valueCols, dashName) {
  const cells = {};
  let r = 0;
  const R = (rr, cc, v, style) => {
    const a = addrOf(rr, cc);
    cells[a] = { v };
    if (style) cells[a].s = style;
  };
  const put = (v, style) => { R(r, 0, v, style); };
  const skip = (n = 1) => { r += n; };

  const totalRecords = records.length;
  const sheetsIncluded = new Set(records.map(x => x.__sheet)).size;
  const dateCol = cols.find(c => c.type === 'date')?.name;
  const mainValue = valueCols[0]?.name;

  /* ============== TITLE ============== */
  R(r, 0, '📊  WORKBOOK DASHBOARD', {
    bold: true, fontSize: 18, color: '#065f46'
  });
  r++;
  R(r, 0, `Dibuat: ${new Date().toLocaleString('id-ID')}`, {
    fontSize: 10, color: '#64748b', italic: true
  });
  r += 2;

  /* ============== KPI CARDS (visual) ============== */
  // Tulis label KPI dalam 4 kartu
  const kpiRow = r;
  const kpiData = [
    { label: 'TOTAL RECORDS', value: totalRecords.toLocaleString('id-ID'), bg: '#dbeafe', fg: '#1e40af' },
    { label: 'SHEET DIANALISA', value: sheetsIncluded, bg: '#d1fae5', fg: '#065f46' },
    { label: 'TOTAL KOLOM', value: cols.length, bg: '#fef3c7', fg: '#92400e' },
  ];

  const kpiCols = 4;   // 4 kartu berjejer
  kpiData.forEach((k, i) => {
    const c = i * 2;
    R(kpiRow, c, k.label, { bold: true, fontSize: 9, color: k.fg, bg: k.bg, align: 'center' });
    R(kpiRow + 1, c, k.value, { bold: true, fontSize: 20, color: k.fg, bg: k.bg, align: 'center' });
    R(kpiRow + 2, c, '', { bg: k.bg });
    // padding kolom sebelah
    R(kpiRow, c + 1, '', { bg: '#ffffff' });
    R(kpiRow + 1, c + 1, '', { bg: '#ffffff' });
    R(kpiRow + 2, c + 1, '', { bg: '#ffffff' });
  });
  r += 4;

  /* ============== RINGKASAN KOLOM NUMERIK ============== */
  if (valueCols.length) {
    R(r, 0, '📈  STATISTIK KOLOM NUMERIK', { bold: true, fontSize: 12, color: '#059669' });
    r++;

    const hdr = ['Kolom', 'Total', 'Rata-rata', 'Min', 'Max'];
    hdr.forEach((h, i) => R(r, i, h, {
      bold: true, bg: '#ecfdf5', color: '#065f46', align: 'center',
      borders: { top: true, bottom: true, left: true, right: true, style: 'thin', color: '#10b981' }
    }));
    r++;

    for (const vc of valueCols.slice(0, 8)) {
      const vals = records.map(x => Number(x[vc.name])).filter(n => !isNaN(n));
      if (!vals.length) continue;
      const sum = vals.reduce((a, b) => a + b, 0);
      const avg = sum / vals.length;
      const min = Math.min(...vals);
      const max = Math.max(...vals);

      R(r, 0, vc.name, { bold: true });
      R(r, 1, fmtNum(sum), { align: 'right' });
      R(r, 2, fmtNum(avg, 2), { align: 'right' });
      R(r, 3, fmtNum(min), { align: 'right' });
      R(r, 4, fmtNum(max), { align: 'right' });
      r++;
    }
    r++;
  }

  /* ============== TOP 5 PER CATEGORY (mini pivot) ============== */
  for (const pcol of pivotCols.slice(0, 3)) {
    R(r, 0, `🔎  TOP 5 ${pcol.name.toUpperCase()}`, { bold: true, fontSize: 12, color: '#059669' });
    r++;

    const counts = new Map();
    records.forEach(rec => {
      const k = String(rec[pcol.name] ?? '(kosong)');
      counts.set(k, (counts.get(k) || 0) + 1);
    });
    const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 5);

    R(r, 0, pcol.name, { bold: true, bg: '#ecfdf5', align: 'center' });
    R(r, 1, 'Count', { bold: true, bg: '#ecfdf5', align: 'center' });
    R(r, 2, '%', { bold: true, bg: '#ecfdf5', align: 'center' });
    r++;

    for (const [k, cnt] of sorted) {
      R(r, 0, k);
      R(r, 1, cnt, { align: 'right' });
      R(r, 2, ((cnt / totalRecords) * 100).toFixed(1) + '%', { align: 'right', color: '#64748b' });
      r++;
    }
    r++;
  }

  /* ============== FOOTER ============== */
  R(r, 0, '💡 Tip: lihat sheet "Pivot …" untuk data lengkap per kategori', {
    italic: true, fontSize: 10, color: '#94a3b8'
  });

  /* ============== CREATE SHEET ============== */
  const newSheet = {
    name: dashName,
    cells,
    colWidths: { 0: 220, 1: 120, 2: 120, 3: 120, 4: 120 },
    rowHeight: 26,
    rowHeights: {},   // set KPI rows lebih tinggi
    cf: [], charts: [], merged: [], filters: {},
    images: [], validations: [],
    freezeRow: 0, freezeCol: 0, protected: false
  };

  // KPI row height: lebih tinggi biar kartu terlihat
  newSheet.rowHeights[kpiRow] = 32;
  newSheet.rowHeights[kpiRow + 1] = 40;
  newSheet.rowHeights[kpiRow + 2] = 8;

  // Charts untuk mini pivot
  newSheet.charts = [];

  wbData.sheets.push(newSheet);
  const newIdx = wbData.sheets.length - 1;
  loadedSheetIdxs.add(newIdx);
  return newIdx;
}

/* =========================================================
   CREATE PIVOT SHEET
   ========================================================= */
function createPivotSheet(records, groupCol, valueCol, prefix, dashName) {
  const counts = new Map();

  records.forEach(rec => {
    const k = String(rec[groupCol] ?? '(kosong)');
    if (!counts.has(k)) counts.set(k, { count: 0, sum: 0, values: [] });
    const g = counts.get(k);
    g.count++;
    if (valueCol) {
      const n = Number(rec[valueCol]);
      if (!isNaN(n)) {
        g.sum += n;
        g.values.push(n);
      }
    }
  });

  const rows = [...counts.entries()].map(([k, g]) => ({
    key: k,
    count: g.count,
    sum: g.sum,
    avg: g.values.length ? g.sum / g.values.length : 0
  })).sort((a, b) => b.count - a.count);

  const cells = {};
  let r = 0;
  const R = (rr, cc, v, style) => {
    const a = addrOf(rr, cc);
    cells[a] = { v };
    if (style) cells[a].s = style;
  };

  R(r, 0, `📊  PIVOT: ${groupCol}`, { bold: true, fontSize: 16, color: '#065f46' });
  r++;
  R(r, 0, `Total ${records.length} baris · ${rows.length} kategori`, {
    italic: true, fontSize: 10, color: '#64748b'
  });
  r += 2;

  // Header
  const hdr = [groupCol, 'Count', '%'];
  if (valueCol) {
    hdr.push(`Sum(${valueCol})`, `Avg(${valueCol})`);
  }
  hdr.forEach((h, i) => R(r, i, h, {
    bold: true, bg: '#d1fae5', color: '#065f46', align: 'center',
    borders: { top: true, bottom: true, left: true, right: true, style: 'medium', color: '#10b981' }
  }));
  r++;

  // Data
  rows.forEach((row, i) => {
    const isAlt = i % 2 === 1;
    const bg = isAlt ? '#f8fafc' : '#ffffff';
    R(r, 0, row.key, { bg, bold: true });
    R(r, 1, row.count, { bg, align: 'right' });
    R(r, 2, ((row.count / records.length) * 100).toFixed(1) + '%', {
      bg, align: 'right', color: '#64748b'
    });
    if (valueCol) {
      R(r, 3, fmtNum(row.sum), { bg, align: 'right' });
      R(r, 4, fmtNum(row.avg, 2), { bg, align: 'right' });
    }
    r++;
  });

  // Grand total
  r++;
  R(r, 0, 'TOTAL', {
    bold: true, bg: '#d1fae5', color: '#065f46',
    borders: { top: true, bottom: true, left: true, right: true, style: 'medium', color: '#10b981' }
  });
  R(r, 1, records.length, { bold: true, bg: '#d1fae5', align: 'right', color: '#065f46' });
  R(r, 2, '100%', { bold: true, bg: '#d1fae5', align: 'right', color: '#065f46' });
  if (valueCol) {
    const totalSum = rows.reduce((a, b) => a + b.sum, 0);
    R(r, 3, fmtNum(totalSum), { bold: true, bg: '#d1fae5', align: 'right', color: '#065f46' });
    R(r, 4, '', { bg: '#d1fae5' });
  }

  // Chart data — top 10 di sebelah kanan (kolom G)
  const chartStartCol = 6;
  const chartTop = 4;
  R(chartTop, chartStartCol, 'Label', { bold: true, bg: '#f1f5f9' });
  R(chartTop, chartStartCol + 1, 'Value', { bold: true, bg: '#f1f5f9' });
  const top10 = rows.slice(0, 10);
  top10.forEach((row, i) => {
    R(chartTop + 1 + i, chartStartCol, String(row.key).slice(0, 30));
    R(chartTop + 1 + i, chartStartCol + 1, row.count);
  });

  // Hitung range chart untuk Chart.js
  const chartRange = addrOf(chartTop, chartStartCol) + ':' + addrOf(chartTop + top10.length, chartStartCol + 1);
  const chartX = 200;   // posisi di layar
  const chartY = 80;

  const newSheet = {
    name: `${prefix} — ${groupCol}`.slice(0, 30),
    cells,
    colWidths: { 0: 220, 1: 100, 2: 80, 3: 140, 4: 140 },
    rowHeight: 26,
    cf: [], charts: [], merged: [], filters: {},
    images: [], validations: [],
    freezeRow: 1, freezeCol: 0, protected: false
  };

  // Chart object
  newSheet.charts.push({
    id: 'c' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
    type: 'bar',
    range: chartRange,
    x: chartX,
    y: chartY,
    w: 460,
    h: 320
  });

  wbData.sheets.push(newSheet);
  const newIdx = wbData.sheets.length - 1;
  loadedSheetIdxs.add(newIdx);
  return newIdx;
}

/* =========================================================
   CREATE TREND SHEET (time-series)
   ========================================================= */
function createTrendSheet(records, dateCol, valueCol, dashName) {
  // Group by month
  const byMonth = new Map();

  records.forEach(rec => {
    const dateRaw = rec[dateCol];
    if (!dateRaw) return;

    let d;
    const s = String(dateRaw).trim();
    if (/^\d{5,6}(\.\d+)?$/.test(s)) {
      // Excel serial → date (base 1899-12-30)
      d = new Date(Date.UTC(1899, 11, 30) + Number(s) * 86400000);
    } else {
      d = new Date(s);
    }
    if (isNaN(d.getTime())) return;

    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    if (!byMonth.has(key)) byMonth.set(key, { count: 0, sum: 0 });
    const g = byMonth.get(key);
    g.count++;
    const n = Number(rec[valueCol]);
    if (!isNaN(n)) g.sum += n;
  });

  const sortedMonths = [...byMonth.keys()].sort();

  const cells = {};
  let r = 0;
  const R = (rr, cc, v, style) => {
    const a = addrOf(rr, cc);
    cells[a] = { v };
    if (style) cells[a].s = style;
  };

  R(r, 0, `📈  TREND BULANAN — ${valueCol}`, { bold: true, fontSize: 16, color: '#065f46' });
  r += 2;

  R(r, 0, 'Bulan', { bold: true, bg: '#dbeafe', color: '#1e40af', align: 'center' });
  R(r, 1, 'Count', { bold: true, bg: '#dbeafe', color: '#1e40af', align: 'center' });
  R(r, 2, `Sum(${valueCol})`, { bold: true, bg: '#dbeafe', color: '#1e40af', align: 'center' });
  r++;

  const chartTop = r;
  for (const month of sortedMonths) {
    const g = byMonth.get(month);
    R(r, 0, month, { bold: true });
    R(r, 1, g.count, { align: 'right' });
    R(r, 2, fmtNum(g.sum), { align: 'right' });
    r++;
  }

  const chartRange = addrOf(chartTop, 0) + ':' + addrOf(r - 1, 2);

  const newSheet = {
    name: `Trend — ${valueCol}`.slice(0, 30),
    cells,
    colWidths: { 0: 140, 1: 100, 2: 160 },
    rowHeight: 26,
    cf: [], charts: [], merged: [], filters: {},
    images: [], validations: [],
    freezeRow: 1, freezeCol: 0, protected: false
  };

  newSheet.charts.push({
    id: 'ctrend' + Date.now(),
    type: 'line',
    range: chartRange,
    x: 320,
    y: 80,
    w: 640,
    h: 360
  });

  wbData.sheets.push(newSheet);
  const newIdx = wbData.sheets.length - 1;
  loadedSheetIdxs.add(newIdx);
  return newIdx;
}

/* =========================================================
   Helper: format angka
   ========================================================= */
function fmtNum(n, dec = 0) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (!isFinite(n)) return '—';
  return Number(n).toLocaleString('id-ID', {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec
  });
}

/* =========================================================
   HIDE / SHOW ROWS & COLUMNS
   ========================================================= */
function applyHiddenRowsCols() {
  const sh = sheet();
  if (!sh) return;

  const hiddenCols = new Set(sh.hiddenCols || []);
  const hiddenRows = new Set(sh.hiddenRows || []);

  // 🎯 Sembunyikan kolom — thead + semua td di tbody
  const theadThs = document.querySelectorAll('#grid thead th[data-col]');
  theadThs.forEach(th => {
    const c = parseInt(th.dataset.col, 10);
    if (hiddenCols.has(c)) {
      th.classList.add('col-hidden');
    } else {
      th.classList.remove('col-hidden');
    }
  });

  const tbody = document.getElementById('grid-body');
  if (tbody) {
    Array.from(tbody.children).forEach(tr => {
      const rowTh = tr.firstElementChild;
      const r = rowTh?.dataset?.row ? parseInt(rowTh.dataset.row, 10) : -1;

      // Sembunyikan baris
      if (r >= 0 && hiddenRows.has(r)) {
        tr.classList.add('row-hidden');
      } else {
        tr.classList.remove('row-hidden');
      }

      // Sembunyikan kolom di dalam baris
      const tds = tr.querySelectorAll('td[data-addr]');
      tds.forEach(td => {
        const col = parseInt(td.dataset.col, 10);
        if (hiddenCols.has(col)) {
          td.classList.add('col-hidden');
        } else {
          td.classList.remove('col-hidden');
        }
      });
    });
  }

  // 🎯 Indikator — garis tebal di header sebelum kolom yang hidden
  theadThs.forEach(th => th.classList.remove('col-hidden-indicator'));
  hiddenCols.forEach(c => {
    const prevTh = document.querySelector(`#grid thead th[data-col="${c - 1}"]`);
    if (prevTh) prevTh.classList.add('col-hidden-indicator');
  });

  // Indikator baris
  if (tbody) {
    Array.from(tbody.children).forEach(tr => tr.querySelector('th')?.classList.remove('row-hidden-indicator'));
    hiddenRows.forEach(r => {
      const tr = rowEls[r - 1];
      if (tr?.firstElementChild) tr.firstElementChild.classList.add('row-hidden-indicator');
    });
  }
}

/* Simpan state hide ke sheet & save */
function saveHiddenState() {
  const sh = sheet();
  if (!sh) return;
  markMetaDirty('structure');
  scheduleSave();
}

/* =========================================================
   HEADER CONTEXT MENU — Hide / Unhide Row/Col
   ========================================================= */
let headerCtxMenu = null;

function closeHeaderCtxMenu() {
  if (headerCtxMenu) {
    headerCtxMenu.remove();
    headerCtxMenu = null;
  }
}

function openHeaderCtxMenu(x, y, type, index) {
  closeHeaderCtxMenu();

  const sh = sheet();
  const menu = document.createElement('div');
  menu.className = 'header-context-menu';

  if (type === 'col') {
    const isHidden = (sh.hiddenCols || []).includes(index);
    const hasHiddenNearby = (sh.hiddenCols || []).some(c => c === index - 1 || c === index + 1 || c === index);

    menu.innerHTML = `
      <button data-action="${isHidden ? 'unhide-col' : 'hide-col'}">
        <i class="fa-solid ${isHidden ? 'fa-eye' : 'fa-eye-slash'}"></i>
        ${isHidden ? 'Unhide Column ' + colName(index) : 'Hide Column ' + colName(index)}
      </button>
      ${hasHiddenNearby ? `
        <button data-action="unhide-all-cols">
          <i class="fa-solid fa-eye"></i> Unhide Semua Kolom
        </button>
      ` : ''}
    `;
  } else if (type === 'row') {
    const isHidden = (sh.hiddenRows || []).includes(index);
    const hasHiddenNearby = (sh.hiddenRows || []).some(r => r === index - 1 || r === index + 1 || r === index);

    // 🎯 Cek apakah row ini di dalam group
    const inGroup = getRowGroupAt(index);
    const canGroup = selection.start && selection.end;

    menu.innerHTML = `
      <button data-action="${isHidden ? 'unhide-row' : 'hide-row'}">
        <i class="fa-solid ${isHidden ? 'fa-eye' : 'fa-eye-slash'}"></i>
        ${isHidden ? 'Unhide Row ' + (index + 1) : 'Hide Row ' + (index + 1)}
      </button>
      ${hasHiddenNearby ? `
        <button data-action="unhide-all-rows">
          <i class="fa-solid fa-eye"></i> Unhide Semua Baris
        </button>
      ` : ''}
      <div class="ct-divider"></div>
      <button data-action="group-rows">
        <i class="fa-solid fa-layer-group"></i> Group Baris Terpilih
      </button>
      ${inGroup ? `
        <button data-action="ungroup-rows">
          <i class="fa-solid fa-object-ungroup"></i> Ungroup Baris
        </button>
      ` : ''}
    `;
  }

  document.body.appendChild(menu);
  headerCtxMenu = menu;

  // Position
  const rect = menu.getBoundingClientRect();
  if (x + rect.width > window.innerWidth - 10) x = window.innerWidth - rect.width - 10;
  if (y + rect.height > window.innerHeight - 10) y = window.innerHeight - rect.height - 10;
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';

  // Handler klik
  menu.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.action;
      const sh = sheet();
      sh.hiddenCols = sh.hiddenCols || [];
      sh.hiddenRows = sh.hiddenRows || [];

      if (action === 'hide-col') {
        if (!sh.hiddenCols.includes(index)) sh.hiddenCols.push(index);
      } else if (action === 'unhide-col') {
        sh.hiddenCols = sh.hiddenCols.filter(c => c !== index);
      } else if (action === 'unhide-all-cols') {
        sh.hiddenCols = [];
      } else if (action === 'hide-row') {
        if (!sh.hiddenRows.includes(index)) sh.hiddenRows.push(index);
      } else if (action === 'unhide-row') {
        sh.hiddenRows = sh.hiddenRows.filter(r => r !== index);
      } else if (action === 'unhide-all-rows') {
        sh.hiddenRows = [];
      } else if (action === 'group-rows') {
        // Group dari selection.start sampai selection.end
        const p1 = parseAddr(selection.start);
        const p2 = parseAddr(selection.end);
        if (!p1 || !p2) return;
        const from = Math.min(p1.r, p2.r);
        const to = Math.max(p1.r, p2.r);
        addRowGroup(from, to);

      } else if (action === 'ungroup-rows') {
        removeRowGroup(index);
      }

      saveHiddenState();
      applyHiddenRowsCols();
      closeHeaderCtxMenu();
      toast(`✔ ${action.replace(/-/g, ' ')}`);
    });
  });

  // Tutup klik luar
  setTimeout(() => {
    document.addEventListener('mousedown', function handler(e) {
      if (headerCtxMenu && !headerCtxMenu.contains(e.target)) {
        closeHeaderCtxMenu();
        document.removeEventListener('mousedown', handler, true);
      }
    }, true);
  }, 0);
}

/* =========================================================
   BIND KANAN KLIK DI HEADER
   ========================================================= */
document.addEventListener('contextmenu', e => {
  // Header kolom
  const colTh = e.target.closest('#grid thead th[data-col]');
  if (colTh) {
    e.preventDefault();
    const c = parseInt(colTh.dataset.col, 10);
    openHeaderCtxMenu(e.clientX, e.clientY, 'col', c);
    return;
  }

  // Header baris
  const rowTh = e.target.closest('#grid tbody th[data-row]');
  if (rowTh) {
    e.preventDefault();
    const r = parseInt(rowTh.dataset.row, 10);
    openHeaderCtxMenu(e.clientX, e.clientY, 'row', r);
    return;
  }
}, true);
/* =========================================================
   ROW GROUP BUTTON — Event Delegation
   ========================================================= */
document.getElementById('grid')?.addEventListener('mousedown', e => {
  const btn = e.target.closest('.row-group-btn');
  if (!btn) return;

  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();

  const from = parseInt(btn.dataset.groupFrom, 10);
  const groups = getRowGroups();
  const g = groups.find(x => x.from === from);
  if (!g) return;

  toggleRowGroupCollapse(g);
}, true);
/* =========================================================
   BIND FORMULA AUDITING
   ========================================================= */
document.getElementById('btn-show-formulas')?.addEventListener('click', toggleShowFormulas);
document.getElementById('btn-highlight-formulas')?.addEventListener('click', toggleHighlightFormulas);
document.getElementById('btn-error-check')?.addEventListener('click', runErrorCheck);
document.getElementById('btn-errors-close')?.addEventListener('click', closeErrorModal);
document.getElementById('btn-errors-cancel')?.addEventListener('click', closeErrorModal);
document.getElementById('modal-errors')?.addEventListener('click', e => {
  if (e.target.id === 'modal-errors') closeErrorModal();
});

/* Shortcut Ctrl+` untuk toggle show formulas */
document.addEventListener('keydown', e => {
  if (editor) return;
  const ctrl = e.ctrlKey || e.metaKey;

  // Ctrl+` (backtick)
  if (ctrl && (e.key === '`' || e.key === '~')) {
    e.preventDefault();
    toggleShowFormulas();
    return;
  }
}, true);

console.log('✅ Formula Auditing initialized');
/* =========================================================
   BIND ADVANCED FILTER
   ========================================================= */
document.getElementById('btn-adv-filter')?.addEventListener('click', openAdvancedFilterModal);
document.getElementById('btn-advf-close')?.addEventListener('click', closeAdvancedFilterModal);
document.getElementById('btn-advf-cancel')?.addEventListener('click', closeAdvancedFilterModal);
document.getElementById('modal-adv-filter')?.addEventListener('click', e => {
  if (e.target.id === 'modal-adv-filter') closeAdvancedFilterModal();
});

document.getElementById('btn-advf-add')?.addEventListener('click', () => {
  _advFilterDraft.push({ col: 0, op: 'eq', value: '', logic: 'AND' });
  renderAdvConditions();
});

document.getElementById('btn-advf-apply')?.addEventListener('click', () => {
  // Filter kondisi valid saja
  const valid = _advFilterDraft.filter(c => c.value !== '' || c.op === 'between');
  if (!valid.length) {
    toast('⚠ Tidak ada kondisi valid');
    return;
  }

  sheet().advancedFilters = JSON.parse(JSON.stringify(valid));
  invalidateRowMap();
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  markMetaDirty('structure');
  scheduleSave();
  renderVirtual(false);
  closeAdvancedFilterModal();
  toast(`✔ Advanced filter diterapkan (${valid.length} kondisi)`);
});

document.getElementById('btn-advf-clear')?.addEventListener('click', () => {
  sheet().advancedFilters = [];
  _advFilterDraft = [{ col: 0, op: 'eq', value: '', logic: 'AND' }];
  invalidateRowMap();
  window._lastFrozenKey = null;
  window._lastRenderKey = null;
  visibleStart = -1;
  visibleEnd = -1;
  markMetaDirty('structure');
  scheduleSave();
  renderVirtual(false);
  renderAdvConditions();
  updateAdvPreview();
  toast('✔ Advanced filter dibersihkan');
});

console.log('✅ Advanced Filter initialized');
/* =========================================================
   BIND LOCK/UNLOCK
   ========================================================= */
document.getElementById('btn-lock-cells')?.addEventListener('click', lockSelection);
document.getElementById('btn-unlock-cells')?.addEventListener('click', unlockSelection);

console.log('✅ Lock Cell initialized');