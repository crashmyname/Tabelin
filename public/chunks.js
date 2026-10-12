/* =========================================================
   CHUNK LOADER — lazy load sheet data per block
   ========================================================= */
window.Chunks = (function () {
  const BLOCK_ROWS = 256;

  const state = {
    sheetId: null,
    loadedBlocks: new Map(),      // "sIdx:blockRow" → cells object
    blockList: new Map(),         // "sIdx:blockRow" → true
    loadingBlocks: new Map(),     // "sIdx:blockRow" → Promise
    token: null
  };

  const pendingChanges = new Map(); // "sIdx:blockRow" → { addr: cell|null }

  function blockKey(sheetIdx, blockRow) {
    return sheetIdx + ':' + blockRow;
  }

  function rowToBlock(row) {
    return Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;
  }

  async function init(sheetId, token) {
    state.sheetId = sheetId;
    state.token = token;
    state.loadedBlocks.clear();
    state.blockList.clear();
    state.loadingBlocks.clear();
    pendingChanges.clear();

    try {
      const res = await fetch(`/api/sheets/${sheetId}/blocks`, {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);

      const data = await res.json();
      const blocks = data.blocks || {};
      for (const [sIdx, list] of Object.entries(blocks)) {
        for (const b of list) {
          state.blockList.set(blockKey(+sIdx, b.row), true);
        }
      }
      console.log('📦 Chunk metadata loaded:', state.blockList.size, 'blocks');
      return true;
    } catch (e) {
      console.warn('Chunk init failed:', e);
      return false;
    }
  }

  function ensureBlockLoaded(sheetIdx, blockRow) {
    const key = blockKey(sheetIdx, blockRow);
    if (state.loadedBlocks.has(key)) return Promise.resolve();
    if (state.loadingBlocks.has(key)) return state.loadingBlocks.get(key);

    // Block kosong (tidak ada di metadata) → mark as loaded
    if (!state.blockList.has(key)) {
      state.loadedBlocks.set(key, {});
      return Promise.resolve();
    }

    const promise = fetch(
      `/api/sheets/${state.sheetId}/blocks/${sheetIdx}/${blockRow}`,
      { headers: { 'Authorization': 'Bearer ' + state.token } }
    )
      .then(r => r.json())
      .then(data => {
        state.loadedBlocks.set(key, data.cells || {});
        state.loadingBlocks.delete(key);
      })
      .catch(e => {
        console.warn('Block load failed:', key, e);
        state.loadedBlocks.set(key, {});
        state.loadingBlocks.delete(key);
      });

    state.loadingBlocks.set(key, promise);
    return promise;
  }

  function getCell(sheetIdx, addr) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!m) return null;
    const row = parseInt(m[2], 10) - 1;
    const blockRow = rowToBlock(row);
    const cells = state.loadedBlocks.get(blockKey(sheetIdx, blockRow));
    if (!cells) return null;
    return cells[addr] || null;
  }

  function setCell(sheetIdx, addr, cell) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!m) return;
    const row = parseInt(m[2], 10) - 1;
    const blockRow = rowToBlock(row);
    const key = blockKey(sheetIdx, blockRow);

    // Update in-memory
    let cells = state.loadedBlocks.get(key);
    if (!cells) {
      cells = {};
      state.loadedBlocks.set(key, cells);
    }
    if (cell === null) delete cells[addr];
    else cells[addr] = cell;

    // Track untuk save
    if (!pendingChanges.has(key)) pendingChanges.set(key, {});
    const changes = pendingChanges.get(key);
    changes[addr] = cell;
  }

  async function flushChanges() {
    if (pendingChanges.size === 0) return { saved: 0 };

    const saves = [];
    for (const [key, changes] of pendingChanges.entries()) {
      const [sIdx, blockRow] = key.split(':').map(Number);
      saves.push(
        fetch(
          `/api/sheets/${state.sheetId}/blocks/${sIdx}/${blockRow}`,
          {
            method: 'PATCH',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': 'Bearer ' + state.token
            },
            body: JSON.stringify({ cells: changes })
          }
        )
        .then(r => r.ok ? 1 : 0)
        .catch(() => 0)
      );
    }

    const results = await Promise.all(saves);
    const saved = results.reduce((s, r) => s + r, 0);
    pendingChanges.clear();
    return { saved };
  }

  function getPendingCount() {
    let n = 0;
    for (const changes of pendingChanges.values()) {
      n += Object.keys(changes).length;
    }
    return n;
  }

  /* Ambil semua cell dari block yang sudah di-load (untuk rebuild HF) */
  function getLoadedCells(sheetIdx) {
    const out = {};
    for (const [key, cells] of state.loadedBlocks.entries()) {
      const [sIdx] = key.split(':').map(Number);
      if (sIdx !== sheetIdx) continue;
      Object.assign(out, cells);
    }
    return out;
  }

  function clearAll() {
    state.loadedBlocks.clear();
    state.loadingBlocks.clear();
    pendingChanges.clear();
  }

  return {
    init,
    ensureBlockLoaded,
    getCell,
    setCell,
    flushChanges,
    getPendingCount,
    getLoadedCells,
    clearAll,
    BLOCK_ROWS
  };
})();