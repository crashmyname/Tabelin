/* =========================================================
   CHUNK LOADER — lazy load sheet data per block
   ========================================================= */
window.Chunks = (function () {
  const BLOCK_ROWS = 1000;

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

  /* =========================================================
   LOAD ALL BLOCKS — untuk data kecil/medium
   ========================================================= */
    async function loadAllBlocks(sheetIdx, onProgress) {
      const blocks = [];
      for (const [key] of state.blockList.entries()) {
        const [sIdx, bRow] = key.split(':').map(Number);
        if (sIdx === sheetIdx) blocks.push(bRow);
      }
      blocks.sort((a, b) => a - b);

      let loaded = 0;
      const total = blocks.length;

      // Load sequential (hindari spam network)
      for (const blockRow of blocks) {
        await ensureBlockLoaded(sheetIdx, blockRow);
        loaded++;
        if (onProgress) onProgress(loaded, total);
      }

      return { loaded, total };
    }

    /* Cek jumlah block suatu sheet */
    function getBlockCount(sheetIdx) {
      let count = 0;
      for (const [key] of state.blockList.entries()) {
        const [sIdx] = key.split(':').map(Number);
        if (sIdx === sheetIdx) count++;
      }
      return count;
    }

    /* =========================================================
      GET ALL CELLS FROM SERVER (langsung dari sheets.data)
      Untuk data yang tidak terlalu besar, lebih cepat dari chunk
      ========================================================= */
    async function loadFullSheet(sheetIdx) {
      if (!state.sheetId || !state.token) return null;

      try {
        const res = await fetch(`/api/sheets/${state.sheetId}/full`, {
          headers: { 'Authorization': 'Bearer ' + state.token }
        });
        if (!res.ok) return null;

        const data = await res.json();
        // data = { sheets: [{ cells, colWidths, ... }] }
        return data;
      } catch (e) {
        console.warn('Load full failed:', e);
        return null;
      }
    }

    /* Update state setelah full load */
    function setLoadedCells(sheetIdx, cells) {
      const blockRows = new Set();
      for (const addr of Object.keys(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(addr);
        if (!m) continue;
        const row = parseInt(m[2], 10) - 1;
        const blockRow = rowToBlock(row);
        blockRows.add(blockRow);
      }
      for (const blockRow of blockRows) {
        const key = blockKey(sheetIdx, blockRow);
        if (!state.loadedBlocks.has(key)) {
          state.loadedBlocks.set(key, {});
        }
        // Isi block dengan cells yang sesuai
        const blockCells = state.loadedBlocks.get(key);
        for (const [addr, cell] of Object.entries(cells)) {
          const m = /^([A-Z]+)(\d+)$/.exec(addr);
          if (!m) continue;
          const row = parseInt(m[2], 10) - 1;
          if (rowToBlock(row) === blockRow) {
            blockCells[addr] = cell;
          }
        }
      }
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
    loadAllBlocks,
    getBlockCount,
    loadFullSheet,
    setLoadedCells,
    BLOCK_ROWS
  };
})();