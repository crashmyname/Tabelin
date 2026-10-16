import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

/* =========================================================
   CHECK ACCESS — versi ringan (lokal, tidak butuh import)
   ========================================================= */
async function checkAccessLight(sheetId, userId) {
  const sid = parseInt(sheetId, 10);
  if (isNaN(sid)) return null;

  const { rows } = await pool.query(`
    SELECT s.id, s.user_id,
           CASE WHEN s.user_id = $2::int THEN 'owner' ELSE sh.permission END AS access
      FROM sheets s
      LEFT JOIN shares sh ON sh.sheet_id = s.id AND sh.user_id = $2::int
     WHERE s.id = $1::int
       AND (s.user_id = $2::int OR sh.user_id IS NOT NULL)
     LIMIT 1
  `, [sid, userId]);

  return rows[0] || null;
}

/* =========================================================
   GET /:id/chunks-status
   🎯 FORCE NORMAL MODE — data selalu tampil
   Chunk mode akan diaktifkan setelah bug frontend diperbaiki
   ========================================================= */
router.get('/:id/chunks-status', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    // 🎯 Force normal mode — biar data tampil
    res.json({
      blockCount: 0,
      shouldUseChunks: false,
      source: 'normal'
    });
  } catch (e) {
    console.error('chunks-status error:', e);
    res.status(200).json({
      blockCount: 0,
      shouldUseChunks: false,
      source: 'fallback'
    });
  }
});

/* =========================================================
   GET /:id/blocks — metadata block (TANPA pg_column_size)
   ========================================================= */
router.get('/:id/blocks', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    const { rows } = await pool.query(`
      SELECT sheet_idx, block_row
        FROM sheet_chunks
       WHERE sheet_id = $1::int
       ORDER BY sheet_idx, block_row
    `, [sid]);

    const blocks = {};
    for (const r of rows) {
      if (!blocks[r.sheet_idx]) blocks[r.sheet_idx] = [];
      blocks[r.sheet_idx].push({ row: r.block_row });
    }

    res.json({ blocks });
  } catch (e) {
    console.error('blocks error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   GET /:id/blocks/:sheetIdx/:blockRow — 1 block penuh
   ========================================================= */
router.get('/:id/blocks/:sheetIdx/:blockRow', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    const sIdx = parseInt(req.params.sheetIdx, 10);
    const bRow = parseInt(req.params.blockRow, 10);

    if (isNaN(sid) || isNaN(sIdx) || isNaN(bRow)) {
      return res.status(400).json({ error: 'Invalid params' });
    }

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    const { rows } = await pool.query(`
      SELECT cells FROM sheet_chunks
       WHERE sheet_id = $1::int
         AND sheet_idx = $2::int
         AND block_row = $3::int
    `, [sid, sIdx, bRow]);

    if (!rows.length) {
      return res.json({ cells: {}, block_row: bRow });
    }

    res.json({
      cells: rows[0].cells || {},
      block_row: bRow
    });
  } catch (e) {
    console.error('block load error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   PATCH /:id/blocks/:sheetIdx/:blockRow — delta save per block
   Pakai MERGE operator `||` — anti OOM
   ========================================================= */
router.patch('/:id/blocks/:sheetIdx/:blockRow', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    const sIdx = parseInt(req.params.sheetIdx, 10);
    const bRow = parseInt(req.params.blockRow, 10);

    if (isNaN(sid) || isNaN(sIdx) || isNaN(bRow)) {
      return res.status(400).json({ error: 'Invalid params' });
    }

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });
    if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

    const { cells = {} } = req.body || {};

    const toSet = {};
    const toDelete = [];

    for (const [addr, cell] of Object.entries(cells)) {
      if (!/^[A-Z]+\d+$/.test(addr)) continue;
      if (cell === null || cell === undefined) toDelete.push(addr);
      else toSet[addr] = cell;
    }

    const total = Object.keys(toSet).length + toDelete.length;
    if (total === 0) return res.json({ ok: true, applied: 0 });

    const t0 = Date.now();

    // 1. Upsert dengan MERGE
    if (Object.keys(toSet).length > 0) {
      await pool.query(`
        INSERT INTO sheet_chunks (sheet_id, sheet_idx, block_row, cells)
        VALUES ($1::int, $2::int, $3::int, $4::jsonb)
        ON CONFLICT (sheet_id, sheet_idx, block_row)
        DO UPDATE SET cells = COALESCE(sheet_chunks.cells, '{}'::jsonb) || EXCLUDED.cells
      `, [sid, sIdx, bRow, JSON.stringify(toSet)]);
    }

    // 2. Delete
    if (toDelete.length > 0) {
      await pool.query(`
        UPDATE sheet_chunks
           SET cells = COALESCE(cells, '{}'::jsonb) - $4::text[]
         WHERE sheet_id = $1::int AND sheet_idx = $2::int AND block_row = $3::int
      `, [sid, sIdx, bRow, toDelete]);
    }

    console.log(`💾 Chunk save: ${total} cells in ${Date.now() - t0}ms`);
    res.json({ ok: true, applied: total });
  } catch (e) {
    console.error('chunk save error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   POST /:id/chunk-migrate — migrasi data ke chunks
   ========================================================= */
router.post('/:id/chunk-migrate', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });
    if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

    const BLOCK_ROWS = 1000;

    const { rows } = await pool.query(
      'SELECT data FROM sheets WHERE id = $1::int', [sid]
    );
    if (!rows.length) return res.status(404).json({ error: 'Not found' });

    const wbData = rows[0].data || {};
    const sheets = wbData.sheets || [];
    const totalSheets = sheets.length;

    console.log(`📦 Migrating sheet ${sid}: ${totalSheets} sheets`);

    const t0 = Date.now();
    let totalBlocks = 0;

    // Hapus block lama
    await pool.query('DELETE FROM sheet_chunks WHERE sheet_id = $1::int', [sid]);

    for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
      const cells = sheets[sIdx].cells || {};
      const blocks = {};

      for (const [addr, cell] of Object.entries(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(addr);
        if (!m) continue;
        const row = parseInt(m[2], 10) - 1;
        const blockRow = Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;
        if (!blocks[blockRow]) blocks[blockRow] = {};
        blocks[blockRow][addr] = cell;
      }

      const blockEntries = Object.entries(blocks);
      for (const [blockRow, blockCells] of blockEntries) {
        await pool.query(`
          INSERT INTO sheet_chunks (sheet_id, sheet_idx, block_row, cells)
          VALUES ($1::int, $2::int, $3::int, $4::jsonb)
        `, [sid, sIdx, parseInt(blockRow, 10), JSON.stringify(blockCells)]);
        totalBlocks++;
      }

      console.log(`   ✅ Sheet ${sIdx} (${sheets[sIdx].name}): ${blockEntries.length} blocks`);
    }

    console.log(`✅ Migrated in ${Date.now() - t0}ms: ${totalBlocks} blocks`);
    res.json({ ok: true, blocks: totalBlocks, sheets: totalSheets });
  } catch (e) {
    console.error('migrate error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   GET /:id/full — untuk data kecil (fallback normal mode)
   ========================================================= */
router.get('/:id/full', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    const { rows } = await pool.query(
      'SELECT data FROM sheets WHERE id = $1::int', [sid]
    );
    if (!rows.length) return res.status(404).json({ error: 'Not found' });

    res.json(rows[0].data || {});
  } catch (e) {
    console.error('full load error:', e);
    res.status(500).json({ error: e.message });
  }
});

router.get('/:id/sheets-meta', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    const { rows } = await pool.query(`
      SELECT
        name,
        (password_hash IS NOT NULL) AS has_password,
        COALESCE(data->'activeSheet', '0'::jsonb) AS active_sheet,
        COALESCE(data->'namedRanges', '{}'::jsonb) AS named_ranges,
        COALESCE((
          SELECT jsonb_agg(jsonb_build_object(
            'name',        s->'name',
            'colWidths',   s->'colWidths',
            'rowHeight',   s->'rowHeight',
            'cf',          s->'cf',
            'charts',      s->'charts',
            'merged',      s->'merged',
            'filters',     s->'filters',
            'images',      s->'images',
            'validations', s->'validations',
            'freezeRow',   s->'freezeRow',
            'freezeCol',   s->'freezeCol',
            'protected',   s->'protected',
            'protectHash', s->'protectHash',
            'hidden',      s->'hidden',
            'tabColor',    s->'tabColor',
            'rowGroups',   s->'rowGroups',
            'advancedFilters', s->'advancedFilters',
            'pivots',      s->'pivots'
          ))
          FROM jsonb_array_elements(data->'sheets') s
        ), '[]'::jsonb) AS sheets,
        jsonb_array_length(COALESCE(data->'sheets', '[]'::jsonb)) AS sheet_count
      FROM sheets WHERE id = $1::int
    `, [sid]);

    if (!rows.length) return res.status(404).json({ error: 'Not found' });

    res.json({
      name: rows[0].name,
      hasPassword: rows[0].has_password,
      activeSheet: rows[0].active_sheet,
      namedRanges: rows[0].named_ranges,
      sheets: rows[0].sheets,
      sheetCount: rows[0].sheet_count,
      access: a.access
    });
  } catch (e) {
    console.error('sheets-meta error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   GET SHEET — dengan pagination (maxRow)
   ========================================================= */
router.get('/:id/sheet/:idx', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    const sIdx = parseInt(req.params.idx, 10);
    if (isNaN(sid) || isNaN(sIdx)) return res.status(400).json({ error: 'Invalid' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    const t0 = Date.now();

    // 🎯 Prioritas 1: baca dari sheet_chunks
    const { rows: chunkRows } = await pool.query(`
      SELECT block_row, cells FROM sheet_chunks
      WHERE sheet_id = $1::int AND sheet_idx = $2::int
      ORDER BY block_row
    `, [sid, sIdx]);

    if (chunkRows.length > 0) {
      // Merge chunks
      const chunkCells = {};
      for (const row of chunkRows) Object.assign(chunkCells, row.cells || {});
      const chunkCount = Object.keys(chunkCells).length;

      // Ambil meta + cells dari sheet_data
      const { rows: metaRows } = await pool.query(`
        SELECT sheet FROM sheet_data
        WHERE sheet_id = $1::int AND sheet_idx = $2::int
      `, [sid, sIdx]);

      const metaSheet = metaRows[0]?.sheet || {};
      const metaCells = metaSheet.cells || {};
      const metaCount = Object.keys(metaCells).length;

      let finalCells;
      if (chunkCount >= metaCount * 0.5) {
        // Chunks komprehensif — pakai chunks saja
        finalCells = chunkCells;
        console.log(`📤 Sheet ${sIdx} from CHUNKS: ${chunkCount} cells in ${Date.now()-t0}ms`);
      } else {
        // Chunks partial — merge dengan sheet_data
        finalCells = { ...metaCells, ...chunkCells };
        console.log(`📤 Sheet ${sIdx} MERGED: sheet_data ${metaCount} + chunks ${chunkCount} = ${Object.keys(finalCells).length} cells`);
      }

      const result = { ...metaSheet, cells: finalCells };
      return res.json(result);
    }

    // Fallback: sheet_data
    const { rows } = await pool.query(`
      SELECT sheet FROM sheet_data
      WHERE sheet_id = $1::int AND sheet_idx = $2::int
    `, [sid, sIdx]);

    if (rows.length) {
      const sheet = rows[0].sheet || { cells: {} };
      console.log(`📤 Sheet ${sIdx} from sheet_data: ${Object.keys(sheet.cells || {}).length} cells in ${Date.now()-t0}ms`);
      return res.json(sheet);
    }

    // Fallback: sheets.data
    const { rows: old } = await pool.query(`
      SELECT data->'sheets'->$2::int AS sheet
      FROM sheets WHERE id = $1::int
    `, [sid, sIdx]);

    if (old.length && old[0].sheet) return res.json(old[0].sheet);

    res.json({ cells: {}, name: 'Sheet' });
  } catch (e) {
    console.error('sheet load error:', e);
    res.status(500).json({ error: e.message });
  }
});
/* =========================================================
   GET /:id/chunks/status — cek apakah workbook punya chunks
   ========================================================= */
router.get('/:id/chunks/status', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const { rows } = await pool.query(
      `SELECT COUNT(*)::int AS block_count
       FROM sheet_chunks
       WHERE sheet_id = $1`,
      [sid]
    );

    const blockCount = rows[0]?.block_count || 0;
    res.json({
      hasChunks: blockCount > 0,
      blockCount
    });
  } catch (e) {
    console.error('Chunk status error:', e);
    res.status(500).json({ error: e.message });
  }
});

export default router;