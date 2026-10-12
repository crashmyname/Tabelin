import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';
import { checkAccess } from './shares.js';

const router = Router();
router.use(authRequired);

const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

/* =========================================================
   GET /:id/blocks — metadata semua block (super ringan)
   Return: { sheetIdx: [blockRow1, blockRow2, ...] }
   ========================================================= */
router.get('/:id/blocks', wrap(async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });

  const { rows } = await pool.query(`
    SELECT sheet_idx, block_row, pg_column_size(cells) AS size_bytes
      FROM sheet_chunks
     WHERE sheet_id = $1::int
     ORDER BY sheet_idx, block_row
  `, [req.params.id]);

  // Group by sheet_idx
  const blocks = {};
  for (const r of rows) {
    if (!blocks[r.sheet_idx]) blocks[r.sheet_idx] = [];
    blocks[r.sheet_idx].push({
      row: r.block_row,
      size: r.size_bytes
    });
  }

  res.json({ blocks });
}));

/* =========================================================
   GET /:id/blocks/:sheetIdx/:blockRow — 1 block penuh
   ========================================================= */
router.get('/:id/blocks/:sheetIdx/:blockRow', wrap(async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });

  const { rows } = await pool.query(`
    SELECT cells FROM sheet_chunks
     WHERE sheet_id = $1::int
       AND sheet_idx = $2::int
       AND block_row = $3::int
  `, [req.params.id, req.params.sheetIdx, req.params.blockRow]);

  if (!rows.length) {
    return res.json({ cells: {}, block_row: +req.params.blockRow });
  }

  res.json({
    cells: rows[0].cells || {},
    block_row: +req.params.blockRow
  });
}));

/* =========================================================
   PATCH /:id/blocks/:sheetIdx/:blockRow — delta save per block
   Body: { cells: { "A1": {...}, "B2": null } }
   ========================================================= */
router.patch('/:id/blocks/:sheetIdx/:blockRow', wrap(async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });
  if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

  const { cells = {} } = req.body || {};
  const sid = parseInt(req.params.id, 10);
  const sIdx = parseInt(req.params.sheetIdx, 10);
  const bRow = parseInt(req.params.blockRow, 10);

  if (isNaN(sid) || isNaN(sIdx) || isNaN(bRow)) {
    return res.status(400).json({ error: 'Invalid params' });
  }

  const addrs = Object.keys(cells);
  if (!addrs.length) return res.json({ ok: true, applied: 0 });

  // Bangun jsonb_set berantai
  let expr = "COALESCE(cells, '{}'::jsonb)";
  const params = [sid, sIdx, bRow];
  let pIdx = 4;

  for (const addr of addrs) {
    if (!/^[A-Z]+\d+$/.test(addr)) continue;
    const val = cells[addr];
    const path = `'{${addr}}'`;

    if (val === null) {
      expr = `(${expr} #- ${path}::text[])`;
    } else {
      expr = `jsonb_set(${expr}, ${path}::text[], $${pIdx}::jsonb)`;
      params.push(JSON.stringify(val));
      pIdx++;
    }
  }

  // Upsert block
  await pool.query(`
    INSERT INTO sheet_chunks (sheet_id, sheet_idx, block_row, cells)
    VALUES ($1::int, $2::int, $3::int, ${expr})
    ON CONFLICT (sheet_id, sheet_idx, block_row)
    DO UPDATE SET cells = ${expr}
  `, params);

  res.json({ ok: true, applied: addrs.length });
}));

/* =========================================================
   POST /:id/chunk-migrate — migrasi data lama ke chunks
   ========================================================= */
router.post('/:id/chunk-migrate', wrap(async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });
  if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

  const BLOCK_ROWS = 256;
  const sid = parseInt(req.params.id, 10);

  const { rows } = await pool.query(
    'SELECT data FROM sheets WHERE id = $1::int', [sid]
  );
  if (!rows.length) return res.status(404).json({ error: 'Not found' });

  const wbData = rows[0].data || {};
  const sheets = wbData.sheets || [];

  let totalBlocks = 0;

  for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
    const cells = sheets[sIdx].cells || {};
    const blocks = {};

    // Kelompokkan cells per block
    for (const [addr, cell] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) continue;
      const row = parseInt(m[2], 10) - 1;
      const blockRow = Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;
      if (!blocks[blockRow]) blocks[blockRow] = {};
      blocks[blockRow][addr] = cell;
    }

    // Hapus block lama untuk sheet ini
    await pool.query(`
      DELETE FROM sheet_chunks
       WHERE sheet_id = $1::int AND sheet_idx = $2::int
    `, [sid, sIdx]);

    // Insert block baru
    for (const [blockRow, blockCells] of Object.entries(blocks)) {
      await pool.query(`
        INSERT INTO sheet_chunks (sheet_id, sheet_idx, block_row, cells)
        VALUES ($1::int, $2::int, $3::int, $4::jsonb)
      `, [sid, sIdx, parseInt(blockRow, 10), JSON.stringify(blockCells)]);
      totalBlocks++;
    }
  }

  res.json({ ok: true, blocks: totalBlocks });
}));

/* =========================================================
   GET /:id/chunks-status — cek apakah pakai chunk
   ========================================================= */
router.get('/:id/chunks-status', wrap(async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });

  const { rows } = await pool.query(`
    SELECT
      pg_column_size(s.data) AS data_size,
      (SELECT COUNT(*) FROM sheet_chunks WHERE sheet_id = s.id) AS block_count
    FROM sheets s
    WHERE s.id = $1::int
  `, [req.params.id]);

  if (!rows.length) return res.status(404).json({ error: 'Not found' });

  const r = rows[0];
  const dataSize = +r.data_size;
  const blockCount = +r.block_count;

  // Threshold: pakai chunk kalau data > 5 MB
  const THRESHOLD = 20 * 1024 * 1024;
  const shouldUseChunks = blockCount > 0 && dataSize > THRESHOLD;

  res.json({
    dataSize,
    blockCount,
    shouldUseChunks,
    threshold: THRESHOLD
  });
}));

export default router;