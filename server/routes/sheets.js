import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';
import { checkAccessLight } from './shares.js';
import { snapshotRevision } from './revisions.js';

const router = Router();
router.use(authRequired);

/* =========================================================
   LIST — milik sendiri + yang di-share ke saya
   ========================================================= */
router.get('/', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT s.id, s.name, s.updated_at,
           'owner'::text AS access,
           u.username AS owner_name
    FROM sheets s
    JOIN users u ON u.id = s.user_id
    WHERE s.user_id = $1::int
    UNION
    SELECT s.id, s.name, s.updated_at,
           sh.permission AS access,
           u.username AS owner_name
    FROM sheets s
    JOIN shares sh ON sh.sheet_id = s.id
    JOIN users u ON u.id = s.user_id
    WHERE sh.user_id = $1::int
    ORDER BY updated_at DESC
  `, [req.user.id]);
  res.json(rows);
});

/* =========================================================
   CREATE
   ========================================================= */
router.post('/', async (req, res) => {
  const { name = 'Untitled', data = {} } = req.body || {};
  const trimmed = String(name).trim();
  if (!trimmed) return res.status(400).json({ error: 'Nama wajib diisi' });
  if (trimmed.length < 3) return res.status(400).json({ error: 'Nama minimal 3 karakter' });
  if (trimmed.length > 80) return res.status(400).json({ error: 'Nama maksimal 80 karakter' });

  // Cek duplikat (case-insensitive, per user)
  const dup = await pool.query(
    `SELECT id FROM sheets
      WHERE user_id = $1::int
        AND LOWER(name) = LOWER($2::text)
      LIMIT 1`,
    [req.user.id, trimmed]
  );
  if (dup.rows.length) {
    return res.status(409).json({ error: `Workbook "${trimmed}" sudah ada. Gunakan nama lain.` });
  }

  const { rows } = await pool.query(
    `INSERT INTO sheets (user_id, name, data)
     VALUES ($1::int, $2::text, $3::jsonb)
     RETURNING *`,
    [req.user.id, trimmed, JSON.stringify(data)]
  );
  res.status(201).json(rows[0]);
});

/* =========================================================
   GET ONE
   ========================================================= */
router.get('/:id', async (req, res) => {
  const a = await checkAccessLight(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });
  a.access_mode = a.access;

  if (req.query.light === '1') {
    const data = a.data || {};
    if (data.sheets) {
      data.sheets = data.sheets.map(s => ({
        ...s,
        cells: {}
      }));
    }
    a.data = data;
  }

  res.json(a);
});

/* =========================================================
   FULL UPDATE (PUT)
   ========================================================= */
router.put('/:id', async (req, res) => {
  const a = await checkAccessLight(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });
  if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

  const { name, data, make_revision } = req.body || {};
  const sid = parseInt(req.params.id, 10);

  // Cek duplikat nama (kalau nama diubah)
  if (name !== undefined) {
    const trimmed = String(name).trim();
    if (trimmed.length < 3)
      return res.status(400).json({ error: 'Nama minimal 3 karakter' });
    if (trimmed.length > 80)
      return res.status(400).json({ error: 'Nama maksimal 80 karakter' });

    const dup = await pool.query(
      `SELECT id FROM sheets
        WHERE user_id = $1::int
          AND LOWER(name) = LOWER($2::text)
          AND id <> $3::int
        LIMIT 1`,
      [req.user.id, trimmed, sid]
    );
    if (dup.rows.length)
      return res.status(409).json({ error: `Workbook "${trimmed}" sudah ada.` });

    req.body.name = trimmed;
  }

  // ---- LIMIT SIZE ----
  const MAX_BYTES = 20 * 1024 * 1024;
  if (data !== undefined) {
    const size = JSON.stringify(data).length;
    if (size > MAX_BYTES) {
      return res.status(413).json({
        error: `Workbook terlalu besar: ${(size/1024/1024).toFixed(1)} MB (maks 10 MB). Pecah menjadi beberapa workbook.`
      });
    }
  }

  // Simpan snapshot data LAMA sebelum diupdate
  if (data !== undefined) {
    const lastRev = await pool.query(
      `SELECT created_at FROM revisions
       WHERE sheet_id = $1::int
       ORDER BY created_at DESC LIMIT 1`,
      [sid]
    );
    const shouldSnap = !lastRev.rows.length ||
      (Date.now() - new Date(lastRev.rows[0].created_at).getTime()) > 5 * 60 * 1000;
    if (shouldSnap || make_revision) {
      await snapshotRevision(sid, req.user.id,
        make_revision ? 'Manual' : 'Auto 5-menit');
    }
  }

  const { rows } = await pool.query(
    `UPDATE sheets
        SET name       = COALESCE($2::text, name),
            data       = COALESCE($3::jsonb, data),
            updated_at = NOW()
      WHERE id = $1::int
      RETURNING *`,
    [
      sid,
      name ?? null,
      data !== undefined ? JSON.stringify(data) : null
    ]
  );
  if (!rows.length) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
});

/* =========================================================
   DELETE
   ========================================================= */
router.delete('/:id', async (req, res) => {
  const { rowCount } = await pool.query(
    'DELETE FROM sheets WHERE id = $1::int AND user_id = $2::int',
    [req.params.id, req.user.id]
  );
  if (!rowCount) return res.status(403).json({ error: 'Hanya pemilik bisa hapus' });
  res.status(204).end();
});

router.patch('/:id/delta', async (req, res) => {
  try {
    const a = await checkAccessLight(req.params.id, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });
    if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

    const { sheetIdx, cells = {}, meta = {} } = req.body || {};
    const sid = parseInt(req.params.id, 10);
    const sIdx = parseInt(sheetIdx, 10);

    if (isNaN(sid) || isNaN(sIdx)) {
      return res.status(400).json({ error: 'Invalid id or sheetIdx' });
    }

    const t0 = Date.now();

    // =====================================================
    // 1. Meta — name / activeSheet
    // =====================================================
    if (meta.name !== undefined) {
      await pool.query(
        `UPDATE sheets SET name = $2::text, updated_at = NOW() WHERE id = $1::int`,
        [sid, meta.name]
      );
    }
    if (meta.namedRanges !== undefined) {
      await pool.query(
        `UPDATE sheets
            SET data = jsonb_set(
                  COALESCE(data, '{}'::jsonb),
                  '{namedRanges}',
                  $2::jsonb
                ),
                updated_at = NOW()
          WHERE id = $1::int`,
        [sid, JSON.stringify(meta.namedRanges)]
      );
    }
    if (meta.activeSheet !== undefined) {
      await pool.query(
        `UPDATE sheets
            SET data = jsonb_set(COALESCE(data,'{}'::jsonb),'{activeSheet}',to_jsonb($2::int)),
                updated_at = NOW()
          WHERE id = $1::int`,
        [sid, meta.activeSheet]
      );
    }

    // =====================================================
    // 2. Meta — hiddenCols / hiddenRows (per sheet)
    // =====================================================
    if (meta.hiddenCols !== undefined || meta.hiddenRows !== undefined) {
      const updates = [];
      if (meta.hiddenCols !== undefined) {
        updates.push(pool.query(`
          UPDATE sheet_data
             SET sheet = jsonb_set(
                   COALESCE(sheet, '{}'::jsonb),
                   '{hiddenCols}',
                   $3::jsonb
                 ),
                 updated_at = NOW()
           WHERE sheet_id = $1::int AND sheet_idx = $2::int
        `, [sid, sIdx, JSON.stringify(meta.hiddenCols)]));
      }
      if (meta.hiddenRows !== undefined) {
        updates.push(pool.query(`
          UPDATE sheet_data
             SET sheet = jsonb_set(
                   COALESCE(sheet, '{}'::jsonb),
                   '{hiddenRows}',
                   $3::jsonb
                 ),
                 updated_at = NOW()
           WHERE sheet_id = $1::int AND sheet_idx = $2::int
        `, [sid, sIdx, JSON.stringify(meta.hiddenRows)]));
      }
      await Promise.all(updates);
    }

    // =====================================================
    // 3. 🎯 Meta — sheetOrder / sheetMeta (Duplicate, Reorder, Color, Hide, Rename)
    // =====================================================
    if (meta.sheetOrder !== undefined || meta.sheetMeta !== undefined) {
      const { rows: curRows } = await pool.query(
        `SELECT data FROM sheets WHERE id = $1::int`, [sid]
      );
      if (curRows.length) {
        const cur = curRows[0].data || {};
        const sheets = Array.isArray(cur.sheets) ? cur.sheets : [];

        // 🎯 Reorder + handle SHEET BARU
        if (Array.isArray(meta.sheetOrder)) {
          const reordered = meta.sheetOrder.map((i, position) => {
            const existing = sheets[i];
            if (existing) return existing;

            // 🎯 Sheet baru — buat placeholder
            // Nama diambil dari sheetMeta (posisi yang sama)
            const metaForNew = meta.sheetMeta?.[i] || meta.sheetMeta?.[position] || {};
            return {
              name: metaForNew.name || `Sheet${i + 1}`,
              cells: {},
              colWidths: {},
              rowHeight: 26,
              cf: [], charts: [], merged: [], filters: {},
              images: [], validations: [],
              hidden: !!metaForNew.hidden,
              tabColor: metaForNew.tabColor || null,
              rowGroups: metaForNew.rowGroups || [],
              protected: !!metaForNew.protected,
              protectHash: metaForNew.protectHash || null,
              freezeRow: 0, freezeCol: 0
            };
          });

          cur.sheets = reordered;
        }

        // 🎯 Per-sheet meta (hidden, tabColor, name, dll)
        if (meta.sheetMeta && typeof meta.sheetMeta === 'object') {
          for (const [idxStr, patch] of Object.entries(meta.sheetMeta)) {
            const i = parseInt(idxStr, 10);
            if (!cur.sheets[i]) continue;

            if (patch.hidden !== undefined)      cur.sheets[i].hidden      = patch.hidden;
            if (patch.tabColor !== undefined)    cur.sheets[i].tabColor    = patch.tabColor;
            if (patch.name !== undefined)        cur.sheets[i].name        = patch.name;
            if (patch.rowGroups !== undefined)   cur.sheets[i].rowGroups   = patch.rowGroups;
            if (patch.protected !== undefined)   cur.sheets[i].protected   = patch.protected;
            if (patch.protectHash !== undefined) cur.sheets[i].protectHash = patch.protectHash;
            if (patch.advancedFilters !== undefined) cur.sheets[i].advancedFilters = patch.advancedFilters;
            if (patch.pivots !== undefined) cur.sheets[i].pivots = patch.pivots;
          }
        }

        await pool.query(
          `UPDATE sheets SET data = $2::jsonb, updated_at = NOW() WHERE id = $1::int`,
          [sid, JSON.stringify(cur)]
        );
      }
    }

    // =====================================================
    // 4. Cells — group per block
    // =====================================================
    const BLOCK_ROWS = 1000;
    const blocks = {};

    // 🎯 Track changes untuk audit log
    const auditEntries = [];

    // Ambil nilai lama untuk cell yang diubah (kalau ada)
    const changeAddrs = Object.keys(cells).filter(a => /^[A-Z]+\d+$/.test(a));

    if (changeAddrs.length > 0) {
      try {
        // Ambil old values dari sheet_chunks
        for (const addr of changeAddrs) {
          const m = /^([A-Z]+)(\d+)$/.exec(addr);
          if (!m) continue;
          const row = parseInt(m[2], 10) - 1;
          const blockRow = Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;

          const { rows: oldRows } = await pool.query(
            `SELECT cells->$3 AS old_cell FROM sheet_chunks
            WHERE sheet_id = $1::int AND sheet_idx = $2::int AND block_row = $4::int`,
            [sid, sIdx, addr, blockRow]
          );

          const oldCell = oldRows[0]?.old_cell || null;
          const newCell = cells[addr];

          // Skip kalau sama
          const oldStr = JSON.stringify(oldCell);
          const newStr = JSON.stringify(newCell);
          if (oldStr === newStr) continue;

          // Tentukan action
          let action = 'edit';
          if (oldCell === null && newCell !== null) action = 'create';
          else if (oldCell !== null && (newCell === null || newCell === undefined)) action = 'delete';

          auditEntries.push({
            addr,
            action,
            old_value: oldCell,
            new_value: newCell === undefined ? null : newCell
          });
        }
      } catch (e) {
        console.warn('Audit tracking failed:', e.message);
      }
    }

    // Group per block (seperti sebelumnya)
    for (const [addr, cell] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) continue;
      const row = parseInt(m[2], 10) - 1;
      const blockRow = Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;
      if (!blocks[blockRow]) blocks[blockRow] = { set: {}, del: [] };
      if (cell === null || cell === undefined) blocks[blockRow].del.push(addr);
      else blocks[blockRow].set[addr] = cell;
    }

    const blockKeys = Object.keys(blocks);
    if (blockKeys.length > 0) {
      console.log(`📥 Delta: sheet ${sid}/${sIdx}, ${Object.keys(cells).length} cells in ${blockKeys.length} block(s)`);
    }

    // =====================================================
    // 5. Update setiap block
    // =====================================================
    const updates = blockKeys.map(async (blockRowStr) => {
      const blockRow = parseInt(blockRowStr, 10);
      const { set, del } = blocks[blockRowStr];
      const setCount = Object.keys(set).length;
      const delCount = del.length;

      if (setCount > 0) {
        await pool.query(`
          INSERT INTO sheet_chunks (sheet_id, sheet_idx, block_row, cells)
          VALUES ($1::int, $2::int, $3::int, $4::jsonb)
          ON CONFLICT (sheet_id, sheet_idx, block_row)
          DO UPDATE SET
            cells = COALESCE(sheet_chunks.cells, '{}'::jsonb) || EXCLUDED.cells
        `, [sid, sIdx, blockRow, JSON.stringify(set)]);
      }

      if (delCount > 0) {
        await pool.query(`
          UPDATE sheet_chunks
             SET cells = COALESCE(cells, '{}'::jsonb) - $4::text[]
           WHERE sheet_id = $1::int AND sheet_idx = $2::int AND block_row = $3::int
        `, [sid, sIdx, blockRow, del]);
      }
    });

    await Promise.all(updates);

    // Update timestamp sheets
    await pool.query(
      `UPDATE sheets SET updated_at = NOW() WHERE id = $1::int`,
      [sid]
    );

    const elapsed = Date.now() - t0;
    console.log(`✅ Delta done in ${elapsed}ms`);

    // 🎯 Insert audit log
    if (auditEntries.length > 0) {
      try {
        const values = auditEntries.map(e => [
          sid, sIdx, req.user.id, req.user.username || 'unknown',
          e.addr, e.action,
          JSON.stringify(e.old_value),
          JSON.stringify(e.new_value)
        ]);

        // Batch insert
        const BATCH = 100;
        for (let i = 0; i < values.length; i += BATCH) {
          const batch = values.slice(i, i + BATCH);
          const placeholders = batch.map((_, idx) => {
            const base = idx * 8;
            return `($${base+1}, $${base+2}, $${base+3}, $${base+4}, $${base+5}, $${base+6}, $${base+7}::jsonb, $${base+8}::jsonb)`;
          }).join(',');
          const flat = batch.flat();

          await pool.query(
            `INSERT INTO audit_log (sheet_id, sheet_idx, user_id, username, addr, action, old_value, new_value)
            VALUES ${placeholders}`,
            flat
          );
        }

        console.log(`📝 Audit log: ${auditEntries.length} entries`);
      } catch (e) {
        console.error('Audit log insert failed:', e.message);
      }
    }

    res.json({
      ok: true,
      applied: Object.keys(cells).length,
      blocks: blockKeys.length,
      ms: elapsed
    });

  } catch (e) {
    console.error('❌ Delta error:', e);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   GET /:id/audit-log — riwayat perubahan
   Query: ?limit=100&offset=0&user=xxx&addr=A1&action=edit
   ========================================================= */
router.get('/:id/audit-log', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const a = await checkAccessLight(sid, req.user.id);
    if (!a) return res.status(404).json({ error: 'Not found' });

    const limit = Math.min(parseInt(req.query.limit || '100', 10), 500);
    const offset = parseInt(req.query.offset || '0', 10);
    const userFilter = req.query.user;
    const addrFilter = req.query.addr;
    const actionFilter = req.query.action;
    const sheetIdxFilter = req.query.sheetIdx;

    let where = ['sheet_id = $1'];
    const params = [sid];

    if (userFilter) {
      params.push(userFilter);
      where.push(`username = $${params.length}`);
    }
    if (addrFilter) {
      params.push(addrFilter.toUpperCase());
      where.push(`addr = $${params.length}`);
    }
    if (actionFilter) {
      params.push(actionFilter);
      where.push(`action = $${params.length}`);
    }
    if (sheetIdxFilter !== undefined && sheetIdxFilter !== '') {
      params.push(parseInt(sheetIdxFilter, 10));
      where.push(`sheet_idx = $${params.length}`);
    }

    params.push(limit);
    const limitIdx = params.length;
    params.push(offset);
    const offsetIdx = params.length;

    const { rows } = await pool.query(`
      SELECT id, sheet_idx, user_id, username, addr, action, old_value, new_value, created_at
      FROM audit_log
      WHERE ${where.join(' AND ')}
      ORDER BY created_at DESC
      LIMIT $${limitIdx} OFFSET $${offsetIdx}
    `, params);

    // Count total
    const { rows: countRows } = await pool.query(
      `SELECT COUNT(*)::int AS total FROM audit_log WHERE ${where.join(' AND ')}`,
      params.slice(0, params.length - 2)
    );

    res.json({
      total: countRows[0]?.total || 0,
      entries: rows
    });
  } catch (e) {
    console.error('audit-log error:', e);
    res.status(500).json({ error: e.message });
  }
});

export default router;