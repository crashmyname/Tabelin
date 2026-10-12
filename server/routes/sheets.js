import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';
import { checkAccess } from './shares.js';
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
  const a = await checkAccess(req.params.id, req.user.id);
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
  const a = await checkAccess(req.params.id, req.user.id);
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

/* =========================================================
   DELTA UPDATE (PATCH)
   Kirim hanya sel yang berubah → payload ~KB bukan MB
   ========================================================= */
router.patch('/:id/delta', async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });
  if (a.access === 'view') return res.status(403).json({ error: 'Read-only' });

  const { sheetIdx, cells = {}, meta = {} } = req.body || {};
  const sid = parseInt(req.params.id, 10);
  const sIdx = parseInt(sheetIdx, 10);
  if (isNaN(sid) || isNaN(sIdx)) {
    return res.status(400).json({ error: 'Invalid id or sheetIdx' });
  }

  const MAX_DELTA = 4 * 1024 * 1024;
  const deltaSize = JSON.stringify(cells).length;
  if (deltaSize > MAX_DELTA) {
    return res.status(413).json({
      error: `Delta terlalu besar: ${(deltaSize/1024).toFixed(0)} KB. Coba save bertahap.`
    });
  }

  /* ---- 1. Meta (name / activeSheet) ---- */
  if (meta.name !== undefined || meta.activeSheet !== undefined) {
    await pool.query(
      `UPDATE sheets
          SET name = COALESCE($2::text, name),
              data = CASE
                WHEN $3::int IS NOT NULL
                THEN jsonb_set(
                       COALESCE(data, '{}'::jsonb),
                       '{activeSheet}',
                       to_jsonb($3::int)
                     )
                ELSE data
              END,
              updated_at = NOW()
        WHERE id = $1::int`,
      [sid, meta.name ?? null, meta.activeSheet ?? null]
    );
  }

  /* ---- 2. Delta cells ---- */
  const addrs = Object.keys(cells);
  if (addrs.length) {
    let expr = 'data';
    const params = [sid];   // $1 = sheetId  (satu-satunya param tetap)
    let pIdx = 2;           // param berikutnya mulai dari $2

    for (const addr of addrs) {
      if (!/^[A-Z]+\d+$/.test(addr)) continue;
      const val = cells[addr];
      const path = `'{sheets,${sIdx},cells,${addr}}'`;

      if (val === null) {
        // Hapus key: operator #-
        expr = `(${expr} #- ${path}::text[])`;
      } else {
        // Set / replace: jsonb_set(target, path, value)
        expr = `jsonb_set(COALESCE(${expr}, '{}'::jsonb), ${path}::text[], $${pIdx}::jsonb)`;
        params.push(JSON.stringify(val));
        pIdx++;
      }
    }

    if (expr !== 'data') {
      await pool.query(
        `UPDATE sheets
            SET data = ${expr},
                updated_at = NOW()
          WHERE id = $1::int`,
        params
      );
    }
  }

  res.json({ ok: true, applied: addrs.length });
});

export default router;