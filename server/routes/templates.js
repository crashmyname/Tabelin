import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

/* wrap async biar error tidak crash server */
const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

/* ============ LIST ============ */
router.get('/', wrap(async (req, res) => {
  const { category } = req.query;
  let sql = `
    SELECT id, name, description, category, is_public, used_count,
           created_at, created_by
    FROM templates
    WHERE is_public = true OR created_by = $1::int
  `;
  const params = [req.user.id];
  if (category) { sql += ' AND category = $2::text'; params.push(category); }
  sql += ' ORDER BY is_public DESC, used_count DESC, created_at DESC';
  const { rows } = await pool.query(sql, params);
  res.json(rows);
}));

/* ============ GET ONE ============ */
router.get('/:id', wrap(async (req, res) => {
  const { rows } = await pool.query(
    `SELECT * FROM templates
      WHERE id = $1::int
        AND (is_public = true OR created_by = $2::int)`,
    [req.params.id, req.user.id]
  );
  if (!rows.length) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
}));

/* ============ SAVE AS TEMPLATE ============ */
router.post('/', wrap(async (req, res) => {
  const { name, description, category = 'general', data, is_public = false } = req.body || {};
  if (!name || !data) return res.status(400).json({ error: 'name & data wajib' });
  const { rows } = await pool.query(
    `INSERT INTO templates (name, description, category, data, is_public, created_by)
     VALUES ($1::text, $2::text, $3::text, $4::jsonb, $5::bool, $6::int)
     RETURNING *`,
    [name, description ?? null, category, JSON.stringify(data), !!is_public, req.user.id]
  );
  res.status(201).json(rows[0]);
}));

/* ============ USE TEMPLATE ============ */
router.post('/:id/use', wrap(async (req, res) => {
  const { rows: tpl } = await pool.query(
    `SELECT * FROM templates
      WHERE id = $1::int
        AND (is_public = true OR created_by = $2::int)`,
    [req.params.id, req.user.id]
  );
  if (!tpl.length) return res.status(404).json({ error: 'Not found' });

  let name = String(req.body?.name || `${tpl[0].name} - Copy`).trim();
  if (name.length < 3) name = `${tpl[0].name} - Copy`;

  // Cek duplikat
  const dup = await pool.query(
    `SELECT id FROM sheets
      WHERE user_id = $1::int AND LOWER(name) = LOWER($2::text)
      LIMIT 1`,
    [req.user.id, name]
  );
  if (dup.rows.length) {
    // Auto-rename: "Nama (2)", "Nama (3)", dst
    let n = 2;
    let candidate = `${name} (${n})`;
    while (true) {
      const check = await pool.query(
        `SELECT id FROM sheets WHERE user_id = $1::int AND LOWER(name) = LOWER($2::text) LIMIT 1`,
        [req.user.id, candidate]
      );
      if (!check.rows.length) { name = candidate; break; }
      n++;
      candidate = `${name.replace(/\s*\(\d+\)$/, '')} (${n})`;
      if (n > 100) { name = candidate + ' - ' + Date.now(); break; }
    }
  }

  const { rows } = await pool.query(
    `INSERT INTO sheets (user_id, name, data)
     VALUES ($1::int, $2::text, $3::jsonb)
     RETURNING *`,
    [req.user.id, name, JSON.stringify(tpl[0].data)]
  );
  await pool.query(
    'UPDATE templates SET used_count = used_count + 1 WHERE id = $1::int',
    [tpl[0].id]
  );
  res.status(201).json(rows[0]);
}));

/* ============ DELETE ============ */
router.delete('/:id', wrap(async (req, res) => {
  await pool.query(
    'DELETE FROM templates WHERE id = $1::int AND created_by = $2::int',
    [req.params.id, req.user.id]
  );
  res.status(204).end();
}));

export default router;