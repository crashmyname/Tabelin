import { Router } from 'express';
import crypto from 'crypto';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

function generateKey() {
  return 'wex_' + crypto.randomBytes(24).toString('hex');
}

/* LIST */
router.get('/', wrap(async (req, res) => {
  const { rows } = await pool.query(`
    SELECT k.id, k.key, k.label, k.sheet_id, k.permission,
           k.created_at, k.last_used_at, k.request_count,
           s.name AS sheet_name
    FROM api_keys k
    LEFT JOIN sheets s ON s.id = k.sheet_id
    WHERE k.user_id = $1::int
    ORDER BY k.created_at DESC
  `, [req.user.id]);
  res.json(rows);
}));

/* CREATE */
router.post('/', wrap(async (req, res) => {
  const { label, sheet_id } = req.body || {};
  const key = generateKey();
  const { rows } = await pool.query(`
    INSERT INTO api_keys (user_id, sheet_id, key, label)
    VALUES ($1::int, $2::int, $3::text, $4::text)
    RETURNING *
  `, [req.user.id, sheet_id ?? null, key, label ?? 'Default']);
  res.status(201).json(rows[0]);
}));

/* ROTATE */
router.post('/:id/rotate', wrap(async (req, res) => {
  const key = generateKey();
  const { rows } = await pool.query(`
    UPDATE api_keys SET key = $1::text
    WHERE id = $2::int AND user_id = $3::int
    RETURNING *
  `, [key, req.params.id, req.user.id]);
  if (!rows.length) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
}));

/* DELETE */
router.delete('/:id', wrap(async (req, res) => {
  await pool.query(
    'DELETE FROM api_keys WHERE id = $1::int AND user_id = $2::int',
    [req.params.id, req.user.id]
  );
  res.status(204).end();
}));

export default router;