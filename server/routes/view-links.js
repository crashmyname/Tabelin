import { Router } from 'express';
import crypto from 'crypto';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

router.get('/:id/view-links', wrap(async (req, res) => {
  const { rows } = await pool.query(`
    SELECT id, token, label, created_at, expires_at, view_count
      FROM view_links
     WHERE sheet_id = $1::int
     ORDER BY created_at DESC
  `, [req.params.id]);
  res.json(rows);
}));

router.post('/:id/view-links', wrap(async (req, res) => {
  const { label } = req.body || {};
  // Verify ownership
  const own = await pool.query(
    'SELECT id FROM sheets WHERE id=$1::int AND user_id=$2::int',
    [req.params.id, req.user.id]
  );
  if (!own.rows.length) return res.status(403).json({ error: 'Forbidden' });

  const token = crypto.randomBytes(16).toString('hex');
  const { rows } = await pool.query(`
    INSERT INTO view_links (sheet_id, token, label, created_by)
    VALUES ($1::int, $2::text, $3::text, $4::int)
    RETURNING id, token, label, created_at, expires_at, view_count
  `, [req.params.id, token, label || 'View Link', req.user.id]);
  res.status(201).json(rows[0]);
}));

router.delete('/:id/view-links/:linkId', wrap(async (req, res) => {
  await pool.query(`
    DELETE FROM view_links
    WHERE id = $1::int AND sheet_id = $2::int
      AND sheet_id IN (SELECT id FROM sheets WHERE user_id = $3::int)
  `, [req.params.linkId, req.params.id, req.user.id]);
  res.status(204).end();
}));

export default router;