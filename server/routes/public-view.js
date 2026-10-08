import { Router } from 'express';
import { pool } from '../db.js';
import { computeSheet } from '../utils/spreadsheet.js';

const router = Router();
const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

/* Endpoint publik untuk akses via token */
router.get('/view/:token', wrap(async (req, res) => {
  const { rows } = await pool.query(`
    SELECT vl.id AS link_id, vl.view_count, vl.expires_at,
           s.id, s.name, s.data, s.updated_at
      FROM view_links vl
      JOIN sheets s ON s.id = vl.sheet_id
     WHERE vl.token = $1::text
  `, [req.params.token]);

  if (!rows.length) return res.status(404).json({ error: 'Link tidak valid' });

  const r = rows[0];
  if (r.expires_at && new Date(r.expires_at) < new Date()) {
    return res.status(410).json({ error: 'Link sudah kadaluarsa' });
  }

  // Update view_count
  await pool.query(
    'UPDATE view_links SET view_count = view_count + 1 WHERE id = $1::int',
    [r.link_id]
  );

  res.json({
    id: r.id,
    name: r.name,
    data: r.data,
    updated_at: r.updated_at,
    view_mode: true
  });
}));

export default router;