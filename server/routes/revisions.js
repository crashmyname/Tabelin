import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';
import { checkAccess } from './shares.js';

const router = Router();
router.use(authRequired);

const MAX_REVISIONS = 50;

/* ---------- Snapshot dipanggil dari sheets.js ---------- */
export async function snapshotRevision(sheetId, userId, label) {
  const sid = parseInt(sheetId, 10);
  const uid = parseInt(userId, 10);
  if (isNaN(sid) || isNaN(uid)) return;

  const { rows } = await pool.query(
    'SELECT data, name FROM sheets WHERE id = $1::int',
    [sid]
  );
  if (!rows.length) return;

  const payload = JSON.stringify(rows[0].data || {});
  await pool.query(
    `INSERT INTO revisions (sheet_id, data, name, label, created_by, size_bytes)
     VALUES ($1::int, $2::jsonb, $3::text, $4::text, $5::int, $6::int)`,
    [sid, JSON.stringify(rows[0].data), rows[0].name, label || null, uid, payload.length]
  );

  // Prune: simpan hanya N terbaru
  await pool.query(`
    DELETE FROM revisions
     WHERE id IN (
       SELECT id FROM revisions
        WHERE sheet_id = $1::int
        ORDER BY created_at DESC
        OFFSET $2::int
     )
  `, [sid, MAX_REVISIONS]);
}

/* ---------- LIST ---------- */
router.get('/:id/revisions', async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });

  const { rows } = await pool.query(`
    SELECT r.id, r.name, r.label, r.created_at, r.size_bytes,
           u.username AS created_by
      FROM revisions r
      LEFT JOIN users u ON u.id = r.created_by
     WHERE r.sheet_id = $1::int
     ORDER BY r.created_at DESC
     LIMIT 100
  `, [req.params.id]);

  res.json(rows);
});

/* ---------- CREATE manual snapshot ---------- */
router.post('/:id/revisions', async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a || a.access === 'view')
    return res.status(403).json({ error: 'Forbidden' });

  const { label } = req.body || {};
  await snapshotRevision(req.params.id, req.user.id, label || 'Manual snapshot');
  res.json({ ok: true });
});

/* ---------- GET ONE ---------- */
router.get('/:id/revisions/:revId', async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a) return res.status(404).json({ error: 'Not found' });

  const { rows } = await pool.query(
    'SELECT * FROM revisions WHERE id = $1::int AND sheet_id = $2::int',
    [req.params.revId, req.params.id]
  );
  if (!rows.length) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
});

/* ---------- RESTORE ---------- */
router.post('/:id/revisions/:revId/restore', async (req, res) => {
  const a = await checkAccess(req.params.id, req.user.id);
  if (!a || a.access === 'view')
    return res.status(403).json({ error: 'Forbidden' });

  const { rows } = await pool.query(
    'SELECT data FROM revisions WHERE id = $1::int AND sheet_id = $2::int',
    [req.params.revId, req.params.id]
  );
  if (!rows.length) return res.status(404).json({ error: 'Not found' });

  // Snapshot state sekarang dulu
  await snapshotRevision(req.params.id, req.user.id, 'Sebelum restore');

  const upd = await pool.query(
    `UPDATE sheets
        SET data = $2::jsonb,
            updated_at = NOW()
      WHERE id = $1::int
      RETURNING *`,
    [req.params.id, JSON.stringify(rows[0].data)]
  );

  res.json(upd.rows[0]);
});

export default router;