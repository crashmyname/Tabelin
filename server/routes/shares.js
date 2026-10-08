import { Router } from 'express';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

/* =========================================================
   CHECK ACCESS — dipakai oleh routes lain juga
   ========================================================= */
export async function checkAccess(sheetId, userId) {
  const sid = parseInt(sheetId, 10);
  const uid = parseInt(userId, 10);
  if (isNaN(sid) || isNaN(uid)) return null;

  const { rows } = await pool.query(`
    SELECT s.*, 'owner'::text AS access
    FROM sheets s
    WHERE s.id = $1::int AND s.user_id = $2::int
    UNION
    SELECT s.*, sh.permission AS access
    FROM sheets s
    JOIN shares sh ON sh.sheet_id = s.id
    WHERE s.id = $1::int AND sh.user_id = $2::int
    LIMIT 1
  `, [sid, uid]);
  return rows[0] || null;
}

/* =========================================================
   LIST — user yang punya akses ke sheet
   ========================================================= */
router.get('/:id/shares', async (req, res) => {
  const access = await checkAccess(req.params.id, req.user.id);
  if (!access) return res.status(404).json({ error: 'Not found' });

  const { rows } = await pool.query(`
    SELECT sh.id,
           sh.permission,
           sh.created_at,
           u.id       AS user_id,
           u.username,
           sb.username AS shared_by_username
    FROM shares sh
    JOIN users u       ON u.id  = sh.user_id
    LEFT JOIN users sb ON sb.id = sh.shared_by
    WHERE sh.sheet_id = $1::int
    ORDER BY sh.created_at
  `, [req.params.id]);

  res.json(rows);
});

/* =========================================================
   ADD SHARE
   ========================================================= */
router.post('/:id/shares', async (req, res) => {
  const { username, permission = 'view' } = req.body || {};
  if (!['view', 'edit', 'admin'].includes(permission))
    return res.status(400).json({ error: 'Permission tidak valid' });

  // 1. Cek owner sheet
  const { rows: owner } = await pool.query(
    'SELECT user_id FROM sheets WHERE id = $1::int',
    [req.params.id]
  );
  if (!owner.length) return res.status(404).json({ error: 'Sheet not found' });
  if (owner[0].user_id !== req.user.id)
    return res.status(403).json({ error: 'Hanya pemilik bisa membagikan' });

  // 2. Cari user target
  const { rows: u } = await pool.query(
    'SELECT id FROM users WHERE username = $1::text',
    [username]
  );
  if (!u.length) return res.status(404).json({ error: 'User tidak ditemukan' });
  if (u[0].id === req.user.id)
    return res.status(400).json({ error: 'Tidak bisa share ke diri sendiri' });

  // 3. Insert/update
  await pool.query(`
    INSERT INTO shares (sheet_id, user_id, permission, shared_by)
    VALUES ($1::int, $2::int, $3::text, $4::int)
    ON CONFLICT (sheet_id, user_id)
    DO UPDATE SET permission = EXCLUDED.permission,
                  shared_by  = EXCLUDED.shared_by,
                  created_at = NOW()
  `, [req.params.id, u[0].id, permission, req.user.id]);

  res.json({ ok: true });
});

/* =========================================================
   UPDATE PERMISSION
   ========================================================= */
router.put('/:id/shares/:shareId', async (req, res) => {
  const { permission } = req.body || {};
  if (!['view', 'edit', 'admin'].includes(permission))
    return res.status(400).json({ error: 'Permission tidak valid' });

  const { rows } = await pool.query(`
    UPDATE shares sh
       SET permission = $1::text
      FROM sheets s
     WHERE sh.id       = $2::int
       AND sh.sheet_id = $3::int
       AND s.id        = sh.sheet_id
       AND s.user_id   = $4::int
    RETURNING sh.id
  `, [permission, req.params.shareId, req.params.id, req.user.id]);

  if (!rows.length) return res.status(404).json({ error: 'Not found' });
  res.json({ ok: true });
});

/* =========================================================
   REMOVE SHARE
   ========================================================= */
router.delete('/:id/shares/:shareId', async (req, res) => {
  const { rowCount } = await pool.query(`
    DELETE FROM shares sh
     USING sheets s
     WHERE sh.id       = $1::int
       AND sh.sheet_id = $2::int
       AND s.id        = sh.sheet_id
       AND s.user_id   = $3::int
  `, [req.params.shareId, req.params.id, req.user.id]);

  if (!rowCount) return res.status(404).json({ error: 'Not found' });
  res.status(204).end();
});

/* =========================================================
   SEARCH USERS — untuk autocomplete di share modal
   ========================================================= */
router.get('/users/search', async (req, res) => {
  const q = String(req.query.q || '').trim();
  if (q.length < 1) return res.json([]);

  try {
    const { rows } = await pool.query(`
      SELECT id, username
        FROM users
       WHERE username ILIKE $1::text
         AND id <> $2::int
       ORDER BY username
       LIMIT 10
    `, ['%' + q + '%', req.user.id]);
    res.json(rows);
  } catch (e) {
    console.error('user search failed:', e);
    res.json([]);
  }
});
export default router;