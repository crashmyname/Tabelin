import { Router } from 'express';
import bcrypt from 'bcryptjs';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';
import { checkAccessLight } from './shares.js';

const router = Router();
router.use(authRequired);

/* =========================================================
   GET /:id/password-status
   Cek apakah workbook punya password
   ========================================================= */
router.get('/:id/password-status', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const { rows } = await pool.query(
      `SELECT password_hash FROM sheets WHERE id = $1::int`,
      [sid]
    );
    if (!rows.length) return res.status(404).json({ error: 'Not found' });

    res.json({
      hasPassword: !!rows[0].password_hash
    });
  } catch (e) {
    console.error('password-status error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   POST /:id/verify-password
   Verifikasi password untuk buka workbook
   ========================================================= */
router.post('/:id/verify-password', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    const { password } = req.body || {};
    if (!password) return res.status(400).json({ error: 'Password wajib diisi' });

    // Cek akses — user harus punya akses ke sheet ini (owner atau shared)
    const access = await checkAccessLight(sid, req.user.id);
    if (!access) return res.status(404).json({ error: 'Not found' });

    const { rows } = await pool.query(
      `SELECT password_hash FROM sheets WHERE id = $1::int`,
      [sid]
    );
    if (!rows.length) return res.status(404).json({ error: 'Not found' });

    const hash = rows[0].password_hash;
    if (!hash) return res.json({ ok: true, hasPassword: false });

    const valid = await bcrypt.compare(password, hash);
    if (!valid) return res.status(401).json({ error: 'Password salah' });

    res.json({ ok: true });
  } catch (e) {
    console.error('verify-password error:', e);
    res.status(500).json({ error: e.message });
  }
});

/* =========================================================
   POST /:id/password
   Set / ganti / hapus password
   Body: { currentPassword?, newPassword }
   - newPassword = null/kosong → hapus password
   ========================================================= */
router.post('/:id/password', async (req, res) => {
  try {
    const sid = parseInt(req.params.id, 10);
    if (isNaN(sid)) return res.status(400).json({ error: 'Invalid id' });

    // Hanya owner yang bisa set password
    const { rows: ownerRows } = await pool.query(
      `SELECT user_id, password_hash FROM sheets WHERE id = $1::int`,
      [sid]
    );
    if (!ownerRows.length) return res.status(404).json({ error: 'Not found' });
    if (ownerRows[0].user_id !== req.user.id) {
      return res.status(403).json({ error: 'Hanya pemilik yang bisa ubah password' });
    }

    const { currentPassword, newPassword } = req.body || {};
    const hasExisting = !!ownerRows[0].password_hash;

    // Kalau ada password lama, verifikasi dulu
    if (hasExisting) {
      if (!currentPassword) {
        return res.status(400).json({ error: 'Password lama wajib diisi' });
      }
      const valid = await bcrypt.compare(currentPassword, ownerRows[0].password_hash);
      if (!valid) {
        return res.status(401).json({ error: 'Password lama salah' });
      }
    }

    // Kalau newPassword kosong → hapus password
    if (!newPassword) {
      await pool.query(
        `UPDATE sheets SET password_hash = NULL, updated_at = NOW() WHERE id = $1::int`,
        [sid]
      );
      return res.json({ ok: true, removed: true });
    }

    // Validasi panjang
    if (newPassword.length < 4) {
      return res.status(400).json({ error: 'Password minimal 4 karakter' });
    }

    // Hash & simpan
    const hash = await bcrypt.hash(newPassword, 10);
    await pool.query(
      `UPDATE sheets SET password_hash = $2, updated_at = NOW() WHERE id = $1::int`,
      [sid, hash]
    );

    res.json({ ok: true });
  } catch (e) {
    console.error('set-password error:', e);
    res.status(500).json({ error: e.message });
  }
});

export default router;