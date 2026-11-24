// routes/folders.js  (FILE BARU)
import { Router } from 'express';
import { pool } from '../db.js';                // named export, sama seperti di index.js
import { authRequired } from '../auth.js';      // <-- SESUAIKAN: lihat cara routes/sheets.js meng-import authRequired

const router = Router();

// authRequired dipasang PER ROUTE (bukan router.use) supaya tidak ikut
// memproteksi endpoint lain di /api seperti /api/auth/login.

const toInt = v => {
  const n = Number(v);
  return Number.isInteger(n) && n > 0 ? n : null;
};

/* ---------- Daftar folder milik user ---------- */
router.get('/folders', authRequired, async (req, res) => {
  try {
    const { rows } = await pool.query(
      'SELECT id, name FROM folders WHERE owner_id = $1 ORDER BY LOWER(name)',
      [req.user.id]
    );
    res.json(rows);
  } catch (e) {
    console.error('GET /folders', e);
    res.status(500).json({ error: 'Gagal memuat folder' });
  }
});

/* ---------- Peta { sheet_id: folder_id } milik user ---------- */
router.get('/folders/map', authRequired, async (req, res) => {
  try {
    const { rows } = await pool.query(
      'SELECT sheet_id, folder_id FROM sheet_folder WHERE user_id = $1 AND folder_id IS NOT NULL',
      [req.user.id]
    );
    const map = {};
    for (const r of rows) map[r.sheet_id] = r.folder_id;
    res.json(map);
  } catch (e) {
    console.error('GET /folders/map', e);
    res.status(500).json({ error: 'Gagal memuat penempatan folder' });
  }
});

/* ---------- Buat folder ---------- */
router.post('/folders', authRequired, async (req, res) => {
  try {
    const name = String(req.body?.name || '').trim();
    if (!name) return res.status(400).json({ error: 'Nama wajib diisi' });
    if (name.length > 60) return res.status(400).json({ error: 'Nama maksimal 60 karakter' });

    const { rows } = await pool.query(
      'INSERT INTO folders (owner_id, name) VALUES ($1, $2) RETURNING id, name',
      [req.user.id, name]
    );
    res.json(rows[0]);
  } catch (e) {
    if (e.code === '23505') return res.status(409).json({ error: 'Nama folder sudah ada' });
    console.error('POST /folders', e);
    res.status(500).json({ error: 'Gagal membuat folder' });
  }
});

/* ---------- Ganti nama folder ---------- */
router.patch('/folders/:id', authRequired, async (req, res) => {
  try {
    const id = toInt(req.params.id);
    const name = String(req.body?.name || '').trim();
    if (!id) return res.status(400).json({ error: 'ID tidak valid' });
    if (!name) return res.status(400).json({ error: 'Nama wajib diisi' });
    if (name.length > 60) return res.status(400).json({ error: 'Nama maksimal 60 karakter' });

    const r = await pool.query(
      'UPDATE folders SET name = $1 WHERE id = $2 AND owner_id = $3',
      [name, id, req.user.id]
    );
    if (!r.rowCount) return res.status(404).json({ error: 'Folder tidak ditemukan' });
    res.json({ ok: true });
  } catch (e) {
    if (e.code === '23505') return res.status(409).json({ error: 'Nama folder sudah ada' });
    console.error('PATCH /folders/:id', e);
    res.status(500).json({ error: 'Gagal mengubah folder' });
  }
});

/* ---------- Hapus folder (workbook di dalamnya kembali ke "Tanpa Folder") ---------- */
router.delete('/folders/:id', authRequired, async (req, res) => {
  try {
    const id = toInt(req.params.id);
    if (!id) return res.status(400).json({ error: 'ID tidak valid' });

    // ON DELETE SET NULL di tabel sheet_folder yang mengurus workbook-nya
    const r = await pool.query(
      'DELETE FROM folders WHERE id = $1 AND owner_id = $2',
      [id, req.user.id]
    );
    if (!r.rowCount) return res.status(404).json({ error: 'Folder tidak ditemukan' });
    res.json({ ok: true });
  } catch (e) {
    console.error('DELETE /folders/:id', e);
    res.status(500).json({ error: 'Gagal menghapus folder' });
  }
});

/* ---------- Pindahkan workbook ke folder (folder_id null = Tanpa Folder) ---------- */
router.put('/sheets/:id/folder', authRequired, async (req, res) => {
  try {
    const sheetId = toInt(req.params.id);
    const folderId = req.body?.folder_id == null ? null : toInt(req.body.folder_id);
    if (!sheetId) return res.status(400).json({ error: 'ID workbook tidak valid' });
    if (req.body?.folder_id != null && !folderId) {
      return res.status(400).json({ error: 'ID folder tidak valid' });
    }

    // Folder tujuan harus milik user ini
    if (folderId !== null) {
      const f = await pool.query(
        'SELECT 1 FROM folders WHERE id = $1 AND owner_id = $2',
        [folderId, req.user.id]
      );
      if (!f.rowCount) return res.status(404).json({ error: 'Folder tidak ditemukan' });
    }

    // Workbook harus ada. (Data yang disimpan hanya penempatan folder milik user ini sendiri,
    // jadi tidak membuka akses ke isi workbook orang lain.)
    const s = await pool.query('SELECT 1 FROM sheets WHERE id = $1', [sheetId]);
    if (!s.rowCount) return res.status(404).json({ error: 'Workbook tidak ditemukan' });

    await pool.query(
      `INSERT INTO sheet_folder (user_id, sheet_id, folder_id)
       VALUES ($1, $2, $3)
       ON CONFLICT (user_id, sheet_id)
       DO UPDATE SET folder_id = EXCLUDED.folder_id`,
      [req.user.id, sheetId, folderId]
    );
    res.json({ ok: true });
  } catch (e) {
    console.error('PUT /sheets/:id/folder', e);
    res.status(500).json({ error: 'Gagal memindahkan workbook' });
  }
});

export default router;