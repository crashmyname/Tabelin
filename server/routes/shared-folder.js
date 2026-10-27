/* =========================================================
   server/routes/shared-folder.js
   Browse & save file dari folder sharing — TANPA DB
   ========================================================= */
import { Router } from 'express';
import express from 'express';
import fs from 'fs';
import path from 'path';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

/* =========================================================
   CONFIG
   ========================================================= */
// 🔧 DEFAULT_BASE dipakai kalau frontend tidak kirim ?base=
//    (kompatibel dengan frontend lama yang panggil /list saja)
const RAW_BASE = process.env.SHARED_BASE || '\\\\ISE-IRPTR-01\\Document_IReporter\\';
const DEFAULT_BASE = (() => {
  const n = RAW_BASE.replace(/\//g, '\\');
  return n.endsWith('\\') ? n : n + '\\';
})();

const ALLOW_ANY_PATH = true;
const ALLOWED_SERVERS = [];
const MAX_FILE_SIZE = 50 * 1024 * 1024;   // 50 MB
const BACKUP_KEEP   = 10;                  // 🔧 simpan maksimal 10 backup terakhir
const IS_PROD       = process.env.NODE_ENV === 'production';

/* =========================================================
   HELPERS
   ========================================================= */
// 🔧 Fallback base — supaya frontend lama (tanpa ?base=) tetap jalan
function getBase(req) {
  const raw = req.query.base || req.body?.base || DEFAULT_BASE;
  return validatePath(raw);
}

function validatePath(inputPath) {
  if (!inputPath || typeof inputPath !== 'string') {
    throw new Error('Path wajib diisi');
  }
  const trimmed = inputPath.trim();

  const isUNC      = /^\\\\[^\\]+\\[^\\]+/.test(trimmed);
  const isAbsolute = /^[A-Za-z]:[\\\/]/.test(trimmed);

  if (!isUNC && !isAbsolute) {
    throw new Error('Path harus UNC (\\\\server\\share\\...) atau drive (C:\\...)');
  }

  if (trimmed.split(/[\\\/]/).some(seg => seg === '..')) {
    throw new Error('Path tidak boleh mengandung ".."');
  }

  if (isUNC && !ALLOW_ANY_PATH) {
    const m = /^\\\\?([^\\]+)/.exec(trimmed);
    if (m) {
      const server = m[1].toUpperCase();
      if (!ALLOWED_SERVERS.includes(server)) {
        throw new Error(`Server "${server}" tidak diizinkan. Server yang diizinkan: ${ALLOWED_SERVERS.join(', ')}`);
      }
    }
  }

  let normalized = trimmed.replace(/\//g, '\\');
  if (!normalized.endsWith('\\')) normalized += '\\';
  return normalized;
}

// 🔧 Fix path traversal: path.relative, bukan startsWith
function resolveSafePath(basePath, relPath) {
  const baseResolved = path.resolve(basePath);
  if (!relPath) return baseResolved;

  const cleaned = String(relPath).replace(/[\\\/]+/g, path.sep).replace(/^[\\\/]+/, '');
  const full = path.resolve(baseResolved, cleaned);

  const rel = path.relative(baseResolved, full);
  if (rel === '..' || rel.startsWith('..' + path.sep) || path.isAbsolute(rel)) {
    throw new Error('Path di luar folder sharing');
  }
  return full;
}

// 🔧 Sanitize error message supaya path internal tidak bocor di prod
function safeErr(e) {
  if (!IS_PROD) return e.message;
  if (e.code === 'ENOENT')                  return 'File/folder tidak ditemukan';
  if (e.code === 'EACCES' || e.code === 'EPERM') return 'Akses ditolak';
  return 'Terjadi kesalahan server';
}

// 🔧 Rotasi backup: nama timestamp, keep N terbaru
function rotateBackups(fullPath) {
  try {
    const dir  = path.dirname(fullPath);
    const base = path.basename(fullPath);
    const prefix = base + '.';
    const suffix = '.backup';

    const files = fs.readdirSync(dir)
      .filter(n => n.startsWith(prefix) && n.endsWith(suffix))
      .map(n => ({ name: n, mtime: fs.statSync(path.join(dir, n)).mtimeMs }))
      .sort((a, b) => b.mtime - a.mtime);

    for (let i = BACKUP_KEEP; i < files.length; i++) {
      try { fs.unlinkSync(path.join(dir, files[i].name)); } catch {}
    }
  } catch { /* best-effort */ }
}

// 🔧 Hidden dari listing: backup & file test
function isHiddenEntry(name) {
  return name.endsWith('.backup') || name.startsWith('.tabelin_test_');
}

/* =========================================================
   POST /test — test folder bisa diakses
   🔧 Tambah express.json() inline supaya tidak bergantung global
   ========================================================= */
router.post('/test', express.json({ limit: '1mb' }), async (req, res) => {
  try {
    const validated = validatePath(req.body?.path);

    if (!fs.existsSync(validated)) {
      return res.json({ ok: false, error: 'Folder tidak ditemukan / tidak bisa diakses', normalizedPath: validated });
    }

    const stat = fs.statSync(validated);
    if (!stat.isDirectory()) {
      return res.json({ ok: false, error: 'Path bukan folder', normalizedPath: validated });
    }

    try { fs.readdirSync(validated, { withFileTypes: true }); }
    catch (e) { return res.json({ ok: false, error: 'Tidak bisa baca: ' + safeErr(e) }); }

    let writable = false;
    try {
      const testFile = path.join(validated, '.tabelin_test_' + Date.now());
      fs.writeFileSync(testFile, 'x');
      fs.unlinkSync(testFile);
      writable = true;
    } catch {}

    res.json({ ok: true, normalizedPath: validated, writable, readonly: !writable });
  } catch (e) {
    res.status(400).json({ ok: false, error: safeErr(e) });
  }
});

/* =========================================================
   GET /list?base=path&path=subfolder
   ========================================================= */
router.get('/list', async (req, res) => {
  try {
    const basePath = getBase(req);              // 🔧 fallback ke DEFAULT_BASE
    const relPath  = req.query.path || '';
    const fullPath = resolveSafePath(basePath, relPath);

    console.log(`📂 [${req.user?.username || '?'}] List: ${fullPath}`);

    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'Folder tidak ditemukan' });
    }
    const stat = fs.statSync(fullPath);
    if (!stat.isDirectory()) return res.status(400).json({ error: 'Bukan folder' });

    const entries = fs.readdirSync(fullPath, { withFileTypes: true })
      .filter(d => !isHiddenEntry(d.name));     // 🔧 sembunyikan .backup

    const items = entries.map(d => {
      const itemPath = path.join(fullPath, d.name);
      let stats = null;
      try { stats = fs.statSync(itemPath); } catch {}

      return {
        name: d.name,
        isDir: d.isDirectory(),
        size: stats?.size || 0,
        mtime: stats?.mtime?.toISOString() || null,
        ext: d.isDirectory() ? null : path.extname(d.name).toLowerCase().slice(1)
      };
    }).sort((a, b) => {
      if (a.isDir !== b.isDir) return a.isDir ? -1 : 1;
      return a.name.localeCompare(b.name);
    });

    res.json({ basePath, path: relPath, items });
  } catch (e) {
    console.error('list error:', e);
    res.status(400).json({ error: safeErr(e) });
  }
});

/* =========================================================
   GET /file?base=path&path=file.xlsx
   ========================================================= */
router.get('/file', async (req, res) => {
  try {
    const basePath = getBase(req);              // 🔧
    const relPath  = req.query.path;
    if (!relPath) return res.status(400).json({ error: 'Path file wajib' });

    const fullPath = resolveSafePath(basePath, relPath);
    console.log(`📥 [${req.user?.username || '?'}] Read: ${fullPath}`);

    if (!fs.existsSync(fullPath)) return res.status(404).json({ error: 'File tidak ditemukan' });

    const stat = fs.statSync(fullPath);
    if (!stat.isFile()) return res.status(400).json({ error: 'Bukan file' });
    if (stat.size > MAX_FILE_SIZE) return res.status(413).json({ error: 'File > 50 MB' });

    const ext = path.extname(fullPath).toLowerCase();
    const mimeMap = {
      '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      '.xls':  'application/vnd.ms-excel',
      '.csv':  'text/csv'
    };

    // 🔧 Content-Disposition dengan encoding RFC 5987 (aman untuk nama non-ASCII)
    const fname = path.basename(fullPath);
    const ascii = fname.replace(/[^\x20-\x7E]/g, '_').replace(/"/g, '');
    res.setHeader('Content-Disposition',
      `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(fname)}`);

    res.setHeader('Content-Type', mimeMap[ext] || 'application/octet-stream');
    res.setHeader('Content-Length', stat.size);
    res.setHeader('X-Last-Modified', stat.mtime.toISOString());

    const stream = fs.createReadStream(fullPath);
    stream.on('error', (e) => {
      console.error('stream error:', e);
      if (!res.headersSent) res.status(500).json({ error: safeErr(e) });
    });
    stream.pipe(res);
  } catch (e) {
    console.error('read error:', e);
    if (!res.headersSent) res.status(400).json({ error: safeErr(e) });
  }
});

/* =========================================================
   POST /file?base=path&path=file.xlsx
   ========================================================= */
router.post('/file', async (req, res) => {
  try {
    const basePath = getBase(req);              // 🔧
    const relPath  = req.query.path;
    if (!relPath) return res.status(400).json({ error: 'Path file wajib' });

    const fullPath = resolveSafePath(basePath, relPath);
    console.log(`💾 [${req.user?.username || '?'}] Write: ${fullPath}`);

    // Lock check via header
    const expectedMtime = req.headers['x-expected-mtime'];
    if (expectedMtime && fs.existsSync(fullPath)) {
      const currentMtime = fs.statSync(fullPath).mtime.toISOString();
      if (currentMtime !== expectedMtime) {
        return res.status(409).json({
          error: 'File sudah diubah user lain. Reload dulu.',
          currentMtime
        });
      }
    }

    // 🔧 Baca body dengan size limit (cegah OOM)
    const chunks = [];
    let total = 0;
    for await (const chunk of req) {
      total += chunk.length;
      if (total > MAX_FILE_SIZE) {
        // drain sisa stream biar koneksi tidak macet
        try { for await (const _ of req) { /* discard */ } } catch {}
        return res.status(413).json({ error: 'File > 50 MB' });
      }
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    if (!buffer.length) return res.status(400).json({ error: 'File kosong' });

    // 🔧 Backup timestamped + retensi (tidak menimpa backup lama)
    if (fs.existsSync(fullPath)) {
      try {
        const ts = new Date().toISOString().replace(/[:.]/g, '-');
        fs.copyFileSync(fullPath, `${fullPath}.${ts}.backup`);
        rotateBackups(fullPath);
      } catch (e) {
        console.warn('backup failed (lanjut):', e.message);
      }
    }

    fs.writeFileSync(fullPath, buffer);
    const newStat = fs.statSync(fullPath);
    console.log(`✅ Saved ${buffer.length} bytes`);

    res.json({ ok: true, size: buffer.length, mtime: newStat.mtime.toISOString() });
  } catch (e) {
    console.error('write error:', e);
    res.status(400).json({ error: safeErr(e) });
  }
});

export default router;