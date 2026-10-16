import { Router } from 'express';
import { authRequired } from '../auth.js';

const router = Router();
router.use(authRequired);

/* =========================================================
   POST /api/sheets/fetch-url
   Body: { url, headers? }
   Fetch konten dari URL (server-side untuk hindari CORS)
   ========================================================= */
router.post('/fetch-url', async (req, res) => {
  try {
    const { url, headers = {} } = req.body || {};

    if (!url) return res.status(400).json({ error: 'URL wajib diisi' });

    // 🎯 Validasi URL — hanya http/https
    let parsedUrl;
    try {
      parsedUrl = new URL(url);
    } catch {
      return res.status(400).json({ error: 'URL tidak valid' });
    }

    if (!['http:', 'https:'].includes(parsedUrl.protocol)) {
      return res.status(400).json({ error: 'Hanya HTTP/HTTPS yang didukung' });
    }

    // 🎯 Blokir URL internal (SSRF protection)
    const hostname = parsedUrl.hostname.toLowerCase();
    const blockedHosts = [
      'localhost', '127.0.0.1', '0.0.0.0', '::1',
      '169.254.', '10.', '192.168.', '172.16.', '172.17.', '172.18.',
      '172.19.', '172.20.', '172.21.', '172.22.', '172.23.', '172.24.',
      '172.25.', '172.26.', '172.27.', '172.28.', '172.29.', '172.30.',
      '172.31.'
    ];

    for (const b of blockedHosts) {
      if (hostname === b || hostname.startsWith(b)) {
        return res.status(400).json({ error: 'URL tidak diizinkan (internal network)' });
      }
    }

    // 🎯 Fetch dengan timeout 30 detik
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);

    const t0 = Date.now();
    const fetchRes = await fetch(url, {
      headers: {
        'User-Agent': 'Tabelin/2.0 (+https://tabelin.app)',
        ...headers
      },
      signal: controller.signal,
      redirect: 'follow'
    });

    clearTimeout(timeout);

    if (!fetchRes.ok) {
      return res.status(400).json({
        error: `HTTP ${fetchRes.status} ${fetchRes.statusText}`
      });
    }

    // 🎯 Batasi ukuran 10 MB
    const contentLength = parseInt(fetchRes.headers.get('content-length') || '0', 10);
    if (contentLength > 10 * 1024 * 1024) {
      return res.status(413).json({ error: 'File terlalu besar (maks 10 MB)' });
    }

    const text = await fetchRes.text();

    if (text.length > 10 * 1024 * 1024) {
      return res.status(413).json({ error: 'File terlalu besar (maks 10 MB)' });
    }

    const elapsed = Date.now() - t0;
    console.log(`🌐 Fetched URL: ${url} (${(text.length / 1024).toFixed(1)} KB, ${elapsed}ms)`);

    res.json({
      ok: true,
      content: text,
      contentType: fetchRes.headers.get('content-type') || '',
      size: text.length,
      elapsed
    });

  } catch (e) {
    if (e.name === 'AbortError') {
      return res.status(408).json({ error: 'Timeout — server tidak merespons dalam 30 detik' });
    }
    console.error('fetch-url error:', e);
    res.status(500).json({ error: e.message });
  }
});

export default router;