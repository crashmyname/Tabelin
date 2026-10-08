import { Router } from 'express';
import { pool } from '../db.js';
import { computeSheet } from '../utils/spreadsheet.js';

const router = Router();

const wrap = fn => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

/* Auth via X-API-Key atau ?api_key= */
async function apiKeyAuth(req, res, next) {
  const key = req.headers['x-api-key'] || req.query.api_key;
  if (!key) return res.status(401).json({ error: 'API key required (X-API-Key header atau ?api_key=)' });

  const { rows } = await pool.query(
    `SELECT * FROM api_keys WHERE key = $1::text`,
    [String(key)]
  );
  if (!rows.length) return res.status(401).json({ error: 'Invalid API key' });

  const k = rows[0];
  // update usage — fire-and-forget
  pool.query(
    `UPDATE api_keys SET last_used_at = NOW(), request_count = request_count + 1 WHERE id = $1::int`,
    [k.id]
  ).catch(() => {});

  req.apiKey = k;
  next();
}

/* Helper: cek akses sheet */
async function getAccessibleSheet(apiKey, sheetId) {
  const sid = parseInt(sheetId, 10);
  if (isNaN(sid)) return null;
  if (apiKey.sheet_id && apiKey.sheet_id !== sid) return null;
  const { rows } = await pool.query(
    `SELECT id, name, data, updated_at FROM sheets
     WHERE id = $1::int AND user_id = $2::int`,
    [sid, apiKey.user_id]
  );
  return rows[0] || null;
}

router.use(apiKeyAuth);

/* ============ LIST SHEETS ============ */
router.get('/sheets', wrap(async (req, res) => {
  const k = req.apiKey;
  const sql = k.sheet_id
    ? `SELECT id, name, updated_at FROM sheets WHERE id = $1::int AND user_id = $2::int`
    : `SELECT id, name, updated_at FROM sheets WHERE user_id = $1::int ORDER BY updated_at DESC`;
  const params = k.sheet_id ? [k.sheet_id, k.user_id] : [k.user_id];
  const { rows } = await pool.query(sql, params);
  res.json({ sheets: rows });
}));

/* ============ SHEET METADATA ============ */
router.get('/sheets/:id', wrap(async (req, res) => {
  const sheet = await getAccessibleSheet(req.apiKey, req.params.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });
  const data = sheet.data || {};
  res.json({
    id: sheet.id,
    name: sheet.name,
    updated_at: sheet.updated_at,
    sheets: (data.sheets || []).map(s => ({
      name: s.name,
      cellCount: Object.keys(s.cells || {}).length,
      charts: (s.charts || []).length,
      merged: (s.merged || []).length
    }))
  });
}));

/* ============ COMPUTED VALUES (formula-evaluated) ============ */
router.get('/sheets/:id/data', wrap(async (req, res) => {
  const sheet = await getAccessibleSheet(req.apiKey, req.params.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const data = sheet.data || {};
  const sheets = data.sheets || [];
  const idx = parseInt(req.query.sheet_idx ?? data.activeSheet ?? 0, 10);
  const s = sheets[idx] || sheets[0];
  if (!s) return res.json({ rows: 0, cols: 0, values: [] });

  const { values, rows, cols } = computeSheet(s.cells || {});
  res.json({
    sheet: { index: idx, name: s.name },
    rows, cols,
    values
  });
}));

/* ============ RAW CELLS ============ */
router.get('/sheets/:id/cells', wrap(async (req, res) => {
  const sheet = await getAccessibleSheet(req.apiKey, req.params.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const data = sheet.data || {};
  const sheets = data.sheets || [];
  const idx = parseInt(req.query.sheet_idx ?? data.activeSheet ?? 0, 10);
  const s = sheets[idx] || sheets[0];
  if (!s) return res.json({ cells: {} });

  res.json({
    sheet: { index: idx, name: s.name },
    cells: s.cells || {},
    merged: s.merged || []
  });
}));

/* ============ ROW-OBJECTS (header row = keys) ============ */
router.get('/sheets/:id/json', wrap(async (req, res) => {
  const sheet = await getAccessibleSheet(req.apiKey, req.params.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const data = sheet.data || {};
  const sheets = data.sheets || [];
  const idx = parseInt(req.query.sheet_idx ?? data.activeSheet ?? 0, 10);
  const s = sheets[idx] || sheets[0];
  if (!s) return res.json({ data: [] });

  const { values, rows, cols } = computeSheet(s.cells || {});
  if (rows === 0) return res.json({ data: [] });

  // Baris pertama = header, sisanya = data
  const header = values[0].map((v, i) => String(v || `col_${i + 1}`));
  const out = [];
  for (let r = 1; r < rows; r++) {
    const obj = {};
    for (let c = 0; c < cols; c++) {
      obj[header[c]] = values[r][c];
    }
    out.push(obj);
  }
  res.json({ sheet: { index: idx, name: s.name }, data: out });
}));

/* ============ CSV ============ */
router.get('/sheets/:id/csv', wrap(async (req, res) => {
  const sheet = await getAccessibleSheet(req.apiKey, req.params.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const data = sheet.data || {};
  const sheets = data.sheets || [];
  const idx = parseInt(req.query.sheet_idx ?? data.activeSheet ?? 0, 10);
  const s = sheets[idx] || sheets[0];

  if (!s) {
    res.setHeader('Content-Type', 'text/csv');
    return res.send('');
  }

  const { values } = computeSheet(s.cells || {});
  const csv = values.map(row =>
    row.map(v => {
      const t = v === null || v === undefined ? '' : String(v);
      return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
    }).join(',')
  ).join('\r\n');

  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition',
    `inline; filename="${encodeURIComponent(sheet.name)}.csv"`);
  res.send('\uFEFF' + csv);
}));

/* Root — API docs ringkas */
router.get('/', (req, res) => {
  res.json({
    name: 'Web Excel API',
    version: 'v1',
    auth: 'Header: X-API-Key: wex_xxx   atau   ?api_key=wex_xxx',
    endpoints: [
      'GET /api/v1/sheets',
      'GET /api/v1/sheets/:id',
      'GET /api/v1/sheets/:id/data',
      'GET /api/v1/sheets/:id/cells',
      'GET /api/v1/sheets/:id/json',
      'GET /api/v1/sheets/:id/csv'
    ]
  });
});

export default router;