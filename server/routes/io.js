import { Router } from 'express';
import multer from 'multer';
import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';
import { Readable } from 'stream';
import { pool } from '../db.js';
import { authRequired } from '../auth.js';
import { computeSheet, rcToA1 } from '../utils/spreadsheet.js';

const router = Router();
router.use(authRequired);
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

async function getSheet(id, uid) {
  const { rows } = await pool.query(
    'SELECT * FROM sheets WHERE id=$1 AND user_id=$2', [id, uid]);
  return rows[0] || null;
}

/* ---------- IMPORT ---------- */
router.post('/:id/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file' });

    const sheet = await getSheet(req.params.id, req.user.id);
    if (!sheet) return res.status(404).json({ error: 'Not found' });

    const wb = new ExcelJS.Workbook();
    const fname = req.file.originalname.toLowerCase();

    if (fname.endsWith('.csv')) {
      await wb.csv.read(Readable.from(req.file.buffer.toString('utf8')));
    } else {
      await wb.xlsx.load(req.file.buffer);
    }

    const ws = wb.worksheets[0];
    if (!ws) return res.status(400).json({ error: 'No worksheet' });

    const cells = {};
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const addr = rcToA1(r - 1, c - 1);
        const val = cell.value;
        if (val === null || val === undefined || val === '') return;
        if (val instanceof Date) return void (cells[addr] = { v: val.toISOString().slice(0, 10) });
        if (typeof val === 'object') {
          if (val.formula) return void (cells[addr] = { f: '=' + val.formula });
          if (val.result !== undefined) return void (cells[addr] = { v: val.result });
          if (val.text !== undefined) return void (cells[addr] = { v: val.text });
          return;
        }
        cells[addr] = { v: val };
      });
    });

    const data = { sheets: [{ name: sheet.name, cells }], activeSheet: 0 };

    // ← FIX: stringify + cast eksplisit
    await pool.query(
      `UPDATE sheets SET data = $2::jsonb, updated_at = NOW() WHERE id = $1::int`,
      [parseInt(req.params.id, 10), JSON.stringify(data)]
    );

    res.json({ ok: true, cells: Object.keys(cells).length });
  } catch (e) {
    console.error('Import error:', e);
    if (!res.headersSent) {
      res.status(500).json({ error: e.message || 'Import gagal' });
    }
  }
});

/* ---------- XLSX ---------- */
router.get('/:id/export/xlsx', async (req, res) => {
  const sheet = await getSheet(req.params.id, req.user.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const wbData = normalize(sheet.data);
  const wb = new ExcelJS.Workbook();
  for (const s of wbData.sheets) {
    const { values } = computeSheet(s.cells || {});
    const ws = wb.addWorksheet((s.name || 'Sheet').slice(0, 30));
    values.forEach(row => ws.addRow(row));
    ws.columns.forEach(c => (c.width = 18));
    // formatting ringan
    for (const [addr, cell] of Object.entries(s.cells || {})) {
      if (!cell.s) continue;
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) continue;
      let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
      const excelCell = ws.getCell(addr);
      const st = cell.s;
      if (st.bold) excelCell.font = { ...excelCell.font, bold: true };
      if (st.italic) excelCell.font = { ...excelCell.font, italic: true };
      if (st.color) excelCell.font = { ...excelCell.font, color: { argb: 'FF' + st.color.replace('#', '').toUpperCase() } };
      if (st.bg) excelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + st.bg.replace('#', '').toUpperCase() } };
      if (st.align) excelCell.alignment = { horizontal: st.align };
    }
  }
  const buf = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(sheet.name)}.xlsx"`);
  res.send(Buffer.from(buf));
});

/* ---------- CSV (sheet aktif) ---------- */
router.get('/:id/export/csv', async (req, res) => {
  const sheet = await getSheet(req.params.id, req.user.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });
  const wb = normalize(sheet.data);
  const activeIdx = wb.activeSheet || 0;
  const s = wb.sheets[activeIdx] || wb.sheets[0] || { cells: {} };
  const { values } = computeSheet(s.cells || {});
  const csv = values.map(row => row.map(v => {
    const t = v === null || v === undefined ? '' : String(v);
    return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
  }).join(',')).join('\r\n');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(sheet.name)}.csv"`);
  res.send('\uFEFF' + csv);
});

/* ---------- PDF (sheet aktif) ---------- */
router.get('/:id/export/pdf', async (req, res) => {
  const sheet = await getSheet(req.params.id, req.user.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });
  const wb = normalize(sheet.data);
  const activeIdx = wb.activeSheet || 0;
  const s = wb.sheets[activeIdx] || wb.sheets[0] || { cells: {} };
  const { values, rows, cols } = computeSheet(s.cells || {});

  const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 30 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(sheet.name)}.pdf"`);
  doc.pipe(res);
  doc.fontSize(16).fillColor('#111').text(`${sheet.name} — ${s.name || 'Sheet'}`);
  doc.moveDown(0.5);
  if (!rows || !cols) { doc.fontSize(10).fillColor('#888').text('(kosong)'); doc.end(); return; }

  const pageW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const cw = Math.max(40, Math.min(110, pageW / cols));
  const chh = 20;
  let y = doc.y;
  const x0 = doc.page.margins.left;
  for (let r = 0; r < rows; r++) {
    if (y + chh > doc.page.height - doc.page.margins.bottom) {
      doc.addPage(); y = doc.page.margins.top;
    }
    let x = x0;
    for (let c = 0; c < cols; c++) {
      doc.rect(x, y, cw, chh).strokeColor('#d0d0d0').lineWidth(0.5).stroke();
      doc.fontSize(8).fillColor('#111').text(String(values[r][c] ?? ''), x + 3, y + 6,
        { width: cw - 6, height: chh - 6, ellipsis: true, lineBreak: false });
      x += cw;
    }
    y += chh;
  }
  doc.end();
});

function normalize(data) {
  if (data && Array.isArray(data.sheets)) return data;
  return { sheets: [{ name: 'Sheet1', cells: data || {} }], activeSheet: 0 };
}

export default router;