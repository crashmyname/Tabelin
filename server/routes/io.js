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
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }   // 50 MB
});

async function getSheet(id, uid) {
  const { rows } = await pool.query(
    'SELECT * FROM sheets WHERE id=$1 AND user_id=$2', [id, uid]);
  return rows[0] || null;
}

/* ---------- IMPORT ---------- */
/* =========================================================
   IMPORT — dengan auto-chunking untuk data besar
   ========================================================= */
/* =========================================================
   IMPORT — Multi-Sheet dengan preserve struktur Excel
   ========================================================= */
router.post('/:id/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file' });

    const sheet = await getSheet(req.params.id, req.user.id);
    if (!sheet) return res.status(404).json({ error: 'Not found' });

    console.log(`📥 Import mulai: ${req.file.originalname} (${(req.file.size / 1024 / 1024).toFixed(2)} MB)`);

    const wb = new ExcelJS.Workbook();
    const fname = req.file.originalname.toLowerCase();

    if (fname.endsWith('.csv')) {
      await wb.csv.read(Readable.from(req.file.buffer.toString('utf8')));
    } else {
      await wb.xlsx.load(req.file.buffer);
    }

    if (!wb.worksheets || !wb.worksheets.length) {
      return res.status(400).json({ error: 'File tidak punya worksheet' });
    }

    console.log(`📊 Total sheet di Excel: ${wb.worksheets.length}`);

    // =====================================================
    // 🎯 LOOP SEMUA WORKSHEET
    // =====================================================
    const allSheetsData = [];
    let totalCells = 0;

    wb.worksheets.forEach((ws, sheetIdx) => {
      const sheetName = ws.name || `Sheet${sheetIdx + 1}`;
      const cells = {};
      const merged = [];
      const colWidths = {};
      const rowHeights = {};

      let maxRow = 0;
      let maxCol = 0;

      // -----------------------------------------------
      // 🎯 Parse cells + formatting
      // -----------------------------------------------
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        if (r > maxRow) maxRow = r;

        // Row height
        if (row.height && row.height > 15) {
          rowHeights[r - 1] = Math.round(row.height * 1.33); // pt → px approximate
        }

        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if (c > maxCol) maxCol = c;

          const addr = rcToA1(r - 1, c - 1);
          const val = cell.value;
          if (val === null || val === undefined || val === '') {
            // Cek format saja
            if (cell.style && hasStyle(cell.style)) {
              cells[addr] = { s: extractStyle(cell.style) };
            }
            return;
          }

          const cellObj = {};

          // Nilai
          if (val instanceof Date) {
            cellObj.v = val.toISOString().slice(0, 10);
          } else if (typeof val === 'object') {
            if (val.formula) {
              cellObj.f = '=' + val.formula;
            } else if (val.result !== undefined) {
              cellObj.v = val.result;
            } else if (val.text !== undefined) {
              cellObj.v = val.text;
            } else if (val.richText) {
              cellObj.v = val.richText.map(r => r.text).join('');
            } else if (val.hyperlink) {
              cellObj.v = val.text || val.hyperlink;
              cellObj.link = { url: val.hyperlink, text: val.text || val.hyperlink };
              cellObj.s = { color: '#2563eb', underline: true };
            } else {
              cellObj.v = String(val);
            }
          } else {
            cellObj.v = val;
          }

          // Style
          if (cell.style && hasStyle(cell.style)) {
            cellObj.s = { ...(cellObj.s || {}), ...extractStyle(cell.style) };
          }

          // Number format
          if (cell.numFmt) {
            cellObj.s = cellObj.s || {};
            if (/^[#0,]+\.?\d*$/.test(cell.numFmt)) cellObj.s.numFmt = 'number';
            else if (/Rp|\\"Rp\\"/.test(cell.numFmt) || /[$€£]/.test(cell.numFmt)) cellObj.s.numFmt = 'currency';
            else if (/%/.test(cell.numFmt)) cellObj.s.numFmt = 'percent';
            else if (/[ymd]/i.test(cell.numFmt) && /\d/.test(cell.numFmt)) cellObj.s.numFmt = 'date';
          }

          cells[addr] = cellObj;
          totalCells++;
        });
      });

      // -----------------------------------------------
      // 🎯 Merged cells
      // -----------------------------------------------
      if (ws.model && ws.model.merges && ws.model.merges.length) {
        for (const mergeRange of ws.model.merges) {
          // Format: "A1:C1" atau "A1"
          if (mergeRange && mergeRange.includes(':')) {
            merged.push(mergeRange.toUpperCase());
          } else if (mergeRange) {
            // Single cell merge tidak valid, skip
          }
        }
      }

      // -----------------------------------------------
      // 🎯 Column widths
      // -----------------------------------------------
      if (ws.columns) {
        ws.columns.forEach((col, idx) => {
          if (col && col.width) {
            // Excel width (chars) → pixel approximate
            const px = Math.round(col.width * 7 + 5);
            colWidths[idx] = Math.max(40, Math.min(px, 500));
          }
        });
      }

      console.log(`   📄 Sheet "${sheetName}": ${Object.keys(cells).length} cells, ${merged.length} merges`);

      allSheetsData.push({
        name: sheetName,
        cells,
        colWidths,
        rowHeight: 26,
        cf: [],
        charts: [],
        merged,
        filters: {},
        images: []
      });
    });

    // =====================================================
    // 🎯 Build workbook data
    // =====================================================
    const data = {
      sheets: allSheetsData,
      activeSheet: 0
    };

    const dataSize = JSON.stringify(data).length;
    console.log(`💾 Total data: ${(dataSize / 1024 / 1024).toFixed(2)} MB, ${totalCells} cells, ${allSheetsData.length} sheets`);

    // =====================================================
    // 🎯 Save ke server
    // =====================================================
    await pool.query(
      `UPDATE sheets
          SET data = $2::jsonb,
              updated_at = NOW()
        WHERE id = $1::int`,
      [parseInt(req.params.id, 10), JSON.stringify(data)]
    );

    // Optional: auto-chunk kalau data besar
    const THRESHOLD = 5 * 1024 * 1024;
    if (dataSize > THRESHOLD) {
      console.log('📦 Data > 5 MB — auto-chunking...');
      // Trigger chunk migration async (tidak blocking response)
      setImmediate(async () => {
        try {
          const BLOCK_ROWS = 256;
          await pool.query('DELETE FROM sheet_chunks WHERE sheet_id = $1::int',
            [parseInt(req.params.id, 10)]);

          for (let sIdx = 0; sIdx < allSheetsData.length; sIdx++) {
            const cells = allSheetsData[sIdx].cells;
            const blocks = {};
            for (const [addr, cell] of Object.entries(cells)) {
              const m = /^([A-Z]+)(\d+)$/.exec(addr);
              if (!m) continue;
              const row = parseInt(m[2], 10) - 1;
              const blockRow = Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;
              if (!blocks[blockRow]) blocks[blockRow] = {};
              blocks[blockRow][addr] = cell;
            }
            for (const [blockRow, blockCells] of Object.entries(blocks)) {
              await pool.query(
                `INSERT INTO sheet_chunks (sheet_id, sheet_idx, block_row, cells)
                 VALUES ($1::int, $2::int, $3::int, $4::jsonb)`,
                [parseInt(req.params.id, 10), sIdx, parseInt(blockRow, 10), JSON.stringify(blockCells)]
              );
            }
          }
          console.log('✅ Chunks created');
        } catch (e) {
          console.error('Chunk migration failed:', e);
        }
      });
    }

    // =====================================================
    // 🎯 Response
    // =====================================================
    res.json({
      ok: true,
      cells: totalCells,
      sheets: allSheetsData.length,
      sheetNames: allSheetsData.map(s => s.name),
      dataSize
    });

  } catch (e) {
    console.error('❌ Import error:', e);
    if (!res.headersSent) {
      res.status(500).json({ error: e.message || 'Import gagal' });
    }
  }
});

/* =========================================================
   HELPER: Cek & extract style dari ExcelJS cell
   ========================================================= */
function hasStyle(style) {
  if (!style) return false;
  return !!(
    style.font?.bold ||
    style.font?.italic ||
    style.font?.underline ||
    style.font?.strike ||
    style.font?.color?.argb ||
    style.font?.size ||
    style.font?.name ||
    style.fill?.fgColor?.argb ||
    style.alignment?.horizontal ||
    style.alignment?.wrapText ||
    style.border?.top?.style ||
    style.border?.bottom?.style ||
    style.border?.left?.style ||
    style.border?.right?.style
  );
}

function extractStyle(style) {
  const s = {};

  /* ---------- FONT ---------- */
  if (style.font?.bold)      s.bold = true;
  if (style.font?.italic)    s.italic = true;
  if (style.font?.underline) s.underline = true;
  if (style.font?.strike)    s.strike = true;

  if (style.font?.name) {
    s.fontFamily = style.font.name;
  }

  if (style.font?.color?.argb) {
    const argb = style.font.color.argb;
    if (argb && argb.length === 8) {
      const hex = argb.slice(2);
      // skip hitam default biar tidak overwrite
      if (hex.toUpperCase() !== '000000') {
        s.color = '#' + hex;
      }
    }
  }

  if (style.font?.size) {
    s.fontSize = style.font.size;
  }

  /* ---------- FILL ---------- */
  if (style.fill) {
    if (
      (style.fill.type === 'pattern' || style.fill.type === 'solid') &&
      style.fill.fgColor?.argb
    ) {
      const argb = style.fill.fgColor.argb;
      if (argb && argb.length === 8) {
        const hex = argb.slice(2);
        // skip putih default
        if (hex.toUpperCase() !== 'FFFFFF') {
          s.bg = '#' + hex;
        }
      }
    }
  }

  /* ---------- ALIGNMENT ---------- */
  if (style.alignment?.horizontal) {
    const h = style.alignment.horizontal;
    s.align = h === 'centerContinuous' ? 'center' : h;
  }

  if (style.alignment?.wrapText) {
    s.wrap = true;
  }

  /* ---------- BORDER ---------- */
  if (style.border) {
    const b = {};
    let hasBorder = false;
    let borderStyle = null;
    let borderColor = null;

    ['top', 'bottom', 'left', 'right'].forEach(side => {
      const bd = style.border[side];
      if (bd && bd.style) {
        b[side] = true;
        hasBorder = true;
        if (!borderStyle) {
          // map Excel style → app style
          switch (bd.style) {
            case 'medium': borderStyle = 'medium'; break;
            case 'thick':  borderStyle = 'thick';  break;
            case 'dashed': borderStyle = 'dashed'; break;
            case 'dotted': borderStyle = 'dotted'; break;
            case 'double': borderStyle = 'double'; break;
            default:       borderStyle = 'thin';
          }
          if (bd.color?.argb && bd.color.argb.length === 8) {
            const hex = bd.color.argb.slice(2);
            borderColor = hex.toUpperCase() === '000000' ? '#111111' : '#' + hex;
          } else {
            borderColor = '#111111';
          }
        }
      }
    });

    if (hasBorder) {
      b.style = borderStyle || 'thin';
      b.color = borderColor || '#111111';
      s.borders = b;
    }
  }

  return s;
}

/* ---------- XLSX ---------- */
/* server/routes/io.js — GANTI route /export/xlsx dengan ini */
router.get('/:id/export/xlsx', async (req, res) => {
  const sheet = await getSheet(req.params.id, req.user.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const wbData = normalize(sheet.data);
  const fname = (sheet.name || 'workbook').replace(/[^\w\d-]+/g, '_');

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fname)}.xlsx"`);
  res.setHeader('Cache-Control', 'no-cache');

  // 🎯 STREAMING writer — hemat memory
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: res,
    useStyles: true,
    useSharedStrings: false   // ← penting: hemat RAM besar
  });

  try {
    for (const s of wbData.sheets) {
      const ws = workbook.addWorksheet((s.name || 'Sheet').slice(0, 30));

      // Column widths
      if (s.colWidths) {
        for (const [i, w] of Object.entries(s.colWidths)) {
          ws.getColumn(+i + 1).width = Math.max(8, (w - 5) / 7.5);
        }
      }

      // Cari max row/col
      const cells = s.cells || {};
      let maxR = 0, maxC = 0;
      for (const addr of Object.keys(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(addr);
        if (!m) continue;
        let c = 0;
        for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
        const r = parseInt(m[2], 10);
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      }

      // 🎯 Streaming: tulis per baris, langsung flush
      for (let r = 1; r <= maxR; r++) {
        const row = ws.getRow(r);

        for (let c = 1; c <= maxC; c++) {
          const addr = colLetter(c) + r;
          const cl = cells[addr];
          if (!cl) continue;

          const excelCell = row.getCell(c);
          // Value: pakai v mentah (bukan HF-evaluated — biar cepat)
          // Kalau perlu HF-eval, batasi 10.000 baris pertama saja
          if (cl.f) {
            excelCell.value = {
              formula: cl.f.slice(1),  // buang '='
              result: cl.v ?? ''       // pakai cached result kalau ada
            };
          } else if (cl.v !== undefined) {
            excelCell.value = cl.v;
          }

          // Style — light: cuma font bold + bg
          if (cl.s) {
            if (cl.s.bold) excelCell.font = { bold: true };
            if (cl.s.bg) {
              const hex = cl.s.bg.replace('#', '').toUpperCase();
              excelCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF' + hex }
              };
            }
          }
        }

        // Commit row → flush ke stream
        row.commit();
      }

      // Merges
      (s.merged || []).forEach(range => {
        try { ws.mergeCells(range); } catch {}
      });

      ws.commit();
    }

    await workbook.commit();
  } catch (e) {
    console.error('XLSX stream error:', e);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* ---------- CSV (sheet aktif) ---------- */
/* server/routes/io.js — GANTI route /export/csv */
router.get('/:id/export/csv', async (req, res) => {
  const sheet = await getSheet(req.params.id, req.user.id);
  if (!sheet) return res.status(404).json({ error: 'Not found' });

  const wbData = normalize(sheet.data);
  const activeIdx = wbData.activeSheet || 0;
  const s = wbData.sheets[activeIdx] || wbData.sheets[0] || { cells: {} };

  const fname = (sheet.name || 'export').replace(/[^\w\d-]+/g, '_');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fname)}.csv"`);

  // 🎯 BOM untuk Excel
  res.write('\uFEFF');

  // Cari max
  const cells = s.cells || {};
  let maxR = 0, maxC = 0;
  for (const addr of Object.keys(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    if (!m) continue;
    let c = 0;
    for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    const r = parseInt(m[2], 10);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }

  // 🎯 Streaming per baris — no HF, no computeSheet
  const BATCH = 5000;
  const esc = (t) => {
    t = t == null ? '' : String(t);
    return /[",\n\r]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
  };

  for (let startR = 1; startR <= maxR; startR += BATCH) {
    const endR = Math.min(startR + BATCH - 1, maxR);
    const lines = [];

    for (let r = startR; r <= endR; r++) {
      const row = [];
      for (let c = 1; c <= maxC; c++) {
        const addr = colLetter(c) + r;
        const cl = cells[addr];
        row.push(cl ? esc(cl.v ?? '') : '');
      }
      lines.push(row.join(','));
    }

    res.write(lines.join('\r\n') + '\r\n');

    // Backpressure — tunggu drain kalau buffer penuh
    if (!res.write('')) {
      await new Promise(r => res.once('drain', r));
    }
  }

  res.end();
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