/* =========================================================
   server/utils/excel-import.js
   Parse Excel (.xlsx) via ExcelJS → schema wbData
   ========================================================= */
import ExcelJS from 'exceljs';

/* Kolom angka → huruf */
function colLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* ARGB (AARRGGBB) → #RRGGBB */
function argbToHex(argb) {
  if (!argb || typeof argb !== 'string') return null;
  // ExcelJS kadang kasih 'FFRRGGBB', kadang 'RRGGBB'
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  if (hex.length !== 6) return null;
  if (hex.toUpperCase() === 'FFFFFF') return null; // skip putih
  return '#' + hex;
}

/* Map Excel number format → tag simpel */
function mapNumFmt(fmt) {
  if (!fmt) return null;
  const f = String(fmt).toLowerCase();
  if (f.includes('"') || f.includes('[') || f.length > 40) return { custom: fmt };
  if (f.includes('%')) return 'percent';
  if (/[$€£]|rp/.test(f)) return 'currency';
  if (/[ymd]/.test(f) && /[/-]/.test(f)) return 'date';
  if (f.includes('#,##0') || f === '0.00' || f === '0') return 'number';
  return 'general';
}

/* Map border style Excel → tag aplikasi */
function mapBorderStyle(s) {
  switch (s) {
    case 'medium': return 'medium';
    case 'thick':  return 'thick';
    case 'dashed': return 'dashed';
    case 'dotted': return 'dotted';
    case 'double': return 'double';
    default:       return 'thin';
  }
}

/* =========================================================
   MAIN: parse XLSX buffer → array of sheet objects
   ========================================================= */
export async function parseExcelBuffer(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const sheets = [];

  wb.eachSheet((ws) => {
    const cells = {};
    const merged = [];
    const colWidths = {};
    const rowHeights = {};
    const cf = [];
    const validations = [];

    /* ---------- CELLS + STYLE ---------- */
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
      row.eachCell({ includeEmpty: false }, (cell, colNum) => {
        const addr = colLetter(colNum) + rowNum;
        const v = cell.value;

        if (v == null || v === '') return;

        const obj = {};

        /* --- Value & Formula --- */
        if (v && typeof v === 'object') {
          if (v.formula) {
            // Formula: prepend '=' dan biarkan absolute $ tetap
            obj.f = '=' + v.formula;
            // Simpan result sebagai v (biar kelihatan sebelum HF evaluate)
            if (v.result !== undefined && v.result !== null) {
              obj.v = v.result;
            }
          } else if (v.richText) {
            obj.v = v.richText.map(r => r.text).join('');
          } else if (v.hyperlink) {
            obj.v = v.text || v.hyperlink;
            obj.link = { url: v.hyperlink, text: v.text || v.hyperlink };
          } else if (v.sharedFormula) {
            obj.f = '=' + v.sharedFormula;
          } else if (v.error) {
            obj.v = v.error;
          } else if (v.result !== undefined) {
            obj.v = v.result;
          } else {
            obj.v = String(v);
          }
        } else if (v instanceof Date) {
          obj.v = v.getTime();
        } else {
          obj.v = v;
        }

        /* --- Style --- */
        const s = {};

        // Font
        if (cell.font) {
          if (cell.font.bold)      s.bold = true;
          if (cell.font.italic)    s.italic = true;
          if (cell.font.underline) s.underline = true;
          if (cell.font.strike)    s.strike = true;
          if (cell.font.name)      s.fontFamily = cell.font.name;
          if (cell.font.size)      s.fontSize = cell.font.size;
          const fg = argbToHex(cell.font.color?.argb);
          if (fg) s.color = fg;
        }

        // Fill
        if (cell.fill && cell.fill.type === 'pattern') {
          const bg = argbToHex(cell.fill.fgColor?.argb);
          if (bg) s.bg = bg;
        }

        // Alignment
        if (cell.alignment) {
          if (cell.alignment.horizontal) {
            const h = cell.alignment.horizontal;
            s.align = h === 'centerContinuous' ? 'center' : h;
          }
          if (cell.alignment.wrapText) s.wrap = true;
          if (cell.alignment.vertical && !s.align) {
            // biarkan default, hanya info
          }
        }

        // Border
        if (cell.border) {
          const b = {};
          let hasBorder = false;
          let style = null;
          let color = null;

          ['top', 'bottom', 'left', 'right'].forEach(side => {
            const bd = cell.border[side];
            if (bd && bd.style) {
              b[side] = true;
              hasBorder = true;
              if (!style) {
                style = mapBorderStyle(bd.style);
                const c = argbToHex(bd.color?.argb);
                color = c || '#000000';
              }
            }
          });

          if (hasBorder) {
            b.style = style || 'thin';
            b.color = color || '#000000';
            s.borders = b;
          }
        }

        // Number format
        if (cell.numFmt) {
          const nf = mapNumFmt(cell.numFmt);
          if (nf) s.numFmt = nf;
        }

        if (Object.keys(s).length) obj.s = s;

        cells[addr] = obj;
      });
    });

    /* ---------- MERGES ---------- */
    // ExcelJS menyimpan merges di ws.model.merges — format "A1:B2"
    const mergesModel = ws.model?.merges || [];
    mergesModel.forEach(m => {
      if (typeof m === 'string' && m.includes(':')) merged.push(m);
    });

    /* ---------- COLUMN WIDTHS ---------- */
    if (ws.columns && Array.isArray(ws.columns)) {
      ws.columns.forEach((col, i) => {
        if (col.width) {
          // Excel width unit: 1 unit ≈ 7 px + 5 base
          colWidths[i] = Math.max(30, Math.round(col.width * 7.5));
        }
      });
    }

    /* ---------- ROW HEIGHTS ---------- */
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
      if (row.height && row.height !== 15) { // 15 = default Excel
        rowHeights[rowNum - 1] = Math.round(row.height * 1.33);
      }
    });

    /* ---------- FREEZE PANES ---------- */
    let freezeRow = 0, freezeCol = 0;
    const view = ws.views?.[0];
    if (view && view.state === 'frozen') {
      freezeRow = view.ySplit || 0;
      freezeCol = view.xSplit || 0;
    }

    /* ---------- DATA VALIDATION (list only) ---------- */
    // ExcelJS dataValidations di ws.dataValidations (Map)
    try {
      if (ws.dataValidations && ws.dataValidations.model) {
        Object.entries(ws.dataValidations.model).forEach(([addr, dv]) => {
          if (dv.type === 'list' && dv.formulae) {
            const opts = dv.formulae[0]
              ? String(dv.formulae[0]).replace(/^"|"$/g, '').split(',')
              : [];
            if (opts.length) {
              validations.push({
                range: addr,
                type: 'list',
                options: opts,
                errorMsg: dv.error || 'Nilai tidak valid'
              });
            }
          }
        });
      }
    } catch (e) {
      // skip if ExcelJS API beda
    }

    sheets.push({
      name: ws.name,
      cells,
      colWidths,
      rowHeights,
      rowHeight: 26,
      merged,
      freezeRow,
      freezeCol,
      cf,
      charts: [],
      filters: {},
      images: [],
      validations,
      protected: false
    });
  });

  return sheets;
}

/* =========================================================
   Export: generate XLSX dari wbData (opsional)
   ========================================================= */
export async function buildExcelBuffer(wbData, hfValues = null) {
  const wb = new ExcelJS.Workbook();

  for (const s of wbData.sheets || []) {
    const ws = wb.addWorksheet(s.name || 'Sheet');

    // Column widths
    if (s.colWidths) {
      Object.entries(s.colWidths).forEach(([i, w]) => {
        ws.getColumn(+i + 1).width = Math.max(8, (w - 5) / 7.5);
      });
    }

    // Cells
    const cells = s.cells || {};
    Object.entries(cells).forEach(([addr, cl]) => {
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) return;

      let v = cl.v;
      if (cl.f && hfValues) {
        // Pakai nilai HF-evaluated kalau tersedia
        v = hfValues[addr] ?? v ?? '';
      }

      const cell = ws.getCell(addr);
      cell.value = v ?? '';
    });

    // Merges
    (s.merged || []).forEach(range => {
      try { ws.mergeCells(range); } catch {}
    });
  }

  return await wb.xlsx.writeBuffer();
}