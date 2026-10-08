import { HyperFormula } from 'hyperformula';

export function a1ToRC(addr) {
  const m = /^([A-Z]+)(\d+)$/.exec(addr);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: +m[2] - 1, c: c - 1 };
}
export function rcToA1(r, c) {
  let s = '', n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s + (r + 1);
}
export function usedRange(cells) {
  let maxR = -1, maxC = -1;
  for (const a of Object.keys(cells || {})) {
    const p = a1ToRC(a); if (!p) continue;
    if (p.r > maxR) maxR = p.r;
    if (p.c > maxC) maxC = p.c;
  }
  return { rows: maxR + 1, cols: maxC + 1 };
}

/** sheets = [{name, cells}, ...]; aktif = index */
export function computeSheet(cells) {
  const { rows: R, cols: C } = usedRange(cells);
  if (!R || !C) return { values: [], rows: 0, cols: 0 };
  const matrix = Array.from({ length: R }, () => new Array(C).fill(null));
  for (const [a, cell] of Object.entries(cells || {})) {
    const p = a1ToRC(a); if (!p || p.r >= R || p.c >= C) continue;
    if (cell.f) matrix[p.r][p.c] = cell.f;
    else if (cell.v !== undefined && cell.v !== null && cell.v !== '')
      matrix[p.r][p.c] = cell.v;
  }
  let hf;
  try { hf = HyperFormula.buildFromArray(matrix, { licenseKey: 'gpl-v3' }); }
  catch { return { values: matrix.map(r => r.map(v => v ?? '')), rows: R, cols: C }; }

  const values = [];
  for (let r = 0; r < R; r++) {
    const row = [];
    for (let c = 0; c < C; c++)
      row.push(norm(hf.getCellValue({ sheet: 0, row: r, col: c })));
    values.push(row);
  }
  hf.destroy();
  return { values, rows: R, cols: C };
}
function norm(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return 'value' in v ? String(v.value) : '';
  return v;
}