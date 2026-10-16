import pg from 'pg';
import 'dotenv/config';

export const pool = new pg.Pool({
  connectionString: process.env.DATABASE_URL ||
    'postgres://postgres:postgres@localhost:5432/webexcel',
  max: 20,
  idleTimeoutMillis: 30000,
  statement_timeout: 60000,
  query_timeout: 60000,
  connectionTimeoutMillis: 10000
});

export async function initDb() {
  await pool.query(`
    /* ---------- USERS ---------- */
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );

    /* ---------- SHEETS ---------- */
    CREATE TABLE IF NOT EXISTS sheets (
      id SERIAL PRIMARY KEY,
      user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
      name TEXT NOT NULL,
      data JSONB NOT NULL DEFAULT '{}'::jsonb,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
    CREATE INDEX IF NOT EXISTS idx_sheets_user ON sheets(user_id, updated_at DESC);

    /* ---------- SHARES ---------- */
    CREATE TABLE IF NOT EXISTS shares (
      id SERIAL PRIMARY KEY,
      sheet_id INTEGER NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
      user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      permission TEXT NOT NULL DEFAULT 'view'
        CHECK (permission IN ('view','edit','admin')),
      shared_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      UNIQUE (sheet_id, user_id)
    );
    CREATE INDEX IF NOT EXISTS idx_shares_user ON shares(user_id);
    CREATE INDEX IF NOT EXISTS idx_shares_sheet ON shares(sheet_id);

    /* ---------- REVISIONS ---------- */
    CREATE TABLE IF NOT EXISTS revisions (
      id SERIAL PRIMARY KEY,
      sheet_id INTEGER NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
      data JSONB NOT NULL,
      name TEXT,
      label TEXT,
      created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      size_bytes INTEGER
    );
    CREATE INDEX IF NOT EXISTS idx_rev_sheet ON revisions(sheet_id, created_at DESC);

    /* ---------- TEMPLATES ---------- */
    CREATE TABLE IF NOT EXISTS templates (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      description TEXT,
      category TEXT DEFAULT 'general',
      data JSONB NOT NULL,
      is_public BOOLEAN NOT NULL DEFAULT false,
      created_by INTEGER REFERENCES users(id) ON DELETE CASCADE,
      used_count INTEGER NOT NULL DEFAULT 0,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
     CREATE INDEX IF NOT EXISTS idx_tpl_cat ON templates(category, is_public);

    /* ---------- API KEYS ---------- */
    CREATE TABLE IF NOT EXISTS api_keys (
      id SERIAL PRIMARY KEY,
      user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      sheet_id INTEGER REFERENCES sheets(id) ON DELETE CASCADE,
      key TEXT UNIQUE NOT NULL,
      label TEXT,
      permission TEXT NOT NULL DEFAULT 'read'
        CHECK (permission IN ('read')),
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      last_used_at TIMESTAMPTZ,
      request_count INTEGER NOT NULL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS idx_api_keys_key   ON api_keys(key);
    CREATE INDEX IF NOT EXISTS idx_api_keys_user  ON api_keys(user_id);
    CREATE TABLE IF NOT EXISTS view_links (
      id SERIAL PRIMARY KEY,
      sheet_id INTEGER NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
      token TEXT UNIQUE NOT NULL,
      label TEXT,
      created_by INTEGER REFERENCES users(id) ON DELETE SET NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
      expires_at TIMESTAMPTZ,
      view_count INTEGER NOT NULL DEFAULT 0
    );
    CREATE INDEX IF NOT EXISTS idx_view_links_token ON view_links(token);
    ALTER TABLE sheets ADD COLUMN IF NOT EXISTS password_hash TEXT;
  `);

  await seedTemplates();
  console.log('DB ready');
}

/* ---------- Seed template bawaan (hanya sekali) ---------- */
async function seedTemplates() {
  const { rows } = await pool.query('SELECT COUNT(*)::int AS n FROM templates WHERE is_public = true');
  if (rows[0].n > 0) return;

  const P = 'public'; // is_public
  const templates = [
    {
      name: 'Invoice Sederhana',
      description: 'Invoice dengan item, qty, harga, subtotal, PPN, dan total',
      category: 'Keuangan',
      data: {
        activeSheet: 0,
        sheets: [{
          name: 'Invoice',
          cells: {
            A1: { v: 'INVOICE',        s: { bold: true, fontSize: 20, color: '#059669' } },
            A3: { v: 'No. Invoice:' }, B3: { v: 'INV-001' },
            A4: { v: 'Tanggal:' },     B4: { v: '=TODAY()', s: { numFmt: 'date' } },
            A5: { v: 'Pelanggan:' },   B5: { v: 'PT Contoh Jaya' },
            A7: { v: 'Deskripsi',  s: { bold: true, bg: '#d1fae5' } },
            B7: { v: 'Qty',        s: { bold: true, bg: '#d1fae5', align: 'right' } },
            C7: { v: 'Harga',      s: { bold: true, bg: '#d1fae5', align: 'right' } },
            D7: { v: 'Subtotal',   s: { bold: true, bg: '#d1fae5', align: 'right' } },
            A8: { v: 'Produk A' }, B8: { v: 2 },  C8: { v: 500000, s: { numFmt: 'currency' } }, D8: { f: '=B8*C8', s: { numFmt: 'currency' } },
            A9: { v: 'Produk B' }, B9: { v: 1 },  C9: { v: 750000, s: { numFmt: 'currency' } }, D9: { f: '=B9*C9', s: { numFmt: 'currency' } },
            A10:{ v: 'Produk C' }, B10:{ v: 5 },  C10:{ v: 100000, s: { numFmt: 'currency' } }, D10:{ f: '=B10*C10', s: { numFmt: 'currency' } },
            C12:{ v: 'Subtotal', s: { bold: true } }, D12:{ f: '=SUM(D8:D10)', s: { numFmt: 'currency', bold: true } },
            C13:{ v: 'PPN 11%' },                     D13:{ f: '=D12*0.11',    s: { numFmt: 'currency' } },
            C14:{ v: 'TOTAL',    s: { bold: true, bg: '#d1fae5' } },
            D14:{ f: '=D12+D13', s: { numFmt: 'currency', bold: true, bg: '#d1fae5' } }
          },
          colWidths: { 0: 200, 1: 60, 2: 120, 3: 130 },
          cf: [], charts: []
        }]
      }
    },
    {
      name: 'Budget Bulanan',
      description: 'Rencana anggaran bulanan dengan kategori dan selisih',
      category: 'Keuangan',
      data: {
        activeSheet: 0,
        sheets: [{
          name: 'Budget',
          cells: {
            A1: { v: 'BUDGET BULANAN', s: { bold: true, fontSize: 18, color: '#059669' } },
            A3: { v: 'Kategori',   s: { bold: true, bg: '#dbeafe' } },
            B3: { v: 'Rencana',    s: { bold: true, bg: '#dbeafe' } },
            C3: { v: 'Realisasi',  s: { bold: true, bg: '#dbeafe' } },
            D3: { v: 'Selisih',    s: { bold: true, bg: '#dbeafe' } },
            A4: { v: 'Makan' },    B4: { v: 3000000, s: { numFmt: 'currency' } }, C4: { v: 2800000, s: { numFmt: 'currency' } }, D4: { f: '=B4-C4', s: { numFmt: 'currency' } },
            A5: { v: 'Transport' },B5: { v: 1500000, s: { numFmt: 'currency' } }, C5: { v: 1700000, s: { numFmt: 'currency' } }, D5: { f: '=B5-C5', s: { numFmt: 'currency' } },
            A6: { v: 'Hiburan' },  B6: { v: 800000,  s: { numFmt: 'currency' } }, C6: { v: 600000,  s: { numFmt: 'currency' } }, D6: { f: '=B6-C6', s: { numFmt: 'currency' } },
            A7: { v: 'Lainnya' },  B7: { v: 500000,  s: { numFmt: 'currency' } }, C7: { v: 700000,  s: { numFmt: 'currency' } }, D7: { f: '=B7-C7', s: { numFmt: 'currency' } },
            A9: { v: 'TOTAL', s: { bold: true } },
            B9: { f: '=SUM(B4:B7)', s: { numFmt: 'currency', bold: true } },
            C9: { f: '=SUM(C4:C7)', s: { numFmt: 'currency', bold: true } },
            D9: { f: '=SUM(D4:D7)', s: { numFmt: 'currency', bold: true } }
          },
          colWidths: { 0: 160, 1: 120, 2: 120, 3: 120 },
          cf: [{ range: 'D4:D7', type: 'lessThan', value: '0',
                 style: { bg: '#fee2e2', color: '#991b1b' } }],
          charts: []
        }]
      }
    },
    {
      name: 'Todo List',
      description: 'Daftar tugas dengan prioritas, deadline, dan status',
      category: 'Produktivitas',
      data: {
        activeSheet: 0,
        sheets: [{
          name: 'Todo',
          cells: {
            A1: { v: 'TODO LIST', s: { bold: true, fontSize: 18, color: '#059669' } },
            A3: { v: 'No',      s: { bold: true, bg: '#fef3c7' } },
            B3: { v: 'Tugas',   s: { bold: true, bg: '#fef3c7' } },
            C3: { v: 'Prioritas', s: { bold: true, bg: '#fef3c7' } },
            D3: { v: 'Deadline',  s: { bold: true, bg: '#fef3c7' } },
            E3: { v: 'Status',    s: { bold: true, bg: '#fef3c7' } },
            A4: { v: 1 }, B4: { v: 'Rapat tim marketing' }, C4: { v: 'Tinggi' }, E4: { v: 'Belum' },
            A5: { v: 2 }, B5: { v: 'Review desain UI' },    C5: { v: 'Sedang' }, E5: { v: 'Proses' },
            A6: { v: 3 }, B6: { v: 'Update dokumentasi' },  C6: { v: 'Rendah' }, E6: { v: 'Belum' }
          },
          colWidths: { 0: 50, 1: 240, 2: 100, 3: 120, 4: 100 },
          cf: [{ range: 'E4:E100', type: 'equal', value: 'Selesai',
                 style: { bg: '#d1fae5', color: '#065f46' } }],
          charts: []
        }]
      }
    },
    {
      name: 'Tabel Perkalian',
      description: 'Tabel perkalian 1-10 dengan formula dinamis',
      category: 'Edukasi',
      data: {
        activeSheet: 0,
        sheets: [{
          name: 'Kali',
          cells: (() => {
            const c = { A1: { v: '×', s: { bold: true, bg: '#e0e7ff' } } };
            for (let i = 1; i <= 10; i++) {
              c[String.fromCharCode(64 + i + 1) + '1'] = { v: i, s: { bold: true, bg: '#e0e7ff', align: 'center' } };
              c['A' + (i + 1)] = { v: i, s: { bold: true, bg: '#e0e7ff', align: 'center' } };
              for (let j = 1; j <= 10; j++)
                c[String.fromCharCode(64 + j + 1) + (i + 1)] = { f: `=A${i+1}*${String.fromCharCode(64+j+1)}1`, s: { align: 'center' } };
            }
            return c;
          })(),
          colWidths: {}, cf: [], charts: []
        }]
      }
    },
    {
      name: 'Dashboard Penjualan',
      description: 'Ringkasan penjualan dengan chart bawaan',
      category: 'Bisnis',
      data: {
        activeSheet: 0,
        sheets: [{
          name: 'Dashboard',
          cells: {
            A1: { v: 'DASHBOARD PENJUALAN', s: { bold: true, fontSize: 18, color: '#059669' } },
            A3: { v: 'Bulan',   s: { bold: true, bg: '#dbeafe' } },
            B3: { v: 'Penjualan', s: { bold: true, bg: '#dbeafe' } },
            A4: { v: 'Jan' }, B4: { v: 45000000, s: { numFmt: 'currency' } },
            A5: { v: 'Feb' }, B5: { v: 52000000, s: { numFmt: 'currency' } },
            A6: { v: 'Mar' }, B6: { v: 48000000, s: { numFmt: 'currency' } },
            A7: { v: 'Apr' }, B7: { v: 61000000, s: { numFmt: 'currency' } },
            A8: { v: 'Mei' }, B8: { v: 55000000, s: { numFmt: 'currency' } },
            A10:{ v: 'Total',   s: { bold: true } }, B10:{ f: '=SUM(B4:B8)', s: { numFmt: 'currency', bold: true } },
            A11:{ v: 'Rata-rata', s: { bold: true } }, B11:{ f: '=AVERAGE(B4:B8)', s: { numFmt: 'currency' } },
            A12:{ v: 'Tertinggi', s: { bold: true } }, B12:{ f: '=MAX(B4:B8)', s: { numFmt: 'currency' } }
          },
          colWidths: { 0: 120, 1: 150 },
          cf: [],
          charts: [{ id: 'seed-1', type: 'bar', range: 'A3:B8', x: 340, y: 100, w: 420, h: 280 }]
        }]
      }
    }
  ];

  for (const t of templates) {
    await pool.query(
      `INSERT INTO templates (name, description, category, data, is_public)
       VALUES ($1, $2, $3, $4, true)`,
      [t.name, t.description, t.category, t.data]
    );
  }
  console.log(`📦 Seeded ${templates.length} templates`);
}