import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import 'dotenv/config';
import compression from 'compression';
import { initDb, pool } from './db.js';
import authRouter from './routes/auth.js';
import sheetsRouter from './routes/sheets.js';
import ioRouter from './routes/io.js';
import sharesRouter from './routes/shares.js';
import revisionsRouter from './routes/revisions.js';
import templatesRouter from './routes/templates.js';
import apiKeysRouter from './routes/api-keys.js';
import apiV1Router from './routes/api-v1.js';
import viewLinksRouter from './routes/view-links.js';
import publicViewRouter from './routes/public-view.js';
import chunksRouter from './routes/chunks.js';
import importUrlRouter from './routes/import-url.js';
import passwordRouter from './routes/password.js';
import { setupWebSocket } from './collab.js';
import sharedFolderRouter from './routes/shared-folder.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();

/* =========================================
   1. COMPRESSION — WAJIB DI PALING ATAS
   ========================================= */
app.use(compression({
  level: 1,
  threshold: 1024,
  filter: (req, res) => {
    if (req.headers['x-no-compression']) return false;
    if (req.path.includes('/export/xlsx')) return false;
    if (req.path.includes('/export/pdf')) return false;
    return compression.filter(req, res);
  }
}));

/* =========================================
   2. BODY PARSERS
   ========================================= */
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

/* =========================================
   3. STATIC FILES
   ========================================= */
app.use(express.static(path.join(__dirname, '..', 'public')));

/* =========================================
   4. API ROUTES
   ========================================= */
app.use('/api/auth', authRouter);
app.use('/api/sheets', sheetsRouter);
app.use('/api/sheets', ioRouter);
app.use('/api/sheets', sharesRouter);
app.use('/api/sheets', revisionsRouter);
app.use('/api/templates', templatesRouter);
app.use('/api/keys', apiKeysRouter);
app.use('/api/v1', apiV1Router);
app.use('/api/sheets', viewLinksRouter);
app.use('/api/sheets', chunksRouter);
app.use('/api/sheets', importUrlRouter);
app.use('/api/sheets', passwordRouter);
app.use('/api/public', publicViewRouter);
app.use('/api/shared-folder', sharedFolderRouter);

/* =========================================
   5. ERROR HANDLER — cegah server mati
   ========================================= */
app.use((err, req, res, next) => {
  console.error('❌ Server error:', err.message);

  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({ error: 'File terlalu besar' });
  }
  if (err.code === 'LIMIT_UNEXPECTED_FILE') {
    return res.status(400).json({ error: 'Field file tidak valid' });
  }
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({ error: err.message || 'Internal error' });
});

/* =========================================
   6. GLOBAL PROCESS HANDLERS
   ========================================= */
process.on('unhandledRejection', (reason, promise) => {
  console.error('⚠️ Unhandled rejection at:', promise);
  console.error('   Reason:', reason);
});

process.on('uncaughtException', (err) => {
  console.error('⚠️ Uncaught exception:', err);
  console.error('   Stack:', err.stack);
});

pool.on('error', (err) => {
  console.error('⚠️ PG pool error:', err.message);
});

/* =========================================
   7. START SERVER
   ========================================= */
const PORT = process.env.PORT || 3000;

initDb()
  .then(() => {
    const server = app.listen(PORT, () => {
      console.log(`🚀 Server ready: http://localhost:${PORT}`);
    });
    server.timeout = 10 * 60 * 1000;         // 10 menit
    server.keepAliveTimeout = 60 * 1000;     // 60 detik
    setupWebSocket(server);
  })
  .catch(e => {
    console.error('❌ DB init failed:', e);
    process.exit(1);
  });