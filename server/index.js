import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import 'dotenv/config';
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
import compression from 'compression';
import { setupWebSocket } from './collab.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();

app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, '..', 'public')));

app.use('/api/auth', authRouter);
app.use('/api/sheets', sheetsRouter);
app.use('/api/sheets', ioRouter);
app.use('/api/sheets', sharesRouter);
app.use('/api/sheets', revisionsRouter);
app.use('/api/templates', templatesRouter);
app.use('/api/keys', apiKeysRouter);
app.use('/api/v1', apiV1Router);
app.use('/api/sheets', viewLinksRouter);
app.use('/api/public', publicViewRouter);

/* =========================================
   ERROR HANDLER — WAJIB! cegah server mati
   ========================================= */
app.use((err, req, res, next) => {
  console.error('API error:', req.method, req.url);
  console.error('   ', err.message);
  if (err.stack) console.error(err.stack.split('\n')[1]);
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({
    error: err.message || 'Internal server error'
  });
});
app.use(compression({
  level: 6,
  threshold: 1024,
  filter: (req, res) => {
    if (req.headers['x-no-compression']) return false;
    return compression.filter(req, res);
  }
}));

/* Cegah process mati karena unhandled rejection */
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled rejection at:', promise);
  console.error('   Reason:', reason);
});

process.on('uncaughtException', (err) => {
  console.error('Uncaught exception:', err);
  console.error('   Stack:', err.stack);
});


pool.on('error', (err) => {
  console.error('PG pool error:', err.message);
});

const PORT = process.env.PORT || 3000;

initDb()
  .then(() => {
    const server = app.listen(PORT, () => console.log(`http://localhost:${PORT}`));
    setupWebSocket(server);
  })
  .catch(e => { console.error('DB init failed:', e); process.exit(1); });