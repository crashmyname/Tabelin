import { Router } from 'express';
import bcrypt from 'bcryptjs';
import { pool } from '../db.js';
import { signToken } from '../auth.js';

const router = Router();

router.post('/register', async (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password || password.length < 4)
    return res.status(400).json({ error: 'Username & password (min 4) required' });
  const hash = await bcrypt.hash(password, 10);
  try {
    const { rows } = await pool.query(
      'INSERT INTO users (username, password_hash) VALUES ($1,$2) RETURNING id, username',
      [username, hash]
    );
    res.status(201).json({ token: signToken(rows[0]), user: rows[0] });
  } catch (e) {
    if (e.code === '23505') return res.status(409).json({ error: 'Username sudah dipakai' });
    res.status(500).json({ error: e.message });
  }
});

router.post('/login', async (req, res) => {
  const { username, password } = req.body || {};
  const { rows } = await pool.query('SELECT * FROM users WHERE username=$1', [username]);
  if (!rows.length) return res.status(401).json({ error: 'Username/password salah' });
  const ok = await bcrypt.compare(password, rows[0].password_hash);
  if (!ok) return res.status(401).json({ error: 'Username/password salah' });
  res.json({ token: signToken(rows[0]), user: { id: rows[0].id, username: rows[0].username } });
});

export default router;