import jwt from 'jsonwebtoken';
const SECRET = process.env.JWT_SECRET || 'change-me-dev-secret';

export const signToken = u =>
  jwt.sign({ id: u.id, username: u.username }, SECRET, { expiresIn: '7d' });

export function authRequired(req, res, next) {
  const h = req.headers.authorization || '';
  const token = h.startsWith('Bearer ') ? h.slice(7) : null;
  if (!token) return res.status(401).json({ error: 'Unauthorized' });
  try { req.user = jwt.verify(token, SECRET); next(); }
  catch { res.status(401).json({ error: 'Invalid token' }); }
}