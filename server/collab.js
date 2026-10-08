import { WebSocketServer } from 'ws';
import jwt from 'jsonwebtoken';
import { pool } from './db.js';

const JWT_SECRET = process.env.JWT_SECRET || 'change-me-dev-secret';
const COLORS = ['#ef4444','#f59e0b','#10b981','#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f97316'];

const rooms = new Map();       // Map<sheetId, Set<ws>>
const clients = new Map();     // Map<ws, clientInfo>

export function setupWebSocket(server) {
  const wss = new WebSocketServer({ server, path: '/ws' });

  wss.on('connection', async (ws, req) => {
    const url = new URL(req.url, 'http://x');
    const token = url.searchParams.get('token');
    const sheetId = url.searchParams.get('sheet');

    if (!sheetId) {
      ws.close(4002, 'Missing sheet');
      return;
    }

    const client = {
      ws,
      userId: null,
      username: null,
      sheetId: Number(sheetId),
      role: 'viewer',
      color: COLORS[Math.floor(Math.random() * COLORS.length)],
      cursor: null,
      lastSeen: Date.now()
    };

    // Deteksi tipe token: JWT punya 2 titik (xxx.yyy.zzz)
    const isJWT = token && token.split('.').length === 3;

    if (isJWT) {
      try {
        const user = jwt.verify(token, JWT_SECRET);
        client.userId = user.id;
        client.username = user.username;
        client.role = 'editor';
      } catch {
        ws.close(4001, 'Invalid token');
        return;
      }
    } else if (token) {
      // Cek view token di DB
      try {
        const { rows } = await pool.query(
          `SELECT vl.id AS link_id, s.id AS sheet_id
             FROM view_links vl
             JOIN sheets s ON s.id = vl.sheet_id
            WHERE vl.token = $1::text`,
          [token]
        );
        if (!rows.length) {
          ws.close(4001, 'Invalid view token');
          return;
        }
        if (rows[0].sheet_id !== Number(sheetId)) {
          ws.close(4003, 'Token does not match sheet');
          return;
        }
        client.userId = 'viewer_' + rows[0].link_id;
        client.username = 'Viewer';
        client.role = 'viewer';
      } catch (e) {
        console.error('View token check:', e.message);
        ws.close(4001, 'Token error');
        return;
      }
    } else {
      ws.close(4001, 'No token');
      return;
    }

    clients.set(ws, client);

    if (!rooms.has(client.sheetId)) rooms.set(client.sheetId, new Set());
    rooms.get(client.sheetId).add(ws);

    ws.send(JSON.stringify({
      type: 'welcome',
      userId: client.userId,
      username: client.username,
      color: client.color,
      role: client.role
    }));

    broadcastPresence(client.sheetId);
    console.log(`🔌 WS connect: ${client.username} (${client.role}) → sheet ${client.sheetId}`);

    ws.on('message', (data) => {
      try {
        const msg = JSON.parse(data);
        client.lastSeen = Date.now();
        handleMessage(client, msg);
      } catch (e) {
        console.error('WS msg error:', e.message);
      }
    });

    ws.on('close', () => {
      clients.delete(ws);
      const room = rooms.get(client.sheetId);
      if (room) {
        room.delete(ws);
        if (!room.size) rooms.delete(client.sheetId);
      }
      broadcastPresence(client.sheetId);
    });

    ws.on('error', (e) => console.error('WS error:', e.message));
  });

  /* Cleanup stale connections */
  setInterval(() => {
    const now = Date.now();
    clients.forEach((c, ws) => {
      if (now - c.lastSeen > 60000) {
        try { ws.close(); } catch {}
      }
    });
  }, 30000);

  console.log('✅ WebSocket ready (editor + viewer)');
}

function handleMessage(client, msg) {
  switch (msg.type) {
    case 'edit':
      // Hanya editor yang boleh broadcast edit
      if (client.role !== 'editor') return;
      broadcastToRoom(client.sheetId, {
        type: 'edit',
        from: client.userId,
        username: client.username,
        sheetIdx: msg.sheetIdx,
        cells: msg.cells
      }, client.ws);
      break;

    case 'cursor':
      if (client.role !== 'editor') return;
      client.cursor = msg.cursor;
      broadcastToRoom(client.sheetId, {
        type: 'cursor',
        from: client.userId,
        username: client.username,
        color: client.color,
        cursor: msg.cursor
      }, client.ws);
      break;

    case 'ping':
      if (client.ws.readyState === 1) {
        client.ws.send(JSON.stringify({ type: 'pong' }));
      }
      break;
  }
}

function broadcastToRoom(sheetId, message, exclude = null) {
  const room = rooms.get(sheetId);
  if (!room) return;
  const data = JSON.stringify(message);
  room.forEach(ws => {
    if (ws !== exclude && ws.readyState === 1) {
      try { ws.send(data); } catch {}
    }
  });
}

function broadcastPresence(sheetId) {
  const room = rooms.get(sheetId);
  if (!room) return;
  const list = [];
  room.forEach(ws => {
    const c = clients.get(ws);
    // Hanya tampilkan editor di presence (viewer tidak muncul sebagai avatar)
    if (c && c.role === 'editor') {
      list.push({
        userId: c.userId,
        username: c.username,
        color: c.color,
        cursor: c.cursor
      });
    }
  });
  const data = JSON.stringify({ type: 'presence', users: list });
  room.forEach(ws => {
    if (ws.readyState === 1) {
      try { ws.send(data); } catch {}
    }
  });
}