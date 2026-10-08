/* =========================================================
   COLLABORATION CLIENT
   ========================================================= */
(function () {
  let socket = null;
  let myUserId = null;
  let myColor = null;
  let currentSheetId = null;
  let reconnectTimer = null;
  let pingTimer = null;
  let isConnected = false;

  // Cache user lain: Map<userId, {username, color, cursor}>
  const remoteUsers = new Map();

  function connect(sheetId, token) {
    if (socket && currentSheetId === sheetId) return;

    disconnect();

    currentSheetId = sheetId;
    const protocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
    const url = `${protocol}//${location.host}/ws?token=${encodeURIComponent(token)}&sheet=${sheetId}`;

    try {
      socket = new WebSocket(url);
    } catch (e) {
      console.error('WS connect failed:', e);
      return;
    }

    socket.addEventListener('open', () => {
      console.log('🔌 Collaboration connected');
      isConnected = true;
      updatePresenceUI();
      startPing();
    });

    socket.addEventListener('message', (e) => {
      try {
        const msg = JSON.parse(e.data);
        handleMessage(msg);
      } catch (err) {
        console.error('WS msg error:', err);
      }
    });

    socket.addEventListener('close', () => {
      isConnected = false;
      updatePresenceUI();
      stopPing();
      // Auto reconnect 3 detik
      if (currentSheetId) {
        reconnectTimer = setTimeout(() => {
          reconnectTimer = null;
          if (currentSheetId === sheetId) {
            connect(sheetId, token);
          }
        }, 3000);
      }
    });

    socket.addEventListener('error', (e) => {
      console.warn('WS error');
    });
  }

  function disconnect() {
    if (reconnectTimer) { clearTimeout(reconnectTimer); reconnectTimer = null; }
    stopPing();
    if (socket) {
      try { socket.close(); } catch {}
      socket = null;
    }
    isConnected = false;
    remoteUsers.clear();
    updatePresenceUI();
    removeAllCursors();
  }

  function handleMessage(msg) {
    switch (msg.type) {
      case 'welcome':
        myUserId = msg.userId;
        myColor = msg.color;
        break;

      case 'edit':
        // Apply remote edit ke local state
        applyRemoteEdit(msg);
        break;

      case 'cursor':
        updateRemoteCursor(msg.from, msg.username, msg.color, msg.cursor);
        break;

      case 'presence':
        updatePresenceList(msg.users);
        break;

      case 'pong':
        // heartbeat ok
        break;
    }
  }

  function applyRemoteEdit(msg) {
    const { sheetIdx, cells, from, username } = msg;
    if (typeof wbData === 'undefined' || !wbData?.sheets?.[sheetIdx]) return;

    const targetSheet = wbData.sheets[sheetIdx];
    if (!targetSheet.cells) targetSheet.cells = {};

    for (const [addr, cell] of Object.entries(cells || {})) {
      if (cell === null) {
        delete targetSheet.cells[addr];
      } else {
        targetSheet.cells[addr] = cell;
      }
      // Update HF
      if (sheetIdx === activeSheetIdx && typeof hf !== 'undefined' && hf) {
        const p = parseAddr(addr);
        if (p) {
          try {
            hf.setCellContents({ sheet: 0, row: p.r, col: p.c },
              [[cell?.f || cell?.v || null]]);
          } catch {}
        }
      }
    }

    // Re-render kalau sheet aktif
    if (sheetIdx === activeSheetIdx) {
      if (typeof renderVisibleCells === 'function') renderVisibleCells();
      if (typeof applySelectionStyles === 'function') applySelectionStyles();
    }

    // Toast kecil
    if (typeof toast === 'function' && username) {
      // Jangan spam — hanya log
      console.log(`📝 ${username} mengedit ${Object.keys(cells).length} sel`);
    }
  }

  function updatePresenceList(users) {
    const oldIds = new Set(remoteUsers.keys());
    remoteUsers.clear();

    users.forEach(u => {
      if (u.userId === myUserId) return;   // skip diri sendiri
      remoteUsers.set(u.userId, u);
      oldIds.delete(u.userId);
    });

    // Hapus cursor user yang sudah pergi
    oldIds.forEach(uid => removeCursor(uid));

    updatePresenceUI();
  }

  function updatePresenceUI() {
    let container = document.getElementById('presence-bar');
    if (!container) {
      container = document.createElement('div');
      container.id = 'presence-bar';
      container.className = 'presence-bar';
      const titlebar = document.querySelector('.titlebar');
      if (titlebar) {
        // Sisipkan sebelum user-chip
        const spacer = titlebar.querySelector('.spacer');
        if (spacer && spacer.nextSibling) {
          titlebar.insertBefore(container, spacer.nextSibling);
        } else {
          titlebar.appendChild(container);
        }
      }
    }

    if (!isConnected) {
      container.innerHTML = '<span class="presence-offline" title="Tidak terkoneksi"><i class="fa-solid fa-plug-circle-xmark"></i></span>';
      return;
    }

    let html = '<span class="presence-live" title="Terhubung real-time"><i class="fa-solid fa-circle"></i></span>';

    if (remoteUsers.size === 0) {
      html += '<span class="presence-solo">Sendiri</span>';
    } else {
      const maxShow = 4;
      let shown = 0;
      remoteUsers.forEach(u => {
        if (shown >= maxShow) return;
        html += `<span class="presence-avatar" style="background:${u.color}" title="${escapeHtml(u.username)}">
          ${escapeHtml(u.username.charAt(0).toUpperCase())}
        </span>`;
        shown++;
      });
      if (remoteUsers.size > maxShow) {
        html += `<span class="presence-more">+${remoteUsers.size - maxShow}</span>`;
      }
    }

    container.innerHTML = html;
  }

  function updateRemoteCursor(userId, username, color, cursor) {
    if (!cursor) { removeCursor(userId); return; }

    let el = document.querySelector(`.remote-cursor[data-uid="${userId}"]`);
    if (!el) {
      el = document.createElement('div');
      el.className = 'remote-cursor';
      el.dataset.uid = userId;
      el.innerHTML = `<div class="rc-dot"></div><div class="rc-label"></div>`;
      document.body.appendChild(el);
    }

    const label = el.querySelector('.rc-label');
    label.textContent = username;
    label.style.background = color;
    el.querySelector('.rc-dot').style.background = color;

    // Posisi: cari td di DOM
    const td = document.querySelector(`#grid td[data-addr="${cursor}"]`);
    if (!td) { el.style.display = 'none'; return; }

    const rect = td.getBoundingClientRect();
    el.style.display = '';
    el.style.left = rect.left + 'px';
    el.style.top = rect.top + 'px';
    el.style.width = rect.width + 'px';
    el.style.height = rect.height + 'px';
    el.style.borderColor = color;
  }

  function removeCursor(userId) {
    const el = document.querySelector(`.remote-cursor[data-uid="${userId}"]`);
    if (el) el.remove();
  }

  function removeAllCursors() {
    document.querySelectorAll('.remote-cursor').forEach(el => el.remove());
  }

  function sendEdit(cells, sheetIdx) {
    if (!socket || socket.readyState !== 1) return;
    try {
      socket.send(JSON.stringify({
        type: 'edit',
        sheetIdx: sheetIdx ?? activeSheetIdx,
        cells
      }));
    } catch (e) {
      console.error('Send edit failed:', e);
    }
  }

  function sendCursor(addr) {
    if (!socket || socket.readyState !== 1) return;
    try {
      socket.send(JSON.stringify({
        type: 'cursor',
        cursor: addr
      }));
    } catch {}
  }

  function startPing() {
    stopPing();
    pingTimer = setInterval(() => {
      if (socket && socket.readyState === 1) {
        try { socket.send(JSON.stringify({ type: 'ping' })); } catch {}
      }
    }, 25000);
  }

  function stopPing() {
    if (pingTimer) { clearInterval(pingTimer); pingTimer = null; }
  }

  // Update cursor position tiap 500ms kalau berubah
  let lastSentCursor = null;
  setInterval(() => {
    if (!isConnected) return;
    if (typeof activeCell === 'undefined') return;
    if (activeCell !== lastSentCursor) {
      lastSentCursor = activeCell;
      sendCursor(activeCell);
    }
  }, 500);

  // Expose ke global
  window.Collab = {
    connect,
    disconnect,
    sendEdit,
    isConnected: () => isConnected,
    getUserCount: () => remoteUsers.size + (isConnected ? 1 : 0)
  };
})();