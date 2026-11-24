// public/folders.js  (GANTI ISI folders.js lama dengan ini)
// Dimuat SEBELUM app.js:  <script src="/folders.js"></script>
(() => {
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
  const notify = msg => { if (typeof toast === 'function') toast(msg); else alert(msg); };

  // `token` adalah variabel global di app.js (dipakai juga oleh renderHomeList)
  const getToken = () =>
    (typeof token !== 'undefined' && token) || localStorage.getItem('token') || '';

  async function api(url, method = 'GET', body) {
    const r = await fetch(url, {
      method,
      headers: {
        'Authorization': 'Bearer ' + getToken(),
        ...(body ? { 'Content-Type': 'application/json' } : {})
      },
      body: body ? JSON.stringify(body) : undefined
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(data.error || ('HTTP ' + r.status));
    return data;
  }

  const F = window.Folders = {
    list: [],          // [{id, name}]
    map: {},           // { sheetId: folderId }
    current: 'all',    // 'all' | 'none' | <folderId number>
    editing: null
  };

  /* ---------- DATA ---------- */
  F.load = async () => {
    const [list, map] = await Promise.all([api('/api/folders'), api('/api/folders/map')]);
    F.list = list;
    F.map = map;
    // kalau folder yang sedang dibuka sudah tidak ada, balik ke "Semua"
    if (typeof F.current === 'number' && !list.some(f => f.id === F.current)) F.current = 'all';
  };

  // tempelkan folder_id ke tiap workbook dari /api/sheets
  F.attach = sheets => {
    for (const s of sheets) s.folder_id = F.map[s.id] ?? null;
  };

  F.matches = s => {
    if (F.current === 'all') return true;
    if (F.current === 'none') return !s.folder_id;
    return Number(s.folder_id) === F.current;
  };

  /* ---------- SIDEBAR ---------- */
  F.renderTree = () => {
    const tree = $('folder-tree');
    if (!tree) return;
    const sheets = window.homeSheets || [];
    const count = id => sheets.filter(s =>
      id === 'none' ? !s.folder_id : Number(s.folder_id) === id
    ).length;

    const row = (id, icon, name, cnt, actions = '') => `
      <div class="folder-item ${F.current === id ? 'active' : ''}" data-fid="${id}">
        <i class="fa-solid ${icon}"></i>
        <span class="f-name" title="${esc(name)}">${esc(name)}</span>
        <span class="f-count">${cnt}</span>
        ${actions}
      </div>`;

    const acts = `<span class="f-actions">
      <button data-act="rename" title="Ganti nama"><i class="fa-solid fa-pen"></i></button>
      <button data-act="delete" title="Hapus folder"><i class="fa-solid fa-trash"></i></button>
    </span>`;

    tree.innerHTML =
      row('all', 'fa-layer-group', 'Semua Workbook', sheets.length) +
      row('none', 'fa-inbox', 'Tanpa Folder', count('none')) +
      F.list.map(f => row(f.id, 'fa-folder', f.name, count(f.id), acts)).join('');

    const cur = F.list.find(f => f.id === F.current);
    const bc = $('folder-breadcrumb');
    if (bc) {
      bc.innerHTML =
        `<a data-fid="all">Semua</a>` +
        (F.current === 'none' ? ' / <b>Tanpa Folder</b>' : '') +
        (cur ? ` / <b>${esc(cur.name)}</b>` : '');
    }
  };

  F.select = id => {
    F.current = (id === 'all' || id === 'none') ? id : Number(id);
    if (typeof drawHomeList === 'function') drawHomeList();   // render ulang tanpa fetch
  };

  /* ---------- PINDAHKAN WORKBOOK ---------- */
  F.move = async (sheetId, folderId) => {
    try {
      await api(`/api/sheets/${sheetId}/folder`, 'PUT', { folder_id: folderId });
      if (folderId) F.map[sheetId] = folderId; else delete F.map[sheetId];
      const s = (window.homeSheets || []).find(x => String(x.id) === String(sheetId));
      if (s) s.folder_id = folderId;
      const name = folderId ? F.list.find(f => f.id === folderId)?.name : 'Tanpa Folder';
      notify('✔ Dipindahkan ke ' + name);
      if (typeof drawHomeList === 'function') drawHomeList();
    } catch (e) {
      notify('Gagal memindahkan: ' + e.message);
    }
  };

  F.showMoveMenu = (sheetId, anchorEl) => {
    const m = $('move-menu');
    const current = (window.homeSheets || []).find(s => String(s.id) === String(sheetId))?.folder_id ?? null;

    m.innerHTML =
      `<button data-to="" ${!current ? 'disabled' : ''}><i class="fa-solid fa-inbox"></i> Tanpa Folder</button>` +
      (F.list.length
        ? F.list.map(f =>
            `<button data-to="${f.id}" ${Number(current) === f.id ? 'disabled' : ''}>
               <i class="fa-solid fa-folder"></i> ${esc(f.name)}</button>`).join('')
        : `<div class="move-empty">Belum ada folder.<br>Klik "Folder Baru" dulu.</div>`);

    const rect = anchorEl.getBoundingClientRect();
    m.classList.remove('hidden');
    const w = m.offsetWidth, h = m.offsetHeight;
    m.style.left = Math.max(8, Math.min(rect.left, innerWidth - w - 8)) + 'px';
    m.style.top = Math.max(8, Math.min(rect.bottom + 4, innerHeight - h - 8)) + 'px';

    m.onclick = e => {
      const b = e.target.closest('button[data-to]');
      if (!b || b.disabled) return;
      m.classList.add('hidden');
      F.move(sheetId, b.dataset.to ? Number(b.dataset.to) : null);
    };
  };

  /* ---------- MODAL BUAT / RENAME ---------- */
  F.openModal = (folder = null) => {
    F.editing = folder;
    $('folder-modal-title').textContent = folder ? 'Ganti Nama Folder' : 'Folder Baru';
    $('folder-name-input').value = folder ? folder.name : '';
    $('modal-folder').classList.remove('hidden');
    setTimeout(() => $('folder-name-input').focus(), 30);
  };
  F.closeModal = () => $('modal-folder').classList.add('hidden');

  F.save = async () => {
    const name = $('folder-name-input').value.trim();
    if (!name) return;
    try {
      if (F.editing) {
        await api(`/api/folders/${F.editing.id}`, 'PATCH', { name });
      } else {
        const created = await api('/api/folders', 'POST', { name });
        F.current = created.id;           // langsung buka folder yang baru dibuat
      }
      F.closeModal();
      await F.load();
      if (typeof drawHomeList === 'function') drawHomeList();
    } catch (e) {
      notify(e.message);
    }
  };

  /* ---------- INIT EVENT ---------- */
  function init() {
    $('btn-folder-new')?.addEventListener('click', () => F.openModal());
    $('btn-folder-save')?.addEventListener('click', F.save);
    $('btn-folder-cancel')?.addEventListener('click', F.closeModal);
    $('folder-name-input')?.addEventListener('keydown', e => {
      if (e.key === 'Enter') F.save();
      if (e.key === 'Escape') F.closeModal();
    });

    $('folder-breadcrumb')?.addEventListener('click', e => {
      const a = e.target.closest('[data-fid]');
      if (a) F.select(a.dataset.fid);
    });

    const tree = $('folder-tree');
    if (!tree) return;

    tree.addEventListener('click', async e => {
      const item = e.target.closest('.folder-item');
      if (!item) return;
      const act = e.target.closest('[data-act]')?.dataset.act;
      const folder = F.list.find(f => String(f.id) === item.dataset.fid);

      if (act === 'rename' && folder) return F.openModal(folder);

      if (act === 'delete' && folder) {
        const ok = typeof customConfirm === 'function'
          ? await customConfirm(`Hapus folder "${folder.name}"? Workbook di dalamnya tidak ikut terhapus.`,
                                { type: 'danger', okText: 'Hapus' })
          : confirm(`Hapus folder "${folder.name}"?`);
        if (!ok) return;
        try {
          await api(`/api/folders/${folder.id}`, 'DELETE');
          if (F.current === folder.id) F.current = 'all';
          await F.load();
          // workbook di folder tsb otomatis jadi "tanpa folder"
          (window.homeSheets || []).forEach(s => { s.folder_id = F.map[s.id] ?? null; });
          if (typeof drawHomeList === 'function') drawHomeList();
        } catch (err) { notify(err.message); }
        return;
      }
      F.select(item.dataset.fid);
    });

    // drop workbook ke folder di sidebar
    tree.addEventListener('dragover', e => {
      const it = e.target.closest('.folder-item');
      if (it && it.dataset.fid !== 'all') { e.preventDefault(); it.classList.add('drop-over'); }
    });
    tree.addEventListener('dragleave', e => {
      e.target.closest('.folder-item')?.classList.remove('drop-over');
    });
    tree.addEventListener('drop', e => {
      const it = e.target.closest('.folder-item');
      if (!it || it.dataset.fid === 'all') return;
      e.preventDefault();
      it.classList.remove('drop-over');
      const sid = e.dataTransfer.getData('text/sheet-id');
      if (sid) F.move(sid, it.dataset.fid === 'none' ? null : Number(it.dataset.fid));
    });

    // tutup menu "Pindahkan" kalau klik di luar
    document.addEventListener('click', e => {
      if (!e.target.closest('#move-menu, [data-move]')) $('move-menu')?.classList.add('hidden');
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();