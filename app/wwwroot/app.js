// Mail Monitor — front controller
const missingApiWarning = '[Mail Monitor] window.api indisponible : exécution hors WebView2 ou preload non injecté.';
const fallbackApi = new Proxy({
  version: async () => '?',
  outlookStatus: async () => ({ connected: false }),
  monitoringStatus: async () => ({ running: false }),
  autostartGet: async () => false,
  statsSummary: async () => ({ folders: 0, emails: 0, unread: 0, last7days: 0 }),
}, {
  get(target, prop) {
    if (prop in target) return target[prop];
    return async () => {
      console.warn(missingApiWarning, `Appel ignoré: api.${String(prop)}()`);
      return null;
    };
  }
});
const api = window.api ?? fallbackApi;
if (!window.api) {
  console.warn(missingApiWarning);
}
const $  = (s) => document.querySelector(s);
const $$ = (s) => Array.from(document.querySelectorAll(s));

const state = {
  folders: [],
  selectedFolder: null,
  stores: [],
};

function fmtDate(ts) {
  if (!ts) return '—';
  const d = new Date(ts * 1000);
  return d.toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
}

function setChip(el, label, stateAttr) {
  el.textContent = label;
  el.dataset.state = stateAttr;
}

async function refreshHeader() {
  const v = await api.version().catch(() => '?');
  $('#appVersion').textContent = 'v' + v;

  const ol = await api.outlookStatus().catch(() => ({ connected: false }));
  setChip($('#outlookChip'), ol.connected ? 'Outlook OK' : 'Outlook KO', ol.connected ? 'on' : 'err');

  const mo = await api.monitoringStatus().catch(() => ({ running: false }));
  setChip($('#monitorChip'), mo.running ? 'Monitoring actif' : 'Monitoring inactif', mo.running ? 'on' : 'warn');

  const as = await api.autostartGet().catch(() => false);
  $('#autostartCb').checked = !!as;
}

async function refreshKpis() {
  try {
    const s = await api.statsSummary();
    $('#kpiFolders').textContent = s.folders ?? 0;
    $('#kpiEmails').textContent  = (s.emails  ?? 0).toLocaleString('fr-FR');
    $('#kpiUnread').textContent  = (s.unread  ?? 0).toLocaleString('fr-FR');
    $('#kpiRecent').textContent  = (s.last7days ?? 0).toLocaleString('fr-FR');
  } catch (e) { console.warn('KPI KO', e); }
}

async function refreshFolders() {
  try {
    state.folders = await api.foldersList();
    const box = $('#foldersList');
    if (!state.folders.length) {
      box.innerHTML = '<div class="empty">Aucun dossier en monitoring.</div>';
      return;
    }
    box.innerHTML = '';
    for (const f of state.folders) {
      const row = document.createElement('div');
      row.className = 'folder-row';
      row.innerHTML = `
        <span class="path">${escapeHtml(f.path)}</span>
        <span class="cat cat-${f.category}">${labelCat(f.category)}</span>
        <button class="remove" data-eid="${f.entryId}">Retirer</button>`;
      box.appendChild(row);
    }
    box.querySelectorAll('.remove').forEach(btn => btn.addEventListener('click', async () => {
      const eid = btn.dataset.eid;
      if (!confirm('Retirer ce dossier du monitoring ?')) return;
      try { await api.folderRemove(eid); refreshFolders(); refreshKpis(); }
      catch (e) { alert('Échec : ' + e.message); }
    }));
  } catch (e) { console.warn('folders KO', e); }
}

function labelCat(c) {
  return c === 'declarations' ? 'Déclarations'
       : c === 'reglements'   ? 'Règlements'
       : 'Mails';
}

async function refreshByCategory() {
  try {
    const list = await api.statsByCategory();
    const total = list.reduce((a, b) => a + Number(b.count), 0) || 1;
    const box = $('#catChart');
    box.innerHTML = '';
    if (!list.length) { box.innerHTML = '<div class="empty">Aucune donnée.</div>'; return; }
    for (const r of list) {
      const pct = Math.round(Number(r.count) * 100 / total);
      const row = document.createElement('div');
      row.className = 'bar-row';
      row.innerHTML = `
        <span class="lbl">${labelCat(r.category)}</span>
        <span class="bar"><span class="cat-${r.category}" style="width:${pct}%"></span></span>
        <span class="val">${Number(r.count).toLocaleString('fr-FR')}</span>`;
      box.appendChild(row);
    }
  } catch (e) { console.warn('byCategory KO', e); }
}

async function refreshWeekly() {
  try {
    const rows = await api.statsWeekly(12);
    const agg = new Map(); // key=year-week -> total
    for (const r of rows) {
      const k = `${r.year}-W${String(r.week).padStart(2, '0')}`;
      agg.set(k, (agg.get(k) ?? 0) + Number(r.count));
    }
    const ordered = [...agg.entries()].sort((a, b) => a[0] > b[0] ? -1 : 1).slice(0, 12).reverse();
    const max = Math.max(1, ...ordered.map(([,v]) => v));
    const box = $('#weeklyChart');
    box.innerHTML = '';
    if (!ordered.length) { box.innerHTML = '<div class="empty">Aucune donnée.</div>'; return; }
    for (const [k, v] of ordered) {
      const pct = Math.round(v * 100 / max);
      const row = document.createElement('div');
      row.className = 'bar-row';
      row.innerHTML = `
        <span class="lbl">${k}</span>
        <span class="bar"><span style="width:${pct}%"></span></span>
        <span class="val">${v.toLocaleString('fr-FR')}</span>`;
      box.appendChild(row);
    }
  } catch (e) { console.warn('weekly KO', e); }
}

async function refreshEmails() {
  try {
    const list = await api.emailsRecent(20);
    const box = $('#emailsList');
    box.innerHTML = '';
    if (!list.length) { box.innerHTML = '<div class="empty">Aucun mail.</div>'; return; }
    for (const m of list) {
      const row = document.createElement('div');
      row.className = 'email-row' + (m.isUnread ? ' unread' : '');
      row.innerHTML = `
        <span class="when">${fmtDate(m.receivedTs)}</span>
        <span class="who">${escapeHtml(m.sender || '—')}</span>
        <span class="sub">${escapeHtml(m.subject || '(sans objet)')}</span>
        <span class="cat-tag">${labelCat(m.category)}</span>`;
      box.appendChild(row);
    }
  } catch (e) { console.warn('emails KO', e); }
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

// Logs (live)
function appendLog(entry) {
  const box = $('#logsBox');
  const dt = new Date(entry.ts).toLocaleTimeString('fr-FR');
  const lvl = (entry.level || 0);
  const lvlNames = ['DEBUG', 'INFO', 'WARN', 'ERROR'];
  const cls = ['log-debug', 'log-info', 'log-warn', 'log-error'][lvl] || 'log-info';
  const line = document.createElement('div');
  line.className = cls;
  line.textContent = `${dt} [${lvlNames[lvl]}] ${entry.category}: ${entry.message}`;
  box.appendChild(line);
  while (box.childElementCount > 400) box.removeChild(box.firstChild);
  box.scrollTop = box.scrollHeight;
  $('#logsCount').textContent = box.childElementCount + ' lignes';
}

async function loadInitialLogs() {
  try {
    const entries = await api.logsRecent();
    $('#logsBox').innerHTML = '';
    for (const e of entries) appendLog(e);
  } catch (e) { console.warn('logs KO', e); }
}

// ---- Add folder modal ----
const modal = $('#addModal');
$('#addFolderBtn').addEventListener('click', openAddModal);
$('#closeModalBtn').addEventListener('click', () => modal.hidden = true);
$('#cancelAddBtn').addEventListener('click', () => modal.hidden = true);
$('#confirmAddBtn').addEventListener('click', confirmAdd);

async function openAddModal() {
  modal.hidden = false;
  state.selectedFolder = null;
  $('#confirmAddBtn').disabled = true;
  $('#folderTree').textContent = 'Chargement…';

  try {
    state.stores = await api.outlookListStores();
    const sel = $('#storeSelect');
    sel.innerHTML = '';
    for (const s of state.stores) {
      const opt = document.createElement('option');
      opt.value = s.id; opt.textContent = s.name;
      sel.appendChild(opt);
    }
    sel.onchange = () => loadFoldersFor(sel.value);
    if (state.stores.length) loadFoldersFor(state.stores[0].id);
    else $('#folderTree').textContent = 'Aucune boîte détectée.';
  } catch (e) {
    $('#folderTree').textContent = 'Outlook non joignable : ' + e.message;
  }
}

async function loadFoldersFor(storeId) {
  $('#folderTree').textContent = 'Chargement…';
  try {
    const folders = await api.outlookListFolders(storeId);
    const box = $('#folderTree');
    box.innerHTML = '';
    if (!folders.length) { box.textContent = 'Vide.'; return; }
    for (const f of folders) {
      const depth = (f.path.match(/\//g) || []).length;
      const node = document.createElement('div');
      node.className = 'tree-node';
      node.innerHTML = `<span class="indent">${'│  '.repeat(depth)}</span>📂 ${escapeHtml(f.name)} <span class="indent">(${f.itemCount})</span>`;
      node.title = f.path;
      node.addEventListener('click', () => {
        box.querySelectorAll('.tree-node.selected').forEach(n => n.classList.remove('selected'));
        node.classList.add('selected');
        state.selectedFolder = { storeId, entryId: f.entryId, path: f.path, displayName: f.name };
        $('#confirmAddBtn').disabled = false;
      });
      box.appendChild(node);
    }
  } catch (e) { $('#folderTree').textContent = 'KO : ' + e.message; }
}

async function confirmAdd() {
  if (!state.selectedFolder) return;
  const cat = $('#categorySelect').value || null;
  const payload = { ...state.selectedFolder, category: cat };
  $('#confirmAddBtn').disabled = true;
  try {
    await api.folderAdd(payload);
    modal.hidden = true;
    refreshFolders(); refreshKpis(); refreshByCategory(); refreshEmails();
  } catch (e) {
    alert('Échec ajout : ' + e.message);
    $('#confirmAddBtn').disabled = false;
  }
}

// Window controls
$('#minBtn').addEventListener('click', () => api.windowMinimize());
$('#closeBtn').addEventListener('click', () => api.windowClose());
$('#autostartCb').addEventListener('change', e => api.autostartSet(e.target.checked));
$('#checkUpdatesBtn').addEventListener('click', async () => {
  setChip($('#updateChip'), 'Recherche…', 'warn');
  try { await api.checkUpdates(); }
  catch { setChip($('#updateChip'), 'Erreur MAJ', 'err'); }
});

// Live events
api.on('mail.add',     () => { refreshKpis(); refreshEmails(); refreshByCategory(); });
api.on('mail.change',  () => { refreshKpis(); refreshEmails(); });
api.on('mail.remove',  () => { refreshKpis(); refreshEmails(); });
api.on('stats.changed',() => { refreshKpis(); refreshByCategory(); refreshWeekly(); });
api.on('log.entry',    appendLog);
api.on('update.available', d => setChip($('#updateChip'), 'MAJ ' + d.version + ' dispo', 'warn'));
api.on('update.progress',  d => setChip($('#updateChip'), 'DL ' + d.percent + ' %',  'warn'));
api.on('update.ready',     ()=> setChip($('#updateChip'), 'MAJ prête (redémarrer)', 'warn'));

// Boot
(async () => {
  await refreshHeader();
  await Promise.all([
    refreshKpis(), refreshFolders(), refreshByCategory(), refreshWeekly(), refreshEmails(), loadInitialLogs()
  ]);
})();
