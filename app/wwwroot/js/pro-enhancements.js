/* Mail Monitor — pro-enhancements.js
 * Surcouche UI sans toucher au markup généré par app.js :
 *  - injecte un page-header par onglet (titre + sous-titre + pills meta)
 *  - met à jour le statut connexion et la semaine ISO
 *  - expose window.MMToast pour notifications discrètes
 *  - raccourcis clavier : Ctrl+1..9 (changement d'onglet), Ctrl+K (focus search)
 *  - badge "dernière synchro" temps réel via monitoring.cycle-complete
 *
 * Ce script est défensif : pas d'erreur si l'API n'est pas dispo.
 */
(function () {
  if (window.__mmProInit) return; window.__mmProInit = true;

  const TABS = [
    { id: 'dashboard',            icon: 'bi-speedometer2',          title: 'Tableau de bord',        subtitle: 'Vue d\u2019ensemble du portefeuille mails surveill\u00e9s' },
    { id: 'emails',               icon: 'bi-envelope',              title: 'Emails r\u00e9cents',    subtitle: 'Flux entrant cat\u00e9goris\u00e9 et trac\u00e9' },
    { id: 'weekly',               icon: 'bi-calendar-week',         title: 'Suivi hebdomadaire',     subtitle: 'Arriv\u00e9es / trait\u00e9s / stock par semaine ISO' },
    { id: 'personal-performance', icon: 'bi-person-lines-fill',     title: 'Performances personnelles', subtitle: 'Indicateurs gestionnaire sur 6 semaines glissantes' },
    { id: 'import-activity',      icon: 'bi-file-earmark-spreadsheet', title: 'Import Activit\u00e9 (.xlsb)', subtitle: 'R\u00e9conciliation avec le suivi Excel historique' },
    { id: 'monitoring',           icon: 'bi-eye',                   title: 'Monitoring',             subtitle: 'Dossiers Outlook surveill\u00e9s et arborescence' },
    { id: 'logs',                 icon: 'bi-journal-text',          title: 'Journal applicatif',     subtitle: 'Historique d\u00e9taill\u00e9 et flux temps r\u00e9el' },
    { id: 'db',                   icon: 'bi-database',              title: 'Base de donn\u00e9es',   subtitle: 'Inspection SQLite (lecture seule)' },
    { id: 'settings',             icon: 'bi-gear',                  title: 'Param\u00e8tres',        subtitle: 'Pr\u00e9f\u00e9rences application et compte' }
  ];

  function isoWeekLabel(d) {
    d = d || new Date();
    const tmp = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const day = tmp.getUTCDay() || 7;
    tmp.setUTCDate(tmp.getUTCDate() + 4 - day);
    const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((tmp - yearStart) / 86400000) + 1) / 7);
    return `S${String(weekNo).padStart(2, '0')} \u2022 ${tmp.getUTCFullYear()}`;
  }

  // -------------------- Toasts --------------------
  function ensureToastHost() {
    let host = document.getElementById('mm-toast-host');
    if (!host) {
      host = document.createElement('div');
      host.id = 'mm-toast-host';
      document.body.appendChild(host);
    }
    return host;
  }
  window.MMToast = function (message, kind, timeoutMs) {
    const host = ensureToastHost();
    const t = document.createElement('div');
    t.className = 'mm-toast ' + (kind || '');
    t.textContent = message || '';
    host.appendChild(t);
    setTimeout(() => { t.style.opacity = '0'; t.style.transition = 'opacity .25s'; setTimeout(() => t.remove(), 260); }, timeoutMs || 3500);
  };

  // -------------------- Page headers --------------------
  function injectHeaders() {
    TABS.forEach(t => {
      const pane = document.getElementById(t.id);
      if (!pane) return;
      if (pane.querySelector(':scope > .page-header')) return;
      const h = document.createElement('div');
      h.className = 'page-header';
      h.dataset.tab = t.id;
      h.innerHTML = `
        <div class="ph-title">
          <span class="ph-icon"><i class="bi ${t.icon}"></i></span>
          <div>
            <h1>${t.title}</h1>
            <div class="ph-subtitle">${t.subtitle}</div>
          </div>
        </div>
        <div class="ph-meta">
          <span class="ph-pill" data-role="week"><i class="bi bi-calendar3"></i><span>${isoWeekLabel()}</span></span>
          <span class="ph-pill" data-role="conn"><i class="bi bi-plug"></i><span>Connexion\u2026</span></span>
          <span class="ph-pill" data-role="sync"><i class="bi bi-arrow-repeat"></i><span>Synchro --:--</span></span>
        </div>`;
      pane.insertBefore(h, pane.firstChild);
    });
  }

  function setMetaAll(role, html, cls) {
    document.querySelectorAll(`.page-header .ph-pill[data-role="${role}"]`).forEach(p => {
      p.classList.remove('ok', 'warn', 'bad');
      if (cls) p.classList.add(cls);
      const span = p.querySelector('span'); if (span) span.innerHTML = html;
    });
  }

  // -------------------- Connexion / synchro --------------------
  async function refreshConnection() {
    try {
      const s = await window.electronAPI.outlookStatus();
      if (s && s.connected) {
        setMetaAll('conn', 'Outlook connect\u00e9', 'ok');
      } else if (s && s.status === 'connecting') {
        setMetaAll('conn', 'Connexion\u2026', 'warn');
      } else {
        setMetaAll('conn', 'Outlook indisponible', 'bad');
      }
    } catch { setMetaAll('conn', 'Statut inconnu', 'warn'); }
  }

  function setLastSync(date) {
    const d = date || new Date();
    const hh = String(d.getHours()).padStart(2,'0');
    const mm = String(d.getMinutes()).padStart(2,'0');
    const ss = String(d.getSeconds()).padStart(2,'0');
    setMetaAll('sync', `Synchro ${hh}:${mm}:${ss}`);
    const ls = document.getElementById('last-sync');
    if (ls) ls.textContent = `${hh}:${mm}:${ss}`;
  }

  // -------------------- Raccourcis clavier --------------------
  function bindShortcuts() {
    document.addEventListener('keydown', (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
      if (e.key >= '1' && e.key <= '9') {
        const idx = parseInt(e.key, 10) - 1;
        const tab = TABS[idx];
        if (!tab) return;
        const link = document.querySelector(`[data-bs-target="#${tab.id}"]`);
        if (link) { e.preventDefault(); link.click(); }
      } else if (e.key.toLowerCase() === 'k') {
        const active = document.querySelector('.tab-pane.active, .tab-pane.show.active');
        if (!active) return;
        const inp = active.querySelector('input[type="search"], input[type="text"]');
        if (inp) { e.preventDefault(); inp.focus(); inp.select && inp.select(); }
      }
    });
  }

  // -------------------- Boot --------------------
  function boot() {
    try { injectHeaders(); } catch (e) { console.warn('[pro] header inject KO', e); }
    bindShortcuts();
    // Live updates
    try {
      if (window.electronAPI?.onMonitoringCycleComplete) {
        window.electronAPI.onMonitoringCycleComplete(() => setLastSync());
      }
      if (window.electronAPI?.onMonitoringStatus) {
        window.electronAPI.onMonitoringStatus(() => refreshConnection());
      }
      if (window.electronAPI?.onCOMListeningStarted) {
        window.electronAPI.onCOMListeningStarted(() => setMetaAll('conn', 'Outlook connect\u00e9', 'ok'));
      }
      if (window.electronAPI?.onCOMListeningFailed) {
        window.electronAPI.onCOMListeningFailed(() => setMetaAll('conn', 'Outlook indisponible', 'bad'));
      }
    } catch (e) { console.warn('[pro] events KO', e); }
    // Premier tick
    refreshConnection();
    setLastSync();
    setInterval(refreshConnection, 30000);
    // Met à jour la semaine ISO chaque heure
    setInterval(() => setMetaAll('week', isoWeekLabel()), 60 * 60 * 1000);

    // Toast d'accueil discret (une fois)
    if (!sessionStorage.getItem('mm-welcomed')) {
      sessionStorage.setItem('mm-welcomed', '1');
      setTimeout(() => window.MMToast('Interface pr\u00eate \u2022 Ctrl+1\u20269 pour naviguer, Ctrl+K pour rechercher', 'ok', 5000), 600);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
