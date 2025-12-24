/**
 * Mail Monitor - Application refactorisée
 * 
 * Copyright (c) 2025 Tanguy Raingeard. Tous droits réservés.
 * 
 * Architecture claire avec Bootstrap 5 et persistance des configurations
 */

class MailMonitor {
  constructor() {
    this.state = {
      isConnected: false,
      isMonitoring: false,
      folderCategories: {},
      stats: {
        emailsToday: 0,
        treatedToday: 0,  // Changé de sentToday à treatedToday
        unreadTotal: 0,
        avgResponseTime: '0.0'
      },
      recentEmails: [],
      databaseStats: {},
      lastUpdate: null,
  currentWeekInfo: null,
      weeklyHistory: {
        page: 1,
  pageSize: 5,
        totalWeeks: 0,
        totalPages: 1
      },
      ui: {
  foldersSort: 'alpha', // 'alpha' | 'unread'
  theme: 'light' // 'light' | 'dark'
      }
    };
  // Logs view state
  this.logsState = { level: 'all', search: '', lastId: 0, autoScroll: true };
    
    // Gestionnaire de chargement centralisé
    this.loadingManager = {
      tasks: new Map(),
      totalTasks: 0,
      completedTasks: 0,
      isLoading: false
    };
    
    this.updateInterval = null;
  // Timers for animated counters to prevent overlaps
  this.counterTimers = new Map();
  // Prevent duplicate auto-refresh wiring
  this._autoRefreshSetup = false;
  // Weekly throttling/single-flight state
  this._weeklyRefreshTimer = null;
  this._weeklyInFlight = { stats: null, history: null };
  this._weeklyLastCall = { stats: 0, history: 0 };
  // DB tab state
  this._db = { initialized: false, active: false, timer: null };
    this.charts = {};
    this.init();
  }

  // === PREMIER LANCEMENT: forcer le choix "lu = traité" avant première sync ===
  async ensureFirstRunReadAsTreatedChoice() {
    try {
      const res = await window.electronAPI.invoke('api-settings-count-read-as-treated');
      if (!res?.success) return;

      // Si le setting existe déjà en base => rien à faire
      if (res.exists) return;

      // Flag: après la config initiale, rediriger vers l'ajout d'un dossier
      this._firstRunGoToAddFolder = true;

      const modalEl = document.getElementById('firstRunReadAsTreatedModal');
      const btnYes = document.getElementById('first-run-read-as-treated-yes');
      const btnNo = document.getElementById('first-run-read-as-treated-no');
      if (!modalEl || !btnYes || !btnNo || !window.bootstrap) return;

      await new Promise((resolve) => {
        const modal = new bootstrap.Modal(modalEl, { backdrop: 'static', keyboard: false });

        const cleanup = () => {
          try { btnYes.disabled = false; } catch (_) {}
          try { btnNo.disabled = false; } catch (_) {}
          btnYes.onclick = null;
          btnNo.onclick = null;
        };

        const choose = async (value) => {
          try {
            btnYes.disabled = true;
            btnNo.disabled = true;

            const save = await window.electronAPI.invoke('api-settings-count-read-as-treated', { value });
            if (!save?.success) {
              this.showNotification('Erreur', save?.error || save?.message || 'Impossible d\'enregistrer le paramètre', 'danger');
              btnYes.disabled = false;
              btnNo.disabled = false;
              return;
            }

            // Mettre à jour la checkbox du panneau Paramètres hebdo si elle est déjà dans le DOM
            const weeklyCheckbox = document.getElementById('count-read-as-treated');
            if (weeklyCheckbox) weeklyCheckbox.checked = !!value;

            // Autoriser le main process à démarrer le monitoring/baseline
            await window.electronAPI.invoke('api-first-run-complete');

            modal.hide();
            cleanup();
            resolve();
          } catch (e) {
            console.error('Erreur choix premier lancement:', e);
            this.showNotification('Erreur', 'Impossible d\'enregistrer le paramètre', 'danger');
            btnYes.disabled = false;
            btnNo.disabled = false;
          }
        };

        btnYes.onclick = () => choose(true);
        btnNo.onclick = () => choose(false);

        modal.show();
      });
    } catch (error) {
      console.error('Erreur ensureFirstRunReadAsTreatedChoice:', error);
    }
  }

  async goToMonitoringAndOpenAddFolder() {
    try {
      const nav = document.getElementById('nav-monitoring');
      if (nav && window.bootstrap?.Tab) {
        const tab = new bootstrap.Tab(nav);
        tab.show();
      }

      // Laisser Bootstrap appliquer le changement d'onglet
      await new Promise(r => setTimeout(r, 50));

      if (typeof this.showAddFolderModal === 'function') {
        await this.showAddFolderModal();
      }
    } catch (e) {
      console.warn('Impossible de rediriger vers l\'ajout de dossier:', e);
    }
  }

  // ---------- Performances personnelles: filtres ----------
  getPersonalPerformanceFilters() {
    const defaults = {
      periodMode: 'last-6',
      year: null,
      categories: {
        declarations: true,
        reglements: true,
        mails: true
      },
      includeAdg: false
    };

    let stored = null;
    try {
      stored = JSON.parse(localStorage.getItem('ppFilters') || 'null');
    } catch (_) {
      stored = null;
    }

    const merged = {
      ...defaults,
      ...(stored || {}),
      categories: {
        ...defaults.categories,
        ...((stored && stored.categories) || {})
      }
    };

    const periodModeEl = document.getElementById('pp-period-mode');
    const yearEl = document.getElementById('pp-year');
    const decEl = document.getElementById('pp-cat-declarations');
    const regEl = document.getElementById('pp-cat-reglements');
    const mailEl = document.getElementById('pp-cat-mails');
    const includeAdgEl = document.getElementById('pp-include-adg');

    const periodMode = periodModeEl?.value || merged.periodMode;
    const year = (yearEl?.value && String(yearEl.value).trim() !== '') ? Number(yearEl.value) : (merged.year || null);
    const categories = {
      declarations: decEl ? !!decEl.checked : !!merged.categories.declarations,
      reglements: regEl ? !!regEl.checked : !!merged.categories.reglements,
      mails: mailEl ? !!mailEl.checked : !!merged.categories.mails
    };
    const includeAdg = includeAdgEl ? !!includeAdgEl.checked : !!merged.includeAdg;

    return { periodMode, year, categories, includeAdg };
  }

  applyPersonalPerformanceFiltersToUI(filters) {
    const periodModeEl = document.getElementById('pp-period-mode');
    const yearEl = document.getElementById('pp-year');
    const yearWrapEl = document.getElementById('pp-year-wrap');
    const decEl = document.getElementById('pp-cat-declarations');
    const regEl = document.getElementById('pp-cat-reglements');
    const mailEl = document.getElementById('pp-cat-mails');
    const includeAdgEl = document.getElementById('pp-include-adg');

    if (periodModeEl && filters.periodMode) periodModeEl.value = filters.periodMode;
    const isYear = (filters.periodMode === 'year');
    if (yearWrapEl) yearWrapEl.style.display = isYear ? '' : 'none';
    if (yearEl && filters.year) yearEl.value = String(filters.year);
    if (decEl) decEl.checked = !!filters.categories?.declarations;
    if (regEl) regEl.checked = !!filters.categories?.reglements;
    if (mailEl) mailEl.checked = !!filters.categories?.mails;
    if (includeAdgEl) includeAdgEl.checked = !!filters.includeAdg;
  }

  persistPersonalPerformanceFilters(filters) {
    try {
      localStorage.setItem('ppFilters', JSON.stringify(filters));
    } catch (_) {}
  }

  // === INITIALISATION ===
  async init() {
    console.log('🚀 Initialisation de Mail Monitor...');
    
    try {
      // Premier lancement: demander le choix avant de charger le reste
      await this.ensureFirstRunReadAsTreatedChoice();

      // Injecter la version app dans le footer (source: app.getVersion)
      try {
              const ver = (window.electronAPI && await window.electronAPI.getAppVersion()) || 'unknown';
        const el = document.getElementById('app-version');
        if (el && ver) el.textContent = `v${ver}`;
      } catch {}
      // Appliquer le dernier thème connu immédiatement (fallback rapide) avant l'UI
      try {
        const t = localStorage.getItem('uiTheme');
        if (t === 'dark' || t === 'light') {
          document.body?.setAttribute('data-theme', t);
          this.state.ui.theme = t;
        }
      } catch(_) {}

      // Démarrer le gestionnaire de chargement
      this.startLoading();
      
      // Enregistrer toutes les tâches de chargement
      this.registerLoadingTask('configuration', 'Chargement de la configuration...');
      this.registerLoadingTask('connection', 'Vérification de la connexion...');
      this.registerLoadingTask('stats', 'Chargement des statistiques...');
      this.registerLoadingTask('emails', 'Chargement des emails récents...');
      this.registerLoadingTask('categories', 'Chargement des catégories...');
      this.registerLoadingTask('folders', 'Chargement des dossiers...');
      this.registerLoadingTask('vba', 'Chargement des métriques VBA...');
      this.registerLoadingTask('monitoring', 'Vérification du monitoring...');
      this.registerLoadingTask('settings', 'Chargement des paramètres...');
      this.registerLoadingTask('weekly', 'Initialisation du suivi hebdomadaire...');
      
      this.setupEventListeners();
      
      await this.completeLoadingTask('configuration', this.loadConfiguration());
      await this.completeLoadingTask('connection', this.checkConnection());
      
      // Charger les données en parallèle
      await Promise.allSettled([
        this.completeLoadingTask('stats', this.loadStats()),
        this.completeLoadingTask('emails', this.loadRecentEmails()),
        this.completeLoadingTask('categories', this.loadCategoryStats()),
        this.completeLoadingTask('folders', this.loadFoldersStats()),
        this.completeLoadingTask('vba', this.loadVBAMetrics()),
        this.completeLoadingTask('monitoring', this.checkMonitoringStatus()),
        this.completeLoadingTask('settings', this.loadSettings())
      ]);
      
      await this.completeLoadingTask('weekly', this.initWeeklyTracking());
  // Initialize logs tab
  this.initLogsTab();
  // Initialize DB tab (if enabled/visible, activation happens on tab change)
  this.initDbTab();
      
      this.startPeriodicUpdates();
  // setupAutoRefresh is already invoked in setupEventListeners(); avoid double wiring
      
      // Terminer le chargement
      this.finishLoading();

      // Sidebar responsive: appliquer et binder
      try {
        this.applyResponsiveSidebar();
        this.bindSidebarToggle();
        window.addEventListener('resize', (e) => {
          // Si le resize est programmatique, on ignore pour éviter la boucle
          if (e && e.isProgrammaticResize) return;
          this.applyResponsiveSidebar();
        });
      } catch(_) {}

      // Premier lancement: après choix "lu = traité", guider vers l'ajout de dossier
      if (this._firstRunGoToAddFolder) {
        this._firstRunGoToAddFolder = false;
        await this.goToMonitoringAndOpenAddFolder();
      }
      
      console.log('✅ MailMonitor app initialisée avec succès');
    } catch (error) {
      console.error('❌ Erreur lors de l\'initialisation:', error);
      this.showNotification('Erreur d\'initialisation', error.message, 'danger');
      this.finishLoading();
    }
  }

  // ====== BDD TAB (lecture brute) ======
  initDbTab() {
    const tableEl = document.getElementById('db-table');
    const limitEl = document.getElementById('db-limit');
    const autoEl = document.getElementById('db-auto');
    const refreshBtn = document.getElementById('db-refresh');
    const headEl = document.getElementById('db-table-head');
    const bodyEl = document.getElementById('db-table-body');
    const statusEl = document.getElementById('db-status');
    if (!tableEl || !limitEl || !autoEl || !refreshBtn || !headEl || !bodyEl || !statusEl) return;

    // State: filters + last data (client-side filtering)
    this._db = this._db || { initialized: false, active: false, timer: null };
    this._db.filters = this._db.filters || {};
    this._db.lastColumnsKey = this._db.lastColumnsKey || '';
    this._db.lastData = this._db.lastData || { columns: [], rows: [] };

    const escapeHtml = (text) => String(text ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');

    const formatCell = (val) => {
      if (val === null || val === undefined) return '<span class="text-muted">NULL</span>';
      if (typeof val === 'object') {
        try { return escapeHtml(JSON.stringify(val)); } catch { return escapeHtml(String(val)); }
      }
      return escapeHtml(String(val));
    };

    const setStatus = (msg) => {
      statusEl.textContent = msg || '—';
    };

    const renderTable = (columns, rows) => {
      const cols = Array.isArray(columns) ? columns : [];
      const safeCols = cols.length ? cols : (rows && rows[0] ? Object.keys(rows[0]) : []);

      // Persist last data for filtering
      this._db.lastData = { columns: safeCols, rows: Array.isArray(rows) ? rows : [] };

      const normalizeForSearch = (val) => {
        if (val === null || val === undefined) return '';
        if (typeof val === 'string') return val;
        if (typeof val === 'number' || typeof val === 'boolean') return String(val);
        try { return JSON.stringify(val); } catch { return String(val); }
      };

      const getActiveFilters = () => {
        const f = this._db.filters || {};
        const entries = Object.entries(f)
          .map(([k, v]) => [k, String(v ?? '').trim()])
          .filter(([, v]) => v.length > 0);
        return Object.fromEntries(entries);
      };

      const applyFilters = (allRows) => {
        const filters = getActiveFilters();
        const filterKeys = Object.keys(filters);
        if (filterKeys.length === 0) return allRows;
        return allRows.filter(r => {
          for (const key of filterKeys) {
            const q = filters[key].toLowerCase();
            const cell = normalizeForSearch(r?.[key]).toLowerCase();
            if (!cell.includes(q)) return false;
          }
          return true;
        });
      };

      const ensureHeaderWithFilters = (colsToUse) => {
        const key = (colsToUse || []).join('|');
        if (key && key === this._db.lastColumnsKey) return;
        this._db.lastColumnsKey = key;

        headEl.innerHTML = `
          <tr>${colsToUse.map(c => `<th>${escapeHtml(c)}</th>`).join('')}</tr>
          <tr>${colsToUse.map(c => {
            const current = String(this._db.filters?.[c] ?? '');
            return `
              <th>
                <input
                  type="text"
                  class="form-control form-control-sm"
                  data-db-filter-col="${escapeHtml(c)}"
                  placeholder="Filtrer…"
                  value="${escapeHtml(current)}"
                />
              </th>`;
          }).join('')}</tr>
        `;

        const inputs = headEl.querySelectorAll('input[data-db-filter-col]');
        for (const input of inputs) {
          input.addEventListener('input', () => {
            const col = input.getAttribute('data-db-filter-col');
            if (!col) return;
            this._db.filters[col] = input.value;
            // Re-render rows only, do not rebuild header
            renderRows();
          });
        }
      };

      const renderRows = () => {
        const allRows = this._db.lastData?.rows || [];
        const colsToUse = this._db.lastData?.columns || [];
        const filtered = applyFilters(allRows);

        if (!filtered || filtered.length === 0) {
          const active = Object.values(getActiveFilters()).some(v => v && String(v).trim().length > 0);
          const msg = active ? 'Aucun résultat (filtres actifs)' : 'Aucune donnée';
          bodyEl.innerHTML = `<tr><td class="text-muted py-3" colspan="${Math.max(1, colsToUse.length)}">${escapeHtml(msg)}</td></tr>`;
          return;
        }

        const html = filtered.map(r => {
          const cells = colsToUse.map(c => `<td>${formatCell(r?.[c])}</td>`).join('');
          return `<tr>${cells}</tr>`;
        }).join('');
        bodyEl.innerHTML = html;
      };

      ensureHeaderWithFilters(safeCols);
      renderRows();
    };

    const fetchTables = async () => {
      try {
        setStatus('Chargement des tables…');
        const res = await window.electronAPI.getDbTables?.();
        const tables = (res && res.success && Array.isArray(res.tables)) ? res.tables : [];
        tableEl.innerHTML = tables.length
          ? tables.map(t => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('')
          : '<option value="">(aucune table)</option>';
        // Sélectionner la première table si aucune sélection
        if (!tableEl.value && tables[0]) tableEl.value = tables[0];
        setStatus(`Tables: ${tables.length} • ${new Date().toLocaleTimeString('fr-FR')}`);
      } catch (e) {
        tableEl.innerHTML = '<option value="">(erreur)</option>';
        setStatus('Erreur chargement tables');
      }
    };

    const fetchPreview = async () => {
      const table = tableEl.value;
      const limit = Number(limitEl.value) || 200;
      if (!table) {
        renderTable([], []);
        setStatus('Sélectionnez une table');
        return;
      }
      try {
        setStatus(`Chargement ${table}…`);
        const res = await window.electronAPI.getDbTablePreview?.({ table, limit, offset: 0 });
        if (!(res && res.success)) {
          renderTable([], []);
          setStatus(res?.error || 'Erreur chargement table');
          return;
        }
        const data = res.data || {};
        renderTable(data.columns || [], data.rows || []);
        setStatus(`${escapeHtml(table)} • ${Array.isArray(data.rows) ? data.rows.length : 0} ligne(s) • ${new Date().toLocaleTimeString('fr-FR')}`);
      } catch (e) {
        renderTable([], []);
        setStatus('Erreur chargement table');
      }
    };

    const startAuto = () => {
      if (this._db.timer) return;
      this._db.timer = setInterval(() => {
        // Auto-refresh uniquement si onglet actif ET option activée
        if (!this._db.active) return;
        if (!autoEl.checked) return;
        fetchPreview();
      }, 2000);
    };

    const stopAuto = () => {
      if (this._db.timer) {
        clearInterval(this._db.timer);
        this._db.timer = null;
      }
    };

    // Public-ish hooks for tab lifecycle
    this._dbStart = () => { this._db.active = true; startAuto(); };
    this._dbStop = () => { this._db.active = false; stopAuto(); };
    this._dbRefreshNow = async () => {
      if (!this._db.initialized) {
        await fetchTables();
        this._db.initialized = true;
      }
      await fetchPreview();
    };

    // Bind UI events
    refreshBtn.addEventListener('click', () => this._dbRefreshNow());
    tableEl.addEventListener('change', () => this._dbRefreshNow());
    limitEl.addEventListener('change', () => this._dbRefreshNow());
    autoEl.addEventListener('change', () => {
      if (autoEl.checked) startAuto();
    });

    // Initial tables load is deferred until first activation to avoid extra IPC on startup
  }

  // ---------- Sidebar responsive ----------
  bindSidebarToggle() {
    const toggle = (e) => {
      if (e) e.preventDefault();
      const body = document.body;
      const collapsed = body.classList.toggle('sidebar-collapsed');
      try { localStorage.setItem('sidebarCollapsed', collapsed ? '1' : '0'); } catch(_) {}
  // Force layout recalculation so content adapts to new width
  this.deferLayoutResize();
    };
    const burger = document.getElementById('sidebar-burger');
    if (burger) burger.addEventListener('click', toggle);
    const fab = document.getElementById('sidebar-fab');
    if (fab) fab.addEventListener('click', toggle);
  }

  applyResponsiveSidebar() {
    const body = document.body;
    const w = window.innerWidth || document.documentElement.clientWidth;
    let pref = null;
    try { pref = localStorage.getItem('sidebarCollapsed'); } catch(_) {}
    if (pref === '1') body.classList.add('sidebar-collapsed');
    else if (pref === '0') body.classList.remove('sidebar-collapsed');
    else {
      if (w < 1100) body.classList.add('sidebar-collapsed');
      else body.classList.remove('sidebar-collapsed');
    }
    // After applying state, ensure charts and tables resize to fit
    this.deferLayoutResize();
  }

  // Trigger a layout resize on next frame (and again after transition) to ensure responsive components adapt
  deferLayoutResize() {
    // Immediate resize
    this.forceLayoutResize();
    // After CSS transition (300ms in CSS vars), run again
    setTimeout(() => this.forceLayoutResize(), 320);
  }

  forceLayoutResize() {
    try {
      // Notify responsive libs (Chart.js listens to window resize)
      const evt = new Event('resize');
      evt.isProgrammaticResize = true;
      window.dispatchEvent(evt);
    } catch(_) {}
    try {
      if (this.charts) {
        Object.values(this.charts).forEach(c => {
          if (c && typeof c.resize === 'function') c.resize();
          if (c && c.update) c.update('none');
        });
      }
    } catch(_) {}
  }

  // === GESTIONNAIRE DE CHARGEMENT ===
  startLoading() {
    this.loadingManager.isLoading = true;
    this.loadingManager.totalTasks = 0;
    this.loadingManager.completedTasks = 0;
    this.loadingManager.tasks.clear();
    this.updateLoadingUI();
  }

  registerLoadingTask(taskId, description) {
    this.loadingManager.tasks.set(taskId, {
      description,
      completed: false,
      startTime: Date.now()
    });
    this.loadingManager.totalTasks++;
    this.updateLoadingUI();
  }

  async completeLoadingTask(taskId, promise) {
    try {
      const result = await promise;
      if (this.loadingManager.tasks.has(taskId)) {
        this.loadingManager.tasks.get(taskId).completed = true;
        this.loadingManager.completedTasks++;
        this.updateLoadingUI();
      }
      return result;
    } catch (error) {
      console.error(`❌ Erreur tâche ${taskId}:`, error);
      if (this.loadingManager.tasks.has(taskId)) {
        this.loadingManager.tasks.get(taskId).completed = true;
        this.loadingManager.tasks.get(taskId).error = error.message;
        this.loadingManager.completedTasks++;
        this.updateLoadingUI();
      }
      throw error;
    }
  }

  updateLoadingUI() {
    if (!this.loadingManager.isLoading) return;
    
    const percentage = this.loadingManager.totalTasks > 0 
      ? Math.round((this.loadingManager.completedTasks / this.loadingManager.totalTasks) * 100)
      : 0;
    
    // Mettre à jour la barre de progression s'il y en a une
    const progressBar = document.querySelector('.loading-progress-bar');
    if (progressBar) {
      progressBar.style.width = `${percentage}%`;
      progressBar.textContent = `${percentage}%`;
    }
    
    // Mettre à jour le texte de statut
    const statusText = document.querySelector('.loading-status-text');
    if (statusText) {
      const currentTask = Array.from(this.loadingManager.tasks.values())
        .find(task => !task.completed);
      
      if (currentTask) {
        statusText.textContent = currentTask.description;
      } else {
        statusText.textContent = 'Finalisation...';
      }
    }
    
    // Mettre à jour le compteur
    const counter = document.querySelector('.loading-counter');
    if (counter) {
      counter.textContent = `${this.loadingManager.completedTasks}/${this.loadingManager.totalTasks}`;
    }
    
    console.log(`📊 Chargement: ${this.loadingManager.completedTasks}/${this.loadingManager.totalTasks} (${percentage}%)`);
  }

  finishLoading() {
    this.loadingManager.isLoading = false;
    
    // Masquer la page de chargement
    const loadingScreen = document.getElementById('loading-screen');
    if (loadingScreen) {
      loadingScreen.style.display = 'none';
    }
    
    // Afficher l'interface principale
    const mainContent = document.getElementById('main-content');
    if (mainContent) {
      mainContent.style.display = 'block';
    }
    
    console.log('✅ Chargement terminé');
  }

  /**
   * Configuration de l'actualisation automatique
   */
  setupAutoRefresh() {
  // Guard against double wiring
  if (this._autoRefreshSetup) return;
  this._autoRefreshSetup = true;
  // Log d'initialisation - conservé pour débogage important
  console.log('🔄 Configuration de l\'actualisation automatique...');
    
    // Actualisation des statistiques principales toutes les 1 seconde
    this.statsRefreshInterval = setInterval(() => {
      this.performStatsRefresh();
    }, 1000);
    
    // Actualisation des emails récents toutes les 1 seconde
    this.emailsRefreshInterval = setInterval(() => {
      this.performEmailsRefresh();
    }, 1000);
    
    // Actualisation des statistiques de dossiers toutes les 2 secondes
    this.foldersRefreshInterval = setInterval(() => {
      this.loadFoldersStats();
    }, 2000);
    
    // Actualisation complète toutes les 5 secondes
    this.fullRefreshInterval = setInterval(() => {
      this.performFullRefresh();
    }, 5000);
    
    // Log de confirmation configuration - conservé pour débogage
    console.log('✅ Auto-refresh configuré (1s stats, 1s emails, 2s dossiers, 5s complet)');
  }

  // Debounced weekly refresh to coalesce multiple triggers
  scheduleWeeklyRefresh(delayMs = 300) {
    if (this._weeklyRefreshTimer) clearTimeout(this._weeklyRefreshTimer);
    this._weeklyRefreshTimer = setTimeout(async () => {
      try {
        await this.loadCurrentWeekStats();
        await this.loadWeeklyHistory();
      } catch (_) {}
    }, Math.max(0, delayMs));
  }

  // Debounced folders refresh to reflect count changes quickly
  scheduleFoldersStatsRefresh(delayMs = 250, forceFoldersTree = false) {
    if (forceFoldersTree) this._forceFoldersTreeNext = true;
    if (this._foldersStatsRefreshTimer) clearTimeout(this._foldersStatsRefreshTimer);
    this._foldersStatsRefreshTimer = setTimeout(async () => {
      try {
        await this.loadFoldersStats({ forceFoldersTree: !!this._forceFoldersTreeNext });
        this._forceFoldersTreeNext = false;
      } catch (_) {}
    }, Math.max(0, delayMs));
  }

  /**
   * Configuration des listeners d'événements temps réel
   */
  setupRealtimeEventListeners() {
    console.log('🔔 Configuration des événements temps réel...');
    
    // Écouter les mises à jour de statistiques en temps réel
    if (window.electronAPI.onStatsUpdate) {
      window.electronAPI.onStatsUpdate((stats) => {
        // Stats temps réel - log désactivé pour réduire le spam
        // console.log('📊 Événement stats temps réel reçu:', stats);
        this.state.stats = { ...this.state.stats, ...stats };
        this.updateStatsDisplay();
        this.updateLastRefreshTime();
      });
    }
    
    // Écouter les mises à jour d'emails en temps réel
    if (window.electronAPI.onEmailUpdate) {
      window.electronAPI.onEmailUpdate((emailData) => {
        console.log('📧 Événement email temps réel reçu:', emailData);
        this.handleRealtimeEmailUpdate(emailData);
      });
    }
    
    // Écouter les nouveaux emails en temps réel
    if (window.electronAPI.onNewEmail) {
      window.electronAPI.onNewEmail((emailData) => {
        console.log('📬 Nouvel email temps réel reçu:', emailData);
        this.handleRealtimeNewEmail(emailData);
      });
    }
    
    // Écouter la fin des cycles de monitoring
    if (window.electronAPI.onMonitoringCycleComplete) {
      window.electronAPI.onMonitoringCycleComplete((cycleData) => {
        // Cycle monitoring - log désactivé pour réduire le spam
        // console.log('🔄 Cycle de monitoring terminé:', cycleData);
        this.handleMonitoringCycleComplete(cycleData);
      });
    }

    // Changements de compteur dossiers (déclenche généralement une sync partielle)
    if (window.electronAPI.onFolderCountUpdated) {
      window.electronAPI.onFolderCountUpdated((payload) => {
        // api-folders-tree est caché côté main (5 min). Sur événement COUNT on force un refresh non-caché.
        this.scheduleFoldersStatsRefresh(350, true);
        // Rafraîchir aussi les cartes dashboard (emailsToday/unreadTotal) sans attendre.
        try { this.performStatsRefresh(); } catch (_) {}
      });
    }

    // Écouter les changements de statut du monitoring
    if (window.electronAPI.onMonitoringStatus) {
      window.electronAPI.onMonitoringStatus((status) => {
        const isActive = status?.status === 'active';
        this.state.isMonitoring = isActive;
        this.updateMonitoringStatus(isActive);
      });
    }

    // NOUVEAU: Écouter les événements COM Outlook
    if (window.electronAPI.onCOMListeningStarted) {
      window.electronAPI.onCOMListeningStarted((data) => {
        console.log('🔔 COM listening started:', data);
        this.handleCOMListeningStarted(data);
      });
    }

    if (window.electronAPI.onCOMListeningFailed) {
      window.electronAPI.onCOMListeningFailed((error) => {
        console.log('❌ COM listening failed:', error);
        this.handleCOMListeningFailed(error);
      });
    }

    // Événements temps réel pour les emails COM
    if (window.electronAPI.onRealtimeEmailUpdate) {
      window.electronAPI.onRealtimeEmailUpdate((emailData) => {
        console.log('📧 Mise à jour email temps réel COM:', emailData);
        this.handleRealtimeEmailUpdate(emailData);
      });
    }

    if (window.electronAPI.onRealtimeNewEmail) {
      window.electronAPI.onRealtimeNewEmail((emailData) => {
        console.log('📬 Nouvel email temps réel COM:', emailData);
        this.handleRealtimeNewEmail(emailData);
      });
    }
    
    // NOUVEAU: Rafraîchissement du dashboard quand le cache est invalidé
    if (window.electronAPI.onStatsCacheInvalidated) {
      window.electronAPI.onStatsCacheInvalidated((payload) => {
        console.log('📊 Événement stats-cache-invalidated reçu (cache main invalide)', payload);
        // Rafraîchir immédiatement le dashboard et les emails récents
        try {
          this.performStatsRefresh();
          this.performEmailsRefresh();
        } catch (e) {
          console.warn('⚠️ Erreur refresh rapide après invalidation cache:', e);
        }
      });
    }
    
    // Rafraîchissement auto des stats hebdo après import/ajustements
    if (window.electronAPI.onWeeklyStatsUpdated) {
      window.electronAPI.onWeeklyStatsUpdated(async (payload) => {
        console.log('📅 Événement weekly-stats-updated reçu:', payload);
        try {
          // Coalesce refreshes to avoid duplicate calls
          this.scheduleWeeklyRefresh(200);
          // Mettre à jour aussi l'onglet Performances personnelles si utilisé
          await this.loadPersonalPerformance();
          this.updateLastRefreshTime();
        } catch (e) {
          console.warn('⚠️ Erreur refresh weekly après event:', e);
        }
      });
    }
    
    // Événements de mise à jour automatique
    this.setupUpdateListeners();
    
    console.log('✅ Événements temps réel configurés (y compris COM)');
  }

  // Configuration des listeners pour les mises à jour
  setupUpdateListeners() {
    if (!window.electronAPI) return;
    
    // Toast pour afficher les notifications
    const showUpdateToast = (message, type = 'info', duration = 5000) => {
      const toastContainer = document.getElementById('toast-container') || (() => {
        const container = document.createElement('div');
        container.id = 'toast-container';
        container.className = 'position-fixed top-0 end-0 p-3';
        container.style.zIndex = '9999';
        document.body.appendChild(container);
        return container;
      })();
      
      const toastId = `toast-${Date.now()}`;
      const toast = document.createElement('div');
      toast.id = toastId;
      toast.className = `toast align-items-center text-white bg-${type} border-0`;
      toast.setAttribute('role', 'alert');
      toast.setAttribute('aria-live', 'assertive');
      toast.setAttribute('aria-atomic', 'true');
      toast.innerHTML = `
        <div class="d-flex">
          <div class="toast-body">${message}</div>
          <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
      `;
      
      toastContainer.appendChild(toast);
      const bsToast = new bootstrap.Toast(toast, { delay: duration });
      bsToast.show();
      
      toast.addEventListener('hidden.bs.toast', () => {
        toast.remove();
      });
    };
    
    // Vérification en cours
    if (window.electronAPI.onUpdateChecking) {
      window.electronAPI.onUpdateChecking(() => {
        console.log('🔍 Vérification des mises à jour...');
      });
    }
    
    // Mise à jour disponible
    if (window.electronAPI.onUpdateAvailable) {
      window.electronAPI.onUpdateAvailable((info) => {
        console.log('🎉 Mise à jour disponible:', info);
        showUpdateToast(
          `<i class="bi bi-download me-2"></i>Mise à jour v${info.version} disponible ! Téléchargement en cours...`,
          'success',
          8000
        );
      });
    }
    
    // Aucune mise à jour
    if (window.electronAPI.onUpdateNotAvailable) {
      window.electronAPI.onUpdateNotAvailable(() => {
        console.log('✅ Application à jour');
      });
    }
    
    // Erreur de mise à jour
    if (window.electronAPI.onUpdateError) {
      window.electronAPI.onUpdateError((data) => {
        console.error('❌ Erreur de mise à jour:', data.error);
        showUpdateToast(
          `<i class="bi bi-exclamation-triangle me-2"></i>Erreur de mise à jour : ${data.error}`,
          'danger',
          10000
        );
      });
    }
    
    // Progression du téléchargement
    let lastProgressToast = null;
    if (window.electronAPI.onUpdateDownloadProgress) {
      window.electronAPI.onUpdateDownloadProgress((progress) => {
        const percent = Math.floor(progress.percent);
        console.log(`⬇️ Téléchargement: ${percent}%`);
        
        // Mettre à jour le toast existant ou en créer un nouveau
        if (percent % 10 === 0 || percent === 100) { // Tous les 10%
          if (lastProgressToast) {
            const toastEl = document.getElementById(lastProgressToast);
            if (toastEl) {
              const body = toastEl.querySelector('.toast-body');
              if (body) {
                body.innerHTML = `<i class="bi bi-download me-2"></i>Téléchargement: ${percent}%`;
              }
            }
          } else {
            lastProgressToast = `toast-${Date.now()}`;
            showUpdateToast(
              `<i class="bi bi-download me-2"></i>Téléchargement: ${percent}%`,
              'info',
              percent === 100 ? 3000 : 30000
            );
          }
        }
      });
    }
    
    // Mise à jour en attente de redémarrage
    if (window.electronAPI.onUpdatePendingRestart) {
      window.electronAPI.onUpdatePendingRestart((data) => {
        console.log('🔄 Mise à jour en attente de redémarrage:', data);
        showUpdateToast(
          `<i class="bi bi-info-circle me-2"></i>La mise à jour v${data.version} sera installée au prochain démarrage`,
          'info',
          10000
        );
      });
    }
  }

  // ====== LOGS TAB ======
  initLogsTab() {
    const view = document.getElementById('logs-view');
    const countEl = document.getElementById('logs-count');
    const bufferedEl = document.getElementById('logs-buffered');
    const statusEl = document.getElementById('logs-status');
    const searchEl = document.getElementById('logs-search');
    const levelEl = document.getElementById('logs-level');
    const categoryEl = document.getElementById('logs-category');
  const refreshBtn = document.getElementById('logs-refresh');
    const exportBtn = document.getElementById('logs-export');
  const openFolderBtn = document.getElementById('logs-open-folder');
    const copyAllBtn = document.getElementById('logs-copy-all');
    const container = document.getElementById('logs-container') || view?.parentElement;
  if (!view || !countEl || !bufferedEl || !statusEl || !searchEl || !levelEl || !categoryEl || !refreshBtn || !exportBtn || !openFolderBtn || !copyAllBtn) return;

    // Update autoScroll based on user scroll position
    container.addEventListener('scroll', () => {
      const nearBottom = container.scrollTop + container.clientHeight >= container.scrollHeight - 24;
      this.logsState.autoScroll = nearBottom;
    });

    const levelToClass = (level) => {
      switch (String(level || '').toUpperCase()) {
        case 'ERROR':
          return 'text-danger';
        case 'WARN':
          return 'text-warning';
        case 'DEBUG':
          return 'text-secondary';
        case 'SUCCESS':
          return 'text-success';
        default:
          return 'text-body';
      }
    };

    const formatTime = (iso) => {
      try {
        return new Date(iso).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
      } catch {
        return String(iso || '--');
      }
    };

    const escapeHtml = (text) => String(text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    const renderHistory = (entries) => {
      const lines = (entries || []).map((e) => {
        const lvl = String(e.level || '').toUpperCase();
        const cls = levelToClass(lvl);
        const ts = formatTime(e.timestamp);
        const cat = escapeHtml(e.categoryLabel || e.category || '—');
        const msg = escapeHtml(e.message || '');
        const data = e.data ? `<div class="text-muted" style="white-space: pre-wrap; margin-left: 1.2rem;">${escapeHtml(e.data)}</div>` : '';
        return `<div class="${cls}">[${ts}] [${cat}] [${lvl}] ${msg}${data}</div>`;
      });
      view.innerHTML = lines.join('');
      countEl.textContent = String((entries || []).length);
      bufferedEl.textContent = '—';
      if (this.logsState.autoScroll) container.scrollTop = container.scrollHeight;
    };

    const buildFilters = () => {
      const level = levelEl.value || 'ALL';
      const category = categoryEl.value || 'ALL';
      const search = (searchEl.value || '').trim();
      return { level, category, search, limit: 1000 };
    };

    const fetchHistory = async () => {
      try {
        statusEl.textContent = 'Chargement...';
        const filters = buildFilters();
        const history = await window.electronAPI.invoke('api-get-log-history', filters);
        renderHistory(history || []);
        statusEl.textContent = `Dernière mise à jour: ${new Date().toLocaleTimeString()}`;
      } catch (e) {
        statusEl.textContent = 'Erreur de chargement';
        view.textContent = `Erreur lors du chargement des logs: ${e?.message || String(e)}`;
        countEl.textContent = '0';
        bufferedEl.textContent = '—';
      }
    };

    // Bind filters
    levelEl.addEventListener('change', async () => {
      this.logsState.level = levelEl.value || 'ALL';
      await fetchHistory();
    });
    categoryEl.addEventListener('change', async () => {
      this.logsState.category = categoryEl.value || 'ALL';
      await fetchHistory();
    });
    let searchTimer = null;
    searchEl.addEventListener('input', () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(async () => {
        this.logsState.search = searchEl.value || '';
        await fetchHistory();
      }, 250);
    });
    refreshBtn.addEventListener('click', () => fetchHistory());
    exportBtn.addEventListener('click', async () => {
      try {
        const filters = buildFilters();
        // Pour l'export, on préfère le maximum du buffer mémoire
        filters.limit = 2000;
        const res = await window.electronAPI.invoke('api-export-log-history', filters);
        if (res && res.success) {
          this.showNotification('Export des logs', `Fichier enregistré (${res.exported || 0} entrées)`, 'success');
        } else if (!(res && res.canceled)) {
          this.showNotification('Export des logs', 'Erreur lors de l\'export', 'danger');
        }
      } catch (e) {
        this.showNotification('Export des logs', 'Erreur lors de l\'export', 'danger');
      }
    });
    openFolderBtn.addEventListener('click', async () => {
      const res = await window.electronAPI.openLogsFolder();
      if (!(res && res.success)) this.showNotification('Dossier de logs', 'Impossible d\'ouvrir le dossier', 'warning');
    });
    copyAllBtn.addEventListener('click', async () => {
      try {
        // Récupérer le contenu texte de tous les logs affichés
        const logsText = view.textContent || '';
        if (!logsText.trim()) {
          this.showNotification('Copie des logs', 'Aucun log à copier', 'warning');
          return;
        }
        // Copier dans le presse-papier
        await navigator.clipboard.writeText(logsText);
        // Feedback visuel temporaire
        const originalHTML = copyAllBtn.innerHTML;
        copyAllBtn.innerHTML = '<i class="bi bi-check-lg me-1"></i>Copié!';
        copyAllBtn.classList.remove('btn-outline-success');
        copyAllBtn.classList.add('btn-success');
        setTimeout(() => {
          copyAllBtn.innerHTML = originalHTML;
          copyAllBtn.classList.remove('btn-success');
          copyAllBtn.classList.add('btn-outline-success');
        }, 2000);
        this.showNotification('Copie des logs', `${logsText.split('\n').length} lignes copiées dans le presse-papier`, 'success');
      } catch (err) {
        console.error('Erreur lors de la copie:', err);
        this.showNotification('Copie des logs', 'Erreur lors de la copie dans le presse-papier', 'danger');
      }
    });

    // Stream en temps réel: on accepte uniquement les entrées logService (avec timestamp ISO)
    if (window.electronAPI.onLogEntry) {
      window.electronAPI.onLogEntry(async (entry) => {
        if (!entry || !entry.timestamp || !entry.level) return;

        const currentLevel = (levelEl.value || 'ALL').toUpperCase();
        const currentCategory = (categoryEl.value || 'ALL').toUpperCase();
        const currentSearch = (searchEl.value || '').trim().toLowerCase();

        const lvl = String(entry.level).toUpperCase();
        const cat = String(entry.category || '').toUpperCase();

        // Niveau
        if (currentLevel !== 'ALL') {
          const levelValues = { DEBUG: 0, INFO: 1, SUCCESS: 1, WARN: 2, ERROR: 3 };
          if ((levelValues[lvl] ?? -1) < (levelValues[currentLevel] ?? 0)) return;
        }
        // Catégorie
        if (currentCategory !== 'ALL' && cat !== currentCategory) return;
        // Recherche
        if (currentSearch) {
          const hay = `${entry.message || ''}\n${entry.data || ''}`.toLowerCase();
          if (!hay.includes(currentSearch)) return;
        }

        const ts = formatTime(entry.timestamp);
        const catLabel = escapeHtml(entry.categoryLabel || entry.category || '—');
        const msg = escapeHtml(entry.message || '');
        const data = entry.data ? `\n${entry.data}` : '';

        const div = document.createElement('div');
        div.className = levelToClass(lvl);
        div.textContent = `[${ts}] [${catLabel}] [${lvl}] ${entry.message || ''}${data ? `\n${data}` : ''}`;
        view.appendChild(div);
        countEl.textContent = String((Number(countEl.textContent) || 0) + 1);
        if (this.logsState.autoScroll) container.scrollTop = container.scrollHeight;
      });
    }

    // Load initial content
    fetchHistory();
  }

  /**
   * Gestion des mises à jour d'emails en temps réel
   */
  handleRealtimeEmailUpdate(emailData) {
    // Actualiser immédiatement les statistiques
    this.performStatsRefresh();
    
    // Actualiser la liste des emails si on est sur l'onglet emails
    const activeTab = document.querySelector('.nav-pills .nav-link.active');
    const activeTabId = activeTab ? activeTab.getAttribute('data-bs-target') : '';
    
    if (activeTabId === '#emails') {
      this.loadRecentEmails();
    }
  }

  /**
   * Gestion des nouveaux emails en temps réel
   */
  handleRealtimeNewEmail(emailData) {
    // Actualiser immédiatement toutes les données
    this.performStatsRefresh();
    this.performEmailsRefresh();
    
    // Afficher une notification discrète
    this.showNotification(
      'Nouvel email détecté', 
      `Email: ${emailData.subject ? emailData.subject.substring(0, 50) : 'Sans sujet'}...`, 
      'info'
    );
  }

  /**
   * Gestion de la fin d'un cycle de monitoring
   */
  handleMonitoringCycleComplete(cycleData) {
    // Actualisation légère après chaque cycle
    this.performStatsRefresh();
    this.updateLastRefreshTime();
    // Rafraîchir aussi les compteurs dossiers (dashboard) après cycle
    try { this.scheduleFoldersStatsRefresh(600); } catch (_) {}
    // Rafraîchir aussi le suivi hebdomadaire à chaque cycle (traitements/reçus évoluent)
    // On ne spamme pas: appels idempotents et peu coûteux côté IPC
    try {
  // Coalesce weekly refresh to avoid double/triple logs
  this.scheduleWeeklyRefresh(400);
      this.loadPersonalPerformance?.();
    } catch(_) {}
  }

  /**
   * NOUVEAU: Gestion du démarrage de l'écoute COM
   */
  handleCOMListeningStarted(data) {
    console.log('🔔 COM Outlook écoute démarrée:', data);
    
    // Mettre à jour le statut de monitoring pour afficher le mode COM
    const statusContainer = document.getElementById('monitoring-status');
    if (statusContainer) {
      const comBadge = `<span class="badge bg-success ms-2">COM Actif</span>`;
      const existingContent = statusContainer.innerHTML;
      if (!existingContent.includes('COM Actif')) {
        statusContainer.innerHTML = existingContent.replace('</div>', comBadge + '</div>');
      }
    }
    
    // Notification discrète
    this.showNotification(
      'Écoute COM activée', 
      `Surveillance temps réel active sur ${data.folders} dossier(s)`, 
      'success'
    );
  }

  /**
   * NOUVEAU: Gestion de l'échec de l'écoute COM
   */
  handleCOMListeningFailed(error) {
    console.log('❌ COM Outlook écoute échouée:', error);
    
    // Mettre à jour le statut pour indiquer le fallback
    const statusContainer = document.getElementById('monitoring-status');
    if (statusContainer) {
      const fallbackBadge = `<span class="badge bg-warning ms-2">Mode Polling</span>`;
      const existingContent = statusContainer.innerHTML;
      statusContainer.innerHTML = existingContent.replace(/COM Actif/g, 'Mode Polling');
    }
    
    // Notification d'avertissement
    this.showNotification(
      'Basculement vers polling', 
      'L\'écoute COM a échoué, utilisation du polling de secours', 
      'warning'
    );
  }

  /**
   * Actualisation des statistiques (rapide)
   */
  async performStatsRefresh() {
    try {
      await Promise.allSettled([
        this.checkConnection(),
        this.loadStats()
      ]);
      this.updateLastRefreshTime();
    } catch (error) {
      console.warn('⚠️ Erreur actualisation stats:', error);
    }
  }

  /**
   * Actualisation des emails récents
   */
  async performEmailsRefresh() {
    try {
      // Actualiser seulement si on est sur l'onglet emails ou dashboard
      const activeTab = document.querySelector('.nav-pills .nav-link.active');
      const activeTabId = activeTab ? activeTab.getAttribute('data-bs-target') : '';
      
      if (activeTabId === '#emails' || activeTabId === '#dashboard') {
        await this.loadRecentEmails();
      }
    } catch (error) {
      console.warn('⚠️ Erreur actualisation emails:', error);
    }
  }

  /**
   * Actualisation légère (statuts et compteurs rapides)
   */
  async performLightRefresh() {
    try {
      await Promise.allSettled([
        this.checkConnection(),
        this.loadStats(),
        this.updateLastRefreshTime()
      ]);
    } catch (error) {
      console.warn('⚠️ Erreur actualisation légère:', error);
    }
  }

  /**
   * Actualisation complète (tous les données)
   */
  async performFullRefresh() {
    try {
      // Actualisation auto - log simplifié
      // console.log('🔄 Actualisation complète automatique...');
      await Promise.allSettled([
        this.checkConnection(),
        this.loadStats(),
        this.loadRecentEmails(),
        this.loadCategoryStats(),
        this.loadFoldersStats(),
        this.loadVBAMetrics()
      ]);
      this.updateLastRefreshTime();
      console.log('✅ Actualisation complète terminée');
    } catch (error) {
      console.warn('⚠️ Erreur actualisation complète:', error);
    }
  }

  /**
   * Méthode utilitaire pour mettre à jour le contenu d'un élément
   * @param {string} elementId - L'ID de l'élément à mettre à jour
   * @param {string} content - Le nouveau contenu
   */
  updateElement(elementId, content) {
    const element = document.getElementById(elementId);
    if (element) {
      element.textContent = content;
    }
  }

  /**
   * Mise à jour de l'heure de dernière actualisation
   */
  updateLastRefreshTime() {
    const now = new Date();
    const timeString = now.toLocaleTimeString('fr-FR');
    this.updateElement('last-sync', timeString);
  }

  setupEventListeners() {
    // Barre de titre personnalisée
    document.getElementById('minimize-btn')?.addEventListener('click', async () => {
      try {
        await window.electronAPI.minimize();
      } catch (error) {
        console.error('Erreur minimisation:', error);
      }
    });
    
    document.getElementById('close-btn')?.addEventListener('click', async () => {
      try {
        await window.electronAPI.close();
      } catch (error) {
        console.error('Erreur fermeture:', error);
        // Fallback: pas de window.close() car ça ne marche pas dans Electron
      }
    });
    
  // Auto-refresh configuration (removed manual refresh buttons)
  this.setupAutoRefresh();
    
    // Système d'écoute temps réel des événements de monitoring
    this.setupRealtimeEventListeners();
    
    // Raccourcis clavier
    document.addEventListener('keydown', (e) => {
      // Ctrl+I pour afficher les informations/copyright
      if (e.ctrlKey && e.key === 'i') {
        e.preventDefault();
        this.showAbout();
      }
      // F1 pour l'aide/à propos
      if (e.key === 'F1') {
        e.preventDefault();
        this.showAbout();
      }
    });
    
    // Onglets avec Bootstrap 5
    document.querySelectorAll('[data-bs-toggle="pill"]').forEach(tab => {
      tab.addEventListener('shown.bs.tab', (e) => {
        const target = e.target.getAttribute('data-bs-target');
        this.handleTabChange(target);
        // Si on ouvre l'onglet monitoring, afficher la vue active (par défaut: tableau)
        if (target === '#monitoring') {
          if (window.foldersTree && typeof window.foldersTree.renderCurrentView === 'function') {
            window.foldersTree.renderCurrentView();
          } else if (window.foldersTree && typeof window.foldersTree.renderBoard === 'function') {
            window.foldersTree.renderBoard();
          }
        }
      });
    });

    // Performances personnelles - filtres
    const ppPeriod = document.getElementById('pp-period-mode');
    const ppYear = document.getElementById('pp-year');
    const ppDec = document.getElementById('pp-cat-declarations');
    const ppReg = document.getElementById('pp-cat-reglements');
    const ppMail = document.getElementById('pp-cat-mails');
    const ppAdg = document.getElementById('pp-include-adg');
    const onPPChange = async () => {
      const filters = this.getPersonalPerformanceFilters();
      // Toggle year selector visibility immediately
      const yearWrapEl = document.getElementById('pp-year-wrap');
      if (yearWrapEl) yearWrapEl.style.display = (filters.periodMode === 'year') ? '' : 'none';
      this.persistPersonalPerformanceFilters(filters);
      await this.loadPersonalPerformance?.();
    };
    ppPeriod?.addEventListener('change', onPPChange);
    ppYear?.addEventListener('change', onPPChange);
    ppDec?.addEventListener('change', onPPChange);
    ppReg?.addEventListener('change', onPPChange);
    ppMail?.addEventListener('change', onPPChange);
    ppAdg?.addEventListener('change', onPPChange);
    
    // Emails - Event listeners améliorés
    document.getElementById('refresh-emails')?.addEventListener('click', () => this.loadRecentEmails());
    
    // Filtres d'emails
    document.getElementById('email-search')?.addEventListener('input', (e) => this.filterEmails());
    document.getElementById('folder-filter')?.addEventListener('change', (e) => this.filterEmails());
    document.getElementById('date-filter')?.addEventListener('change', (e) => this.filterEmails());
    document.getElementById('status-filter')?.addEventListener('change', (e) => this.filterEmails());
    
    // Filtres rapides des emails
    document.querySelectorAll('input[name="email-filter"]').forEach(filter => {
      filter.addEventListener('change', (e) => this.applyQuickEmailFilter(e.target.id));
    });
    
    // Monitoring
  // Boutons start/stop retirés: la surveillance est automatique
    document.getElementById('add-folder')?.addEventListener('click', () => this.showAddFolderModal());
    document.getElementById('refresh-folders')?.addEventListener('click', () => this.refreshFoldersDisplay());
    document.getElementById('force-resync')?.addEventListener('click', () => this.forceFullResync());
    // Filtres Monitoring (recherche + catégorie)
  document.getElementById('folder-search')?.addEventListener('input', () => this.updateFolderConfigDisplay());
  document.getElementById('category-filter')?.addEventListener('change', () => this.updateFolderConfigDisplay());
    // Raccourcis chips
    document.querySelectorAll('[data-filter]')?.forEach(chip => {
      chip.addEventListener('click', () => {
        const val = chip.getAttribute('data-filter') || '';
        const select = document.getElementById('category-filter');
        if (select) select.value = val;
        this.updateFolderConfigDisplay();
      });
    });
    
    // Tri dossiers (dashboard)
    document.getElementById('sort-folders-alpha')?.addEventListener('click', () => {
      this.state.ui.foldersSort = 'alpha';
      document.getElementById('sort-folders-alpha')?.classList.add('active');
      document.getElementById('sort-folders-unread')?.classList.remove('active');
      // Re-render sans refetch
      this.updateFoldersStatsDisplay(this._lastFoldersItems || []);
    });
    document.getElementById('sort-folders-unread')?.addEventListener('click', () => {
      this.state.ui.foldersSort = 'unread';
      document.getElementById('sort-folders-unread')?.classList.add('active');
      document.getElementById('sort-folders-alpha')?.classList.remove('active');
      this.updateFoldersStatsDisplay(this._lastFoldersItems || []);
    });

    // Bascule de thème (soleil/lune)
    const themeToggle = document.getElementById('theme-toggle');
    const themeToggleIcon = document.getElementById('theme-toggle-icon');
    const themeToggleLabel = document.getElementById('theme-toggle-label');
    const syncToggleUI = (theme) => {
      if (!themeToggleIcon || !themeToggleLabel) return;
      const isDark = theme === 'dark';
      themeToggleIcon.className = isDark ? 'bi bi-moon' : 'bi bi-sun';
      themeToggleLabel.textContent = isDark ? 'Sombre' : 'Clair';
    };
    if (themeToggle) {
      themeToggle.addEventListener('click', async () => {
        const current = this.state?.ui?.theme === 'dark' ? 'dark' : 'light';
        const next = current === 'dark' ? 'light' : 'dark';
        this.applyTheme(next);
        syncToggleUI(next);
        await this.persistTheme(next);
  // Recolor charts if present
  try { await this.loadPersonalPerformance?.(); } catch(_) {}
      });
    }

    // Paramètres
    document.getElementById('settings-form')?.addEventListener('submit', (e) => this.saveSettings(e));
    document.getElementById('reset-settings')?.addEventListener('click', () => this.resetSettings());

    // Import Activité (.xlsb)
    const pickBtn = document.getElementById('btn-pick-xlsb');
    const previewBtn = document.getElementById('btn-preview-xlsb');
    const importBtn = document.getElementById('btn-import-xlsb');
    const pathInput = document.getElementById('xlsb-path');
    const previewSection = document.getElementById('preview-section');
    const previewBody = document.getElementById('preview-body');
    const previewSummary = document.getElementById('preview-summary');
    const importProgress = document.getElementById('import-progress');
    const importStatus = document.getElementById('import-status');
    const importLogs = document.getElementById('import-logs');

    // Helper: enable/disable buttons
    const setButtonsDisabled = (disabled) => {
      if (previewBtn) previewBtn.disabled = disabled;
      if (importBtn) importBtn.disabled = disabled;
      if (pickBtn) pickBtn.disabled = disabled;
    };

    // Helper: render preview rows
    const renderPreview = (rows) => {
      if (!previewBody) return;
      if (!rows || rows.length === 0) {
        previewBody.innerHTML = `<tr><td colspan="9" class="text-center text-muted">Aucune donnée détectée</td></tr>`;
        return;
      }
      const fmt = new Intl.NumberFormat('fr-FR');
      const html = rows.map(r => {
        const d = r.week_start_date ? new Date(r.week_start_date) : null;
        const dStr = d ? d.toLocaleDateString('fr-FR') : '';
        return `<tr>
          <td>${r.year}</td>
          <td>${r.week_number}</td>
          <td>${dStr}</td>
          <td>${r.category}</td>
          <td class="text-end">${fmt.format(r.recu || 0)}</td>
          <td class="text-end">${fmt.format(r.traite || 0)}</td>
          <td class="text-end">${fmt.format(r.traite_adg || 0)}</td>
          <td class="text-end">${fmt.format(r.stock_debut || 0)}</td>
          <td class="text-end">${fmt.format(r.stock_fin || 0)}</td>
        </tr>`;
      }).join('');
      previewBody.innerHTML = html;
    };

    pickBtn?.addEventListener('click', async () => {
      try {
        const filePath = await window.electronAPI.openXlsbFile();
        if (filePath && pathInput) {
          pathInput.value = filePath;
          previewSection?.classList.add('d-none');
          importProgress?.classList.add('d-none');
        }
      } catch (error) {
        this.showNotification('Erreur', `Sélection du fichier: ${error.message}`, 'danger');
      }
    });

    previewBtn?.addEventListener('click', async () => {
      try {
        const filePath = pathInput?.value?.trim();
        if (!filePath) {
          this.showNotification('Fichier requis', 'Choisissez un fichier .xlsb', 'warning');
          return;
        }
        setButtonsDisabled(true);
        previewSection?.classList.remove('d-none');
        previewSummary && (previewSummary.textContent = 'Analyse en cours...');
        renderPreview([]);
        const res = await window.electronAPI.importActivityPreview(filePath, undefined);
        if (res?.error) throw new Error(res.error);
        renderPreview(res.preview || []);
        if (previewSummary) {
          const count = res?.count || 0;
          const skipped = (res?.skippedWeeks && res.skippedWeeks.length) ? res.skippedWeeks.length : 0;
          previewSummary.textContent = `Année ${res?.year || ''} • ${count} lignes pré-calculées • ${skipped} semaine(s) ignorée(s)`;
        }
      } catch (error) {
        this.showNotification('Erreur', `Aperçu import: ${error.message}`, 'danger');
      } finally {
        setButtonsDisabled(false);
      }
    });

    importBtn?.addEventListener('click', async () => {
      try {
        const filePath = pathInput?.value?.trim();
        if (!filePath) {
          this.showNotification('Fichier requis', 'Choisissez un fichier .xlsb', 'warning');
          return;
        }
        setButtonsDisabled(true);
        importProgress?.classList.remove('d-none');
        if (importStatus) importStatus.textContent = 'Import en cours...';
        if (importLogs) importLogs.textContent = '';
        const res = await window.electronAPI.importActivityRun(filePath, undefined, false);
        if (res?.error) throw new Error(res.error);
        if (importStatus) importStatus.textContent = `Import terminé: ${res.inserted || 0} ligne(s) insérées • ${res.skippedWeeks?.length || 0} semaine(s) ignorée(s)`;
        if (importLogs) importLogs.textContent = `Année: ${res.year || ''}${res.csvPath ? ` • CSV: ${res.csvPath}` : ''}`;
        this.showNotification('Import terminé', 'Les données hebdomadaires ont été enregistrées', 'success');

        // Rafraîchir automatiquement l'affichage hebdomadaire (semaine courante + historique)
        try {
          await this.loadCurrentWeekStats();
          await this.loadWeeklyHistory();
          this.updateLastRefreshTime();
        } catch (refreshErr) {
          console.warn('⚠️ Rafraîchissement hebdomadaire après import: ', refreshErr);
        }
      } catch (error) {
        if (importStatus) importStatus.textContent = 'Erreur lors de l\'import';
        this.showNotification('Erreur', `Import: ${error.message}`, 'danger');
      } finally {
        setButtonsDisabled(false);
      }
    });
  }

  // === CHARGEMENT DES DONNÉES ===
  async loadConfiguration() {
    try {
      console.log('📁 Chargement de la configuration...');
      const result = await window.electronAPI.loadFoldersConfig();
      
      if (result.success) {
        this.state.folderCategories = result.folderCategories || {};
        console.log(`✅ Configuration chargée: ${Object.keys(this.state.folderCategories).length} dossiers configurés`);
        console.log('🔍 DEBUG - Dossiers configurés:', this.state.folderCategories);
        this.updateFolderConfigDisplay();
      } else {
        console.warn('⚠️ Aucune configuration trouvée, utilisation des valeurs par défaut');
        this.state.folderCategories = {};
        this.updateFolderConfigDisplay();
      }
    } catch (error) {
      console.error('❌ Erreur chargement configuration:', error);
      this.state.folderCategories = {};
      this.updateFolderConfigDisplay();
    }
  }

  async loadFoldersConfiguration() {
    try {
      console.log('🔄 Rechargement de la configuration des dossiers...');
      const result = await window.electronAPI.loadFoldersConfig();
      
      // DEBUG: Vérifier les données reçues de l'API
      console.log('🔍 Réponse API loadFoldersConfig:', result);
      
      if (result.success) {
        this.state.folderCategories = result.folderCategories || {};
        console.log(`✅ Configuration rechargée: ${Object.keys(this.state.folderCategories).length} dossiers configurés`);
        console.log('🔍 Détail des dossiers:', this.state.folderCategories);
      } else {
        console.warn('⚠️ Aucune configuration trouvée lors du rechargement');
        console.warn('🔍 Réponse complète:', result);
        this.state.folderCategories = {};
      }
    } catch (error) {
      console.error('❌ Erreur rechargement configuration:', error);
      this.state.folderCategories = {};
    }
  }

  async loadInitialData() {
    console.log('📊 Chargement des données initiales...');
    
    await Promise.allSettled([
      this.loadStats(),
      this.loadRecentEmails(),
      this.loadCategoryStats(),
      this.loadFoldersStats(), // Nouvelle méthode pour stats dossiers
      this.loadVBAMetrics(), // Métriques VBA
      this.checkMonitoringStatus(),
      this.loadSettings()
    ]);
  }

  async checkConnection() {
    try {
      const status = await window.electronAPI.outlookStatus();
      this.state.isConnected = status.status;
      this.updateConnectionStatus(status);
    } catch (error) {
      console.error('❌ Erreur vérification connexion:', error);
      this.state.isConnected = false;
      this.updateConnectionStatus({ status: false, error: error.message });
    }
  }

  async loadStats() {
    try {
      const stats = await window.electronAPI.getStatsSummary();
      this.state.stats = { ...this.state.stats, ...stats };
      this.state.lastUpdate = new Date();
      this.updateStatsDisplay();
    } catch (error) {
      console.error('❌ Erreur chargement stats:', error);
    }
  }

  // Méthode supprimée - utilise maintenant la nouvelle version loadRecentEmails() plus bas dans le fichier

  async loadCategoryStats() {
    try {
      const data = await window.electronAPI.getStatsByCategory();
      this.updateCategoryDisplay(data.categories || {});
    } catch (error) {
      console.error('❌ Erreur chargement stats catégories:', error);
    }
  }

  // === MÉTRIQUES VBA ===
  async loadVBAMetrics() {
    try {
      console.log('📊 Chargement métriques VBA...');
      
      // Charger le résumé des métriques VBA
      const vbaMetrics = await window.electronAPI.getVBAMetricsSummary();
      if (vbaMetrics) {
        this.updateVBAMetricsDisplay(vbaMetrics);
      }
      
      // Charger la distribution des dossiers
      const folderDistribution = await window.electronAPI.getVBAFolderDistribution();
      if (folderDistribution) {
        this.updateVBAFolderDistribution(folderDistribution);
      }
      
      // Charger l'évolution hebdomadaire
      const weeklyEvolution = await window.electronAPI.getVBAWeeklyEvolution();
      if (weeklyEvolution) {
        this.updateVBAWeeklyEvolution(weeklyEvolution);
      }
      
    } catch (error) {
      console.error('❌ Erreur chargement métriques VBA:', error);
    }
  }

  // Nouvelle méthode pour charger et afficher les statistiques par dossier
  async loadFoldersStats(opts = {}) {
    try {
      console.log('📁 Chargement statistiques dossiers...');
      
      // Récupérer l'arborescence complète puis réduire aux dossiers surveillés (plus pertinent et compact)
      const forceFoldersTree = !!opts.forceFoldersTree;
      const foldersTreePromise = forceFoldersTree
        ? window.electronAPI.invoke('api-folders-tree', { force: true })
        : window.electronAPI.getFoldersTree();

      const [foldersData, dbFolderStatsResp] = await Promise.all([
        foldersTreePromise,
        window.electronAPI.getFolderStats?.({ force: true }).catch(() => null)
      ]);

      // Index des stats DB par chemin (normalisé)
      const dbStats = (dbFolderStatsResp && (dbFolderStatsResp.stats || dbFolderStatsResp)) || [];
      const toKey = (s) => (s || '').toLowerCase();
      const statsByPath = new Map();
      if (Array.isArray(dbStats)) {
        for (const row of dbStats) {
          const p = toKey(row.path || row.folder_name);
          if (p) statsByPath.set(p, row);
        }
      }

      // Conserver la vue "brute" (toute l'arborescence) + index DB pour les calculs dashboard (ex: par client)
      // Sans impacter le rendu de l'onglet Monitoring (qui reste limité aux dossiers surveillés).
      this._lastFoldersTreeRaw = (foldersData && Array.isArray(foldersData.folders)) ? foldersData.folders : [];
      this._lastDbFolderStatsRaw = Array.isArray(dbStats) ? dbStats : [];
      this._lastDbFolderStatsByPath = statsByPath;

      const monitoredConfigs = this.state.folderCategories || {};
      const monitoredPaths = Object.keys(monitoredConfigs);

      let items = [];

      if (foldersData && Array.isArray(foldersData.folders) && foldersData.folders.length > 0) {
        // IMPORTANT: api-folders-tree renvoie déjà uniquement les dossiers monitorés.
        // Un filtrage supplémentaire sur monitoredPaths peut exclure des dossiers (mismatch de chemin/prefix).
        const lc = (s) => (s || '').toLowerCase();
        items = foldersData.folders.map(f => {
          const fPath = f.path || f.FolderPath || f.Path || '';

          // Optionnel: tenter de récupérer la config UI correspondante pour compléter (nom/catégorie)
          const matchKey = monitoredPaths.find(p => lc(fPath) === lc(p) || lc(fPath).endsWith(lc(p))) || fPath;
          const cfg = monitoredConfigs[matchKey] || {};
          const name = f.name || cfg.name || this.extractFolderName(fPath) || 'Dossier';
          const category = f.category || cfg.category || '';

          // Comptages robustes (valeurs natives Outlook > fallback BDD)
          const totalRaw =
            f.TotalItems ?? f.ItemsCount ?? f.TotalItemCount ?? f.ItemCount ?? f.Count ?? f.total ?? f.totalEmails ?? f.emailCount;
          let unreadRaw =
            f.UnReadItemCount ?? f.UnreadItemCount ?? f.UnreadCount ?? f.unreadCount ?? f.UnreadItems ?? f.unreadItems ?? f.unread ?? f.unreadEmails;

          // Fusionner avec la BDD si compteurs incomplets
          const db = statsByPath.get(toKey(fPath))
            || statsByPath.get(toKey(this.extractFolderName(fPath)))
            || [...statsByPath.entries()].find(([k]) => toKey(fPath).endsWith(k))?.[1];
          const dbUnread = db?.unreadCount ?? db?.unread_count;
          const dbTotal = db?.emailCount ?? db?.email_count;

          const total = Number.isFinite(+totalRaw) ? +totalRaw : (Number.isFinite(+dbTotal) ? +dbTotal : null);
          const unread = Number.isFinite(+unreadRaw) ? +unreadRaw : (Number.isFinite(+dbUnread) ? +dbUnread : null);

          const counts = this._computeFolderCounts({ total, unread, path: fPath, name });
          return { name, path: fPath, category, total: counts.total, unread: counts.unread };
        });
      }

      // Fallback: s'il n'y a pas d'arborescence, construire depuis la configuration surveillée
      if ((!items || items.length === 0) && monitoredPaths.length > 0) {
        items = monitoredPaths.map(p => {
          const cfg = monitoredConfigs[p] || {};
          const name = cfg.name || this.extractFolderName(p) || 'Dossier';
          // Essayer de récupérer d'abord depuis la BDD
          const db = statsByPath.get(toKey(p))
            || statsByPath.get(toKey(name))
            || [...statsByPath.entries()].find(([k]) => toKey(p).endsWith(k))?.[1];
          const dbTotal = db?.emailCount ?? db?.email_count;
          const dbUnread = db?.unreadCount ?? db?.unread_count;
          let total = Number.isFinite(+dbTotal) ? +dbTotal : null;
          let unread = Number.isFinite(+dbUnread) ? +dbUnread : null;
          const counts = this._computeFolderCounts({ path: p, name, total, unread });
          return { name, path: p, category: cfg.category || '', total: counts.total, unread: counts.unread };
        });
      }

  // Conserver les items pour re-render local
  this._lastFoldersItems = items || [];
  // Mise à jour de l'affichage (compact)
  this.updateFoldersStatsDisplay(this._lastFoldersItems);

      // Mettre à jour le badge total (affichage courant)
      const totalFoldersEl = document.getElementById('total-folders-badge');
      if (totalFoldersEl) {
        const n = items?.length || 0;
        totalFoldersEl.textContent = `${n} dossier${n > 1 ? 's' : ''}`;
      }
      
    } catch (error) {
      console.error('❌ Erreur chargement stats dossiers:', error);
    }
  }

  // Nouvelle méthode pour afficher les statistiques des dossiers
  updateFoldersStatsDisplay(folders) {
    // Dashboard : grille classique
    const grid = document.getElementById('folders-stats-grid');
    if (grid) {
      if (!folders || folders.length === 0) {
        grid.innerHTML = `
          <div class="col-12 text-center text-muted py-4">
            <i class="bi bi-folder-x fs-1 mb-3"></i>
            <p>Aucun dossier configuré pour le monitoring</p>
            <small>Ajoutez des dossiers dans l'onglet Configuration</small>
          </div>
        `;
      } else {
        // Préparer données (restants = non lus)
        const safeFolders = (folders || []).map(f => ({
          name: f.name || 'Dossier',
          path: f.path || '',
          category: f.category || '',
          total: Number.isFinite(+f.total) ? +f.total : 0,
          unread: Number.isFinite(+f.unread) ? +f.unread : 0
        }));

        const normalizeCategoryKey = (category) => {
          const v = String(category || '').toLowerCase();
          if (v.includes('déclar') || v.includes('declar') || v === 'declarations') return 'declarations';
          if (v.includes('règle') || v.includes('regle') || v.includes('reglement') || v === 'reglements') return 'reglements';
          if (v.includes('mail')) return 'mails';
          return 'mails';
        };

        const categoryMeta = {
          declarations: { label: 'Déclarations', icon: 'bi-file-earmark-text', color: 'text-danger' },
          mails: { label: 'Mails simples', icon: 'bi-envelope', color: 'text-info' },
          reglements: { label: 'Règlements', icon: 'bi-credit-card', color: 'text-success' }
        };

        // ------------------------------------------------------------------
        // Regroupement par client :
        // Prendre en compte les ENFANTS des dossiers racine :
        //  - 1-Déclaration
        //  - 2-courriers
        //  - 3-règlements
        // (et ignorer ces dossiers racine eux-mêmes)
        // ------------------------------------------------------------------

        const normalizeSeg = (s) => String(s || '')
          .trim()
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '');

        const rootToCategoryKey = (seg) => {
          const v = normalizeSeg(seg);
          if (!v) return null;
          if ((v.startsWith('1-') || v.startsWith('1 ')) && v.includes('declar')) return 'declarations';
          if ((v.startsWith('2-') || v.startsWith('2 ')) && (v.includes('courrier') || v.includes('courriers'))) return 'mails';
          if ((v.startsWith('3-') || v.startsWith('3 ')) && v.includes('regl')) return 'reglements';
          return null;
        };

        const tryExtractClientFromPath = (folderPath) => {
          if (!folderPath) return null;
          const parts = String(folderPath).split(/\\|\//).map(p => p.trim()).filter(Boolean);
          for (let i = 0; i < parts.length; i++) {
            const catKey = rootToCategoryKey(parts[i]);
            if (!catKey) continue;
            const client = parts[i + 1];
            // Si on est sur le dossier racine (pas d'enfant), on ignore
            if (!client) return null;
            return { clientName: client, categoryKey: catKey };
          }
          return null;
        };

        const getCountsFromTreeNode = (node) => {
          const fPath = node?.path || node?.FolderPath || node?.Path || '';
          const name = node?.name || node?.Name || this.extractFolderName(fPath) || 'Dossier';

          const totalRaw =
            node?.TotalItems ?? node?.ItemsCount ?? node?.TotalItemCount ?? node?.ItemCount ?? node?.Count ?? node?.total ?? node?.totalEmails ?? node?.emailCount;
          let unreadRaw =
            node?.UnReadItemCount ?? node?.UnreadItemCount ?? node?.UnreadCount ?? node?.unreadCount ?? node?.UnreadItems ?? node?.unreadItems ?? node?.unread ?? node?.unreadEmails;

          // Fallback DB unread si absent côté Outlook
          const statsByPath = this._lastDbFolderStatsByPath;
          const toKey = (s) => (s || '').toLowerCase();
          if (!(Number.isFinite(+unreadRaw)) && statsByPath && typeof statsByPath.get === 'function') {
            const db = statsByPath.get(toKey(fPath))
              || statsByPath.get(toKey(this.extractFolderName(fPath)))
              || [...statsByPath.entries()].find(([k]) => toKey(fPath).endsWith(k))?.[1];
            const dbUnread = db?.unreadCount ?? db?.unread_count;
            if (Number.isFinite(+dbUnread)) unreadRaw = +dbUnread;
          }

          const counts = this._computeFolderCounts({ total: totalRaw, unread: unreadRaw, path: fPath, name });
          return { name, path: fPath, total: counts.total, unread: counts.unread };
        };

        // Source: la répartition "par client" doit couvrir TOUS les dossiers (pas seulement ceux monitorés).
        // On privilégie donc les stats BDD (emails.groupBy folder_name) qui couvrent l'ensemble des chemins.
        const clientSource = (Array.isArray(this._lastDbFolderStatsRaw) && this._lastDbFolderStatsRaw.length)
          ? this._lastDbFolderStatsRaw
          : ((Array.isArray(this._lastFoldersTreeRaw) && this._lastFoldersTreeRaw.length)
            ? this._lastFoldersTreeRaw
            : safeFolders);

        const byClient = new Map();
        const totalsByCategory = { declarations: 0, mails: 0, reglements: 0 };

        for (const raw of clientSource) {
          const fPath = raw?.path || raw?.FolderPath || raw?.Path || raw?.path || '';
          const extracted = tryExtractClientFromPath(fPath);
          if (!extracted) {
            continue;
          }

          const { clientName, categoryKey } = extracted;
          const key = String(clientName || 'Client').trim() || 'Client';

          let counts;
          if (raw && raw.unread !== undefined && raw.total !== undefined) {
            counts = {
              unread: Number.isFinite(+raw.unread) ? +raw.unread : 0,
              total: Number.isFinite(+raw.total) ? +raw.total : 0
            };
          } else if (raw && (raw.unreadCount !== undefined || raw.unread_count !== undefined || raw.emailCount !== undefined || raw.email_count !== undefined)) {
            const unread = raw.unreadCount ?? raw.unread_count;
            const total = raw.emailCount ?? raw.email_count;
            counts = {
              unread: Number.isFinite(+unread) ? +unread : 0,
              total: Number.isFinite(+total) ? +total : 0
            };
          } else {
            counts = getCountsFromTreeNode(raw);
          }

          // On ignore le dossier racine catégorie lui-même (par sécurité)
          // => extrait ne renvoie pas de client si root, mais on garde ce garde-fou.
          if (!counts || !Number.isFinite(+counts.unread)) continue;

          if (!byClient.has(key)) byClient.set(key, { declarations: 0, mails: 0, reglements: 0 });
          const bucket = byClient.get(key);
          bucket[categoryKey] += (counts.unread || 0);
          totalsByCategory[categoryKey] += (counts.unread || 0);
        }

        // Fallback: si rien n'a été extrait (structure différente), reprendre l'ancienne heuristique
        if (byClient.size === 0) {
          for (const f of safeFolders) {
            const clientName = this.extractClientNameFromFolder?.(f) || f.name;
            const key = String(clientName || 'Client').trim() || 'Client';
            if (!byClient.has(key)) byClient.set(key, { declarations: 0, mails: 0, reglements: 0 });
            const bucket = byClient.get(key);
            const k = normalizeCategoryKey(f.category);
            bucket[k] += (f.unread || 0);
            totalsByCategory[k] += (f.unread || 0);
          }
        }

        const clients = Array.from(byClient.entries()).map(([client, counts]) => {
          const total = (counts.declarations || 0) + (counts.mails || 0) + (counts.reglements || 0);
          return { client, ...counts, total };
        });

        const sortMode = this.state?.ui?.foldersSort || 'alpha';
        clients.sort((a, b) => {
          if (sortMode === 'unread') {
            const d = (b.total || 0) - (a.total || 0);
            if (d !== 0) return d;
          }
          return String(a.client).localeCompare(String(b.client), 'fr', { sensitivity: 'base', numeric: true });
        });

        const fmt = new Intl.NumberFormat('fr-FR');
        const totalAll = clients.reduce((s, c) => s + Number(c.total || 0), 0);

        const summaryHtml = `
          <div class="col-12">
            <div class="d-flex flex-wrap gap-2 align-items-center">
              <span class="small text-muted me-2">Restants:</span>
              ${['declarations','mails','reglements'].map(k => {
                const m = categoryMeta[k];
                const val = totalsByCategory[k] || 0;
                return `
                  <div class="d-flex align-items-center gap-2 px-2 py-1 rounded border bg-surface">
                    <span class="${m.color}"><i class="bi ${m.icon}"></i></span>
                    <span class="small">${m.label}</span>
                    <span class="badge bg-warning text-dark" title="Non lus">${fmt.format(val)}</span>
                  </div>
                `;
              }).join('')}
              <div class="ms-auto small text-muted">Total: <span class="fw-semibold">${fmt.format(totalAll)}</span></div>
            </div>
          </div>
        `;

        const tableHtml = `
          <div class="col-12">
            <div class="table-responsive">
              <table class="table table-sm align-middle mb-0">
                <thead class="table-light">
                  <tr>
                    <th>Client</th>
                    <th class="text-end">Déclarations</th>
                    <th class="text-end">Mails</th>
                    <th class="text-end">Règlements</th>
                    <th class="text-end">Total</th>
                  </tr>
                </thead>
                <tbody>
                  ${clients.length ? clients.map(c => {
                    const cell = (n) => (n > 0 ? `<span class="badge bg-warning text-dark">${fmt.format(n)}</span>` : `<span class="text-muted">0</span>`);
                    return `
                      <tr>
                        <td class="fw-semibold">${this.escapeHtml(c.client)}</td>
                        <td class="text-end">${cell(c.declarations || 0)}</td>
                        <td class="text-end">${cell(c.mails || 0)}</td>
                        <td class="text-end">${cell(c.reglements || 0)}</td>
                        <td class="text-end fw-semibold">${fmt.format(c.total || 0)}</td>
                      </tr>
                    `;
                  }).join('') : `<tr><td colspan="5" class="text-center text-muted py-3">Aucune donnée</td></tr>`}
                </tbody>
              </table>
            </div>
            <div class="mt-2 small text-muted">
              Tri: ${sortMode === 'unread' ? 'non lus (total) décroissant' : 'A→Z'}
            </div>
          </div>
        `;

        grid.innerHTML = summaryHtml + tableHtml;
      }
    }
    // Onglet monitoring : arborescence
    const tree = document.getElementById('monitoring-folder-tree');
    if (tree) {
      if (!folders || folders.length === 0) {
        tree.innerHTML = `
          <div class="text-center text-muted py-4">
            <i class="bi bi-folder-x fs-1 mb-3"></i>
            <p>Aucun dossier configuré pour le monitoring</p>
            <small>Ajoutez des dossiers dans l'onglet Configuration</small>
          </div>
        `;
      } else {
        let treeHtml = '';
        for (const folder of folders) {
          treeHtml += this.createFolderTree(folder, 0);
        }
        tree.innerHTML = treeHtml;
      }
    }
  }

  // Méthodes utilitaires pour les catégories
  getCategoryColor(category) {
    const colors = {
      'Déclarations': 'text-danger',
      'Règlements': 'text-success', 
      'mails_simples': 'text-info',
      'Mails simples': 'text-info',
      'test': 'text-secondary'
    };
    return colors[category] || 'text-muted';
  }

  getCategoryIcon(category) {
    const icons = {
      'Déclarations': 'bi-file-earmark-text',
      'Règlements': 'bi-credit-card',
      'mails_simples': 'bi-envelope',
      'Mails simples': 'bi-envelope',
      'test': 'bi-folder'
    };
    return icons[category] || 'bi-folder';
  }

  // Déduire un "client" à partir d'un dossier (heuristique légère)
  extractClientNameFromFolder(folder) {
    const name = String(folder?.name || '').trim();
    const category = String(folder?.category || '').trim();
    const path = String(folder?.path || '').trim();

    const knownCats = new Set(['déclarations', 'declarations', 'règlements', 'reglements', 'mails simples', 'mails_simples', 'mails', 'mail']);
    const looksLikeCategoryName = (s) => {
      const v = String(s || '').toLowerCase();
      if (!v) return false;
      if (knownCats.has(v)) return true;
      if (v.includes('déclar') || v.includes('declar')) return true;
      if (v.includes('règle') || v.includes('regle') || v.includes('reglement')) return true;
      if (v.includes('mail')) return true;
      return false;
    };

    // Si le nom du dossier ressemble à une catégorie, on prend le parent dans le chemin
    if (path) {
      const parts = path.split(/\\|\//).map(p => p.trim()).filter(Boolean);
      if (parts.length >= 2) {
        const last = parts[parts.length - 1];
        const prev = parts[parts.length - 2];
        if (looksLikeCategoryName(last) || looksLikeCategoryName(name) || looksLikeCategoryName(category)) {
          return prev || name || last;
        }
      }
    }

    // Fallback: nom du dossier
    return name || 'Client';
  }

  // Calcule les compteurs d'un dossier (total / non lus) avec fallback sur les emails récents
  _computeFolderCounts({ total, unread, path, name }) {
    let t = Number.isFinite(+total) ? +total : null;
    let u = Number.isFinite(+unread) ? +unread : null;

    // If both are known, return immediately
    if (t !== null && u !== null) return { total: t, unread: u };

    // Fallback: derive missing piece from recent emails, but keep any known value
    const emails = Array.isArray(this.state.recentEmails) ? this.state.recentEmails : [];
    const toLc = (s) => (s || '').toLowerCase();
    const pathLc = toLc(path);
    const nameLc = toLc(name);
    const byFolder = emails.filter(e => {
      const eName = toLc(e.folder_name || e.Folder || '');
      const ePath = toLc(e.folder_path || e.FolderPath || '');
      return (ePath && (ePath === pathLc || ePath.endsWith(pathLc))) || (eName && eName === nameLc);
    });

    if (t === null) t = byFolder.length;
    if (u === null) u = byFolder.filter(e => (!e.is_read && e.is_read !== undefined) || e.UnRead === true).length;
    return { total: t || 0, unread: u || 0 };
  }

  async checkMonitoringStatus() {
    try {
      const serviceStatus = await window.electronAPI.getMonitoringStatus?.();
      const isActive = serviceStatus?.active === true;
      let foldersCount = typeof serviceStatus?.foldersMonitored === 'number' ? serviceStatus.foldersMonitored : 0;
      if (foldersCount === 0 && window.electronAPI.getFoldersTree) {
        try {
          const foldersData = await window.electronAPI.getFoldersTree();
          foldersCount = foldersData?.folders?.length || 0;
        } catch (_) {}
      }
      this.state.isMonitoring = isActive;
      this.state.foldersMonitoredCount = foldersCount;
  // Mettre à jour l'indicateur de statut en mode réussi
  this.updateMonitoringStatus(isActive);
    } catch (error) {
      console.error('❌ Erreur vérification statut monitoring:', error);
      const fallbackCount = Object.keys(this.state.folderCategories || {}).length;
      const fallbackActive = fallbackCount > 0;
      this.state.isMonitoring = fallbackActive;
      this.state.foldersMonitoredCount = fallbackCount;
      this.updateMonitoringStatus(fallbackActive);
    }
  }

  // === MONITORING ===
  // Les méthodes start/stop manuelles ne sont plus nécessaires (surveillance auto)

  // === CONFIGURATION DES DOSSIERS ===
  async showAddFolderModal() {
    // Ouvre immédiatement le modal avec un état de chargement, puis remplit les boîtes mail
    try {
      const ok = await this.showDefaultFolderTreeModal();
      if (!ok) {
        // Fallback ultime: ajout manuel
        this.showNotification('Aucune boîte mail', 'Impossible de charger l\'arborescence Outlook. Vous pouvez ajouter un dossier manuellement.', 'warning');
        this.showManualAddFolderModal();
      }
    } catch (error) {
      console.error('❌ Erreur ouverture modal d\'ajout:', error);
      this.showNotification('Erreur', 'Impossible d\'ouvrir le sélecteur de dossier', 'danger');
    }
  }

  // Fallback prioritaire: tenter d'afficher l'arbo du Store par défaut
  async showDefaultFolderTreeModal() {
    try {
      // Nettoyer un ancien modal éventuel
      const existingModal = document.getElementById('folderModal');
      if (existingModal) existingModal.remove();

      const modalHtml = `
        <div class="modal fade" id="folderModal" tabindex="-1">
          <div class="modal-dialog modal-lg">
            <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title">Ajouter un dossier à surveiller</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
              </div>
              <div class="modal-body">
                <div id="folder-modal-alert" class="alert alert-danger d-none" role="alert"></div>
                <form id="add-folder-form">
                  <div class="mb-3">
                    <label class="form-label">Chemin du dossier racine</label>
                    <div class="input-group">
                      <input type="text" class="form-control" id="root-folder-input" placeholder="Ex: testboitepartagee\\Boîte de réception" autocomplete="off" />
                      <button class="btn btn-outline-primary" type="button" id="load-root-tree">
                        <span class="load-text">Explorer</span>
                        <span class="spinner-border spinner-border-sm d-none ms-2" role="status" aria-hidden="true"></span>
                      </button>
                    </div>
                    <div class="form-text" id="root-folder-hint">Indiquez le chemin complet tel qu'affiché dans Outlook (nom de la boîte partagée inclus).</div>
                  </div>
                  <div class="mb-3">
                    <label class="form-label">Dossier</label>
                    <div class="border rounded p-3 bg-surface" style="max-height: 300px; overflow-y: auto;">
                      <div id="folder-tree" class="folder-tree">
                        <div class="text-muted"><i class="bi bi-info-circle me-2"></i>Saisissez un chemin et cliquez sur « Explorer »</div>
                      </div>
                    </div>
                    <input type="hidden" id="selected-folder-path" required>
                    <div class="form-text mt-2">
                      <small>Cliquez sur un dossier pour le sélectionner</small>
                    </div>
                  </div>
                  <div class="mb-3">
                    <label class="form-label">Dossiers sélectionnés</label>
                    <div class="border rounded bg-surface">
                      <table class="table table-sm align-middle mb-0">
                        <thead class="table-light">
                          <tr>
                            <th style="width: 18%;">Dossier</th>
                            <th>Chemin complet</th>
                            <th style="width: 190px;">Catégorie</th>
                            <th style="width: 52px;"></th>
                          </tr>
                        </thead>
                        <tbody id="selected-folders-body">
                          <tr id="selected-folders-empty">
                            <td colspan="4" class="text-muted small py-2 ps-2">Cochez un ou plusieurs dossiers dans l'arborescence.</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <div class="form-text">Chaque dossier peut être affecté à une catégorie différente.</div>
                  </div>
                  <div class="mb-3">
                    <label class="form-label">Catégorie</label>
                    <select class="form-control" id="category-input" required>
                      <option value="">-- Sélectionner une catégorie --</option>
                      <option value="Déclarations">Déclarations</option>
                      <option value="Règlements">Règlements</option>
                      <option value="Mails simples">Mails simples</option>
                    </select>
                    <div class="form-text">Choisissez la catégorie appropriée pour classer les emails de ce dossier</div>
                  </div>
                </form>
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Annuler</button>
                <button type="button" class="btn btn-primary" id="save-folder-config">
                  <span class="save-text">Ajouter</span>
                  <span class="spinner-border spinner-border-sm d-none ms-2" role="status" aria-hidden="true"></span>
                </button>
              </div>
            </div>
          </div>
        </div>`;

      document.body.insertAdjacentHTML('beforeend', modalHtml);
      const modal = new bootstrap.Modal(document.getElementById('folderModal'));
      modal.show();
      // Préselectionner une catégorie par défaut pour éviter l'état vide
      try {
        const catEl = document.getElementById('category-input');
        if (catEl && !catEl.value) catEl.value = 'Mails simples';
        catEl?.classList.remove('is-invalid');
      } catch(_) {}

      const folderTree = document.getElementById('folder-tree');
      const rootInput = document.getElementById('root-folder-input');
      const loadButton = document.getElementById('load-root-tree');
      const alertBox = document.getElementById('folder-modal-alert');

      // Gestion multi-sélection pour l'ajout en masse
      const selectionCtrl = this.createFolderSelectionController({
        defaultCategorySelectId: 'category-input',
        selectedBodyId: 'selected-folders-body',
        hiddenInputId: 'selected-folder-path'
      });
      this._folderSelectionController = selectionCtrl;

      if (rootInput && this.state?.lastRootFolderPath) {
        rootInput.value = this.state.lastRootFolderPath;
      }

      const setLoadingState = (loading) => {
        try {
          const spinner = loadButton?.querySelector('.spinner-border');
          const label = loadButton?.querySelector('.load-text');
          if (loading) {
            spinner?.classList.remove('d-none');
            if (label) label.textContent = 'Chargement';
            if (loadButton) loadButton.disabled = true;
          } else {
            spinner?.classList.add('d-none');
            if (label) label.textContent = 'Explorer';
            if (loadButton) loadButton.disabled = false;
          }
        } catch (_) {}
      };

      const resetSelection = () => {
        this._folderSelectionController?.clear?.();
        document.getElementById('selected-folder-path').value = '';
      };

      const renderTree = (rootNode, meta) => {
        if (!rootNode) {
          folderTree.innerHTML = '<div class="text-warning"><i class="bi bi-info-circle me-2"></i>Aucun dossier trouvé pour ce chemin</div>';
          return;
        }
        folderTree.dataset.mailbox = meta?.storeName || '';
        folderTree.dataset.mailboxDisplay = meta?.storeName || '';
        folderTree.dataset.smtp = meta?.smtp || '';
        folderTree.dataset.storeId = meta?.storeId || '';
        folderTree.dataset.mode = 'root-path';
        const html = this.createFolderTree(rootNode, 0, true);
        folderTree.innerHTML = html;
        this.initializeFolderTreeEvents({
          onToggle: (payload) => selectionCtrl?.toggle(payload),
          isSelected: (path) => selectionCtrl?.isSelected(path),
          highlightSelected: true,
          withCheckbox: true
        });
        selectionCtrl?.syncCheckboxes?.();
      };

      const loadTreeFromRoot = async () => {
        const rootPath = rootInput?.value?.trim() || '';
        if (alertBox) {
          alertBox.classList.add('d-none');
          alertBox.textContent = '';
        }
        rootInput?.classList.remove('is-invalid');
        resetSelection();

        if (!rootPath) {
          rootInput?.classList.add('is-invalid');
          return;
        }

        folderTree.innerHTML = '<div class="text-muted"><i class="bi bi-hourglass-split me-2"></i>Chargement de l\'arborescence…</div>';
        setLoadingState(true);

        try {
          // Limiter la profondeur pour éviter les gels Outlook sur les grandes boîtes
          const result = await window.electronAPI.getFolderTreeFromRoot(rootPath, 4);
          if (!result || result.success === false || !result.root) {
            const message = (result && result.error) ? result.error : 'Impossible de récupérer cette arborescence';
            folderTree.innerHTML = `<div class="text-danger"><i class="bi bi-exclamation-triangle me-2"></i>${this.escapeHtml(message)}</div>`;
            if (alertBox) {
              alertBox.textContent = message;
              alertBox.classList.remove('d-none');
            }
            return;
          }

          this.state = this.state || {};
          this.state.lastRootFolderPath = rootPath;

          renderTree(result.root, {
            storeName: result.store?.name || '',
            storeId: result.store?.id || '',
            smtp: result.store?.smtp || ''
          });

          // Par défaut, sélectionner le dossier racine pour guider l'utilisateur
          try {
            const rootSelectable = document.querySelector('#folder-tree .folder-selectable');
            if (rootSelectable) {
              rootSelectable.click();
            }
          } catch (_) {}
        } catch (error) {
          console.error('❌ Erreur chargement arbo racine:', error);
          const message = error?.message || 'Erreur de chargement';
          folderTree.innerHTML = `<div class="text-danger"><i class="bi bi-exclamation-triangle me-2"></i>${this.escapeHtml(message)}</div>`;
          if (alertBox) {
            alertBox.textContent = message;
            alertBox.classList.remove('d-none');
          }
        } finally {
          setLoadingState(false);
        }
      };

      loadButton?.addEventListener('click', loadTreeFromRoot);
      rootInput?.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          loadTreeFromRoot();
        }
      });

      // Focus sur le champ à l'ouverture
      setTimeout(() => rootInput?.focus(), 150);

      if (rootInput?.value) {
        loadTreeFromRoot();
      }

      // Sauvegarde
      document.getElementById('save-folder-config').addEventListener('click', () => {
        this.saveFolderConfiguration(modal);
      });

      return true;
    } catch (error) {
      console.error('❌ showDefaultFolderTreeModal:', error);
      return false;
    }
  }

  // Fallback lorsqu'aucune boîte Outlook n'est détectée: ajout manuel
  showManualAddFolderModal() {
    const existingModal = document.getElementById('manualFolderModal');
    if (existingModal) existingModal.remove();

    const modalHtml = `
      <div class="modal fade" id="manualFolderModal" tabindex="-1">
        <div class="modal-dialog">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title">Ajouter un dossier (manuel)</h5>
              <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
              <form id="manual-folder-form">
                <div class="mb-3">
                  <label class="form-label">Nom du dossier</label>
                  <input type="text" class="form-control" id="manual-folder-name" placeholder="Ex: Déclarations 2025" required />
                </div>
                <div class="mb-3">
                  <label class="form-label">Chemin complet</label>
                  <input type="text" class="form-control" id="manual-folder-path" placeholder="Ex: Boîte de réception\\Déclarations" required />
                  <div class="form-text">Utilisez le chemin tel qu'il apparaît dans Outlook</div>
                </div>
                <div class="mb-3">
                  <label class="form-label">Catégorie</label>
                  <select class="form-select" id="manual-folder-category" required>
                    <option value="">-- Sélectionner une catégorie --</option>
                    <option value="Déclarations">Déclarations</option>
                    <option value="Règlements">Règlements</option>
                    <option value="Mails simples">Mails simples</option>
                  </select>
                </div>
              </form>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Annuler</button>
              <button type="button" class="btn btn-primary" id="manual-folder-save">
                <span class="save-text">Ajouter</span>
                <span class="spinner-border spinner-border-sm d-none ms-2" role="status" aria-hidden="true"></span>
              </button>
            </div>
          </div>
        </div>
      </div>`;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    const modal = new bootstrap.Modal(document.getElementById('manualFolderModal'));
    modal.show();

    // Préselectionner une catégorie par défaut pour éviter l'état vide
    try {
      const catEl = document.getElementById('manual-folder-category');
      if (catEl && !catEl.value) catEl.value = 'Mails simples';
      catEl?.classList.remove('is-invalid');
    } catch(_) {}

    document.getElementById('manual-folder-save').addEventListener('click', async (ev) => {
      const btn = ev.currentTarget;
      const spinner = btn.querySelector('.spinner-border');
      const text = btn.querySelector('.save-text');
      const name = document.getElementById('manual-folder-name').value.trim();
      const path = document.getElementById('manual-folder-path').value.trim();
      const category = document.getElementById('manual-folder-category').value.trim();
      if (!name || !path || !category) {
        this.showNotification('Champs requis', 'Merci de renseigner le nom, le chemin, et la catégorie', 'warning');
        return;
      }
      try { spinner?.classList.remove('d-none'); text && (text.textContent = 'Ajout...'); btn.disabled = true; } catch(_) {}
      try {
        const res = await window.electronAPI.addFolderToMonitoring({ folderPath: path, category });
        if (res && res.success) {
          // Rafraîchir configuration depuis BDD
          await this.loadFoldersConfiguration();
          this.updateFolderConfigDisplay();
          const el = document.getElementById('manualFolderModal');
          const instance = el ? bootstrap.Modal.getInstance(el) : null;
          if (instance) instance.hide();
          const count = res.count || 1;
          this.showNotification('Dossier(s) ajouté(s)', `${count} élément(s) catégorisé(s) en "${category}"`, 'success');
        } else {
          this.showNotification('Erreur', (res && res.error) || 'Échec de l\'ajout', 'danger');
        }
      } catch (e) {
        console.error('❌ Erreur ajout manuel:', e);
        this.showNotification('Erreur', e.message, 'danger');
      } finally { try { spinner?.classList.add('d-none'); text && (text.textContent = 'Ajouter'); btn.disabled = false; } catch(_) {} }
    });
  }

  createFolderSelectionModal(mailboxes) {
    // Supprimer le modal existant s'il y en a un
    const existingModal = document.getElementById('folderModal');
    if (existingModal) {
      existingModal.remove();
    }

    const modalHtml = `
      <div class="modal fade" id="folderModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title">Ajouter un dossier à surveiller</h5>
              <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
              <div id="folder-modal-alert" class="alert alert-danger d-none" role="alert"></div>
              <form id="add-folder-form">
                <div class="mb-3">
                  <label class="form-label">Boîte mail</label>
                  <select class="form-select" id="mailbox-select" required>
                    <option value="">Sélectionnez une boîte mail</option>
                    ${mailboxes.map(mb => {
                      const label = mb.SmtpAddress ? `${mb.Name} (${mb.SmtpAddress})` : mb.Name;
                      return `<option value="${this.escapeHtml(mb.StoreID)}" data-smtp="${this.escapeHtml(mb.SmtpAddress || '')}" data-name="${this.escapeHtml(mb.Name || '')}">${this.escapeHtml(label)}</option>`;
                    }).join('')}
                  </select>
                </div>
                <div class="mb-3">
                  <label class="form-label">Dossier</label>
                  <div class="border rounded p-3 bg-surface" style="max-height: 300px; overflow-y: auto;">
                    <div id="folder-tree" class="folder-tree">
                      <div class="text-muted">Sélectionnez d'abord une boîte mail</div>
                    </div>
                  </div>
                  <input type="hidden" id="selected-folder-path" required>
                  <div class="form-text mt-2">
                    <small>Cliquez sur un dossier pour le sélectionner</small>
                  </div>
                </div>
                <div class="mb-3">
                  <label class="form-label">Dossiers sélectionnés</label>
                  <div class="border rounded bg-surface">
                    <table class="table table-sm align-middle mb-0">
                      <thead class="table-light">
                        <tr>
                          <th style="width: 18%;">Dossier</th>
                          <th>Chemin complet</th>
                          <th style="width: 190px;">Catégorie</th>
                          <th style="width: 52px;"></th>
                        </tr>
                      </thead>
                      <tbody id="selected-folders-body">
                        <tr id="selected-folders-empty">
                          <td colspan="4" class="text-muted small py-2 ps-2">Cochez un ou plusieurs dossiers dans l'arborescence.</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <div class="form-text">Chaque dossier peut être affecté à une catégorie différente.</div>
                </div>
                <div class="mb-3">
                  <label class="form-label">Catégorie</label>
                  <select class="form-control" id="category-input" required>
                    <option value="">-- Sélectionner une catégorie --</option>
                    <option value="Déclarations">Déclarations</option>
                    <option value="Règlements">Règlements</option>
                    <option value="Mails simples">Mails simples</option>
                  </select>
                  <div class="form-text">Choisissez la catégorie appropriée pour classer les emails de ce dossier</div>
                </div>
              </form>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Annuler</button>
              <button type="button" class="btn btn-primary" id="save-folder-config">
                <span class="save-text">Ajouter</span>
                <span class="spinner-border spinner-border-sm d-none ms-2" role="status" aria-hidden="true"></span>
              </button>
            </div>
          </div>
        </div>
      </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    const modal = new bootstrap.Modal(document.getElementById('folderModal'));
    modal.show();
    // Préselectionner une catégorie par défaut pour éviter l'état vide
    try {
      const catEl = document.getElementById('category-input');
      if (catEl && !catEl.value) catEl.value = 'Mails simples';
      catEl?.classList.remove('is-invalid');
    } catch(_) {}

    // Gestion multi-sélection
    const selectionCtrl = this.createFolderSelectionController({
      defaultCategorySelectId: 'category-input',
      selectedBodyId: 'selected-folders-body',
      hiddenInputId: 'selected-folder-path'
    });
    this._folderSelectionController = selectionCtrl;

    // Gérer le changement de boîte mail
    const mailboxSelectEl = document.getElementById('mailbox-select');
    mailboxSelectEl.addEventListener('change', async (e) => {
      const storeId = e.target.value || '';
      this._folderSelectionController?.clear?.();
      await this.loadFoldersForMailbox(storeId);
    });

    // Gérer la sauvegarde
      document.getElementById('save-folder-config').addEventListener('click', () => {
      this.saveFolderConfiguration(modal);
    });

    // Pré-sélectionner la première boîte et charger ses dossiers pour éviter l'état vide
    try {
      if (mailboxSelectEl && mailboxSelectEl.options.length > 1) {
        mailboxSelectEl.selectedIndex = 1;
        (async () => { try { await this.loadFoldersForMailbox(mailboxSelectEl.value || ''); } catch(_) {} })();
      }
    } catch (_) {}
  }

  async loadFoldersForMailbox(storeId) {
    try {
      const folderTree = document.getElementById('folder-tree');
      folderTree.innerHTML = '<div class="text-muted"><i class="bi bi-hourglass-split me-2"></i>Chargement...</div>';

      // Utiliser EWS directement pour ce store (on affiche le libellé comme mailbox)
  const mailboxSelect = document.getElementById('mailbox-select');
  const selectedOpt = mailboxSelect ? mailboxSelect.selectedOptions[0] : null;
  const mailboxDisplay = selectedOpt ? (selectedOpt.getAttribute('data-name') || selectedOpt.text || '') : '';
  const smtp = selectedOpt ? (selectedOpt.getAttribute('data-smtp') || '') : '';
  folderTree.dataset.mailboxDisplay = mailboxDisplay;
  folderTree.dataset.mailbox = mailboxDisplay; // backward compat
  folderTree.dataset.smtp = smtp || '';
      // Si storeId vide, essayer de le résoudre via COM stores d'après le label (DisplayName / SMTP)
      if (!storeId) {
        try {
          const storesRes = await window.electronAPI.olStores();
          if (storesRes && storesRes.ok && Array.isArray(storesRes.data)) {
            const label = (mailboxDisplay || '').trim();
            const match = storesRes.data.find(s => {
              const disp = (s.DisplayName || '').trim();
              const smtp = (s.SmtpAddress || '').trim();
              const combined = smtp ? `${disp} (${smtp})` : disp;
              return combined === label || disp === label || smtp === label;
            });
      if (match && match.StoreId) storeId = match.StoreId;
          }
        } catch (_) {}
      }
      try {
        // COM structure complète d'abord (fonctionne pour boîtes partagées)
        let list = [];
  const diag = { mailbox: mailboxDisplay, storeId: storeId || '(inconnu)', errors: [] };
    // Conserver le storeId pour le lazy-load ultérieur
    folderTree.dataset.storeId = storeId || '';
        try {
          const struct = await window.electronAPI.getFolderStructure(storeId || '');
          if (struct && struct.success && Array.isArray(struct.folders) && struct.folders.length > 0) {
            const mailbox = struct.folders[0];
            const sub = Array.isArray(mailbox.SubFolders) ? mailbox.SubFolders : [];
            // Conserver la structure complète
            list = sub;
          }
        } catch (e) { diag.errors.push(`COM: ${e?.message || String(e)}`); }
        // Fallback COM rapide si nécessaire
        if (!list.length) {
          try {
            const res = await window.electronAPI.olFoldersShallow(storeId, '');
            if (res && res.ok) {
              const arr = Array.isArray(res.data) ? res.data : (res.data?.folders || []);
              if (Array.isArray(arr)) list = arr.map(n => ({ Name: n.Name, Id: n.EntryId, ChildCount: n.ChildCount }));
              // If shallow includes a ParentName, remember and wrap later
              var parentLabel2 = (res.data && (res.data.ParentName || res.data.parentName)) ? String(res.data.ParentName || res.data.parentName) : '';
              if (!list.length && res.data && res.data.error) diag.errors.push(`COM: ${res.data.error}`);
              // Post-process wrap under Inbox-equivalent if parentLabel2 is present and items exist
              if (typeof parentLabel2 === 'string' && parentLabel2.trim() && list.length) {
                const p = parentLabel2.trim();
                const children = list;
                list = [{ Name: p, Id: '', ChildCount: children.length, FolderPath: `${mailboxDisplay}\\${p}`, SubFolders: children }];
              }
            }
            else if (res && res.ok === false) diag.errors.push(`COM: ${res.error || 'erreur inconnue'}`);
          } catch (e) { diag.errors.push(`COM: ${e?.message || String(e)}`); }
        }
        if (!list.length) {
          // Fallback EWS
          try {
            const ewsMailbox = (smtp && smtp.includes('@')) ? smtp : mailboxDisplay;
            const ewsRes = await window.electronAPI.ewsTopLevel(ewsMailbox);
            list = Array.isArray(ewsRes) ? ewsRes : (ewsRes?.folders || []);
            if (ewsRes && ewsRes.success === false && ewsRes.error) diag.errors.push(`EWS: ${ewsRes.error}`);
          } catch (e2) { diag.errors.push(`EWS: ${e2?.message || String(e2)}`); }
        }
        if (!list.length) {
          const ctx = this.escapeHtml(`Boîte: ${diag.mailbox} | StoreId: ${diag.storeId}`);
          const errs = diag.errors.length ? `<br/><small class="text-muted">${this.escapeHtml(diag.errors.join(' | '))}</small>` : '';
          folderTree.innerHTML = `<div class="text-warning"><i class="bi bi-info-circle me-2"></i>Aucun dossier trouvé pour cette boîte mail<br/><small>${ctx}</small>${errs}</div>`;
          return;
        }
        const buildTreeItem2 = (node, parentDisplayPath) => {
          const name = node.Name;
          const entryId = node.Id || node.EntryID || node.EntryId || '';
          const childCount = node.ChildCount || (Array.isArray(node.SubFolders) ? node.SubFolders.length : 0);
          const rawPath = node.FolderPath || node.Path || '';
          const folderPath = (rawPath && rawPath.startsWith(mailboxDisplay)) ? rawPath : `${parentDisplayPath}\\${name}`;
          const children = Array.isArray(node.SubFolders) ? node.SubFolders.map(sf => buildTreeItem2(sf, folderPath)) : [];
          return { Name: name, FolderPath: folderPath, EntryID: entryId, ChildCount: childCount, SubFolders: children };
        };
        const treeItems = list.map(n => buildTreeItem2(n, mailboxDisplay));
        let html = '';
        for (const it of treeItems) html += this.createFolderTree(it, 0, true);
        folderTree.innerHTML = html;
        this.initializeFolderTreeEvents({
          onToggle: (payload) => this._folderSelectionController?.toggle(payload),
          isSelected: (path) => this._folderSelectionController?.isSelected(path),
          highlightSelected: true,
          withCheckbox: true
        });
        this._folderSelectionController?.syncCheckboxes?.();
      } catch (err) {
        console.error('❌ Erreur top-level:', err);
        folderTree.innerHTML = '<div class="text-danger"><i class="bi bi-exclamation-triangle me-2"></i>Erreur de chargement</div>';
      }
    } catch (error) {
      console.error('❌ Erreur chargement dossiers:', error);
      document.getElementById('folder-tree').innerHTML = '<div class="text-danger"><i class="bi bi-exclamation-triangle me-2"></i>Erreur de chargement</div>';
    }
  }

  createFolderTree(structure, level = 0, withCheckbox = false) {
    let html = '';
    
    // Support des deux formats : l'ancien (Path/Subfolders) et le nouveau (FolderPath/SubFolders)
    const folderName = structure.Name;
    const folderPath = structure.Path || structure.FolderPath;
    const rawSubfolders = structure.Subfolders ?? structure.SubFolders;
    let subfolders = rawSubfolders || [];
  const unreadCount = structure.UnreadCount || 0;
  const entryId = structure.EntryID || '';
  const childCount = structure.ChildCount || 0;
    
    // Nettoyer SubFolders si ce n'est pas un tableau (fix pour les cas bizarres de PowerShell)
    if (!Array.isArray(subfolders)) {
      console.log('🔧 DEBUG - SubFolders non-tableau détecté:', typeof subfolders, subfolders);
      if (typeof subfolders === 'string' && (subfolders === '' || subfolders.includes('System.Collections'))) {
        subfolders = [];
      } else if (typeof subfolders === 'object' && subfolders !== null) {
        // Si c'est un objet, essayer de le convertir en tableau
        subfolders = Object.values(subfolders).filter(item => item && typeof item === 'object');
      } else {
        subfolders = [];
      }
    }
    
    if (folderName && folderPath) {
      const hasInlineChildren = Array.isArray(subfolders) && subfolders.length > 0;
      const childCountIsNumber = typeof structure.ChildCount === 'number';
      const assumeExpandable = !hasInlineChildren && (structure.ChildCount === 0 || !childCountIsNumber) && rawSubfolders === undefined;
      const hasChildren = hasInlineChildren || (childCountIsNumber ? childCount > 0 : true) || assumeExpandable;
      const indent = '  '.repeat(level);
      const folderId = `folder_${Math.random().toString(36).substr(2, 9)}`;
      
      const safeName = this.escapeHtml(folderName);
      const safePath = this.escapeHtml(folderPath);
      const safeEntryId = this.escapeHtml(entryId);
      const checkboxHtml = withCheckbox ? `<input type="checkbox" class="form-check-input folder-checkbox me-2" data-path="${safePath}" data-name="${safeName}" data-entry-id="${safeEntryId}">` : '';
      html += `
        <div class="folder-item" data-level="${level}">
          <div class="folder-line folder-selectable" 
               data-path="${safePath}" 
               data-name="${safeName}"
               data-entry-id="${safeEntryId}"
               data-has-children="${hasChildren ? '1' : '0'}"
               style="--depth:${level};">
            ${hasChildren ? 
              `<i class="bi bi-chevron-right folder-toggle" data-target="${folderId}" role="button" aria-label="Afficher les sous-dossiers"></i>` : 
              `<span class="folder-toggle-placeholder"></span>`
            }
            ${checkboxHtml}
            <i class="bi bi-folder text-warning folder-glyph"></i>
            <span class="folder-name" title="${safePath}">${safeName}</span>
            ${unreadCount > 0 ? `<span class="badge bg-light text-dark ms-1">${unreadCount}</span>` : ''}
          </div>
          ${hasChildren ? `<div class="folder-children" id="${folderId}" style="display: none;"></div>` : ''}
      `;
      // Si on a déjà des subfolders fournis, les insérer maintenant
      if (Array.isArray(subfolders) && subfolders.length > 0) {
        html = html.replace(`id="${folderId}" style="display: none;">`, `id="${folderId}" style="display: none;">` + subfolders.map(sf => this.createFolderTree(sf, level + 1, withCheckbox)).join(''));
      }

      html += '</div>';
    }

    return html;
  }

  toggleFolderLineHighlight(line, active) {
    if (!line) return;
    line.classList.toggle('bg-primary', !!active);
    line.classList.toggle('text-white', !!active);
    if (!active) line.classList.remove('hover-surface');
  }

  renderCategoryOptions(selected) {
    const categories = ['Déclarations', 'Règlements', 'Mails simples'];
    return categories.map(cat => `<option value="${cat}" ${selected === cat ? 'selected' : ''}>${cat}</option>`).join('');
  }

  createFolderSelectionController(config = {}) {
    const selection = new Map();
    const {
      defaultCategorySelectId = 'category-input',
      selectedBodyId = 'selected-folders-body',
      hiddenInputId = 'selected-folder-path'
    } = config;

    const defaultCategoryEl = document.getElementById(defaultCategorySelectId);
    const body = document.getElementById(selectedBodyId);
    const hiddenInput = document.getElementById(hiddenInputId);

    const getDefaultCategory = () => (defaultCategoryEl?.value?.trim() || 'Mails simples');

    const syncHidden = () => {
      if (!hiddenInput) return;
      hiddenInput.value = selection.size ? [...selection.keys()][0] : '';
    };

    const syncCheckboxes = () => {
      document.querySelectorAll('.folder-checkbox').forEach(cb => {
        const path = cb.getAttribute('data-path');
        const checked = selection.has(path);
        cb.checked = checked;
        this.toggleFolderLineHighlight(cb.closest('.folder-line'), checked);
      });
    };

    const render = () => {
      if (!body) return;
      if (selection.size === 0) {
        body.innerHTML = `<tr id="selected-folders-empty"><td colspan="4" class="text-muted small py-2 ps-2">Cochez un ou plusieurs dossiers dans l'arborescence.</td></tr>`;
      } else {
        body.innerHTML = [...selection.entries()].map(([path, info]) => {
          const cat = info.category || getDefaultCategory();
          const safePath = this.escapeHtml(path);
          const safeName = this.escapeHtml(info.name || 'Dossier');
          return `
            <tr class="selected-folder-row" data-path="${safePath}" data-entry-id="${this.escapeHtml(info.entryId || '')}" data-name="${safeName}">
              <td class="text-nowrap">${safeName}</td>
              <td class="small">${safePath}</td>
              <td style="width: 190px;">
                <select class="form-select form-select-sm selected-folder-category">${this.renderCategoryOptions(cat)}</select>
              </td>
              <td class="text-end" style="width: 52px;">
                <button class="btn btn-outline-secondary btn-sm" type="button" data-remove-folder="${safePath}" aria-label="Retirer"><i class="bi bi-x"></i></button>
              </td>
            </tr>`;
        }).join('');
      }

      syncHidden();

      body.querySelectorAll('.selected-folder-category').forEach(sel => sel.addEventListener('change', (e) => {
        const row = e.target.closest('.selected-folder-row');
        const p = row?.dataset.path;
        if (p && selection.has(p)) selection.get(p).category = e.target.value;
      }));

      body.querySelectorAll('[data-remove-folder]').forEach(btn => btn.addEventListener('click', (ev) => {
        const p = ev.currentTarget.getAttribute('data-remove-folder');
        if (p) selection.delete(p);
        render();
        syncCheckboxes();
      }));

      syncCheckboxes();
    };

    const add = ({ path, name, entryId }) => {
      if (!path) return;
      if (!selection.has(path)) selection.set(path, { name, entryId, category: getDefaultCategory() });
      render();
    };

    const remove = (path) => {
      selection.delete(path);
      render();
    };

    const toggle = ({ path, name, entryId, checked }) => {
      if (checked) add({ path, name, entryId });
      else remove(path);
    };

    render();

    return {
      toggle,
      add,
      remove,
      clear: () => { selection.clear(); render(); },
      isSelected: (path) => selection.has(path),
      syncCheckboxes,
      render,
      getSelections: () => [...selection.entries()].map(([path, info]) => ({
        path,
        name: info.name,
        entryId: info.entryId,
        category: info.category || getDefaultCategory()
      }))
    };
  }

  initializeFolderTreeEvents(options = {}) {
    const { onToggle, isSelected, highlightSelected = false, withCheckbox = false } = options;

    // Gestion dépliage/repliage + lazy-load
    document.querySelectorAll('.folder-toggle').forEach(toggle => {
      toggle.addEventListener('click', async (e) => {
        e.stopPropagation();
        const targetId = toggle.getAttribute('data-target');
        const targetDiv = document.getElementById(targetId);
        const isExpanded = targetDiv.style.display !== 'none';
        if (!isExpanded && targetDiv && targetDiv.children.length === 0) {
          const folderLine = toggle.closest('.folder-item')?.querySelector('.folder-line');
          if (folderLine && folderLine.getAttribute('data-has-children') === '1') {
            const mailbox = document.getElementById('folder-tree')?.dataset.mailbox || '';
            const smtp = document.getElementById('folder-tree')?.dataset.smtp || '';
            const parentEntryId = folderLine.getAttribute('data-entry-id') || '';
            const parentDisplayPath = folderLine.getAttribute('data-path') || '';
            const storeId = document.getElementById('folder-tree')?.dataset.storeId || '';
            if (mailbox && (parentEntryId || parentDisplayPath)) {
              targetDiv.innerHTML = '<div class="text-muted ms-4"><i class="bi bi-hourglass-split me-2"></i>Chargement…</div>';
              try {
                let children = [];
                try {
                  if (storeId) {
                    const res = await window.electronAPI.getSubFolders({ storeId, parentEntryId: parentEntryId || '', parentPath: parentDisplayPath || '' });
                    if (res && res.success && Array.isArray(res.children)) {
                      children = res.children.map(n => ({ Name: n.Name, Id: n.EntryID || n.EntryId || '', ChildCount: n.ChildCount || 0 }));
                    }
                  }
                } catch {}
                if (!children.length) {
                  try {
                    if (storeId && parentEntryId) {
                      const r = await window.electronAPI.olFoldersShallow(storeId, parentEntryId);
                      if (r && r.ok) {
                        const arr = Array.isArray(r.data) ? r.data : (r.data?.folders || []);
                        if (Array.isArray(arr)) children = arr.map(n => ({ Name: n.Name, Id: n.EntryId, ChildCount: n.ChildCount }));
                      }
                    }
                  } catch {}
                }
                if (!children.length && parentEntryId) {
                  const ewsMailbox = (smtp && smtp.includes('@')) ? smtp : (document.getElementById('folder-tree')?.dataset.mailboxDisplay || mailbox);
                  const res = await window.electronAPI.ewsChildren(ewsMailbox, parentEntryId);
                  children = Array.isArray(res) ? res : (res?.folders || []);
                }
                const parentPath = parentDisplayPath || `${(document.getElementById('folder-tree')?.dataset.mailboxDisplay || mailbox)}\\${folderLine.getAttribute('data-name') || ''}`;
                const mapped = children.map(ch => ({
                  Name: ch.Name,
                  FolderPath: `${parentPath}\\${ch.Name}`,
                  EntryID: ch.Id,
                  ChildCount: ch.ChildCount,
                  SubFolders: []
                }));
                const parentItem = toggle.closest('.folder-item');
                const level = parseInt(parentItem?.getAttribute('data-level') || '0', 10) + 1;
                targetDiv.innerHTML = mapped.map(m => this.createFolderTree(m, level, withCheckbox)).join('');
                this.initializeFolderTreeEvents(options);
              } catch (err) {
                console.error('Lazy-load sous-dossiers EWS échoué:', err);
                targetDiv.innerHTML = '<div class="text-danger ms-4"><i class="bi bi-exclamation-triangle me-2"></i>Erreur</div>';
              }
            }
          }
        }
        if (isExpanded) {
          targetDiv.style.display = 'none';
          toggle.classList.remove('bi-chevron-down');
          toggle.classList.add('bi-chevron-right');
        } else {
          targetDiv.style.display = 'block';
          toggle.classList.remove('bi-chevron-right');
          toggle.classList.add('bi-chevron-down');
        }
      });
    });

    // Cases à cocher (multi-sélection)
    document.querySelectorAll('.folder-checkbox').forEach(cb => {
      const line = cb.closest('.folder-line');
      const path = cb.getAttribute('data-path') || line?.getAttribute('data-path') || '';
      const name = cb.getAttribute('data-name') || line?.getAttribute('data-name') || '';
      const entryId = cb.getAttribute('data-entry-id') || line?.getAttribute('data-entry-id') || '';
      if (typeof isSelected === 'function') {
        const preset = !!isSelected(path);
        cb.checked = preset;
        if (preset && highlightSelected) this.toggleFolderLineHighlight(line, true);
      }
      cb.addEventListener('change', () => {
        const payload = { path, name, entryId, checked: cb.checked };
        if (typeof onToggle === 'function') onToggle(payload);
        if (highlightSelected) this.toggleFolderLineHighlight(line, cb.checked);
      });
    });

    // Sélection via clic sur la ligne
    document.querySelectorAll('.folder-selectable').forEach(folder => {
      folder.addEventListener('click', (e) => {
        if (e.target.closest('.folder-toggle')) return;
        if (e.target.closest('.folder-checkbox')) return;
        const checkbox = folder.querySelector('.folder-checkbox');
        if (checkbox) {
          checkbox.checked = !checkbox.checked;
          checkbox.dispatchEvent(new Event('change'));
          return;
        }

        // Mode historique (sélection unique)
        document.querySelectorAll('.folder-selectable').forEach(f => {
          f.classList.remove('bg-primary', 'text-white');
          f.classList.remove('hover-surface');
        });

        folder.classList.add('bg-primary', 'text-white');
        folder.classList.remove('text-dark');

        const path = folder.getAttribute('data-path');
        document.getElementById('selected-folder-path').value = path;
      });

      folder.addEventListener('mouseenter', () => {
        if (!folder.classList.contains('bg-primary')) {
          folder.classList.add('hover-surface');
        }
      });

      folder.addEventListener('mouseleave', () => {
        if (!folder.classList.contains('bg-primary')) {
          folder.classList.remove('hover-surface');
        }
      });
    });
  }

  flattenFolderStructure(structure, prefix = '') {
    let folders = [];
    
    if (structure.Name && structure.Path) {
      folders.push({
        name: prefix + structure.Name,
        path: structure.Path
      });
    }

    if (structure.Subfolders && Array.isArray(structure.Subfolders) && structure.Subfolders.length > 0) {
      for (const subfolder of structure.Subfolders) {
        folders = folders.concat(this.flattenFolderStructure(subfolder, prefix + '  '));
      }
    }

    return folders;
  }

  async saveFolderConfiguration(modal) {
    try {
      const fallbackPath = document.getElementById('selected-folder-path').value;
      const defaultCategory = document.getElementById('category-input').value.trim();
      const storeId = document.getElementById('folder-tree')?.dataset.storeId || '';
      const storeName = document.getElementById('folder-tree')?.dataset.mailboxDisplay || document.getElementById('folder-tree')?.dataset.mailbox || '';

      const selectionList = (this._folderSelectionController?.getSelections?.() || []).filter(s => s.path);
      // Fallback single-selection (legacy) si aucune case à cocher n'est utilisée
      if (!selectionList.length && fallbackPath) {
        const selectedFolder = document.querySelector('.folder-selectable.bg-primary');
        selectionList.push({
          path: fallbackPath,
          category: defaultCategory,
          name: selectedFolder ? selectedFolder.getAttribute('data-name') : 'Dossier',
          entryId: selectedFolder ? (selectedFolder.getAttribute('data-entry-id') || '') : ''
        });
      }

      const catEl = document.getElementById('category-input');
      const treeEl = document.getElementById('folder-tree');
      const treeWrapper = treeEl ? treeEl.parentElement : null;
      catEl?.classList.remove('is-invalid');
      treeWrapper?.classList.remove('border-danger');

      if (!selectionList.length) {
        treeWrapper?.classList.add('border-danger');
        return;
      }

      // Vérifier les catégories par dossier
      const missingCategory = selectionList.some(sel => !sel.category);
      if (missingCategory) {
        catEl?.classList.add('is-invalid');
        return;
      }

      // UI loading state for save button and inline alert area
      const btn = document.getElementById('save-folder-config');
      const spinner = btn?.querySelector('.spinner-border');
      const text = btn?.querySelector('.save-text');
      const alertBox = document.getElementById('folder-modal-alert');
      const logBox = document.getElementById('folder-log-box');
      const pushLog = (msg) => {
        if (!logBox) return;
        const line = document.createElement('div');
        line.textContent = msg;
        logBox.appendChild(line);
        while (logBox.children.length > 200) logBox.removeChild(logBox.firstChild);
        logBox.scrollTop = logBox.scrollHeight;
      };

      const setInlineStatus = (message, variant = 'info') => {
        if (!alertBox) return;
        alertBox.textContent = message || '';
        alertBox.classList.remove('d-none', 'alert-danger', 'alert-info', 'alert-warning', 'alert-success');
        alertBox.classList.add(`alert-${variant}`);
      };

      // Démarrer l'ajout en mode séquentiel (évite les timeouts et la saturation PowerShell/IPC)
      const total = selectionList.length;
      let succeeded = 0;
      let failed = 0;
      const errors = [];

      try {
        if (alertBox) { alertBox.classList.add('d-none'); alertBox.textContent = ''; }
        spinner?.classList.remove('d-none');
        if (text) text.textContent = 'Ajout...';
        if (btn) btn.disabled = true;
      } catch(_) {}

      setInlineStatus(`Ajout en cours… 0/${total}`, 'info');
      pushLog(`[SAVE] Début ajout: ${total} dossier(s)`);

      // Optimisation: si l'IPC bulk est dispo et qu'on ajoute plusieurs dossiers,
      // déléguer au main process (un seul restart monitoring + moins d'aller-retours IPC).
      if (total > 1 && window.electronAPI?.addFoldersToMonitoringBulk) {
        setInlineStatus(`Ajout en cours… ${total} dossier(s) (voir logs)`, 'info');
        const bulkRes = await window.electronAPI.addFoldersToMonitoringBulk({
          items: selectionList.map(sel => ({
            folderPath: sel.path,
            category: sel.category,
            entryId: sel.entryId || '',
            storeId,
            storeName
          })),
          storeId,
          storeName
        });

        const inserted = Number(bulkRes?.inserted || 0);
        const unique = (bulkRes && Object.prototype.hasOwnProperty.call(bulkRes, 'unique'))
          ? Number(bulkRes.unique || 0)
          : null;

        // On ne reçoit plus un résultat par item; on affiche un résumé.
        succeeded = inserted;
        failed = Math.max(0, total - inserted);

        if (!bulkRes?.success) {
          const msg = bulkRes?.error || 'Certains dossiers n\'ont pas pu être ajoutés';
          errors.push(msg);
          pushLog(`[BULK] inserted=${inserted}${unique !== null ? ` unique=${unique}` : ''} (échec)`);
        } else {
          pushLog(`[BULK] inserted=${inserted}${unique !== null ? ` unique=${unique}` : ''}`);
        }
      } else {

      for (let i = 0; i < total; i++) {
        // Si le modal est fermé pendant l'opération, arrêter proprement
        if (!document.getElementById('folderModal')) {
          pushLog('[SAVE] Modal fermé: arrêt de l\'ajout');
          break;
        }

        const sel = selectionList[i];
        const label = `${i + 1}/${total}`;
        setInlineStatus(`Ajout en cours… ${label}`, 'info');
        pushLog(`[ADD ${label}] ${sel.path}  |  catégorie=${sel.category}`);

        try {
          const res = await window.electronAPI.addFolderToMonitoring({
            folderPath: sel.path,
            category: sel.category,
            entryId: sel.entryId || '',
            storeId,
            storeName,
            suppressRestart: i < (total - 1)
          });
          if (res && res.success) {
            succeeded++;
            pushLog(`[OK  ${label}] ajouté`);
          } else {
            failed++;
            const msg = (res && (res.error || res.message)) || 'Échec de l\'ajout';
            errors.push(msg);
            pushLog(`[ERR ${label}] ${msg}`);
          }
        } catch (e) {
          failed++;
          const msg = e?.message || 'Erreur inconnue';
          errors.push(msg);
          pushLog(`[ERR ${label}] ${msg}`);
        }
      }
      }

      // Rafraîchir l'UI si au moins un dossier a été ajouté
      if (succeeded > 0) {
        await this.loadFoldersConfiguration();
        this.updateFolderConfigDisplay();
        if (window.foldersTree && typeof window.foldersTree.loadFolders === 'function') {
          try { await window.foldersTree.loadFolders(true); } catch (_) {}
        }
      }

      if (failed === 0 && succeeded > 0) {
        // Fermer le modal seulement si tout est OK
        const modalElement = document.getElementById('folderModal');
        if (modalElement) {
          const modal = bootstrap.Modal.getInstance(modalElement);
          modal?.hide?.();
        }
        console.log(`✅ Configuration sauvegardée: ${succeeded} dossier(s) ajouté(s)`);
      } else if (succeeded === 0) {
        const msg = errors[0] || 'Impossible de sauvegarder la configuration';
        console.warn('Erreur de sauvegarde', msg);
        setInlineStatus(msg, 'danger');
      } else {
        // Succès partiel: laisser le modal ouvert et afficher un résumé
        setInlineStatus(`Ajout terminé: ${succeeded} OK, ${failed} en échec.`, 'warning');
      }
    } catch (error) {
      console.error('❌ Erreur sauvegarde configuration:', error);
      // Highlight the form to show an error occurred
      try {
        const el = document.getElementById('folder-tree');
        el?.parentElement?.classList.add('border-danger');
        const alertBox = document.getElementById('folder-modal-alert');
        if (alertBox) { alertBox.textContent = error.message || 'Erreur de sauvegarde'; alertBox.classList.remove('d-none'); }
      } catch(_) {}
    } finally {
      const btn = document.getElementById('save-folder-config');
      const spinner = btn?.querySelector('.spinner-border');
      const text = btn?.querySelector('.save-text');
      try { spinner?.classList.add('d-none'); if (text) text.textContent = 'Ajouter'; if (btn) btn.disabled = false; } catch(_) {}
    }
  }

  // Ouvre un sélecteur de catégorie et met à jour en BDD + état local
  async editFolderCategory(folderPath) {
    try {
      const current = this.state.folderCategories[folderPath] || {};
      const currentCat = current.category || '';

      // Construire un petit modal de sélection de catégorie
      const modalId = 'editCategoryModal';
      document.getElementById(modalId)?.remove();
      const html = `
        <div class="modal fade" id="${modalId}" tabindex="-1">
          <div class="modal-dialog">
            <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title">Modifier la catégorie</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
              </div>
              <div class="modal-body">
                <div class="mb-3">
                  <label class="form-label">Catégorie</label>
                  <select class="form-select" id="edit-category-select">
                    <option value="Déclarations" ${currentCat === 'Déclarations' ? 'selected' : ''}>Déclarations</option>
                    <option value="Règlements" ${currentCat === 'Règlements' ? 'selected' : ''}>Règlements</option>
                    <option value="Mails simples" ${currentCat === 'Mails simples' ? 'selected' : ''}>Mails simples</option>
                  </select>
                </div>
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Annuler</button>
                <button type="button" class="btn btn-primary" id="edit-category-save">Enregistrer</button>
              </div>
            </div>
          </div>
        </div>`;
      document.body.insertAdjacentHTML('beforeend', html);
      const modal = new bootstrap.Modal(document.getElementById(modalId));
      modal.show();

      document.getElementById('edit-category-save').addEventListener('click', async () => {
        const newCategory = document.getElementById('edit-category-select').value;
        if (!newCategory) return;
        try {
          // Mise à jour BDD via IPC dédié si dispo, sinon fallback saveFoldersConfig
          if (window.electronAPI.updateFolderCategory) {
            const res = await window.electronAPI.updateFolderCategory({ folderPath, category: newCategory });
            if (!res?.success) throw new Error(res?.error || 'Échec de la mise à jour');
          } else {
            this.state.folderCategories[folderPath] = { ...(this.state.folderCategories[folderPath] || {}), category: newCategory };
            const res = await window.electronAPI.saveFoldersConfig({ folderCategories: this.state.folderCategories });
            if (!res?.success) throw new Error(res?.error || 'Échec de la sauvegarde');
          }

          // Mettre à jour l'état local de façon optimiste
          if (!this.state.folderCategories[folderPath]) this.state.folderCategories[folderPath] = {};
          this.state.folderCategories[folderPath].category = newCategory;
          this.updateFolderConfigDisplay();

          // Forcer la vue Monitoring à refléter la nouvelle catégorie
          if (window.foldersTree && typeof window.foldersTree.loadFolders === 'function') {
            try { await window.foldersTree.loadFolders(true); } catch (_) {}
          }

          bootstrap.Modal.getInstance(document.getElementById(modalId))?.hide();
          this.showNotification('Catégorie mise à jour', `Nouvelle catégorie: "${newCategory}"`, 'success');
        } catch (e) {
          console.error('❌ Erreur maj catégorie:', e);
          this.showNotification('Erreur', e.message, 'danger');
        }
      });
    } catch (error) {
      console.error('❌ Erreur editFolderCategory:', error);
      this.showNotification('Erreur', error.message, 'danger');
    }
  }

  

  extractFolderName(path) {
    if (!path) return '';
    const parts = String(path).split('\\');
    return parts[parts.length - 1] || path;
  }

  // === MISE À JOUR DE L'INTERFACE ===
  updateConnectionStatus(status) {
    const indicator = document.getElementById('connection-indicator');
    const statusText = document.getElementById('connection-status');
    
    if (indicator && statusText) {
      if (status.status) {
        indicator.className = 'status-indicator status-connected me-2';
        statusText.textContent = 'Connecté à Outlook';
      } else {
        indicator.className = 'status-indicator status-disconnected me-2';
        statusText.textContent = status.error || 'Déconnecté';
      }
    }
  }

  updateStatsDisplay() {
    const stats = this.state.stats;
    
    // Mise à jour des compteurs principaux avec animation
  if (typeof stats.totalEmails === 'number') {
    this.animateCounterUpdate('total-emails', stats.totalEmails);
  }
  // Aligner le comportement de "Non lus" sur "Total mails" (ajustement fluide, sans disparition)
  if (typeof stats.unreadTotal === 'number') {
    this.animateCounterUpdate('emails-unread', stats.unreadTotal);
  }
  if (typeof stats.emailsToday === 'number') {
    this.animateCounterUpdate('emails-today', stats.emailsToday);
  }
    
    // Calcul et affichage des pourcentages
  const totalEmails = typeof stats.totalEmails === 'number' ? stats.totalEmails : 0;
  const unreadEmails = typeof stats.unreadTotal === 'number' ? stats.unreadTotal : 0;
    
    if (totalEmails > 0) {
      const unreadPercentage = Math.round((unreadEmails / totalEmails) * 100);
      document.getElementById('unread-percentage').textContent = `${unreadPercentage}%`;
      
      // Coloration en fonction du pourcentage
      const unreadEl = document.getElementById('unread-percentage');
      if (unreadPercentage > 50) {
        unreadEl.className = 'text-danger fw-bold';
      } else if (unreadPercentage > 25) {
        unreadEl.className = 'text-warning fw-bold';
      } else {
        unreadEl.className = 'text-success fw-bold';
      }
    }
    
    // Mise à jour du statut de monitoring
    const monitoredFolders = Object.keys(this.state.folderCategories).length;
    this.animateCounterUpdate('monitored-folders', monitoredFolders);
    
    this.updateMonitoringStatus();
    this.updateActivityMetrics(stats);
    
    if (this.state.lastUpdate) {
      document.getElementById('last-sync').textContent = this.state.lastUpdate.toLocaleTimeString();
    }
  }

  // Nouvelle méthode pour animer les mises à jour de compteurs
  animateCounterUpdate(elementId, newValue) {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    const parsed = parseInt((element.textContent || '').toString().replace(/[^0-9-]/g, ''), 10);
    const currentValue = Number.isFinite(parsed) ? parsed : 0;
    // Éviter l'animation sur des petites variations
    if (currentValue === newValue) return;
    if (Math.abs(newValue - currentValue) <= 1) {
      element.textContent = newValue;
      return;
    }
    
    // Stopper une animation en cours pour cet élément
    if (this.counterTimers.has(elementId)) {
      clearInterval(this.counterTimers.get(elementId));
      this.counterTimers.delete(elementId);
    }

    element.classList.add('updating');
    const duration = 400;
    const steps = 16;
    const increment = (newValue - currentValue) / steps;
    let currentStep = 0;

    const timer = setInterval(() => {
      currentStep++;
      const value = Math.round(currentValue + (increment * currentStep));
      element.textContent = value;
      if (currentStep >= steps) {
        clearInterval(timer);
        this.counterTimers.delete(elementId);
        element.textContent = newValue;
        element.classList.remove('updating');
      }
    }, duration / steps);
    this.counterTimers.set(elementId, timer);
  }

  // Mise à jour immédiate, sans animation, sans reset visuel
  setCounterImmediate(elementId, newValue) {
    const element = document.getElementById(elementId);
    if (!element) return;
    const parsed = parseInt((element.textContent || '').toString().replace(/[^0-9-]/g, ''), 10);
    const currentValue = Number.isFinite(parsed) ? parsed : null;
    if (currentValue === newValue) return;
    // Stopper une éventuelle animation en cours
    if (this.counterTimers.has(elementId)) {
      clearInterval(this.counterTimers.get(elementId));
      this.counterTimers.delete(elementId);
    }
    element.classList.remove('updating');
    element.textContent = newValue;
  }

  // Nouvelle méthode pour mettre à jour les métriques d'activité
  updateActivityMetrics(stats) {
    // Mise à jour des métriques d'activité
    document.getElementById('activity-total').textContent = stats.emailsToday || 0;
    document.getElementById('activity-read').textContent = (stats.totalEmails || 0) - (stats.unreadTotal || 0);
    document.getElementById('activity-pending').textContent = stats.unreadTotal || 0;
    
    // Calcul emails par heure (approximatif)
    const emailsToday = stats.emailsToday || 0;
    const currentHour = new Date().getHours();
    const emailsPerHour = currentHour > 0 ? Math.round(emailsToday / currentHour) : emailsToday;
    document.getElementById('emails-per-hour').textContent = emailsPerHour;
    
    // Simulation de croissance (en attendant les vraies données historiques)
    this.updateTodayGrowth(emailsToday);
    
    // Mise à jour des timestamps
    const now = new Date();
    document.getElementById('activity-last-update').textContent = now.toLocaleTimeString();
    document.getElementById('last-check').textContent = now.toLocaleTimeString();
  }

  // Méthode pour calculer et afficher la croissance d'aujourd'hui
  updateTodayGrowth(todayCount) {
    const growthEl = document.getElementById('today-growth');
    if (!growthEl) return;
    
    // Pour l'instant, simulation basée sur l'heure
    // Dans une vraie implémentation, il faudrait comparer avec les données d'hier
    const currentHour = new Date().getHours();
    const expectedByNow = Math.round(todayCount * (24 / Math.max(currentHour, 1)));
    const growth = todayCount > expectedByNow ? '+' + (todayCount - expectedByNow) : (todayCount - expectedByNow);
    
    if (growth > 0) {
      growthEl.textContent = `+${growth}`;
      growthEl.className = 'text-success fw-bold';
    } else if (growth < 0) {
      growthEl.textContent = growth;
      growthEl.className = 'text-danger fw-bold';
    } else {
      growthEl.textContent = 'Stable';
      growthEl.className = 'text-muted fw-bold';
    }
  }

  // Unifié: indicateur minimal de statut (affiché seulement si actif)
  updateMonitoringStatus() {
    const statusEl = document.getElementById('monitoring-status');
    if (!statusEl) return;
    if (this.state.isMonitoring) {
      statusEl.style.display = 'inline-flex';
      statusEl.innerHTML = `<span class="status-indicator status-connected me-1"></span>Actif`;
    } else {
      statusEl.style.display = 'none';
      statusEl.innerHTML = '';
    }
  }

  updateEmailsTable() {
    const tbody = document.getElementById('emails-table');
    if (!tbody) return;

    const emails = Array.isArray(this.state.recentEmails) ? this.state.recentEmails : [];

    if (emails.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="5" class="text-center text-muted py-5 border-0">
            <div class="d-flex flex-column align-items-center">
              <div class="spinner-border text-primary mb-3" role="status"></div>
              <h6 class="text-muted mb-2">Aucun email à afficher</h6>
              <p class="text-muted mb-0 small">Les nouveaux emails apparaîtront ici dès leur réception</p>
          </td>
        </tr>`;
      return;
    }

    tbody.innerHTML = emails.map(email => {
      const isUnread = (email.is_read === false) || (email.UnRead === true);
      const rowClass = isUnread ? 'email-row email-unread' : 'email-row email-read';
      const dateStr = new Date(email.received_time || email.created_at).toLocaleDateString();
      const senderName = this.escapeHtml(email.sender_name || 'Inconnu');
      const senderEmail = this.escapeHtml(email.sender_email || '');
      const subject = this.escapeHtml(email.subject || '(Sans objet)');
      const folderName = this.escapeHtml(email.folder_name || 'Inconnu');
      const statusBadge = isUnread
        ? '<span class="badge bg-warning text-dark">Non lu</span>'
        : '<span class="badge bg-success">Lu</span>';

      return `
        <tr class="${rowClass}">
          <td>${dateStr}</td>
          <td>
            <div class="fw-bold">${senderName}</div>
            <small class="text-muted">${senderEmail}</small>
          </td>
          <td>${subject}</td>
          <td><span class="badge bg-secondary">${folderName}</span></td>
          <td>${statusBadge}</td>
        </tr>`;
    }).join('');
  }

  updateFolderConfigDisplay() {
  // Ne plus gérer le rendu list/card ici: l'onglet Monitoring est désormais
  // entièrement géré par FoldersTreeManager (public/js/folders-tree.js)
  // Cette méthode conserve uniquement la mise à jour d'UI légère (compteurs)
  // et délègue le filtrage au gestionnaire d'arborescence.
  const container = document.getElementById('folders-tree');
  if (!container) return;

    const folders = Object.entries(this.state.folderCategories);
    // Appliquer les filtres UI
    const search = (document.getElementById('folder-search')?.value || '').toLowerCase().trim();
    const categoryFilter = document.getElementById('category-filter')?.value || '';
    const filtered = folders.filter(([path, config]) => {
      const name = (config?.name || '').toLowerCase();
      const cat = (config?.category || '').toLowerCase();
      const pathLc = (path || '').toLowerCase();
      const searchOk = !search || name.includes(search) || pathLc.includes(search) || cat.includes(search);
      const categoryOk = !categoryFilter || (config?.category === categoryFilter);
      return searchOk && categoryOk;
    });
    
    // Mettre à jour le compteur dans l'en-tête
    const countBadge = document.getElementById('monitored-folders-count');
    if (countBadge) {
      countBadge.textContent = filtered.length;
    }
    // Déléguer l'affichage/filtrage à l'arborescence UNIQUEMENT si les critères ont changé
    if (window.foldersTree) {
      const prevSearch = (window.foldersTree.searchTerm || '').toLowerCase();
      const prevCategory = window.foldersTree.categoryFilter || '';
      const changed = prevSearch !== search || prevCategory !== categoryFilter;
      if (changed && typeof window.foldersTree.filterFolders === 'function') {
        window.foldersTree.searchTerm = search;
        window.foldersTree.categoryFilter = categoryFilter;
        window.foldersTree.filterFolders();
      }
    }
    return;
  }

  // Helper pour générer une carte dossier (utilisé lors d'ajouts)
  _renderFolderCard(path, config, nums) {
    const categoryClass = this.getCategoryClass(config.category);
    const categoryIcon = this.getCategoryIcon(config.category);
    const safeKey = path.replace(/[^a-zA-Z0-9_-]/g, '_');
    const { total = 0 } = nums || {};
    return `
      <div class="folder-card mb-2 animate-slide-up" data-folder-key="${this.escapeHtml(path)}" id="monCard-${safeKey}">
        <div class="folder-header p-2">
          <div class="d-flex justify-content-between align-items-start">
            <div class="flex-grow-1">
              <div class="d-flex align-items-center mb-1">
                <i class="bi bi-folder-fill text-warning me-2 fs-5"></i>
                <h6 class="mb-0 fw-semibold">${this.escapeHtml(config.name || 'Dossier sans nom')}</h6>
                <span class="badge ${categoryClass} modern ms-2">
                  ${categoryIcon} ${this.escapeHtml(config.category)}
                </span>
              </div>
              <div class="folder-path">${this.escapeHtml(this.truncatePath(path))}</div>
            </div>
            <div class="dropdown">
              <button class="btn btn-sm btn-outline-secondary dropdown-toggle" type="button" data-bs-toggle="dropdown">
                <i class="bi bi-three-dots"></i>
              </button>
              <ul class="dropdown-menu dropdown-menu-end">
                <li>
                  <button class="dropdown-item" onclick='app.editFolderCategory(${JSON.stringify(path)})'>
                    <i class="bi bi-pencil me-2"></i>Modifier la catégorie
                  </button>
                </li>
                
                <li><hr class="dropdown-divider"></li>
                <li>
                  <button class="dropdown-item text-danger" onclick='app.removeFolderConfig(${JSON.stringify(path)})'>
                    <i class="bi bi-trash me-2"></i>Supprimer
                  </button>
                </li>
              </ul>
            </div>
          </div>
        </div>
        <!-- Bloc de statistiques détaillées supprimé -->
      </div>
    `;
  }

  // Méthodes utilitaires pour le design moderne
  getCategoryClass(category) {
    switch(category) {
      case 'Déclarations': return 'category-declarations';
      case 'Règlements': return 'category-reglements';
      case 'Mails simples': return 'category-simples';
      default: return 'bg-secondary';
    }
  }

  getCategoryIcon(category) {
    switch(category) {
      case 'Déclarations': return '📋';
      case 'Règlements': return '💰';
      case 'Mails simples': return '📧';
      default: return '📁';
    }
  }

  truncatePath(path, maxLength = 50) {
    if (path.length <= maxLength) return path;
    const parts = path.split('\\');
    if (parts.length > 2) {
      return `${parts[0]}\\...\\${parts[parts.length - 1]}`;
    }
    return path.substring(0, maxLength) + '...';
  }

  // === GESTION DES EMAILS AMÉLIORÉE ===
  
  /**
   * Chargement des emails récents avec filtrage avancé
   */
  async loadRecentEmails() {
    try {
      // Log réduit - supprimé pour éviter le spam
      // console.log('📧 Chargement des emails récents... (NOUVELLE VERSION)');
      
      const emails = await window.electronAPI.getRecentEmails();
      console.log('📧 Emails reçus de l\'API:', emails);
      
      if (emails && Array.isArray(emails) && emails.length > 0) {
        this.state.recentEmails = emails;
        // Email update - log simplifié
        // console.log('📧 État mis à jour avec', emails.length, 'emails');
        
        // Mettre à jour les statistiques des emails
        this.updateEmailsStats();
        
        // Afficher les emails dans le tableau
        this.displayEmailsTable();
  // Rafraîchir les cartes dossiers (compteurs)
  this.updateFolderConfigDisplay();
        
  // Mettre à jour les filtres
        this.updateEmailFilters();
        
        console.log(`✅ ${emails.length} emails chargés`);
  // Recalculer les compteurs par dossier après réception des emails pour éviter les zéros
  try { this.loadFoldersStats(); } catch (_) {}
      } else {
        console.warn('⚠️ Pas d\'emails trouvés');
        this.state.recentEmails = [];
        this.showEmptyEmailsState();
  this.updateFolderConfigDisplay();
      }
    } catch (error) {
      console.error('❌ Erreur lors du chargement des emails:', error);
      this.showEmailsError();
    }
  }

  /**
   * Mise à jour des statistiques d'emails dans la section dédiée
   */
  updateEmailsStats() {
    const emails = this.state.recentEmails;
    if (!emails || emails.length === 0) return;

    const today = new Date();
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
    
    const stats = {
      total: emails.length,
      unread: emails.filter(email => (!email.is_read && email.is_read !== undefined) || email.UnRead).length,
      today: emails.filter(email => {
        const emailDate = new Date(email.received_time || email.ReceivedTime);
        return emailDate.toDateString() === today.toDateString();
      }).length,
      week: emails.filter(email => {
        const emailDate = new Date(email.received_time || email.ReceivedTime);
        return emailDate >= startOfWeek;
      }).length
    };

    // Mettre à jour les compteurs
    this.animateCounterUpdate('emails-stats-total', stats.total);
    this.animateCounterUpdate('emails-stats-unread', stats.unread);
    this.animateCounterUpdate('emails-stats-today', stats.today);
    this.animateCounterUpdate('emails-stats-week', stats.week);

    // Calcul du taux d'emails par heure (approximatif)
    const hoursInDay = 24;
    const emailsPerHour = stats.today > 0 ? Math.round(stats.today / new Date().getHours() || 1) : 0;
    document.getElementById('emails-stats-rate').textContent = `${emailsPerHour}/h`;

    // Dernière synchronisation
    const lastSyncEl = document.getElementById('emails-stats-last-sync');
    if (lastSyncEl) {
      lastSyncEl.textContent = new Date().toLocaleTimeString();
    }
  }

  /**
   * Affichage du tableau des emails avec design moderne
   */
  displayEmailsTable() {
    const tbody = document.getElementById('emails-table');
    if (!tbody) return;

    const emails = this.state.recentEmails;
    
    if (!emails || emails.length === 0) {
      this.showEmptyEmailsState();
      return;
    }

    const emailsHtml = emails.slice(0, 50).map(email => { // Limiter à 50 emails pour les performances
      const receivedDate = new Date(email.received_time || email.ReceivedTime);
      const isToday = receivedDate.toDateString() === new Date().toDateString();
      const timeAgo = this.getTimeAgo(receivedDate);
      const isUnread = (!email.is_read && email.is_read !== undefined) || !!email.UnRead;
      const rowClass = isUnread ? 'email-row email-unread' : 'email-row email-read';
      
      // Déterminer le dossier depuis le chemin
      const folderName = this.extractFolderName(email.folder_name || email.FolderPath || 'Boîte de réception');
      const folderCategory = this.getFolderCategory(email.folder_name || email.FolderPath);
      
      return `
        <tr class="${rowClass}" data-email-id="${email.outlook_id || email.EntryID}">
          <td class="ps-3">
            <div class="d-flex flex-column">
              <span class="fw-medium ${isToday ? 'text-primary' : ''}">${timeAgo}</span>
              <small class="text-muted">${receivedDate.toLocaleDateString()}</small>
            </div>
          </td>
          <td>
            <div class="d-flex align-items-center">
              <div class="me-2">
                <div class="bg-primary rounded-circle d-flex align-items-center justify-content-center" 
                     style="width: 32px; height: 32px; font-size: 0.8rem; color: white;">
                  ${(email.sender_name || email.SenderName || email.sender_email || email.SenderEmailAddress || 'U').charAt(0).toUpperCase()}
                </div>
              </div>
              <div>
                <div class="fw-medium">${this.escapeHtml(email.sender_name || email.SenderName || email.sender_email || email.SenderEmailAddress || 'Expéditeur inconnu')}</div>
                ${(email.sender_name || email.SenderName) ? `<small class="text-muted">${this.escapeHtml(email.sender_email || email.SenderEmailAddress || '')}</small>` : ''}
              </div>
            </div>
          </td>
          <td>
            <div class="d-flex align-items-start">
              ${(email.has_attachment || email.HasAttachments) ? '<i class="bi bi-paperclip text-muted me-2"></i>' : ''}
              <div>
                <div class="fw-medium text-truncate" style="max-width: 300px;" title="${this.escapeHtml(email.subject || email.Subject)}">
                  ${this.escapeHtml(email.subject || email.Subject || 'Pas de sujet')}
                </div>
                ${(email.size_kb || email.Size) ? `<small class="text-muted">${this.formatFileSize((email.size_kb * 1024) || email.Size)}</small>` : ''}
              </div>
            </div>
          </td>
          <td>
            <div class="d-flex align-items-center">
              <span class="me-2">${this.getCategoryIcon(folderCategory)}</span>
              <div>
                <div class="fw-medium">${this.escapeHtml(folderName)}</div>
                ${folderCategory ? `<small class="text-muted">${folderCategory}</small>` : ''}
              </div>
            </div>
          </td>
          <td>
            <div class="d-flex align-items-center">
              ${isUnread ?
                '<span class="badge bg-warning text-dark"><i class="bi bi-envelope me-1"></i>Non lu</span>' :
                '<span class="badge bg-success"><i class="bi bi-envelope-open me-1"></i>Lu</span>'
              }
              ${(email.FlagStatus > 0) ? '<i class="bi bi-flag-fill text-danger ms-2" title="Marqué"></i>' : ''}
            </div>
          </td>
        </tr>
      `;
    }).join('');

    tbody.innerHTML = emailsHtml;
    
    // Mettre à jour les informations de pagination
    this.updateEmailsPagination(emails.length);
  }

  /**
   * Mise à jour des filtres d'emails
   */
  updateEmailFilters() {
    const folderFilter = document.getElementById('folder-filter');
    if (!folderFilter) return;

    // Récupérer les dossiers uniques depuis les emails
    const folders = [...new Set(this.state.recentEmails.map(email => 
      this.extractFolderName(email.FolderPath || 'Boîte de réception')
    ))];

    // Mettre à jour les options du filtre de dossier
    folderFilter.innerHTML = '<option value="">📁 Tous les dossiers</option>' +
      folders.map(folder => `<option value="${this.escapeHtml(folder)}">${this.escapeHtml(folder)}</option>`).join('');
  }

  /**
   * États d'affichage pour les emails
   */
  showEmptyEmailsState() {
    const tbody = document.getElementById('emails-table');
    if (!tbody) return;

    tbody.innerHTML = `
      <tr>
        <td colspan="5" class="text-center text-muted py-5 border-0">
          <div class="d-flex flex-column align-items-center">
            <i class="bi bi-inbox display-1 text-muted mb-3"></i>
            <h6 class="text-muted mb-2">Aucun email trouvé</h6>
            <p class="text-muted mb-0 small">Les emails apparaîtront ici au fur et à mesure qu'ils sont reçus</p>
          </div>
        </td>
      </tr>
    `;
  }

  showEmailsError() {
    const tbody = document.getElementById('emails-table');
    if (!tbody) return;

    tbody.innerHTML = `
      <tr>
        <td colspan="5" class="text-center text-muted py-5 border-0">
          <div class="d-flex flex-column align-items-center">
            <i class="bi bi-exclamation-triangle display-1 text-warning mb-3"></i>
            <h6 class="text-muted mb-2">Erreur de chargement</h6>
            <p class="text-muted mb-3 small">Impossible de charger les emails</p>
            <button class="btn btn-primary btn-sm" onclick="mailMonitor.loadRecentEmails()">
              <i class="bi bi-arrow-clockwise me-1"></i>Réessayer
            </button>
          </div>
        </td>
      </tr>
    `;
  }

  /**
   * Mise à jour des informations de pagination
   */
  updateEmailsPagination(totalCount) {
    const startEl = document.getElementById('emails-range-start');
    const endEl = document.getElementById('emails-range-end');
    const totalEl = document.getElementById('emails-total-count');

    if (startEl && endEl && totalEl) {
      const displayed = Math.min(totalCount, 50);
      startEl.textContent = totalCount > 0 ? '1' : '0';
      endEl.textContent = displayed;
      totalEl.textContent = totalCount;
    }
  }

  // === MÉTHODES UTILITAIRES POUR EMAILS ===

  /**
   * Calculer le temps écoulé depuis la réception
   */
  getTimeAgo(date) {
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMins / 60);
    const diffDays = Math.floor(diffHours / 24);

    if (diffMins < 1) return 'À l\'instant';
    if (diffMins < 60) return `${diffMins}min`;
    if (diffHours < 24) return `${diffHours}h`;
    if (diffDays < 7) return `${diffDays}j`;
    return date.toLocaleDateString();
  }

  /**
   * Extraire le nom du dossier depuis le chemin complet
   */
  extractFolderName(folderPath) {
    if (!folderPath) return 'Boîte de réception';
    const parts = folderPath.split('\\');
    return parts[parts.length - 1] || 'Boîte de réception';
  }

  /**
   * Obtenir la catégorie d'un dossier
   */
  getFolderCategory(folderPath) {
    if (!folderPath || !this.state.folderCategories) return null;
    const config = this.state.folderCategories[folderPath];
    return config ? config.category : null;
  }

  /**
   * Formater la taille des fichiers
   */
  formatFileSize(bytes) {
    if (!bytes) return '';
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
  }

  /**
   * Échapper le HTML pour éviter les injections
   */
  escapeHtml(text) {
    if (!text) return '';
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
  }

  // === ACTIONS SUR LES EMAILS ===

  /**
   * Marquer un email comme lu
   */
  async markAsRead(emailId) {
    try {
      console.log('📖 Marquage email comme lu:', emailId);
      // Cette fonctionnalité nécessiterait une implémentation côté serveur
      this.showNotification('Action en cours', 'Marquage comme lu...', 'info');
    } catch (error) {
      console.error('❌ Erreur marquage comme lu:', error);
      this.showNotification('Erreur', 'Impossible de marquer comme lu', 'danger');
    }
  }

  /**
   * Afficher les détails d'un email
   */
  async showEmailDetails(emailId) {
    const email = this.state.recentEmails.find(e => e.EntryID === emailId);
    if (!email) return;

    const modalHtml = `
      <div class="modal fade" id="emailDetailsModal" tabindex="-1">
        <div class="modal-dialog modal-lg">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title">
                <i class="bi bi-envelope me-2"></i>Détails de l'email
              </h5>
              <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
              <div class="row g-3">
                <div class="col-12">
                  <label class="form-label fw-bold">Sujet</label>
                  <p class="form-control-plaintext">${this.escapeHtml(email.Subject || 'Pas de sujet')}</p>
                </div>
                <div class="col-md-6">
                  <label class="form-label fw-bold">Expéditeur</label>
                  <p class="form-control-plaintext">${this.escapeHtml(email.SenderName || email.SenderEmailAddress || 'Inconnu')}</p>
                </div>
                <div class="col-md-6">
                  <label class="form-label fw-bold">Date de réception</label>
                  <p class="form-control-plaintext">${new Date(email.ReceivedTime).toLocaleString()}</p>
                </div>
                <div class="col-md-6">
                  <label class="form-label fw-bold">Taille</label>
                  <p class="form-control-plaintext">${this.formatFileSize(email.Size)}</p>
                </div>
                <div class="col-md-6">
                  <label class="form-label fw-bold">Statut</label>
                  <p class="form-control-plaintext">
                    ${email.UnRead ? 
                      '<span class="badge bg-warning text-dark">Non lu</span>' : 
                      '<span class="badge bg-success">Lu</span>'
                    }
                    ${email.HasAttachments ? '<span class="badge bg-info ms-2">Pièces jointes</span>' : ''}
                  </p>
                </div>
                <div class="col-12">
                  <label class="form-label fw-bold">Dossier</label>
                  <p class="form-control-plaintext">${this.escapeHtml(email.FolderPath || 'Boîte de réception')}</p>
                </div>
              </div>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Fermer</button>
            </div>
          </div>
        </div>
      </div>
    `;

    // Supprimer l'ancien modal s'il existe
    const existingModal = document.getElementById('emailDetailsModal');
    if (existingModal) existingModal.remove();

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    const modal = new bootstrap.Modal(document.getElementById('emailDetailsModal'));
    modal.show();
  }

  /**
   * Déplacer un email vers la corbeille
   */
  async moveToTrash(emailId) {
    if (!confirm('Êtes-vous sûr de vouloir supprimer cet email ?')) return;
    
    try {
      console.log('🗑️ Suppression email:', emailId);
      // Cette fonctionnalité nécessiterait une implémentation côté serveur
      this.showNotification('Action en cours', 'Suppression...', 'info');
    } catch (error) {
      console.error('❌ Erreur suppression:', error);
      this.showNotification('Erreur', 'Impossible de supprimer l\'email', 'danger');
    }
  }

  /**
   * Filtrage avancé des emails
   */
  filterEmails() {
    const searchTerm = document.getElementById('email-search')?.value.toLowerCase() || '';
    const folderFilter = document.getElementById('folder-filter')?.value || '';
    const dateFilter = document.getElementById('date-filter')?.value || '';
    const statusFilter = document.getElementById('status-filter')?.value || '';

    let filteredEmails = [...this.state.recentEmails];

    // Filtre par texte de recherche
    if (searchTerm) {
      filteredEmails = filteredEmails.filter(email => 
        (email.SenderName || '').toLowerCase().includes(searchTerm) ||
        (email.SenderEmailAddress || '').toLowerCase().includes(searchTerm) ||
        (email.Subject || '').toLowerCase().includes(searchTerm)
      );
    }

    // Filtre par dossier
    if (folderFilter) {
      filteredEmails = filteredEmails.filter(email => 
        this.extractFolderName(email.FolderPath) === folderFilter
      );
    }

    // Filtre par date
    if (dateFilter) {
      const now = new Date();
      filteredEmails = filteredEmails.filter(email => {
        const emailDate = new Date(email.ReceivedTime);
        switch (dateFilter) {
          case 'today':
            return emailDate.toDateString() === now.toDateString();
          case 'week':
            const startOfWeek = new Date(now);
            startOfWeek.setDate(now.getDate() - now.getDay());
            return emailDate >= startOfWeek;
          case 'month':
            return emailDate.getMonth() === now.getMonth() && 
                   emailDate.getFullYear() === now.getFullYear();
          default:
            return true;
        }
      });
    }

    // Filtre par statut
    if (statusFilter) {
      filteredEmails = filteredEmails.filter(email => {
        switch (statusFilter) {
          case 'unread':
            return email.UnRead;
          case 'read':
            return !email.UnRead;
          default:
            return true;
        }
      });
    }

    // Mettre à jour l'affichage avec les emails filtrés
    const originalEmails = this.state.recentEmails;
    this.state.recentEmails = filteredEmails;
    this.displayEmailsTable();
    this.state.recentEmails = originalEmails; // Restaurer la liste complète

    // Mettre à jour le compteur
    this.updateEmailsPagination(filteredEmails.length);
  }

  /**
   * Application de filtres rapides
   */
  applyQuickEmailFilter(filterId) {
    // Réinitialiser les autres filtres
    const filters = ['email-search', 'folder-filter', 'date-filter', 'status-filter'];
    filters.forEach(id => {
      const element = document.getElementById(id);
      if (element) element.value = '';
    });

    // Appliquer le filtre rapide
    switch (filterId) {
      case 'filter-all':
        this.displayEmailsTable();
        break;
      case 'filter-unread':
        document.getElementById('status-filter').value = 'unread';
        this.filterEmails();
        break;
      case 'filter-today':
        document.getElementById('date-filter').value = 'today';
        this.filterEmails();
        break;
    }
  }

  updateMonitoringStats(stats) {
    if (!stats) return;

    // Mettre à jour les compteurs principaux avec animations
    this.animateCounterUpdate('total-folders', stats.total || 0);
    this.animateCounterUpdate('active-folders', stats.active || 0);
    this.animateCounterUpdate('declarations-count', stats.declarations || 0);
    this.animateCounterUpdate('reglements-count', stats.reglements || 0);
    this.animateCounterUpdate('simples-count', stats.simples || 0);

    // Mettre à jour le badge du nombre de dossiers surveillés
    const countBadge = document.getElementById('monitored-folders-count');
    if (countBadge) {
      countBadge.textContent = stats.total || 0;
      countBadge.className = stats.total > 0 ? 'badge bg-primary rounded-pill' : 'badge bg-secondary rounded-pill';
    }

    // Calculer et mettre à jour l'utilisation du système
    const totalFolders = stats.total || 0;
    const activeFolders = stats.active || 0;
    const usage = totalFolders > 0 ? Math.round((activeFolders / totalFolders) * 100) : 0;
    
    const systemUsage = document.getElementById('system-usage');
    const systemProgress = document.getElementById('system-progress');
    
    if (systemUsage) systemUsage.textContent = `${usage}%`;
    if (systemProgress) {
      systemProgress.style.width = `${usage}%`;
      // Changer la couleur selon l'utilisation
      systemProgress.className = `progress-bar ${this.getUsageColorClass(usage)}`;
    }
  }

  // (supprimé) animateCounterUpdate dupliqué — on utilise la version unique définie plus haut

  getUsageColorClass(usage) {
    if (usage >= 80) return 'bg-danger';
    if (usage >= 60) return 'bg-warning';
    if (usage >= 40) return 'bg-info';
    return 'bg-success';
  }

  updateMonitoringStatus(isRunning) {
    const statusContainer = document.getElementById('monitoring-status');
    if (!statusContainer) return;
    // Minimal inline indicator only; hide when inactive
    if (isRunning) {
      statusContainer.style.display = 'inline-flex';
      statusContainer.innerHTML = `<span class="status-indicator status-connected me-1"></span><span>Actif</span>`;
    } else {
      statusContainer.style.display = 'none';
      statusContainer.innerHTML = '';
    }
    this.state.isMonitoring = !!isRunning;
  }

  updatePerformanceMetrics(isActive) {
    const syncFreq = document.getElementById('sync-frequency');
    const lastSync = document.getElementById('last-sync');
    const systemUsage = document.getElementById('system-usage');
    const systemProgress = document.getElementById('system-progress');
    
    if (isActive) {
      if (syncFreq) syncFreq.textContent = '60';
      if (lastSync) lastSync.textContent = 'maintenant';
      if (systemUsage) systemUsage.textContent = '25%';
      if (systemProgress) systemProgress.style.width = '25%';
    } else {
      if (syncFreq) syncFreq.textContent = '--';
      if (lastSync) lastSync.textContent = '--';
      if (systemUsage) systemUsage.textContent = '0%';
      if (systemProgress) systemProgress.style.width = '0%';
    }
  }

  updateCategoryDisplay(categories) {
    const container = document.getElementById('category-summary');
    if (!container) return;

    const categoryEntries = Object.entries(categories);
    
    if (categoryEntries.length === 0) {
      container.innerHTML = `
        <div class="text-center text-muted">
          <i class="bi bi-pie-chart"></i>
          <p class="mb-0">Aucune donnée par catégorie</p>
        </div>
      `;
      return;
    }

    container.innerHTML = categoryEntries.map(([name, data]) => `
      <div class="d-flex justify-content-between align-items-center mb-3">
        <div>
          <div class="fw-bold">${this.escapeHtml(name)}</div>
          <small class="text-muted">${data.emailsReceived || 0} reçus</small>
        </div>
        <span class="badge bg-primary">${data.unreadCount || 0}</span>
      </div>
    `).join('');
  }

  // === AFFICHAGE MÉTRIQUES VBA ===
  updateVBAMetricsDisplay(vbaMetrics) {
    // Log réduit pour métriques VBA
    // console.log('🎯 Mise à jour métriques VBA:', vbaMetrics);
    
    // Mise à jour sécurisée des éléments (vérifier qu'ils existent)
    const safeSetContent = (id, value) => {
      const element = document.getElementById(id);
      if (!element) return;
      // Préserver les compteurs: ne jamais remplacer 0 par "--" et ne pas écraser si valeur absente
      if (value === null || value === undefined) return;
      element.textContent = value;
    };
    
    // Métriques quotidiennes (utiliser les vrais ID du HTML)
    const daily = vbaMetrics.daily || {};
    // Mettre à jour sans réinitialiser: seulement si on a un nombre, et sans placeholder
    if (typeof daily.emailsReceived === 'number') {
      this.setCounterImmediate('emails-today', daily.emailsReceived);
    }
    // Ne pas écraser le compteur "Non lus" du tableau de bord qui est mis à jour en temps réel
    // if (typeof daily.emailsUnread === 'number') this.setCounterImmediate('emails-unread', daily.emailsUnread);
    
    // Note: emails-sent, stock-start, stock-end n'existent pas dans le DOM
    // Utiliser les éléments existants ou ignorer ces mises à jour
    
    // Métriques hebdomadaires - mappées vers éléments existants
    const weekly = vbaMetrics.weekly || {};
    // Ces éléments n'existent pas dans le DOM actuel, on les ignore
    // safeSetContent('stock-start', weekly.stockStart);
    // safeSetContent('stock-end', weekly.stockEnd);
    // safeSetContent('weekly-arrivals', weekly.arrivals);
    // safeSetContent('weekly-treatments', weekly.treatments);
    
    // Évolution
    const evolution = weekly.evolution || 0;
    const evolutionEl = document.getElementById('stock-evolution');
    if (evolutionEl) {
      evolutionEl.textContent = evolution > 0 ? `+${evolution}` : evolution;
      evolutionEl.className = `fs-4 fw-bold ${evolution > 0 ? 'text-success' : evolution < 0 ? 'text-danger' : 'text-muted'}`;
    }
    
    // Numéro de semaine
    const weekEl = document.getElementById('week-number');
    if (weekEl && weekly.weekNumber) {
      weekEl.textContent = `S${weekly.weekNumber} - ${weekly.year || new Date().getFullYear()}`;
    }

    // Affichage des compteurs par catégories VBA
    if (vbaMetrics.categories) {
      this.updateVBACategoryCounters(vbaMetrics.categories);
    }
  }

  updateVBACategoryCounters(categories) {
    console.log('🏷️ Mise à jour compteurs catégories VBA:', categories);
    
    // Mettre à jour chaque catégorie
    const categoryIds = ['declarations', 'reglements', 'mails_simples', 'autres'];
    
    categoryIds.forEach(categoryId => {
      const categoryData = categories[categoryId];
      if (categoryData) {
        // Total d'emails pour cette catégorie
        const totalCount = (categoryData.received || 0) + (categoryData.processed || 0);
        this.updateElement(`${categoryId.replace('_', '-')}-count`, totalCount);
        this.updateElement(`${categoryId.replace('_', '-')}-unread`, `${categoryData.unread || 0} non lus`);
        
        console.log(`📋 ${categoryData.name}: ${totalCount} total, ${categoryData.unread || 0} non lus`);
      } else {
        // Catégorie vide
        this.updateElement(`${categoryId.replace('_', '-')}-count`, 0);
        this.updateElement(`${categoryId.replace('_', '-')}-unread`, '0 non lus');
      }
    });
  }

  updateVBAWeeklyEvolution(weeklyEvolution) {
    console.log('📈 Mise à jour évolution hebdo:', weeklyEvolution);
    
    const percentageEl = document.getElementById('evolution-percentage');
    const indicatorEl = document.getElementById('evolution-indicator');
    
    if (percentageEl && weeklyEvolution.percentage !== undefined) {
      const percentage = parseFloat(weeklyEvolution.percentage);
      percentageEl.textContent = `${percentage > 0 ? '+' : ''}${percentage}%`;
    }
    
    if (indicatorEl && weeklyEvolution.trend !== undefined) {
      const trend = weeklyEvolution.trend;
      if (trend > 0) {
        indicatorEl.className = 'badge bg-success';
        indicatorEl.textContent = 'En hausse';
      } else if (trend < 0) {
        indicatorEl.className = 'badge bg-danger';
        indicatorEl.textContent = 'En baisse';
      } else {
        indicatorEl.className = 'badge bg-secondary';
        indicatorEl.textContent = 'Stable';
      }
    }
  }

  updateVBAFolderDistribution(distribution) {
    console.log('📁 Mise à jour distribution VBA:', distribution);
    
    const container = document.getElementById('vba-category-summary');
    if (!container) return;

    const distributionEntries = Object.entries(distribution);
    
    if (distributionEntries.length === 0) {
      container.innerHTML = `
        <div class="text-center text-muted">
          <i class="bi bi-folder"></i>
          <p class="mb-0">Aucune donnée de distribution</p>
        </div>
      `;
      return;
    }

    container.innerHTML = distributionEntries.map(([category, data]) => `
      <div class="row g-3 mb-3 p-3 border rounded">
        <div class="col-md-3">
          <div class="text-center">
            <div class="fw-bold text-primary">${category}</div>
            <small class="text-muted">${data.folders} dossier(s)</small>
          </div>
        </div>
        <div class="col-md-3">
          <div class="text-center">
            <div class="fs-5 fw-bold">${data.totalEmails || 0}</div>
            <small class="text-muted">Total emails</small>
          </div>
        </div>
        <div class="col-md-3">
          <div class="text-center">
            <div class="fs-5 fw-bold text-warning">${data.unreadEmails || 0}</div>
            <small class="text-muted">Non lus</small>
          </div>
        </div>
        <div class="col-md-3">
          <div class="text-center">
            <div class="fs-5 fw-bold text-success">${data.processedToday || 0}</div>
            <small class="text-muted">Traités aujourd'hui</small>
          </div>
        </div>
      </div>
    `).join('');
  }

  // === ACTIONS ===
  async refreshFoldersDisplay() {
    try {
      // Dossiers refresh - log désactivé
      // console.log('🔄 Actualisation de l\'affichage des dossiers...');
      
      // DEBUG: État avant rechargement
      // État avant - log désactivé pour réduire le spam
      // console.log('🔍 État AVANT rechargement:', Object.keys(this.state.folderCategories).length, 'dossiers');
      
      // Recharger la configuration depuis la base de données
      await this.loadFoldersConfiguration();
      
      // DEBUG: État après rechargement
      // État après - log désactivé pour réduire le spam
      // console.log('🔍 État APRÈS rechargement:', Object.keys(this.state.folderCategories).length, 'dossiers');
      console.log('🔍 Données complètes:', this.state.folderCategories);
      
      // Mettre à jour l'affichage moderne
      this.updateFolderConfigDisplay();
      
      // Recharger les statistiques aussi
      await this.loadStats();
      
      this.showNotification('Actualisation terminée', 'La liste des dossiers a été mise à jour', 'success');
    } catch (error) {
      console.error('❌ Erreur lors de l\'actualisation:', error);
      this.showNotification('Erreur', 'Impossible d\'actualiser la liste', 'danger');
    }
  }

  async forceFullResync() {
    // Demander confirmation avant de lancer la resynchronisation
    const result = await this.showConfirmModal(
      'Resynchronisation complète',
      'Cette opération va réinitialiser l\'état de synchronisation de tous les dossiers et récupérer à nouveau tout l\'historique des emails. Cela peut prendre plusieurs minutes. Continuer ?',
      'Resynchroniser',
      'warning'
    );
    
    if (!result) {
      return;
    }

    try {
      // Désactiver le bouton pendant la resync
      const btn = document.getElementById('force-resync');
      if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="bi bi-hourglass-split me-1"></i> Resync en cours...';
      }

      this.showNotification('Resynchronisation', 'Resynchronisation complète en cours...', 'info');
      
      console.log('🔄 Démarrage resynchronisation complète forcée...');
      const response = await window.api.invoke('api-force-full-resync');
      
      if (response.success) {
        console.log('✅ Resynchronisation terminée:', response.stats);
        this.showNotification(
          'Resynchronisation terminée',
          `${response.stats.emailsAdded} emails ajoutés, ${response.stats.emailsUpdated} mis à jour`,
          'success'
        );
        
        // Rafraîchir l'affichage
        await this.refreshFoldersDisplay();
        await this.loadStats();
      } else {
        throw new Error(response.message || 'Échec de la resynchronisation');
      }
    } catch (error) {
      console.error('❌ Erreur lors de la resynchronisation:', error);
      this.showNotification(
        'Erreur',
        `Erreur lors de la resynchronisation: ${error.message}`,
        'danger'
      );
    } finally {
      // Réactiver le bouton
      const btn = document.getElementById('force-resync');
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-arrow-repeat me-1"></i> Resync complète';
      }
    }
  }

  async removeFolderConfig(folderPath) {
    // Utiliser une modal Bootstrap au lieu du confirm() natif
    const result = await this.showConfirmModal(
      'Supprimer la configuration',
      'Êtes-vous sûr de vouloir supprimer cette configuration de surveillance ?',
      'Supprimer',
      'danger'
    );
    
    if (result) {
      try {
        // CORRECTION: Mise à jour visuelle IMMÉDIATE pour une réactivité instantanée
        // 1. Supprimer immédiatement de l'état local pour l'affichage
        delete this.state.folderCategories[folderPath];
        
        // 2. Mettre à jour l'affichage IMMÉDIATEMENT
        this.updateFolderConfigDisplay();
        
  // Notification en cours de suppression retirée pour éviter le bruit visuel
        
        // 4. Puis faire l'appel API en arrière-plan
        await window.electronAPI.removeFolderFromMonitoring({ folderPath });
        
        // 5. Recharger la configuration depuis la base pour s'assurer de la cohérence
        await this.loadFoldersConfiguration();
        
        // 6. Rafraîchir l'affichage avec les données actualisées
        this.updateFolderConfigDisplay();
        // 6bis. Forcer la vue Monitoring à se recharger pour refléter la suppression
        if (window.foldersTree && typeof window.foldersTree.loadFolders === 'function') {
          try { await window.foldersTree.loadFolders(true); } catch (_) {}
        }
        await this.refreshFoldersDisplay();
        
        // 7. Attendre un peu puis rafraîchir les statistiques
        setTimeout(async () => {
          try {
            if (typeof this.refreshStats === 'function') {
              await this.refreshStats();
            }
            if (typeof this.refreshEmails === 'function') {
              await this.refreshEmails();
            }
          } catch (e) {
            console.warn('Post-delete refresh skipped:', e?.message || e);
          }
        }, 1000);
        
        this.showNotification('Configuration supprimée', 'Le dossier a été retiré de la surveillance et le service redémarré', 'success');
      } catch (error) {
        console.error('❌ Erreur suppression configuration:', error);
        
        // CORRECTION: En cas d'erreur, recharger la configuration pour restaurer l'état correct
        await this.loadFoldersConfiguration();
        this.updateFolderConfigDisplay();
        
        this.showNotification('Erreur', 'Impossible de supprimer la configuration: ' + error.message, 'danger');
      }
    }
  }

  async saveSettings(e) {
    e.preventDefault();
    
    try {
      // Visibilité des onglets
      const tabVisibility = {
        dashboard: !!document.getElementById('tab-dashboard')?.checked,
        emails: !!document.getElementById('tab-emails')?.checked,
        weekly: !!document.getElementById('tab-weekly')?.checked,
        personalPerformance: !!document.getElementById('tab-personal-performance')?.checked,
        importActivity: !!document.getElementById('tab-import-activity')?.checked,
        monitoring: !!document.getElementById('tab-monitoring')?.checked,
        db: !!document.getElementById('tab-db')?.checked,
      };

      // Récupérer les valeurs du formulaire
      const settings = {
        monitoring: {
          // Options de monitoring simplifiées (UI desactivée)
          treatReadEmailsAsProcessed: false,
          scanInterval: parseInt(document.getElementById('sync-interval')?.value) || 30000,
          autoStart: true
        },
        ui: {
          theme: (document.getElementById('theme-select')?.value || this.state.ui?.theme || 'light'),
          language: "fr",
          emailsLimit: parseInt(document.getElementById('emails-limit')?.value) || 20,
          tabs: tabVisibility
        },
        database: {
          purgeOldDataAfterDays: 365,
          enableEventLogging: true
        },
        notifications: {
          showStartupNotification: true,
          showMonitoringStatus: true,
          enableDesktopNotifications: document.getElementById('notifications-enabled')?.checked || false
        }
      };
      
      // Sauvegarder via IPC
      const result = await window.electronAPI.saveAppSettings(settings);
      
      if (result.success) {
        this.showNotification('Paramètres sauvegardés', 'Redémarrage de l’interface pour appliquer les onglets…', 'success');
        console.log('✅ Paramètres sauvegardés:', settings);
        // Recharger l'UI pour appliquer immédiatement la nouvelle visibilité des onglets
        setTimeout(() => { window.location.reload(); }, 400);
      } else {
        throw new Error(result.error || 'Erreur de sauvegarde');
      }
    } catch (error) {
      console.error('❌ Erreur sauvegarde paramètres:', error);
      this.showNotification('Erreur de sauvegarde', error.message, 'danger');
    }
  }

  async resetSettings() {
    if (confirm('Êtes-vous sûr de vouloir réinitialiser tous les paramètres ?')) {
      try {
        // Paramètres par défaut
        const defaultSettings = {
          monitoring: {
            treatReadEmailsAsProcessed: false,
            scanInterval: 30000,
            autoStart: true
          },
          ui: {
            theme: "light",
            language: "fr",
            emailsLimit: 20,
            // Rétablir la visibilité par défaut: tous les onglets visibles
            tabs: {
              dashboard: true,
              emails: true,
              weekly: true,
              personalPerformance: true,
              importActivity: true,
              monitoring: true,
              // IMPORTANT: l'onglet BDD est désactivé par défaut
              db: false
            }
          },
          database: {
            purgeOldDataAfterDays: 365,
            enableEventLogging: true
          },
          notifications: {
            showStartupNotification: true,
            showMonitoringStatus: true,
            enableDesktopNotifications: false
          }
        };
        
        // Sauvegarder les paramètres par défaut
        const result = await window.electronAPI.saveAppSettings(defaultSettings);
        
        if (result.success) {
          // Mettre à jour le formulaire puis recharger l'UI pour appliquer immédiatement
          this.loadSettingsIntoForm(defaultSettings);
          this.showNotification('Paramètres réinitialisés', 'Redémarrage de l’interface pour appliquer les valeurs par défaut…', 'info');
          setTimeout(() => { window.location.reload(); }, 400);
        } else {
          throw new Error(result.error || 'Erreur de réinitialisation');
        }
      } catch (error) {
        console.error('❌ Erreur réinitialisation paramètres:', error);
        this.showNotification('Erreur de réinitialisation', error.message, 'danger');
      }
    }
  }

  async loadSettings() {
    try {
      const result = await window.electronAPI.loadAppSettings();
      
      if (result.success) {
        this.loadSettingsIntoForm(result.settings);
        console.log('📄 Paramètres chargés:', result.settings);
        // Appliquer la visibilité des onglets au chargement
        this.applyTabVisibility(result.settings?.ui?.tabs);
  // Appliquer le thème au chargement
  {
    const tRaw = (result.settings?.ui?.theme) || 'light';
    const t = (tRaw === 'dark') ? 'dark' : 'light';
    this.state.ui.theme = t;
    this.applyTheme(t);
    // Sync toggle button UI
    try {
      const icon = document.getElementById('theme-toggle-icon');
      const label = document.getElementById('theme-toggle-label');
      if (icon && label) {
        icon.className = t === 'dark' ? 'bi bi-moon' : 'bi bi-sun';
        label.textContent = t === 'dark' ? 'Sombre' : 'Clair';
      }
    } catch(_) {}
  }
      } else {
        console.warn('⚠️ Erreur chargement paramètres, utilisation des valeurs par défaut');
        // Valeurs par défaut: tous les onglets visibles
        this.applyTabVisibility({
          dashboard: true,
          emails: true,
          weekly: true,
          personalPerformance: true,
          importActivity: true,
          monitoring: true,
          db: false
        });
  // Thème par défaut
  this.applyTheme('light');
      }
    } catch (error) {
      console.error('❌ Erreur chargement paramètres:', error);
    }
  }

  loadSettingsIntoForm(settings) {
    // Monitoring
    if (settings.monitoring) {
  // UI desactivée pour 'treatReadEmailsAsProcessed'
      
      const syncInterval = document.getElementById('sync-interval');
      if (syncInterval) {
        syncInterval.value = settings.monitoring.scanInterval || 30000;
      }
      
  // UI desactivée pour auto-start (forcée à true)
    }
    
    // UI
    if (settings.ui) {
      const emailsLimit = document.getElementById('emails-limit');
      if (emailsLimit) {
        emailsLimit.value = settings.ui.emailsLimit || 20;
      }
      // Thème
      {
        const allowedThemes = new Set(['light','dark']);
        const themeRaw = settings.ui.theme || 'light';
        const theme = allowedThemes.has(themeRaw) ? themeRaw : 'light';
        this.applyTheme(theme);
        this.state.ui.theme = theme;
        // Sync toggle button UI
        const icon = document.getElementById('theme-toggle-icon');
        const label = document.getElementById('theme-toggle-label');
        if (icon && label) {
          icon.className = theme === 'dark' ? 'bi bi-moon' : 'bi bi-sun';
          label.textContent = theme === 'dark' ? 'Sombre' : 'Clair';
        }
      }
      // Visibilité des onglets
      const tabs = settings.ui.tabs || {};
      const setCk = (id, def=true) => { const el = document.getElementById(id); if (el) el.checked = (tabs?.[id.replace('tab-','').replace('-','')] ?? def); };
      // Set explicit per key to avoid mapping confusion
      const map = [
        ['tab-dashboard','dashboard'],
        ['tab-emails','emails'],
        ['tab-weekly','weekly'],
        ['tab-personal-performance','personalPerformance'],
        ['tab-import-activity','importActivity'],
        ['tab-monitoring','monitoring'],
        ['tab-db','db']
      ];
      for (const [domId, key] of map) {
        const el = document.getElementById(domId);
        if (el) {
          const defaultChecked = (key === 'db') ? false : true;
          el.checked = (tabs?.[key] !== undefined ? !!tabs[key] : defaultChecked);
        }
      }
    }
    
    // Notifications
    if (settings.notifications) {
      const notificationsCheckbox = document.getElementById('notifications-enabled');
      if (notificationsCheckbox) {
        notificationsCheckbox.checked = settings.notifications.enableDesktopNotifications || false;
      }
    }
  }

  // Applique immédiatement le thème en ajoutant l'attribut data-theme au body
  applyTheme(theme) {
    const allowed = new Set(['light','dark']);
    const t = allowed.has(theme) ? theme : 'light';
    try {
      document.body?.setAttribute('data-theme', t);
      this.state.ui.theme = t;
  try { localStorage.setItem('uiTheme', t); } catch(_) {}
    } catch (e) {
      console.warn('⚠️ applyTheme:', e?.message || e);
    }
  }

  // Persiste uniquement le thème dans les paramètres sans modifier le reste
  async persistTheme(theme) {
    try {
  const allowed = new Set(['light','dark']);
      const t = allowed.has(theme) ? theme : 'light';
  try { localStorage.setItem('uiTheme', t); } catch(_) {}
      const current = await window.electronAPI.loadAppSettings();
      const settings = (current?.success && current.settings) ? current.settings : {};
      // Valeurs par défaut minimales
      settings.monitoring = settings.monitoring || { treatReadEmailsAsProcessed: false, scanInterval: 30000, autoStart: true };
      settings.ui = Object.assign({ language: 'fr', emailsLimit: 20, tabs: { dashboard: true, emails: true, weekly: true, personalPerformance: true, importActivity: true, monitoring: true } }, settings.ui || {});
      settings.ui.theme = t;
      settings.database = settings.database || { purgeOldDataAfterDays: 365, enableEventLogging: true };
      settings.notifications = settings.notifications || { showStartupNotification: true, showMonitoringStatus: true, enableDesktopNotifications: false };
      const res = await window.electronAPI.saveAppSettings(settings);
      if (!res?.success) throw new Error(res?.error || 'Échec de sauvegarde');
      this.showNotification('Thème appliqué', 'Votre thème a été enregistré', 'success');
    } catch (e) {
      console.error('❌ persistTheme:', e);
      this.showNotification('Erreur', 'Impossible d’enregistrer le thème', 'danger');
    }
  }

  // Applique la visibilité des onglets (sidebar + contenu)
  applyTabVisibility(tabs) {
    // Defaults: show all if undefined
    const conf = Object.assign({
      dashboard: true,
      emails: true,
      weekly: true,
      personalPerformance: true,
      importActivity: true,
      monitoring: true,
      db: false,
    }, tabs || {});

    const setVisible = (navSelector, paneSelector, visible) => {
      const nav = document.querySelector(`.nav-pills a[data-bs-target="${navSelector}"]`);
      if (nav) nav.parentElement.style.display = visible ? '' : 'none';
      const pane = document.querySelector(paneSelector);
      if (pane) pane.style.display = visible ? '' : 'none';
      // If the active tab is hidden, switch to the first visible tab
      const activeNav = document.querySelector('.nav-pills .nav-link.active');
      if (activeNav && activeNav.getAttribute('data-bs-target') === navSelector && !visible) {
        const firstVisible = Array.from(document.querySelectorAll('.nav-pills .nav-link'))
          .find(a => a.parentElement.style.display !== 'none');
        if (firstVisible) firstVisible.click();
      }
    };

    setVisible('#dashboard', '#dashboard', conf.dashboard);
    setVisible('#emails', '#emails', conf.emails);
    setVisible('#weekly', '#weekly', conf.weekly);
    setVisible('#personal-performance', '#personal-performance', conf.personalPerformance);
    setVisible('#import-activity', '#import-activity', conf.importActivity);
    setVisible('#monitoring', '#monitoring', conf.monitoring);
    setVisible('#db', '#db', conf.db);
  }

  handleTabChange(target) {
    // Recharger les données spécifiques à l'onglet immédiatement
    // Changement d'onglet - log conservé pour navigation
    console.log(`🔄 Changement d'onglet vers: ${target}`);

    // Lifecycle: BDD auto-refresh only when active
    if (target !== '#db') {
      try { this._dbStop?.(); } catch (_) {}
    }
    
    switch (target) {
      case '#dashboard':
        this.loadStats();
        this.loadCategoryStats();
        this.loadVBAMetrics();
        break;
      case '#emails':
        this.loadRecentEmails();
        break;
      case '#personal-performance':
        this.loadPersonalPerformance();
        break;
      case '#monitoring':
        // Configuration déjà chargée
        break;
      case '#db':
        try {
          this._dbStart?.();
          this._dbRefreshNow?.();
        } catch (_) {}
        break;
    }
  }

  // === UTILITAIRES ===
  startPeriodicUpdates() {
    console.log('🔄 Démarrage des mises à jour périodiques unifiées...');
    // Cette méthode n'est plus nécessaire car setupAutoRefresh() gère tout
    // Conservée pour compatibilité mais ne fait rien
    console.log('ℹ️ Actualisation automatique déjà configurée via setupAutoRefresh()');
  }

  stopPeriodicUpdates() {
    console.log('🛑 Arrêt de toutes les actualisations automatiques...');
    
    // Arrêter tous les intervalles d'actualisation
    if (this.statsRefreshInterval) {
      clearInterval(this.statsRefreshInterval);
      this.statsRefreshInterval = null;
    }
    
    if (this.emailsRefreshInterval) {
      clearInterval(this.emailsRefreshInterval);
      this.emailsRefreshInterval = null;
    }
    
    if (this.fullRefreshInterval) {
      clearInterval(this.fullRefreshInterval);
      this.fullRefreshInterval = null;
    }
    
    if (this.vbaRefreshInterval) {
      clearInterval(this.vbaRefreshInterval);
      this.vbaRefreshInterval = null;
    }
    
    // Ancien interval (pour compatibilité)
    if (this.updateInterval) {
      clearInterval(this.updateInterval);
      this.updateInterval = null;
    }
    
    console.log('✅ Toutes les actualisations automatiques arrêtées');
  }

  showNotification(title, message, type = 'info') {
  // Notifications visuelles désactivées selon la préférence utilisateur
  // Conserver une trace console légère pour le debug si nécessaire
  // console.debug('[notification]', { title, message, type });
  return; // no-op
  }

  showConfirmModal(title, message, confirmText = 'Confirmer', variant = 'primary') {
    return new Promise((resolve) => {
      const modalId = 'confirmModal' + Date.now();
      const modalHtml = `
        <div class="modal fade" id="${modalId}" tabindex="-1" aria-labelledby="${modalId}Label" aria-hidden="true">
          <div class="modal-dialog">
            <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="${modalId}Label">${this.escapeHtml(title)}</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
              </div>
              <div class="modal-body">
                ${this.escapeHtml(message)}
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Annuler</button>
                <button type="button" class="btn btn-${variant}" id="${modalId}Confirm">${this.escapeHtml(confirmText)}</button>
              </div>
            </div>
          </div>
        </div>
      `;

      document.body.insertAdjacentHTML('beforeend', modalHtml);
      
      const modalElement = document.getElementById(modalId);
      const modal = new bootstrap.Modal(modalElement);
      
      // Gérer la confirmation
      document.getElementById(modalId + 'Confirm').addEventListener('click', () => {
        modal.hide();
        resolve(true);
      });
      
      // Gérer l'annulation
      modalElement.addEventListener('hidden.bs.modal', () => {
        modalElement.remove();
        resolve(false);
      });
      
      modal.show();
    });
  }

  escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // === NETTOYAGE ===
  destroy() {
    this.stopPeriodicUpdates();
    
    // Nettoyage des intervalles d'auto-refresh
    if (this.lightRefreshInterval) {
      clearInterval(this.lightRefreshInterval);
      this.lightRefreshInterval = null;
    }
    
    if (this.fullRefreshInterval) {
      clearInterval(this.fullRefreshInterval);
      this.fullRefreshInterval = null;
    }
    
    if (this.vbaRefreshInterval) {
      clearInterval(this.vbaRefreshInterval);
      this.vbaRefreshInterval = null;
    }
    
    console.log('🧹 Mail Monitor nettoyé (intervalles auto-refresh arrêtés)');
  }

  // === INFORMATIONS COPYRIGHT ===
  async showAbout() {
    const appVersion = (window.electronAPI && await window.electronAPI.getAppVersion()) || '—';
    const aboutContent = `
      <div class="text-center">
        <i class="bi bi-envelope-check display-4 text-primary mb-3"></i>
        <h4>Mail Monitor</h4>
        <p class="text-muted">Surveillance professionnelle des emails Outlook</p>
        <hr>
  <p><strong>Version:</strong> ${appVersion}</p>
        <p><strong>Auteur:</strong> Tanguy Raingeard</p>
        <p><strong>Copyright:</strong> © 2025 Tanguy Raingeard</p>
        <p><small class="text-muted">Tous droits réservés</small></p>
        <hr>
        <p><small>Ce logiciel est protégé par le droit d'auteur français et international.</small></p>
      </div>
    `;
    
    // Utiliser Bootstrap modal ou une simple alert selon le contexte
    if (typeof bootstrap !== 'undefined' && bootstrap.Modal) {
      // Créer une modal Bootstrap si disponible
      const modalDiv = document.createElement('div');
      modalDiv.innerHTML = `
        <div class="modal fade" id="aboutModal" tabindex="-1">
          <div class="modal-dialog">
            <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title">À propos de Mail Monitor</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
              </div>
              <div class="modal-body">
                ${aboutContent}
              </div>
              <div class="modal-footer">
                <button type="button" class="btn btn-primary" data-bs-dismiss="modal">OK</button>
              </div>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(modalDiv);
      const modal = new bootstrap.Modal(modalDiv.querySelector('#aboutModal'));
      modal.show();
      
      // Supprimer la modal après fermeture
      modalDiv.querySelector('#aboutModal').addEventListener('hidden.bs.modal', () => {
        document.body.removeChild(modalDiv);
      });
    } else {
  // Fallback vers alert
  alert(`Mail Monitor v${appVersion}\n© 2025 Tanguy Raingeard - Tous droits réservés`);
    }
  }

  // === FONCTIONS SUIVI HEBDOMADAIRE ===

  /**
   * Initialise le suivi hebdomadaire
   */
  async initWeeklyTracking() {
    console.log('📅 Initialisation du suivi hebdomadaire...');
    
    try {
      // Charger les paramètres de suivi hebdomadaire
      await this.loadWeeklySettings();
      
      // Charger les statistiques de la semaine actuelle
      await this.loadCurrentWeekStats();
      
      // Charger l'historique
      await this.loadWeeklyHistory();
      
      // Configurer les événements
      this.setupWeeklyEventListeners();

  // Préparer les commentaires hebdomadaires
  await this.populateWeeksForComments?.();
  await this.refreshCommentsList?.();
      
      console.log('✅ Suivi hebdomadaire initialisé');
    } catch (error) {
      console.error('❌ Erreur lors de l\'initialisation du suivi hebdomadaire:', error);
      this.showNotification('Erreur suivi hebdomadaire', error.message, 'danger');
    }
  }

  /**
   * Configure les événements pour le suivi hebdomadaire
   */
  setupWeeklyEventListeners() {
    // Formulaire d'ajustement manuel (submit)
    const manualForm = document.getElementById('manual-adjustment-form');
    if (manualForm) {
      manualForm.addEventListener('submit', (e) => {
        e.preventDefault();
        this.handleManualAdjustment(1);
      });
    }

    // Bouton retirer (ajustement négatif)
    const negativeBtn = document.getElementById('negative-adjustment');
    if (negativeBtn) {
      negativeBtn.addEventListener('click', (e) => {
        e.preventDefault();
        this.handleManualAdjustment(-1);
      });
    }

    // Bouton de sauvegarde des paramètres
    const saveSettingsBtn = document.getElementById('save-weekly-settings');
    if (saveSettingsBtn) {
      saveSettingsBtn.addEventListener('click', () => this.saveWeeklySettings());
    }

    // Charger les paramètres à l'ouverture de la modal
    const weeklySettingsModal = document.getElementById('weeklySettingsModal');
    if (weeklySettingsModal) {
      weeklySettingsModal.addEventListener('show.bs.modal', () => this.loadWeeklySettings());
    }

  // Commentaires hebdo: ajout
    const addCommentBtn = document.getElementById('add-comment-btn');
    if (addCommentBtn) {
      addCommentBtn.addEventListener('click', async (e) => {
        e.preventDefault();
        const weekSel = document.getElementById('comments-week-select');
        const txtEl = document.getElementById('comment-text');
        const week_identifier = weekSel?.value || this.state.currentWeekInfo?.identifier;
        const comment_text = (txtEl?.value || '').trim();
        if (!week_identifier || !comment_text) { this.showNotification('Erreur', 'Semaine ou texte manquant', 'warning'); return; }
        try {
          const m = String(week_identifier).match(/S(\d+).*?(\d{4})/);
          const week_number = m ? parseInt(m[1],10) : undefined;
          const week_year = m ? parseInt(m[2],10) : undefined;
      const res = await window.electronAPI.addWeeklyComment({ week_identifier, week_year, week_number, comment_text });
          if (res?.success) {
            if (txtEl) txtEl.value = '';
            await this.refreshCommentsList?.();
            this.showNotification('Commentaire ajouté', 'Votre note a été enregistrée', 'success');
          } else {
            this.showNotification('Erreur', res?.error || 'Ajout impossible', 'danger');
          }
        } catch (err) { this.showNotification('Erreur', err?.message || 'Ajout impossible', 'danger'); }
      });
    }

    // Commentaires hebdo: refresh
    const refreshBtn = document.getElementById('refresh-comments');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', async () => {
        await this.populateWeeksForComments?.();
        await this.refreshCommentsList?.();
      });
    }

  // Commentaires hebdo: changer de semaine
    const weekSelect = document.getElementById('comments-week-select');
    if (weekSelect) {
      weekSelect.addEventListener('change', async () => {
        await this.refreshCommentsList?.();
    const dup = document.getElementById('comments-week-select-dup');
    if (dup) { dup.innerHTML = weekSelect.innerHTML; dup.value = weekSelect.value; }
    const countersWeek = document.getElementById('counters-week-select');
    if (countersWeek) { countersWeek.innerHTML = weekSelect.innerHTML; countersWeek.value = weekSelect.value; }
      });
    }

  // Bouton de rafraîchissement manuel supprimé: l'onglet se met à jour automatiquement (événements + intervalle)

    // Actualisation automatique des stats hebdomadaires (toujours, même si l'onglet n'est pas actif)
    setInterval(() => {
      this.loadCurrentWeekStats();
      this.loadWeeklyHistory();
      this.loadPersonalPerformance?.();
    }, 30000); // Toutes les 30 secondes

    // Pagination historique hebdomadaire
    const prevBtn = document.getElementById('weekly-prev');
    const nextBtn = document.getElementById('weekly-next');
    if (prevBtn) {
      prevBtn.addEventListener('click', async () => {
        if (this.state.weeklyHistory.page > 1) {
          this.state.weeklyHistory.page -= 1;
          await this.loadWeeklyHistory();
        }
      });
    }
    if (nextBtn) {
      nextBtn.addEventListener('click', async () => {
        if (this.state.weeklyHistory.page < this.state.weeklyHistory.totalPages) {
          this.state.weeklyHistory.page += 1;
          await this.loadWeeklyHistory();
        }
      });
    }
  }

  /**
   * Remplit le sélecteur de semaines pour les commentaires
   */
  async populateWeeksForComments(limit = 24) {
    try {
      const select = document.getElementById('comments-week-select');
      if (!select || !window.electronAPI?.listWeeksForComments) return;

      // Mémoriser la valeur sélectionnée si existante
      const previous = select.value;

      const res = await window.electronAPI.listWeeksForComments(limit);
      if (!res?.success) return;

      const rows = Array.isArray(res.rows) ? res.rows : [];
      if (rows.length === 0) {
        select.innerHTML = '';
        return;
      }

  // Construire les options
      const options = rows.map(r => {
        const id = r.week_identifier || `S${r.week_number}-${r.week_year}`;
        const label = `S${r.week_number} - ${r.week_year}` +
          (r.week_start_date && r.week_end_date ? ` (${r.week_start_date} - ${r.week_end_date})` : '');
        return `<option value="${this.escapeHtml(id)}">${this.escapeHtml(label)}</option>`;
      }).join('');
  select.innerHTML = options;
  // Synchroniser les autres sélecteurs liés (dup + compteurs)
  const dup = document.getElementById('comments-week-select-dup');
  const counters = document.getElementById('counters-week-select');
  if (dup) dup.innerHTML = options;
  if (counters) counters.innerHTML = options;

      // Choisir la meilleure sélection: précédente > semaine courante > première
      const currentId = this.state?.currentWeekInfo?.identifier;
      const desired = previous || currentId || (rows[0]?.week_identifier);
  if (desired) {
        const opt = Array.from(select.options).find(o => o.value === desired);
        if (opt) select.value = desired; else select.selectedIndex = 0;
      } else {
        select.selectedIndex = 0;
      }

  // Appliquer la même sélection aux duplicatas si présents
  if (dup) dup.value = select.value;
  if (counters) counters.value = select.value;

      return select.value;
    } catch (e) {
      console.warn('Erreur populateWeeksForComments:', e);
    }
  }

  /**
   * Rafraîchit l'affichage de la liste des commentaires pour la semaine sélectionnée
   */
  async refreshCommentsList() {
    try {
      const listEl = document.getElementById('weekly-comments-list');
      const select = document.getElementById('comments-week-select');
      if (!listEl) return;

      // Déterminer la semaine ciblée
      let weekId = select?.value || this.state?.currentWeekInfo?.identifier;
      if (!weekId && window.electronAPI?.invoke) {
        try {
          const resp = await window.electronAPI.invoke('api-weekly-current-stats');
          if (resp?.success) weekId = resp.weekInfo?.identifier;
        } catch {}
      }
  if (!weekId) {
        listEl.innerHTML = '<li class="list-group-item text-muted text-center">Aucune semaine sélectionnée</li>';
        return;
      }

      // Etat de chargement
      listEl.innerHTML = '<li class="list-group-item text-center"><span class="spinner-border spinner-border-sm me-2"></span>Chargement…</li>';

      // Charger les commentaires
      const res = await window.electronAPI.listWeeklyComments?.(weekId);
      if (!res?.success) {
        listEl.innerHTML = `<li class="list-group-item text-danger text-center">Erreur: ${this.escapeHtml(res?.error || 'chargement impossible')}</li>`;
        return;
      }

  const rows = Array.isArray(res.rows) ? res.rows : [];
      if (rows.length === 0) {
        listEl.innerHTML = '<li class="list-group-item text-muted text-center">Aucun commentaire pour cette semaine</li>';
        return;
      }

      // Helper pour affichage date courte
      const fmtDate = (s) => {
        if (!s) return '';
        // s attendu comme 'YYYY-MM-DD HH:MM:SS'
        return s.slice(0, 16).replace('T', ' ');
      };

      // Rendu liste
  listEl.innerHTML = rows.map(r => {
        const cat = r.category ? `<span class="badge rounded-pill bg-secondary me-2">${this.escapeHtml(r.category)}</span>` : '';
        const meta = `<small class="text-muted">${this.escapeHtml(fmtDate(r.created_at))}${r.author ? ' • ' + this.escapeHtml(r.author) : ''}</small>`;
        return `
          <li class="list-group-item d-flex justify-content-between align-items-start" data-id="${r.id}">
            <div class="ms-2 me-auto">
              <div class="fw-semibold">${cat}${this.escapeHtml(r.comment_text || '')}</div>
              ${meta}
            </div>
            <div class="btn-group btn-group-sm align-self-center" role="group">
              <button type="button" class="btn btn-outline-secondary btn-edit-comment" title="Modifier"><i class="bi bi-pencil"></i></button>
              <button type="button" class="btn btn-outline-danger btn-delete-comment" title="Supprimer"><i class="bi bi-trash"></i></button>
            </div>
          </li>`;
      }).join('');

      // Attacher les événements des boutons
      const attach = (selector, handler) => {
        listEl.querySelectorAll(selector).forEach(btn => btn.addEventListener('click', handler));
      };

      attach('.btn-edit-comment', async (e) => {
        const li = e.currentTarget.closest('li[data-id]');
        const id = li ? parseInt(li.getAttribute('data-id'), 10) : null;
        if (!id) return;
        const currentText = li.querySelector('.fw-semibold')?.textContent || '';
        const proposed = prompt('Modifier le commentaire:', currentText.trim());
        if (proposed == null) return; // annulé
        const text = proposed.trim();
        if (!text) return;
        try {
          const upd = await window.electronAPI.updateWeeklyComment?.({ id, comment_text: text });
          if (upd?.success) {
            await this.refreshCommentsList();
            this.showNotification('Commentaire modifié', 'La note a été mise à jour', 'success');
          } else {
            this.showNotification('Erreur', upd?.error || 'Modification impossible', 'danger');
          }
        } catch (err) {
          this.showNotification('Erreur', err?.message || 'Modification impossible', 'danger');
        }
      });

      attach('.btn-delete-comment', async (e) => {
        const li = e.currentTarget.closest('li[data-id]');
        const id = li ? parseInt(li.getAttribute('data-id'), 10) : null;
        if (!id) return;
        if (!confirm('Supprimer ce commentaire ?')) return;
        try {
          const del = await window.electronAPI.deleteWeeklyComment?.({ id });
          if (del?.success) {
            await this.refreshCommentsList();
            this.showNotification('Commentaire supprimé', 'La note a été supprimée', 'success');
          } else {
            this.showNotification('Erreur', del?.error || 'Suppression impossible', 'danger');
          }
        } catch (err) {
          this.showNotification('Erreur', err?.message || 'Suppression impossible', 'danger');
        }
      });
    } catch (e) {
      console.warn('Erreur refreshCommentsList:', e);
    }
  }

  /**
   * Ouvre un modal listant les commentaires d'une semaine
   */
  async openCommentModal(weekIdentifier, weekDisplay) {
    try {
      const modalEl = document.getElementById('commentModal');
      const titleEl = document.getElementById('commentModalTitle');
      const bodyEl = document.getElementById('commentModalBody');
      if (!modalEl || !titleEl || !bodyEl) return;

      // Déterminer l'identifiant et le libellé
      let id = weekIdentifier;
      let display = weekDisplay;
      if (!id && display) {
        const m = String(display).match(/S\s*(\d{1,2}).*?(\d{4})/);
        if (m) id = `S${parseInt(m[1],10)}-${m[2]}`;
      }
      if (!display && id) display = id.replace('S', 'S');

      titleEl.textContent = `Commentaires – ${display || 'Semaine'}`;
      bodyEl.innerHTML = '<div class="text-center text-muted py-2"><span class="spinner-border spinner-border-sm me-2"></span>Chargement…</div>';

      // Charger les commentaires
      const res = await window.electronAPI.listWeeklyComments?.(id);
      if (!res?.success) {
        bodyEl.innerHTML = `<div class="alert alert-danger">Erreur: ${this.escapeHtml(res?.error || 'chargement impossible')}</div>`;
      } else {
        const rows = Array.isArray(res.rows) ? res.rows : [];
        if (rows.length === 0) {
          bodyEl.innerHTML = '<div class="text-muted">Aucun commentaire pour cette semaine.</div>';
        } else {
          const fmtDate = (s) => s ? s.slice(0,16).replace('T',' ') : '';
          bodyEl.innerHTML = `
            <ul class="list-group small">
              ${rows.map(r => `
                <li class="list-group-item">
                  <div class="fw-semibold">${this.escapeHtml(r.comment_text || '')}</div>
                  <div class="text-muted">${this.escapeHtml(fmtDate(r.created_at))}${r.author ? ' • ' + this.escapeHtml(r.author) : ''}</div>
                </li>
              `).join('')}
            </ul>`;
        }
      }

      // Afficher le modal
      const modal = new bootstrap.Modal(modalEl);
      modal.show();
    } catch (e) {
      console.warn('openCommentModal:', e);
    }
  }

  /**
   * Charge les statistiques de la semaine actuelle
   */
  async loadCurrentWeekStats() {
    // Single-flight + throttle guard
    if (this._weeklyInFlight && this._weeklyInFlight.stats) return this._weeklyInFlight.stats;
    const now = Date.now();
    this._weeklyLastCall = this._weeklyLastCall || { stats: 0, history: 0 };
    if (now - (this._weeklyLastCall.stats || 0) < 800) return;
    this._weeklyLastCall.stats = now;
    this._weeklyInFlight.stats = (async () => {
      try {
        // Log API weekly stats - simplifié
        // console.log('📅 Appel API api-weekly-current-stats...');
        const response = await window.electronAPI.invoke('api-weekly-current-stats');
        console.log('📅 Réponse API reçue:', response);
        
        if (response.success) {
          // Les données sont dans response.weekInfo et response.categories
          const weekData = {
            weekInfo: response.weekInfo,
            categories: response.categories
          };
          console.log('📅 Données formatées pour affichage:', weekData);
          // Stocker l'information de semaine pour les ajustements
          this.state.currentWeekInfo = response.weekInfo || null;
          this.updateCurrentWeekDisplay(weekData);
        } else {
          console.error('❌ Erreur lors du chargement des stats hebdomadaires:', response.error);
          this.showWeeklyError(response.error);
        }
      } catch (error) {
        console.error('❌ Erreur lors du chargement des stats hebdomadaires:', error);
        this.showWeeklyError('Erreur de communication avec le serveur');
      } finally {
        this._weeklyInFlight.stats = null;
      }
    })();
    return this._weeklyInFlight.stats;
  }

  /**
   * Met à jour l'affichage de la semaine actuelle
   */
  updateCurrentWeekDisplay(weekData) {
    // Weekly stats update - log simplifié (conservé seulement pour erreurs importantes)
    console.log('📅 Mise à jour des stats de la semaine actuelle:', weekData);

    const weekInfo = weekData.weekInfo;
    const categories = weekData.categories || {};

    // Mettre à jour le titre de la semaine
    const weekTitle = document.getElementById('current-week-title');
    if (weekTitle && weekInfo) {
      weekTitle.textContent = `${weekInfo.displayName} (${weekInfo.startDate} - ${weekInfo.endDate})`;
    }

    // Mettre à jour les statistiques par catégorie
    const statsContainer = document.getElementById('current-week-stats');
    if (statsContainer) {
      let statsHtml = '';
      
      if (Object.keys(categories).length > 0) {
        // Générer le HTML pour chaque catégorie
        for (const [category, catStats] of Object.entries(categories)) {
          const treatmentRate = catStats.received > 0 ? (catStats.treated / catStats.received * 100).toFixed(1) : 0;
          
          statsHtml += `
            <div class="col-lg-4 col-md-6">
              <div class="card border-0 bg-light h-100">
                <div class="card-body p-3">
                  <h6 class="card-title mb-3 fw-bold text-primary">${category}</h6>
                  <div class="row g-2 text-center">
                    <div class="col-6">
                      <div class="h5 mb-1 text-primary">${catStats.received}</div>
                      <small class="text-muted">Reçus</small>
                    </div>
                    <div class="col-6">
                      <div class="h5 mb-1 text-success">${catStats.treated}</div>
                      <small class="text-muted">Traités</small>
                    </div>
                  </div>
                  <div class="row g-2 text-center mt-2">
                    <div class="col-6">
                      <div class="h6 mb-1 text-info">${catStats.adjustments}</div>
                      <small class="text-muted">Traité par ADG</small>
                    </div>
                    <div class="col-6">
                      <div class="h6 mb-1 text-secondary">${catStats.total}</div>
                      <small class="text-muted">Total</small>
                    </div>
                  </div>
                  <div class="mt-2">
                    <small class="text-muted">Taux de traitement: ${treatmentRate}%</small>
                  </div>
                </div>
              </div>
            </div>
          `;
        }
      }
      
      if (statsHtml === '') {
        statsHtml = `
          <div class="col-12 text-center text-muted">
            <i class="bi bi-calendar-x fs-1 mb-3"></i>
            <p>Aucune donnée pour cette semaine</p>
          </div>
        `;
      }
      
      statsContainer.innerHTML = statsHtml;
    }

    // Retirer l'état de chargement
    const loadingElement = document.querySelector('#current-week-stats .spinner-border');
    if (loadingElement) {
      loadingElement.parentElement.style.display = 'none';
    }
  }

  /**
   * Affiche une erreur dans l'onglet hebdomadaire
   */
  showWeeklyError(errorMessage) {
    // Retirer l'état de chargement
    const loadingElement = document.querySelector('#weekly-tab .loading-state');
    if (loadingElement) {
      loadingElement.style.display = 'none';
    }
    
    // Afficher l'erreur
    const statsContainer = document.getElementById('weekly-categories-stats');
    if (statsContainer) {
      statsContainer.innerHTML = `
        <div class="alert alert-danger" role="alert">
          <i class="fas fa-exclamation-triangle me-2"></i>
          Erreur de chargement: ${errorMessage}
        </div>
      `;
    }
  }

  /**
   * Rafraîchit les statistiques de la semaine actuelle
   */
  async refreshCurrentWeekStats() {
  // Conserve pour compatibilité: rafraîchit à la demande (même si onglet inactif)
  await this.loadCurrentWeekStats();
  }

  /**
   * Gère les données traitées par ADG
   */
  async handleManualAdjustment(sign = 1) {
    const form = document.getElementById('manual-adjustment-form');
    if (!form) return;
    const categoryEl = document.getElementById('adjustment-category');
    const valueEl = document.getElementById('adjustment-value');
    const typeEl = document.getElementById('adjustment-type');

    const category = categoryEl ? categoryEl.value : '';
    const value = valueEl ? parseInt(valueEl.value, 10) : 0;
    const adjustmentType = typeEl ? typeEl.value : 'manual_adjustments';

    if (!category || !value || value <= 0) {
      this.showNotification('Erreur', 'Veuillez remplir tous les champs requis', 'warning');
      return;
    }

    // Mapper la catégorie affichée vers le folderType attendu en BDD
    let folderType = category;
    if (category.toLowerCase() === 'déclarations' || category.toLowerCase() === 'declarations') {
      folderType = 'declarations';
    } else if (category.toLowerCase() === 'règlements' || category.toLowerCase() === 'reglements') {
      folderType = 'reglements';
    } else if (category.toLowerCase() === 'mail simple' || category.toLowerCase() === 'mails_simples' || category.toLowerCase() === 'mails simples') {
      folderType = 'mails_simples';
    }

  // Identifier de semaine: prioriser la sélection du formulaire "Ajouter aux compteurs"
  const countersWeekSel = document.getElementById('counters-week-select');
  let weekIdentifier = countersWeekSel?.value || this.state.currentWeekInfo?.identifier;
    if (!weekIdentifier) {
      try {
        const resp = await window.electronAPI.invoke('api-weekly-current-stats');
        if (resp?.success) weekIdentifier = resp.weekInfo?.identifier;
      } catch {}
    }
    if (!weekIdentifier) {
      this.showNotification('Erreur', 'Semaine courante introuvable', 'danger');
      return;
    }

    const payload = {
      weekIdentifier,
      folderType,
      adjustmentValue: sign * Math.abs(value),
      adjustmentType
    };

    try {
      const response = await window.electronAPI.invoke('api-weekly-adjust-count', payload);
      if (response.success) {
        this.showNotification('Succès', `${Math.abs(value)} ${adjustmentType} ${sign > 0 ? 'ajouté(s)' : 'retiré(s)'} pour ${folderType}`, 'success');
        form.reset();
        await this.loadCurrentWeekStats();
        await this.loadWeeklyHistory();
        await this.loadPersonalPerformance?.();
      } else {
        this.showNotification('Erreur', response.error || 'Ajustement refusé', 'danger');
      }
    } catch (error) {
      console.error('Erreur lors de l\'ajustement manuel:', error);
      this.showNotification('Erreur', 'Impossible d\'ajuster les données', 'danger');
    }
  }

  /**
   * Charge l'historique hebdomadaire
   */
  async loadWeeklyHistory() {
    // Single-flight + throttle guard
    if (this._weeklyInFlight && this._weeklyInFlight.history) return this._weeklyInFlight.history;
    const now = Date.now();
    this._weeklyLastCall = this._weeklyLastCall || { stats: 0, history: 0 };
    if (now - (this._weeklyLastCall.history || 0) < 800) return;
    this._weeklyLastCall.history = now;
    this._weeklyInFlight.history = (async () => {
      try {
        const { page, pageSize } = this.state.weeklyHistory;
        const response = await window.electronAPI.invoke('api-weekly-history', { page, pageSize, limit: pageSize });
        if (response.success) {
          this.updateWeeklyHistoryDisplay(response.data);
          // Mettre à jour l'état de pagination
          this.state.weeklyHistory.page = response.page || page;
          this.state.weeklyHistory.pageSize = response.pageSize || pageSize;
          this.state.weeklyHistory.totalWeeks = response.totalWeeks || 0;
          this.state.weeklyHistory.totalPages = response.totalPages || 1;
          this.updateWeeklyPaginationControls();
        } else {
          console.error('Erreur lors du chargement de l\'historique:', response.error);
        }
      } catch (error) {
        console.error('Erreur lors du chargement de l\'historique:', error);
      } finally {
        this._weeklyInFlight.history = null;
      }
    })();
    return this._weeklyInFlight.history;
  }

  /**
   * Charge et met à jour l'onglet "Performances personnelles"
   */
  async loadPersonalPerformance() {
    try {
      // Appliquer filtres persistés au chargement (si UI présente)
      const stored = (() => {
        try { return JSON.parse(localStorage.getItem('ppFilters') || 'null'); } catch (_) { return null; }
      })();
      if (stored) this.applyPersonalPerformanceFiltersToUI({
        periodMode: stored.periodMode || 'last-6',
        year: stored.year || null,
        categories: stored.categories || { declarations: true, reglements: true, mails: true },
        includeAdg: !!stored.includeAdg
      });

      const filters = this.getPersonalPerformanceFilters();

      // Charger assez de semaines pour couvrir une année complète (et quelques années si besoin)
      const fetchSize = (filters.periodMode === 'year') ? 260 : 80;
      const response = await window.electronAPI.invoke('api-weekly-history', { page: 1, pageSize: fetchSize, limit: fetchSize });
      if (!response?.success) return;

      const allWeeks = Array.isArray(response.data) ? response.data : [];

      // Remplir la liste d'années disponibles
      const yearEl = document.getElementById('pp-year');
      const yearWrapEl = document.getElementById('pp-year-wrap');
      if (yearWrapEl) yearWrapEl.style.display = (filters.periodMode === 'year') ? '' : 'none';
      if (yearEl) {
        const years = Array.from(new Set(allWeeks.map(w => Number(w.week_year)).filter(Boolean))).sort((a, b) => b - a);
        const currentSelected = (yearEl.value && String(yearEl.value).trim() !== '') ? Number(yearEl.value) : (filters.year || null);
        yearEl.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('') || '<option value="">--</option>';
        const fallbackYear = years[0] || null;
        const finalYear = currentSelected && years.includes(currentSelected) ? currentSelected : fallbackYear;
        if (finalYear) {
          yearEl.value = String(finalYear);
          if (filters.periodMode === 'year' && (!filters.year || filters.year !== finalYear)) {
            const updated = { ...filters, year: finalYear };
            this.persistPersonalPerformanceFilters(updated);
          }
        }
      }

      // Sélectionner les semaines selon la période
      let weeks = allWeeks;
      if (filters.periodMode === 'year') {
        const y = (document.getElementById('pp-year')?.value && String(document.getElementById('pp-year').value).trim() !== '')
          ? Number(document.getElementById('pp-year').value)
          : (filters.year || null);
        weeks = y ? allWeeks.filter(w => Number(w.week_year) === y) : [];
      } else {
        const m = String(filters.periodMode || 'last-6').match(/^last-(\d+)$/);
        const count = m ? Math.max(1, Math.min(104, parseInt(m[1], 10))) : 6;
        weeks = allWeeks.slice(0, count);
      }

      const selectedCategories = new Set([
        filters.categories?.declarations ? 'Déclarations' : null,
        filters.categories?.reglements ? 'Règlements' : null,
        filters.categories?.mails ? 'Mails simples' : null
      ].filter(Boolean));
      const tableBody = document.getElementById('personal-history-body');
      const weekTitleEl = document.getElementById('personal-week-title');
      const arrivalsEl = document.getElementById('personal-arrivals');
      const treatedEl = document.getElementById('personal-treated');
      const stockEl = document.getElementById('personal-stock');
      const trendEl = document.getElementById('personal-trend');
      const periodBadgeEl = document.getElementById('pp-period-badge');

      if (!tableBody) return;

      if (weeks.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-3">Aucune donnée</td></tr>';
        if (trendEl) { trendEl.className = 'badge bg-secondary'; trendEl.textContent = '--'; }
        if (periodBadgeEl) periodBadgeEl.textContent = '--';
        return;
      }

      // Badge période
      try {
        const label = (filters.periodMode === 'year')
          ? `Année ${Number(document.getElementById('pp-year')?.value || filters.year || '') || ''}`.trim()
          : `${weeks.length} sem.`;
        const catsLabel = Array.from(selectedCategories);
        const catsShort = catsLabel.length === 3 ? 'Toutes catégories' : (catsLabel.length ? catsLabel.join(', ') : 'Aucune catégorie');
        const adgLabel = filters.includeAdg ? 'ADG inclus' : 'ADG exclu';
        if (periodBadgeEl) periodBadgeEl.textContent = `${label} • ${catsShort} • ${adgLabel}`;
      } catch (_) {
        if (periodBadgeEl) periodBadgeEl.textContent = '--';
      }

      // Construire lignes et extraire S0/S-1
      const rowsHtml = weeks.map(w => {
        const sums = (w.categories || []).reduce((acc, c) => {
          if (!selectedCategories.has(c.name)) return acc;
          acc.received += Number(c.received || 0);
          const treatedBase = Number(c.treated || 0);
          const adg = Number(c.adjustments || 0);
          acc.treated += (filters.includeAdg ? (treatedBase + adg) : treatedBase);
          acc.stockEnd += Number(c.stockEndWeek || 0);
          return acc;
        }, { received: 0, treated: 0, stockEnd: 0 });
        // Dériver l'identifiant de semaine (Sxx-YYYY) à partir de l'affichage
        let identifier = '';
        const m = String(w.weekDisplay || '').match(/S\s*(\d{1,2}).*?(\d{4})/);
        if (m) identifier = `S${parseInt(m[1], 10)}-${m[2]}`;
        return {
          display: w.weekDisplay,
          received: sums.received,
          treated: sums.treated,
          stockEnd: sums.stockEnd,
          dateRange: w.dateRange,
          identifier
        };
      });

      // Rendre le tableau
      tableBody.innerHTML = rowsHtml.map(r => `
        <tr>
          <td>${this.escapeHtml(r.display)}</td>
          <td class="text-center"><span class="badge bg-primary rounded-pill">${r.received}</span></td>
          <td class="text-center"><span class="badge bg-success rounded-pill">${r.treated}</span></td>
          <td class="text-center"><span class="badge bg-warning rounded-pill">${r.stockEnd}</span></td>
          <td class="text-center">
            <button type="button" class="btn btn-outline-info btn-sm btn-pp-week-notes" data-week-identifier="${this.escapeHtml(r.identifier)}" data-week-display="${this.escapeHtml(r.display)}">
              <i class="bi bi-info-circle me-1"></i> Note
            </button>
          </td>
        </tr>
      `).join('');

      // Brancher les boutons Note pour ouvrir le modal de commentaires
      tableBody.querySelectorAll('.btn-pp-week-notes').forEach(btn => {
        btn.addEventListener('click', () => {
          const id = btn.getAttribute('data-week-identifier');
          const display = btn.getAttribute('data-week-display');
          this.openCommentModal(id, display);
        });
      });

  // KPIs (période): totaux sur la période, et stock fin sur la semaine la plus récente
      const current = rowsHtml[0];
      const previous = rowsHtml[1];
      if (weekTitleEl && current) weekTitleEl.textContent = current.display;
  const totalReceivedPeriod = rowsHtml.reduce((s, r) => s + Number(r.received || 0), 0);
  const totalTreatedPeriod = rowsHtml.reduce((s, r) => s + Number(r.treated || 0), 0);
  if (arrivalsEl) arrivalsEl.textContent = String(totalReceivedPeriod);
  if (treatedEl) treatedEl.textContent = String(totalTreatedPeriod);
      if (stockEl && current) stockEl.textContent = current.stockEnd;

      // Tendance vs S-1 sur le stock fin de semaine
      if (trendEl && current && previous) {
        const delta = current.stockEnd - previous.stockEnd;
        if (delta > 0) {
          trendEl.className = 'badge bg-danger';
          trendEl.textContent = `En hausse (+${delta})`;
        } else if (delta < 0) {
          trendEl.className = 'badge bg-success';
          trendEl.textContent = `En baisse (${delta})`;
        } else {
          trendEl.className = 'badge bg-secondary';
          trendEl.textContent = 'Stable';
        }
      } else if (trendEl) {
        trendEl.className = 'badge bg-secondary';
        trendEl.textContent = '--';
      }

      // Graphiques
      try {
        const labels = weeks.map(w => w.weekDisplay).reverse(); // oldest→newest
        const receivedData = weeks.map(w => (w.categories || []).reduce((s, c) => {
          if (!selectedCategories.has(c.name)) return s;
          return s + Number(c.received || 0);
        }, 0)).reverse();
        const treatedData = weeks.map(w => (w.categories || []).reduce((s, c) => {
          if (!selectedCategories.has(c.name)) return s;
          const treatedBase = Number(c.treated || 0);
          const adg = Number(c.adjustments || 0);
          return s + (filters.includeAdg ? (treatedBase + adg) : treatedBase);
        }, 0)).reverse();
        const stockData = weeks.map(w => (w.categories || []).reduce((s, c) => {
          if (!selectedCategories.has(c.name)) return s;
          return s + Number(c.stockEndWeek || 0);
        }, 0)).reverse();
        const rateData = receivedData.map((rec, i) => rec > 0 ? Math.round((treatedData[i] / rec) * 1000) / 10 : 0);

        // Theme-aware colors
        const isDark = (document.body?.getAttribute('data-theme') === 'dark');
        const col1 = getComputedStyle(document.body).getPropertyValue('--brand-1').trim() || (isDark ? '#60a5fa' : '#0d6efd');
        const col2 = getComputedStyle(document.body).getPropertyValue('--brand-2').trim() || (isDark ? '#34d399' : '#66b2ff');
        const textCol = getComputedStyle(document.body).getPropertyValue('--text-primary').trim() || (isDark ? '#e5e7eb' : '#1f2937');
        const gridCol = getComputedStyle(document.body).getPropertyValue('--border-color').trim() || (isDark ? '#1f2937' : '#e5e7eb');

        // Destroy previous charts if any
        const destroyIf = (c) => { try { if (c && c.destroy) c.destroy(); } catch(_){} };
        destroyIf(this.charts.ppRT);
  destroyIf(this.charts.ppStock);
  destroyIf(this.charts.ppRate);

        const ctx1 = document.getElementById('pp-chart-rt');
        const ctx2 = document.getElementById('pp-chart-stock');
        if (ctx1 && window.Chart) {
          this.charts.ppRT = new Chart(ctx1, {
            type: 'bar',
            data: {
              labels,
              datasets: [
                { label: 'Reçus', data: receivedData, backgroundColor: col1, borderRadius: 6 },
                { label: 'Traités', data: treatedData, backgroundColor: col2, borderRadius: 6 }
              ]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { labels: { color: textCol } } },
              scales: {
                x: { ticks: { color: textCol }, grid: { color: gridCol } },
                y: { ticks: { color: textCol }, grid: { color: gridCol } }
              }
            }
          });
        }
        if (ctx2 && window.Chart) {
          this.charts.ppStock = new Chart(ctx2, {
            type: 'line',
            data: {
              labels,
              datasets: [
                { label: 'Stock fin', data: stockData, borderColor: col1, backgroundColor: col1 + '33', tension: .3, fill: true }
              ]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { labels: { color: textCol } } },
              scales: {
                x: { ticks: { color: textCol }, grid: { color: gridCol } },
                y: { ticks: { color: textCol }, grid: { color: gridCol } }
              }
            }
          });
        }
        const ctx3 = document.getElementById('pp-chart-rate');
        if (ctx3 && window.Chart) {
          this.charts.ppRate = new Chart(ctx3, {
            type: 'line',
            data: {
              labels,
              datasets: [
                { label: 'Taux (%)', data: rateData, borderColor: col2, backgroundColor: col2 + '33', tension: .3, fill: true }
              ]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { labels: { color: textCol } } },
              scales: {
                x: { ticks: { color: textCol }, grid: { color: gridCol } },
                y: { ticks: { color: textCol }, grid: { color: gridCol } }
              }
            }
          });
        }
      } catch(_) { /* ignore charts if missing */ }

      // Analyses complémentaires (moyennes, taux, semaines extrêmes)
      try {
        const avg = (arr) => arr.length ? (arr.reduce((a,b)=>a+b,0) / arr.length) : 0;
        const safePct = (num, den) => den > 0 ? Math.round((num/den)*1000)/10 : 0; // 1 decimal

        const avgReceived = Math.round(avg(rowsHtml.map(r=>r.received)));
        const avgTreated = Math.round(avg(rowsHtml.map(r=>r.treated)));
        const avgRate = safePct(avgTreated, avgReceived);
        const currentRate = safePct((current?.treated)||0, (current?.received)||0);

        // Semaine la plus chargée (par reçus)
        const busiest = rowsHtml.reduce((best, r) => !best || r.received > best.received ? r : best, null);
        // Meilleure semaine par taux de traitement (ignorer petites semaines <10 reçus)
        const bestRate = rowsHtml
          .map(r => ({ r, rate: safePct(r.treated, r.received) }))
          .filter(x => x.r.received >= 10)
          .reduce((best, x) => (!best || x.rate > best.rate) ? x : best, null);

  const setText = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  // Estimation semaines pour vider: si débit net moyen > 0 (traités > reçus), stock / (traités - reçus)
  const netPerWeek = Math.max(0, avgTreated - avgReceived);
  const clearWeeks = (netPerWeek > 0 && current?.stockEnd > 0) ? Math.ceil(current.stockEnd / netPerWeek) : (netPerWeek === 0 && (current?.stockEnd||0) > 0 ? '∞' : 0);
        setText('pp-rate-current', `${currentRate}%`);
        setText('pp-avg-received', `${avgReceived}`);
        setText('pp-avg-treated', `${avgTreated}`);
        setText('pp-avg-rate', `${avgRate}%`);
        setText('pp-busiest-week', busiest ? `${this.escapeHtml(busiest.display)} (${busiest.received})` : '--');
        setText('pp-best-week', bestRate ? `${this.escapeHtml(bestRate.r.display)} (${bestRate.rate}%)` : '--');
  setText('pp-clear-weeks', `${clearWeeks}`);
      } catch (_) { /* ignore */ }
    } catch (e) {
      console.warn('Erreur chargement Performances personnelles:', e);
    }
  }

  /**
   * Met à jour l'affichage de l'historique
   */
  updateWeeklyHistoryDisplay(historyData) {
    const historyTable = document.getElementById('weekly-history-table');
    const loadingElement = document.getElementById('weekly-loading');
    const noDataElement = document.getElementById('weekly-no-data');
  const paginationEl = document.getElementById('weekly-pagination');
    
    if (!historyTable) return;

    const tbody = historyTable.querySelector('tbody');
    if (!tbody) return;

    // Masquer l'indicateur de chargement
    if (loadingElement) {
      loadingElement.classList.add('d-none');
    }

  let historyHtml = '';
    
    if (!historyData || historyData.length === 0) {
      // Afficher le message "aucune donnée" et masquer le tableau
      if (noDataElement) {
        noDataElement.classList.remove('d-none');
      }
      historyTable.style.display = 'none';
      if (paginationEl) paginationEl.classList.add('d-none');
      return;
    } else {
      // Masquer le message "aucune donnée" et afficher le tableau
      if (noDataElement) {
        noDataElement.classList.add('d-none');
      }
      historyTable.style.display = 'table';
      if (paginationEl) paginationEl.classList.remove('d-none');
      
      for (const week of historyData) {
        // Calculer le total des stocks pour la semaine
        const totalStock = week.categories.reduce((sum, category) => sum + category.stockEndWeek, 0);
        
        // Créer le contenu d'évolution avec design 2025
        const evolutionHtml = this.createEvolutionDisplay(week.evolution);
        
        // Créer un ID unique pour la semaine pour gérer le survol
        const weekId = `week-${week.weekDisplay.replace(/[^a-zA-Z0-9]/g, '-')}`;
        // Tenter de dériver l'identifiant de semaine (Sxx-YYYY)
        let weekIdentifier = '';
        const m = String(week.weekDisplay || '').match(/S\s*(\d{1,2}).*?(\d{4})/);
        if (m) weekIdentifier = `S${parseInt(m[1],10)}-${m[2]}`;
        
        // Ligne pour la semaine avec les 3 catégories
        week.categories.forEach((category, index) => {
          const isFirstRow = index === 0;
          
          historyHtml += `
            <tr class="${isFirstRow ? 'week-separator' : ''} week-group" data-week-id="${weekId}">`;
          
          // Cellules fusionnées seulement sur la première ligne
          if (isFirstRow) {
            historyHtml += `
              <td rowspan="3" class="fw-bold text-primary align-middle text-center" style="vertical-align: middle !important;">${week.weekDisplay}</td>
              <td rowspan="3" class="fw-semibold text-muted align-middle text-center" style="vertical-align: middle !important;">${week.dateRange}</td>`;
          }
          
          historyHtml += `
              <td class="fw-medium">${category.name}</td>
              <td class="text-center"><span class="badge bg-primary rounded-pill">${category.received}</span></td>
              <td class="text-center"><span class="badge bg-success rounded-pill">${category.treated}</span></td>
              <td class="text-center"><span class="badge bg-info rounded-pill">${category.adjustments}</span></td>
              <td class="text-center"><span class="badge bg-warning rounded-pill">${category.stockEndWeek}</span></td>`;
          
          // Cellules fusionnées pour Total Stock et Évolution seulement sur la première ligne
          if (isFirstRow) {
            historyHtml += `
              <td rowspan="3" class="text-center align-middle" style="vertical-align: middle !important;"><span class="badge bg-dark rounded-pill fs-6">${totalStock}</span></td>
              <td rowspan="3" class="text-center align-middle" style="vertical-align: middle !important;">${evolutionHtml}</td>
              <td rowspan="3" class="text-center align-middle" style="vertical-align: middle !important;">
                <button type="button" class="btn btn-outline-info btn-sm btn-week-notes" data-week-identifier="${this.escapeHtml(weekIdentifier)}" data-week-display="${this.escapeHtml(week.weekDisplay)}">
                  <i class="bi bi-info-circle me-1"></i> Note
                </button>
              </td>`;
          }
          
          historyHtml += `
            </tr>
          `;
        });
      }
    }
    
    tbody.innerHTML = historyHtml;
    
    // Ajouter les événements de survol pour les semaines complètes
    this.addWeekHoverEvents();

    // Boutons "Note" (ouvrir le modal de commentaires)
    tbody.querySelectorAll('.btn-week-notes').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = btn.getAttribute('data-week-identifier');
        const display = btn.getAttribute('data-week-display');
        this.openCommentModal(id, display);
      });
    });
  }

  /**
   * Met à jour les contrôles et libellés de pagination de l'historique hebdo
   */
  updateWeeklyPaginationControls() {
    const { page, totalPages, totalWeeks } = this.state.weeklyHistory;
    const prevBtn = document.getElementById('weekly-prev');
    const nextBtn = document.getElementById('weekly-next');
    const pageInfo = document.getElementById('weekly-page-info');
    const totalEl = document.getElementById('weekly-total-weeks');
    const paginationEl = document.getElementById('weekly-pagination');

    if (!paginationEl) return;

    paginationEl.classList.remove('d-none');
    if (pageInfo) pageInfo.textContent = `Page ${page} / ${Math.max(totalPages, 1)}`;
    if (totalEl) totalEl.textContent = totalWeeks || 0;
    if (prevBtn) prevBtn.disabled = page <= 1;
    if (nextBtn) nextBtn.disabled = page >= (totalPages || 1);
  }

  /**
   * Ajoute les événements de survol pour les semaines complètes
   */
  addWeekHoverEvents() {
    const weekRows = document.querySelectorAll('.week-group');
    
    weekRows.forEach(row => {
      const weekId = row.getAttribute('data-week-id');
      
      row.addEventListener('mouseenter', () => {
        // Surligner toutes les lignes de cette semaine
        const weekElements = document.querySelectorAll(`[data-week-id="${weekId}"]`);
        weekElements.forEach(element => element.classList.add('week-hover'));
      });
      
      row.addEventListener('mouseleave', () => {
        // Retirer le surlignage de toutes les lignes de cette semaine
        const weekElements = document.querySelectorAll(`[data-week-id="${weekId}"]`);
        weekElements.forEach(element => element.classList.remove('week-hover'));
      });
    });
  }

  /**
   * Crée l'affichage d'évolution avec design 2025
   */
  createEvolutionDisplay(evolution) {
    if (!evolution || evolution.trend === 'stable') {
      return `
        <div class="d-flex align-items-center justify-content-center">
          <div class="evolution-indicator stable">
            <i class="bi bi-dash-lg"></i>
          </div>
          <span class="ms-2 small text-muted">Stable</span>
        </div>
      `;
    }
    
  const isPositive = evolution.trend === 'up';
  const icon = isPositive ? 'bi-arrow-up-right' : 'bi-arrow-down-right';
  // Inverser la logique : rouge si ça monte, vert si ça baisse
  const colorClass = isPositive ? 'text-danger' : 'text-success';
  const bgClass = isPositive ? 'bg-danger' : 'bg-success';
  const sign = isPositive ? '+' : '';
    
    return `
      <div class="d-flex align-items-center justify-content-center">
        <div class="evolution-indicator ${isPositive ? 'negative' : 'positive'}">
          <i class="bi ${icon}"></i>
        </div>
        <div class="ms-2 text-start">
          <div class="small fw-bold ${colorClass}">${sign}${evolution.absolute}</div>
          <div class="tiny text-muted">${sign}${evolution.percent.toFixed(1)}%</div>
        </div>
      </div>
    `;
  }

  /**
   * Ouvre la modal des paramètres hebdomadaires
   */
  async openWeeklySettings() {
    await this.loadWeeklySettings();
    const modal = new bootstrap.Modal(document.getElementById('weeklySettingsModal'));
    modal.show();
  }

  /**
   * Charge les paramètres de suivi hebdomadaire
   */
  async loadWeeklySettings() {
    try {
      const response = await window.electronAPI.invoke('api-settings-count-read-as-treated');
      if (response.success) {
        const checkbox = document.getElementById('count-read-as-treated');
        if (checkbox) {
          checkbox.checked = !!response.value;
        }
      }

      const startupRes = await window.electronAPI.invoke('api-settings-startup-adjustments');
      if (startupRes?.success) {
        const v = startupRes.value || {};
        const decl = document.getElementById('startup-adjust-declarations');
        const regl = document.getElementById('startup-adjust-reglements');
        const mails = document.getElementById('startup-adjust-mails');
        if (decl) decl.value = Number.isFinite(Number(v.declarations)) ? String(parseInt(v.declarations, 10)) : '0';
        if (regl) regl.value = Number.isFinite(Number(v.reglements)) ? String(parseInt(v.reglements, 10)) : '0';
        if (mails) mails.value = Number.isFinite(Number(v.mails_simples)) ? String(parseInt(v.mails_simples, 10)) : '0';
      }
    } catch (error) {
      console.error('Erreur lors du chargement des paramètres:', error);
    }
  }

  /**
   * Sauvegarde les paramètres de suivi hebdomadaire
   */
  async saveWeeklySettings() {
    try {
      const checkbox = document.getElementById('count-read-as-treated');
      const countReadAsTreated = checkbox ? checkbox.checked : false;

      const decl = document.getElementById('startup-adjust-declarations');
      const regl = document.getElementById('startup-adjust-reglements');
      const mails = document.getElementById('startup-adjust-mails');
      const startupValue = {
        declarations: parseInt(decl?.value || '0', 10) || 0,
        reglements: parseInt(regl?.value || '0', 10) || 0,
        mails_simples: parseInt(mails?.value || '0', 10) || 0
      };

      const response = await window.electronAPI.invoke('api-settings-count-read-as-treated', {
  value: countReadAsTreated
      });

      const startupRes = await window.electronAPI.invoke('api-settings-startup-adjustments', {
        value: startupValue
      });

      if (response.success && startupRes?.success) {
        const mig = response.migration;
        const migInfo = (mig && mig.success)
          ? `\nMise à jour rétroactive: ${mig.updatedEmails || 0} emails, ${mig.updatedWeeklyRows || 0} lignes d'historique.`
          : '';

        this.showNotification(
          'Paramètres sauvegardés',
          `Les paramètres du suivi hebdomadaire ont été mis à jour.${migInfo}`,
          'success'
        );
        
        // Fermer la modal
        const modal = bootstrap.Modal.getInstance(document.getElementById('weeklySettingsModal'));
        if (modal) {
          modal.hide();
        }
        
        // Recharger les statistiques pour appliquer les nouveaux paramètres
        await this.loadCurrentWeekStats();
        await this.loadWeeklyHistory();
        await this.loadPersonalPerformance?.();
        
      } else {
        this.showNotification('Erreur', response.error || startupRes?.error || 'Impossible de sauvegarder les paramètres', 'danger');
      }
    } catch (error) {
      console.error('Erreur lors de la sauvegarde des paramètres:', error);
      this.showNotification('Erreur', 'Impossible de sauvegarder les paramètres', 'danger');
    }
  }
}

// Initialiser l'application quand le DOM est prêt
document.addEventListener('DOMContentLoaded', () => {
  window.app = new MailMonitor();
});

// Nettoyage avant fermeture
window.addEventListener('beforeunload', () => {
  if (window.app) {
    window.app.destroy();
  }
});
