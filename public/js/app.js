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
      lastUpdate: null
    };
    
    // Gestionnaire de chargement centralisé
    this.loadingManager = {
      tasks: new Map(),
      totalTasks: 0,
      completedTasks: 0,
      isLoading: false
    };
    
    this.updateInterval = null;
    this.charts = {};
    this.init();
  }

  // === INITIALISATION ===
  async init() {
    console.log('🚀 Initialisation de Mail Monitor...');
    
    try {
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
      
      this.startPeriodicUpdates();
      this.setupAutoRefresh();
      
      // Terminer le chargement
      this.finishLoading();
      
      console.log('✅ MailMonitor app initialisée avec succès');
    } catch (error) {
      console.error('❌ Erreur lors de l\'initialisation:', error);
      this.showNotification('Erreur d\'initialisation', error.message, 'danger');
      this.finishLoading();
    }
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
    
    console.log('✅ Événements temps réel configurés (y compris COM)');
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
    } else {
      console.warn(`⚠️ Élément avec ID '${elementId}' non trouvé`);
    }
  }

  /**
   * Mise à jour de l'heure de dernière actualisation
   */
  updateLastRefreshTime() {
    const now = new Date();
    const timeString = now.toLocaleTimeString('fr-FR');
    this.updateElement('last-update', timeString);
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
      });
    });
    
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
    document.getElementById('start-monitoring')?.addEventListener('click', () => this.startMonitoring());
    document.getElementById('stop-monitoring')?.addEventListener('click', () => this.stopMonitoring());
    document.getElementById('add-folder')?.addEventListener('click', () => this.showAddFolderModal());
    document.getElementById('refresh-folders')?.addEventListener('click', () => this.refreshFoldersDisplay());
    
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
  async loadFoldersStats() {
    try {
      console.log('📁 Chargement statistiques dossiers...');
      
      // Récupérer les données des dossiers avec leurs emails
      const foldersData = await window.electronAPI.getFoldersTree();
      
      if (foldersData && foldersData.folders) {
        this.updateFoldersStatsDisplay(foldersData.folders);
        
        // Mettre à jour le badge total
        const totalFoldersEl = document.getElementById('total-folders-badge');
        if (totalFoldersEl) {
          totalFoldersEl.textContent = `${foldersData.folders.length} dossier${foldersData.folders.length > 1 ? 's' : ''}`;
        }
      }
      
    } catch (error) {
      console.error('❌ Erreur chargement stats dossiers:', error);
    }
  }

  // Nouvelle méthode pour afficher les statistiques des dossiers
  updateFoldersStatsDisplay(folders) {
    const container = document.getElementById('folders-stats-grid');
    if (!container) return;
    
    if (!folders || folders.length === 0) {
      container.innerHTML = `
        <div class="col-12 text-center text-muted py-4">
          <i class="bi bi-folder-x fs-1 mb-3"></i>
          <p>Aucun dossier configuré pour le monitoring</p>
          <small>Ajoutez des dossiers dans l'onglet Configuration</small>
        </div>
      `;
      return;
    }
    
    // Générer les cartes pour chaque dossier
    const foldersHtml = folders.map(folder => {
      const categoryColor = this.getCategoryColor(folder.category);
      const categoryIcon = this.getCategoryIcon(folder.category);
      const readEmails = folder.emailCount || 0;  // Tous les emails pour ce dossier
      
      return `
        <div class="col-xl-3 col-lg-4 col-md-6">
          <div class="folder-stat-card p-3 h-100">
            <div class="d-flex align-items-center mb-2">
              <i class="bi ${categoryIcon} ${categoryColor} fs-4 me-2"></i>
              <h6 class="mb-0 fw-semibold">${this.escapeHtml(folder.name)}</h6>
            </div>
            <div class="row g-2 text-center">
              <div class="col-6">
                <div class="fs-4 fw-bold text-primary">${readEmails}</div>
                <small class="text-muted">Total</small>
              </div>
              <div class="col-6">
                <div class="fs-4 fw-bold text-warning">0</div>
                <small class="text-muted">Non lus</small>
              </div>
            </div>
            <div class="mt-2">
              <div class="d-flex justify-content-between align-items-center">
                <small class="text-muted">${this.escapeHtml(folder.category)}</small>
                <span class="badge bg-light text-dark">${folder.isMonitored ? 'Actif' : 'Inactif'}</span>
              </div>
            </div>
          </div>
        </div>
      `;
    }).join('');
    
    container.innerHTML = foldersHtml;
  }

  // Méthodes utilitaires pour les catégories
  getCategoryColor(category) {
    const colors = {
      'Déclarations': 'text-danger',
      'Règlements': 'text-success', 
      'mails_simples': 'text-info',
      'test': 'text-secondary'
    };
    return colors[category] || 'text-muted';
  }

  getCategoryIcon(category) {
    const icons = {
      'Déclarations': 'bi-file-earmark-text',
      'Règlements': 'bi-credit-card',
      'mails_simples': 'bi-envelope',
      'test': 'bi-folder'
    };
    return icons[category] || 'bi-folder';
  }

  async checkMonitoringStatus() {
    try {
      // Vérifier si des dossiers sont configurés pour le monitoring
      const foldersData = await window.electronAPI.getFoldersTree();
      const foldersCount = foldersData?.folders?.length || 0;
      
      if (foldersCount > 0) {
        console.log(`📁 ${foldersCount} dossier(s) configuré(s) - monitoring probablement actif`);
        
        // Assumer que le monitoring est actif si des dossiers sont configurés
        // (car il se démarre automatiquement au lancement)
        this.state.isMonitoring = true;
        this.updateMonitoringStatus(true);
        
        // this.showNotification(
        //   'Monitoring automatique', 
        //   `Le monitoring automatique est actif sur ${foldersCount} dossier(s)`, 
        //   'info'
        // ); // Notification supprimée
      } else {
        console.log('📁 Aucun dossier configuré - monitoring arrêté');
        this.state.isMonitoring = false;
        this.updateMonitoringStatus(false);
      }
    } catch (error) {
      console.error('❌ Erreur vérification statut monitoring:', error);
      this.state.isMonitoring = false;
      this.updateMonitoringStatus(false);
    }
  }

  // === MONITORING ===
  async startMonitoring() {
    try {
      console.log('🚀 Démarrage du monitoring...');
      
      if (Object.keys(this.state.folderCategories).length === 0) {
        this.showNotification('Configuration requise', 'Veuillez d\'abord configurer des dossiers à surveiller', 'warning');
        return;
      }
      
      const result = await window.electronAPI.startMonitoring();
      
      if (result.success) {
        this.state.isMonitoring = true;
        this.updateMonitoringStatus(true);
        this.showNotification('Monitoring démarré', result.message, 'success');
        console.log('✅ Monitoring démarré avec succès');
      } else {
        this.showNotification('Erreur de monitoring', result.message || 'Impossible de démarrer le monitoring', 'danger');
      }
    } catch (error) {
      console.error('❌ Erreur démarrage monitoring:', error);
      this.showNotification('Erreur', error.message, 'danger');
    }
  }

  async stopMonitoring() {
    try {
      console.log('🛑 Arrêt du monitoring...');
      const result = await window.electronAPI.stopMonitoring();
      
      this.state.isMonitoring = false;
      this.updateMonitoringStatus(false);
      this.showNotification('Monitoring arrêté', result.message || 'Monitoring arrêté avec succès', 'info');
      console.log('✅ Monitoring arrêté');
    } catch (error) {
      console.error('❌ Erreur arrêt monitoring:', error);
      this.showNotification('Erreur', error.message, 'danger');
    }
  }

  // === CONFIGURATION DES DOSSIERS ===
  async showAddFolderModal() {
    try {
      // Récupérer la structure des dossiers depuis Outlook
      const mailboxes = await window.electronAPI.getMailboxes();
      
      if (!mailboxes.mailboxes || mailboxes.mailboxes.length === 0) {
        this.showNotification('Aucune boîte mail', 'Impossible de récupérer les boîtes mail d\'Outlook', 'warning');
        return;
      }

      // Créer le modal pour ajouter un dossier
      this.createFolderSelectionModal(mailboxes.mailboxes);
    } catch (error) {
      console.error('❌ Erreur récupération dossiers:', error);
      this.showNotification('Erreur', 'Impossible de récupérer la liste des dossiers', 'danger');
    }
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
              <form id="add-folder-form">
                <div class="mb-3">
                  <label class="form-label">Boîte mail</label>
                  <select class="form-select" id="mailbox-select" required>
                    <option value="">Sélectionnez une boîte mail</option>
                    ${mailboxes.map(mb => `<option value="${mb.StoreID}">${mb.Name}</option>`).join('')}
                  </select>
                </div>
                <div class="mb-3">
                  <label class="form-label">Dossier</label>
                  <div class="border rounded p-3" style="max-height: 300px; overflow-y: auto; background-color: #f8f9fa;">
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
                  <label class="form-label">Catégorie</label>
                  <select class="form-control" id="category-input" required>
                    <option value="">-- Sélectionner une catégorie --</option>
                    <option value="declarations">Déclarations</option>
                    <option value="reglements">Règlements</option>
                    <option value="mails_simples">Mails simples</option>
                  </select>
                  <div class="form-text">Choisissez la catégorie appropriée pour classer les emails de ce dossier</div>
                </div>
              </form>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Annuler</button>
              <button type="button" class="btn btn-primary" id="save-folder-config">Ajouter</button>
            </div>
          </div>
        </div>
      </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    const modal = new bootstrap.Modal(document.getElementById('folderModal'));
    modal.show();

    // Gérer le changement de boîte mail
    document.getElementById('mailbox-select').addEventListener('change', async (e) => {
      const storeId = e.target.value;
      if (storeId) {
        await this.loadFoldersForMailbox(storeId);
      }
    });

    // Gérer la sauvegarde
    document.getElementById('save-folder-config').addEventListener('click', () => {
      this.saveFolderConfiguration(modal);
    });
  }

  async loadFoldersForMailbox(storeId) {
    try {
      const folderTree = document.getElementById('folder-tree');
      folderTree.innerHTML = '<div class="text-muted"><i class="bi bi-hourglass-split me-2"></i>Chargement...</div>';

      const result = await window.electronAPI.getFolderStructure(storeId);
      console.log('🔍 DEBUG - Recherche de boîte mail pour StoreID:', storeId);
      console.log('🔍 DEBUG - Structure reçue:', result);
      
      if (result.success && result.folders && Array.isArray(result.folders)) {
        // Trouver le bon compte dans le tableau - nouvelle logique plus simple
        let selectedMailbox = null;
        
        // D'abord, récupérer l'adresse email sélectionnée dans la dropdown
        const mailboxSelect = document.getElementById('mailbox-select');
        const selectedEmailAddress = mailboxSelect ? mailboxSelect.selectedOptions[0]?.text : '';
        
        console.log('🔍 DEBUG - Adresse email sélectionnée dans dropdown:', selectedEmailAddress);
        console.log('🔍 DEBUG - StoreID demandé:', storeId);
        
        // Méthode 1: Correspondance par l'adresse email visible dans la dropdown
        if (selectedEmailAddress) {
          selectedMailbox = result.folders.find(mailbox => mailbox.Name === selectedEmailAddress);
          if (selectedMailbox) {
            console.log('✅ Correspondance trouvée (méthode dropdown):', selectedMailbox.Name);
          }
        }
        
        // Méthode 2 (fallback): Correspondance par StoreID si méthode 1 échoue
        if (!selectedMailbox) {
          for (const mailbox of result.folders) {
            console.log('🔍 Test correspondance StoreID pour:', mailbox.Name);
            
            // Correspondance exacte du nom dans le StoreID
            if (storeId.includes(mailbox.Name)) {
              selectedMailbox = mailbox;
              console.log('✅ Correspondance trouvée (méthode StoreID)');
              break;
            }
            
            // Correspondance par parties du nom
            const mailboxUser = mailbox.Name.split('@')[0];
            if (storeId.includes(mailboxUser)) {
              selectedMailbox = mailbox;
              console.log('✅ Correspondance trouvée (méthode partie nom)');
              break;
            }
          }
        }
        
        // Si toujours pas trouvé, prendre le premier (fallback final)
        if (!selectedMailbox && result.folders.length > 0) {
          selectedMailbox = result.folders[0];
          console.log('⚠️ Aucune correspondance - utilisation du premier compte:', selectedMailbox.Name);
        }
        
        console.log('🔍 DEBUG - Boîte mail sélectionnée:', selectedMailbox);
        
        if (selectedMailbox && selectedMailbox.SubFolders) {
          // Créer l'arbre avec les sous-dossiers du compte sélectionné
          let treeHtml = '';
          for (const folder of selectedMailbox.SubFolders) {
            treeHtml += this.createFolderTree(folder, 0);
          }
          folderTree.innerHTML = treeHtml;
          this.initializeFolderTreeEvents();
        } else {
          folderTree.innerHTML = '<div class="text-warning"><i class="bi bi-info-circle me-2"></i>Aucun dossier trouvé pour cette boîte mail</div>';
        }
      } else {
        folderTree.innerHTML = '<div class="text-danger"><i class="bi bi-exclamation-triangle me-2"></i>Erreur de chargement</div>';
        console.error('Erreur récupération dossiers:', result.error || 'Réponse invalide');
      }
    } catch (error) {
      console.error('❌ Erreur chargement dossiers:', error);
      document.getElementById('folder-tree').innerHTML = '<div class="text-danger"><i class="bi bi-exclamation-triangle me-2"></i>Erreur de chargement</div>';
    }
  }

  createFolderTree(structure, level = 0) {
    let html = '';
    
    // Support des deux formats : l'ancien (Path/Subfolders) et le nouveau (FolderPath/SubFolders)
    const folderName = structure.Name;
    const folderPath = structure.Path || structure.FolderPath;
    let subfolders = structure.Subfolders || structure.SubFolders || [];
    const unreadCount = structure.UnreadCount || 0;
    
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
      const hasSubfolders = Array.isArray(subfolders) && subfolders.length > 0;
      const indent = '  '.repeat(level);
      const folderId = `folder_${Math.random().toString(36).substr(2, 9)}`;
      
      html += `
        <div class="folder-item" style="margin-left: ${level * 20}px;">
          <div class="folder-line d-flex align-items-center py-1 folder-selectable" 
               data-path="${folderPath}" 
               data-name="${folderName}"
               style="cursor: pointer; border-radius: 4px; padding: 4px 8px;">
            ${hasSubfolders ? 
              `<i class="bi bi-chevron-right folder-toggle me-2" data-target="${folderId}" style="cursor: pointer; width: 12px;"></i>` : 
              `<span style="width: 12px; margin-right: 8px;"></span>`
            }
            <i class="bi bi-folder me-2 text-warning"></i>
            <span class="folder-name">${folderName}</span>
            ${unreadCount > 0 ? `<span class="badge bg-primary ms-2">${unreadCount}</span>` : ''}
          </div>
          ${hasSubfolders ? `<div class="folder-children" id="${folderId}" style="display: none;">` : ''}
      `;
      
      if (hasSubfolders) {
        for (const subfolder of subfolders) {
          html += this.createFolderTree(subfolder, level + 1);
        }
        html += '</div>';
      }
      
      html += '</div>';
    }

    return html;
  }

  initializeFolderTreeEvents() {
    // Événements pour déplier/replier les dossiers
    document.querySelectorAll('.folder-toggle').forEach(toggle => {
      toggle.addEventListener('click', (e) => {
        e.stopPropagation();
        const targetId = toggle.getAttribute('data-target');
        const targetDiv = document.getElementById(targetId);
        const isExpanded = targetDiv.style.display !== 'none';
        
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

    // Événements pour sélectionner un dossier
    document.querySelectorAll('.folder-selectable').forEach(folder => {
      folder.addEventListener('click', (e) => {
        // Désélectionner tous les autres dossiers
        document.querySelectorAll('.folder-selectable').forEach(f => {
          f.classList.remove('bg-primary', 'text-white');
          f.classList.add('text-dark');
        });
        
        // Sélectionner le dossier cliqué
        folder.classList.add('bg-primary', 'text-white');
        folder.classList.remove('text-dark');
        
        // Mettre à jour le champ caché
        const path = folder.getAttribute('data-path');
        const name = folder.getAttribute('data-name');
        document.getElementById('selected-folder-path').value = path;
        
        console.log('Dossier sélectionné:', { name, path });
      });
      
      // Effet de survol
      folder.addEventListener('mouseenter', (e) => {
        if (!folder.classList.contains('bg-primary')) {
          folder.classList.add('bg-light');
        }
      });
      
      folder.addEventListener('mouseleave', (e) => {
        if (!folder.classList.contains('bg-primary')) {
          folder.classList.remove('bg-light');
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
      const folderPath = document.getElementById('selected-folder-path').value;
      const category = document.getElementById('category-input').value.trim();
      
      // Récupérer le nom du dossier depuis l'élément sélectionné
      const selectedFolder = document.querySelector('.folder-selectable.bg-primary');
      const folderName = selectedFolder ? selectedFolder.getAttribute('data-name') : 'Dossier';

      if (!folderPath || !category) {
        this.showNotification('Champs requis', 'Veuillez sélectionner un dossier et saisir une catégorie', 'warning');
        return;
      }

      // Ajouter à la configuration locale
      this.state.folderCategories[folderPath] = {
        category: category,
        name: folderName.trim()
      };

      // Sauvegarder sur disque
      const result = await window.electronAPI.saveFoldersConfig({
        folderCategories: this.state.folderCategories
      });

      if (result.success) {
        this.updateFolderConfigDisplay();
        
        // Fermer le modal correctement avec Bootstrap
        const modalElement = document.getElementById('folderModal');
        if (modalElement) {
          const modal = bootstrap.Modal.getInstance(modalElement);
          if (modal) {
            modal.hide();
          }
        }
        
        this.showNotification('Configuration sauvegardée', `Dossier "${folderName}" ajouté à la catégorie "${category}"`, 'success');
        console.log('✅ Configuration de dossier sauvegardée');
      } else {
        this.showNotification('Erreur de sauvegarde', result.error || 'Impossible de sauvegarder la configuration', 'danger');
      }
    } catch (error) {
      console.error('❌ Erreur sauvegarde configuration:', error);
      this.showNotification('Erreur', error.message, 'danger');
    }
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
    this.animateCounterUpdate('total-emails', stats.totalEmails || 0);
    this.animateCounterUpdate('emails-unread', stats.unreadTotal || 0);
    this.animateCounterUpdate('emails-today', stats.emailsToday || 0);
    
    // Calcul et affichage des pourcentages
    const totalEmails = stats.totalEmails || 0;
    const unreadEmails = stats.unreadTotal || 0;
    
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
    
    const currentValue = parseInt(element.textContent) || 0;
    
    if (currentValue !== newValue) {
      element.classList.add('updating');
      
      // Animation de comptage
      const duration = 500;
      const steps = 20;
      const increment = (newValue - currentValue) / steps;
      let currentStep = 0;
      
      const counter = setInterval(() => {
        currentStep++;
        const value = Math.round(currentValue + (increment * currentStep));
        element.textContent = value;
        
        if (currentStep >= steps) {
          clearInterval(counter);
          element.textContent = newValue;
          element.classList.remove('updating');
        }
      }, duration / steps);
    }
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

  // Nouvelle méthode pour mettre à jour le statut de monitoring
  updateMonitoringStatus() {
    const statusEl = document.getElementById('monitoring-status');
    const progressEl = document.getElementById('monitoring-progress');
    const performanceEl = document.getElementById('monitoring-performance');
    
    if (!statusEl) return;
    
    const folderCount = Object.keys(this.state.folderCategories).length;
    
    if (folderCount > 0 && this.state.isMonitoring) {
      statusEl.innerHTML = `<span class="status-indicator status-connected me-1"></span>Actif`;
      if (progressEl) {
        progressEl.style.width = '100%';
        progressEl.className = 'progress-bar bg-success';
      }
      if (performanceEl) {
        performanceEl.textContent = 'Excellent';
        performanceEl.className = 'fw-bold text-success';
      }
    } else if (folderCount > 0) {
      statusEl.innerHTML = `<span class="status-indicator status-connecting me-1"></span>Initialisation...`;
      if (progressEl) {
        progressEl.style.width = '60%';
        progressEl.className = 'progress-bar bg-warning';
      }
      if (performanceEl) {
        performanceEl.textContent = 'En cours';
        performanceEl.className = 'fw-bold text-warning';
      }
    } else {
      statusEl.innerHTML = `<span class="status-indicator status-disconnected me-1"></span>Arrêté`;
      if (progressEl) {
        progressEl.style.width = '0%';
        progressEl.className = 'progress-bar bg-danger';
      }
      if (performanceEl) {
        performanceEl.textContent = 'Arrêté';
        performanceEl.className = 'fw-bold text-danger';
      }
    }
  }

  updateEmailsTable() {
    const tbody = document.getElementById('emails-table');
    if (!tbody) return;

    if (this.state.recentEmails.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="5" class="text-center text-muted py-4">
            <i class="bi bi-inbox me-2"></i>Aucun email récent trouvé
          </td>
        </tr>
      `;
      return;
    }

    tbody.innerHTML = this.state.recentEmails.map(email => `
      <tr>
        <td>${new Date(email.received_time || email.created_at).toLocaleDateString()}</td>
        <td>
          <div class="fw-bold">${this.escapeHtml(email.sender_name || 'Inconnu')}</div>
          <small class="text-muted">${this.escapeHtml(email.sender_email || '')}</small>
        </td>
        <td>${this.escapeHtml(email.subject || '(Sans objet)')}</td>
        <td>
          <span class="badge bg-secondary">${this.escapeHtml(email.folder_name || 'Inconnu')}</span>
        </td>
        <td>
          ${email.is_read ? 
            '<span class="badge bg-success">Lu</span>' : 
            '<span class="badge bg-warning">Non lu</span>'
          }
        </td>
      </tr>
    `).join('');
  }

  updateFolderConfigDisplay() {
    const container = document.getElementById('folders-tree');
    if (!container) return;

    const folders = Object.entries(this.state.folderCategories);
    
    // Mettre à jour le compteur dans l'en-tête
    const countBadge = document.getElementById('monitored-folders-count');
    if (countBadge) {
      countBadge.textContent = folders.length;
    }
    
    if (folders.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <i class="bi bi-folder-plus"></i>
          <h4>Aucun dossier configuré</h4>
          <p>Commencez par ajouter un dossier à surveiller</p>
          <button class="btn btn-primary btn-modern mt-3" onclick="document.getElementById('add-folder').click()">
            <i class="bi bi-plus-circle me-2"></i>Ajouter votre premier dossier
          </button>
        </div>
      `;
      return;
    }

    // Générer les cartes modernes pour chaque dossier
    const foldersHtml = folders.map(([path, config]) => {
      const categoryClass = this.getCategoryClass(config.category);
      const categoryIcon = this.getCategoryIcon(config.category);
      
      return `
        <div class="folder-card mb-3 animate-slide-up">
          <div class="folder-header p-3">
            <div class="d-flex justify-content-between align-items-start">
              <div class="flex-grow-1">
                <div class="d-flex align-items-center mb-2">
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
                    <button class="dropdown-item" onclick="app.editFolderCategory('${this.escapeHtml(path)}')">
                      <i class="bi bi-pencil me-2"></i>Modifier la catégorie
                    </button>
                  </li>
                  <li>
                    <button class="dropdown-item" onclick="app.viewFolderStats('${this.escapeHtml(path)}')">
                      <i class="bi bi-graph-up me-2"></i>Voir les statistiques
                    </button>
                  </li>
                  <li><hr class="dropdown-divider"></li>
                  <li>
                    <button class="dropdown-item text-danger" onclick="app.removeFolderConfig('${this.escapeHtml(path)}')">
                      <i class="bi bi-trash me-2"></i>Supprimer
                    </button>
                  </li>
                </ul>
              </div>
            </div>
          </div>
          <div class="p-3 pt-0">
            <div class="row g-3">
              <div class="col-4">
                <div class="text-center">
                  <div class="h6 mb-1 text-primary">0</div>
                  <small class="text-muted">Emails</small>
                </div>
              </div>
              <div class="col-4">
                <div class="text-center">
                  <div class="h6 mb-1 text-success">0</div>
                  <small class="text-muted">Traités</small>
                </div>
              </div>
              <div class="col-4">
                <div class="text-center">
                  <div class="h6 mb-1 text-warning">0</div>
                  <small class="text-muted">En attente</small>
                </div>
              </div>
            </div>
            <div class="mt-3">
              <div class="d-flex justify-content-between align-items-center mb-1">
                <small class="text-muted">Activité</small>
                <small class="text-muted">Faible</small>
              </div>
              <div class="progress" style="height: 6px;">
                <div class="progress-bar bg-success" role="progressbar" style="width: 25%"></div>
              </div>
            </div>
          </div>
        </div>
      `;
    }).join('');

    container.innerHTML = `
      <div class="folders-list">
        ${foldersHtml}
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
        
        // Mettre à jour les filtres
        this.updateEmailFilters();
        
        console.log(`✅ ${emails.length} emails chargés`);
      } else {
        console.warn('⚠️ Pas d\'emails trouvés');
        this.state.recentEmails = [];
        this.showEmptyEmailsState();
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
      
      // Déterminer le dossier depuis le chemin
      const folderName = this.extractFolderName(email.folder_name || email.FolderPath || 'Boîte de réception');
      const folderCategory = this.getFolderCategory(email.folder_name || email.FolderPath);
      
      return `
        <tr class="email-row ${!email.is_read && email.is_read !== undefined ? 'table-warning' : (email.UnRead ? 'table-warning' : '')}" data-email-id="${email.outlook_id || email.EntryID}">
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
              ${(!email.is_read && email.is_read !== undefined) || email.UnRead ? 
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

  animateCounterUpdate(elementId, newValue) {
    const element = document.getElementById(elementId);
    if (!element) return;

    const currentValue = parseInt(element.textContent) || 0;
    if (currentValue === newValue) return;

    // Animation simple de compteur
    const duration = 500;
    const start = performance.now();
    const startValue = currentValue;
    const difference = newValue - startValue;

    const animate = (currentTime) => {
      const elapsed = currentTime - start;
      const progress = Math.min(elapsed / duration, 1);
      const value = Math.round(startValue + (difference * progress));
      
      element.textContent = value;
      
      if (progress < 1) {
        requestAnimationFrame(animate);
      }
    };

    requestAnimationFrame(animate);
  }

  getUsageColorClass(usage) {
    if (usage >= 80) return 'bg-danger';
    if (usage >= 60) return 'bg-warning';
    if (usage >= 40) return 'bg-info';
    return 'bg-success';
  }

  updateMonitoringStatus(isRunning) {
    const statusContainer = document.getElementById('monitoring-status');
    const startBtn = document.getElementById('start-monitoring');
    const stopBtn = document.getElementById('stop-monitoring');
    
    if (statusContainer) {
      const foldersCount = Object.keys(this.state.folderCategories).length;
      const currentTime = new Date().toLocaleTimeString('fr-FR', { 
        hour: '2-digit', 
        minute: '2-digit',
        second: '2-digit'
      });
      
      if (isRunning) {
        statusContainer.innerHTML = `
          <div class="d-flex align-items-center justify-content-between mb-3">
            <div class="d-flex align-items-center">
              <div class="status-indicator status-connected me-3"></div>
              <div>
                <div class="fw-bold text-success">Surveillance active</div>
                <small class="text-muted">${foldersCount} dossier(s) surveillé(s)</small>
              </div>
            </div>
            <i class="bi bi-play-circle-fill fs-3 text-success"></i>
          </div>
          <div class="bg-success bg-opacity-10 p-2 rounded-3 mb-3">
            <div class="d-flex justify-content-between align-items-center">
              <small class="text-success fw-medium">
                <i class="bi bi-clock me-1"></i>Synchronisation active
              </small>
              <small class="text-success">${currentTime}</small>
            </div>
          </div>
        `;
        
        // Mettre à jour les métriques de performance
        this.updatePerformanceMetrics(true);
      } else {
        statusContainer.innerHTML = `
          <div class="d-flex align-items-center justify-content-between mb-3">
            <div class="d-flex align-items-center">
              <div class="status-indicator status-disconnected me-3"></div>
              <div>
                <div class="fw-bold text-muted">Surveillance arrêtée</div>
                <small class="text-muted">${foldersCount} dossier(s) configuré(s)</small>
              </div>
            </div>
            <i class="bi bi-pause-circle fs-3 text-muted"></i>
          </div>
          <div class="bg-light p-2 rounded-3 mb-3">
            <div class="text-center">
              <small class="text-muted">
                <i class="bi bi-info-circle me-1"></i>Cliquez sur "Démarrer" pour activer la surveillance
              </small>
            </div>
          </div>
        `;
        
        // Réinitialiser les métriques de performance
        this.updatePerformanceMetrics(false);
      }
    }
    
    if (startBtn && stopBtn) {
      startBtn.style.display = isRunning ? 'none' : 'block';
      stopBtn.style.display = isRunning ? 'block' : 'none';
      
      // Ajouter les classes modernes
      startBtn.className = 'btn btn-success btn-lg btn-modern';
      stopBtn.className = 'btn btn-danger btn-lg btn-modern';
    }
    
    // Mettre à jour l'état global
    this.state.isMonitoring = isRunning;
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
      if (element) {
        element.textContent = value || '--';
      } else {
        // Log uniquement en mode debug pour éléments manquants
        // console.warn(`⚠️ Élément DOM manquant: ${id}`);
      }
    };
    
    // Métriques quotidiennes (utiliser les vrais ID du HTML)
    const daily = vbaMetrics.daily || {};
    safeSetContent('emails-today', daily.emailsReceived);
    safeSetContent('emails-unread', daily.emailsUnread);
    
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
        
        // 3. Afficher une notification de traitement
        this.showNotification('Suppression en cours...', 'Le dossier est en cours de suppression', 'info');
        
        // 4. Puis faire l'appel API en arrière-plan
        await window.electronAPI.removeFolderFromMonitoring({ folderPath });
        
        // 5. Recharger la configuration depuis la base pour s'assurer de la cohérence
        await this.loadFoldersConfiguration();
        
        // 6. Rafraîchir l'affichage avec les données actualisées
        this.updateFolderConfigDisplay();
        await this.refreshFoldersDisplay();
        
        // 7. Attendre un peu puis rafraîchir les statistiques
        setTimeout(async () => {
          await this.refreshStats();
          await this.refreshEmails();
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
      // Récupérer les valeurs du formulaire
      const settings = {
        monitoring: {
          treatReadEmailsAsProcessed: document.getElementById('treat-read-as-processed')?.checked || false,
          scanInterval: parseInt(document.getElementById('sync-interval')?.value) || 30000,
          autoStart: document.getElementById('auto-start-monitoring')?.checked || true
        },
        ui: {
          theme: "default",
          language: "fr",
          emailsLimit: parseInt(document.getElementById('emails-limit')?.value) || 20
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
        this.showNotification('Paramètres sauvegardés', 'Vos préférences ont été enregistrées avec succès', 'success');
        console.log('✅ Paramètres sauvegardés:', settings);
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
            theme: "default",
            language: "fr",
            emailsLimit: 20
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
          // Mettre à jour le formulaire
          this.loadSettingsIntoForm(defaultSettings);
          this.showNotification('Paramètres réinitialisés', 'Les paramètres par défaut ont été restaurés', 'info');
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
      } else {
        console.warn('⚠️ Erreur chargement paramètres, utilisation des valeurs par défaut');
      }
    } catch (error) {
      console.error('❌ Erreur chargement paramètres:', error);
    }
  }

  loadSettingsIntoForm(settings) {
    // Monitoring
    if (settings.monitoring) {
      const treatReadCheckbox = document.getElementById('treat-read-as-processed');
      if (treatReadCheckbox) {
        treatReadCheckbox.checked = settings.monitoring.treatReadEmailsAsProcessed || false;
      }
      
      const syncInterval = document.getElementById('sync-interval');
      if (syncInterval) {
        syncInterval.value = settings.monitoring.scanInterval || 30000;
      }
      
      const autoStartCheckbox = document.getElementById('auto-start-monitoring');
      if (autoStartCheckbox) {
        autoStartCheckbox.checked = settings.monitoring.autoStart !== false;
      }
    }
    
    // UI
    if (settings.ui) {
      const emailsLimit = document.getElementById('emails-limit');
      if (emailsLimit) {
        emailsLimit.value = settings.ui.emailsLimit || 20;
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

  handleTabChange(target) {
    // Recharger les données spécifiques à l'onglet immédiatement
    // Changement d'onglet - log conservé pour navigation
    console.log(`🔄 Changement d'onglet vers: ${target}`);
    
    switch (target) {
      case '#dashboard':
        this.loadStats();
        this.loadCategoryStats();
        this.loadVBAMetrics();
        break;
      case '#emails':
        this.loadRecentEmails();
        break;
      case '#monitoring':
        // Configuration déjà chargée
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
    // Créer une notification Bootstrap
    const alertClass = `alert-${type === 'error' ? 'danger' : type}`;
    const icon = {
      success: 'check-circle',
      danger: 'exclamation-triangle',
      warning: 'exclamation-circle',
      info: 'info-circle'
    }[type] || 'info-circle';

    const notificationHtml = `
      <div class="alert ${alertClass} alert-dismissible fade show position-fixed" 
           style="top: 20px; right: 20px; z-index: 9999; min-width: 300px;" role="alert">
        <i class="bi bi-${icon} me-2"></i>
        <strong>${this.escapeHtml(title)}</strong>
        <div>${this.escapeHtml(message)}</div>
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
      </div>
    `;

    document.body.insertAdjacentHTML('beforeend', notificationHtml);
    
    // Auto-supprimer après 5 secondes
    setTimeout(() => {
      const alert = document.querySelector('.alert:last-of-type');
      if (alert) {
        alert.remove();
      }
    }, 5000);
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
  showAbout() {
    const aboutContent = `
      <div class="text-center">
        <i class="bi bi-envelope-check display-4 text-primary mb-3"></i>
        <h4>Mail Monitor</h4>
        <p class="text-muted">Surveillance professionnelle des emails Outlook</p>
        <hr>
        <p><strong>Version:</strong> 1.0.0</p>
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
      alert('Mail Monitor v1.0.0\n© 2025 Tanguy Raingeard - Tous droits réservés');
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
    // Bouton traité par ADG
    const adjustBtn = document.getElementById('add-manual-adjustment');
    if (adjustBtn) {
      adjustBtn.addEventListener('click', () => this.handleManualAdjustment());
    }

    // Bouton de sauvegarde des paramètres
    const saveSettingsBtn = document.getElementById('save-weekly-settings');
    if (saveSettingsBtn) {
      saveSettingsBtn.addEventListener('click', () => this.saveWeeklySettings());
    }

    // Bouton d'ouverture des paramètres
    const settingsBtn = document.getElementById('weekly-settings-btn');
    if (settingsBtn) {
      settingsBtn.addEventListener('click', () => this.openWeeklySettings());
    }

    // Actualisation automatique des stats hebdomadaires
    setInterval(() => {
      this.refreshCurrentWeekStats();
    }, 30000); // Toutes les 30 secondes
  }

  /**
   * Charge les statistiques de la semaine actuelle
   */
  async loadCurrentWeekStats() {
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
        this.updateCurrentWeekDisplay(weekData);
      } else {
        console.error('❌ Erreur lors du chargement des stats hebdomadaires:', response.error);
        this.showWeeklyError(response.error);
      }
    } catch (error) {
      console.error('❌ Erreur lors du chargement des stats hebdomadaires:', error);
      this.showWeeklyError('Erreur de communication avec le serveur');
    }
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
    const weeklyTab = document.getElementById('weekly-tab');
    if (weeklyTab && weeklyTab.classList.contains('active')) {
      await this.loadCurrentWeekStats();
    }
  }

  /**
   * Gère les données traitées par ADG
   */
  async handleManualAdjustment() {
    const form = document.getElementById('manual-adjustment-form');
    const formData = new FormData(form);
    
    const adjustmentData = {
      category: formData.get('adjustment-category'),
      quantity: parseInt(formData.get('adjustment-quantity')),
      type: formData.get('adjustment-type'),
      description: formData.get('adjustment-description') || ''
    };

    // Validation
    if (!adjustmentData.category || !adjustmentData.quantity || adjustmentData.quantity <= 0) {
      this.showNotification('Erreur', 'Veuillez remplir tous les champs requis', 'warning');
      return;
    }

    try {
      const response = await window.electronAPI.invoke('api-weekly-adjust-count', adjustmentData);
      
      if (response.success) {
        this.showNotification('Données ajoutées', 
          `${adjustmentData.quantity} ${adjustmentData.type} ajouté(s) pour ${adjustmentData.category}`, 
          'success'
        );
        
        // Réinitialiser le formulaire
        form.reset();
        
        // Recharger les statistiques
        await this.loadCurrentWeekStats();
        await this.loadWeeklyHistory();
        
      } else {
        this.showNotification('Erreur', response.error, 'danger');
      }
    } catch (error) {
      console.error('Erreur lors de l\'ajout de données par ADG:', error);
      this.showNotification('Erreur', 'Impossible d\'ajouter les données', 'danger');
    }
  }

  /**
   * Charge l'historique hebdomadaire
   */
  async loadWeeklyHistory() {
    try {
      const response = await window.electronAPI.invoke('api-weekly-history', { limit: 10 });
      if (response.success) {
        this.updateWeeklyHistoryDisplay(response.data);
      } else {
        console.error('Erreur lors du chargement de l\'historique:', response.error);
      }
    } catch (error) {
      console.error('Erreur lors du chargement de l\'historique:', error);
    }
  }

  /**
   * Met à jour l'affichage de l'historique
   */
  updateWeeklyHistoryDisplay(historyData) {
    const historyTable = document.getElementById('weekly-history-table');
    const loadingElement = document.getElementById('weekly-loading');
    const noDataElement = document.getElementById('weekly-no-data');
    
    if (!historyTable) return;

    const tbody = historyTable.querySelector('tbody');
    if (!tbody) return;

    // Masquer l'indicateur de chargement
    if (loadingElement) {
      loadingElement.classList.add('d-none');
    }

    let historyHtml = '';
    
    if (historyData.length === 0) {
      // Afficher le message "aucune donnée" et masquer le tableau
      if (noDataElement) {
        noDataElement.classList.remove('d-none');
      }
      historyTable.style.display = 'none';
      return;
    } else {
      // Masquer le message "aucune donnée" et afficher le tableau
      if (noDataElement) {
        noDataElement.classList.add('d-none');
      }
      historyTable.style.display = 'table';
      
      for (const week of historyData) {
        // Calculer le total des stocks pour la semaine
        const totalStock = week.categories.reduce((sum, category) => sum + category.stockEndWeek, 0);
        
        // Créer le contenu d'évolution avec design 2025
        const evolutionHtml = this.createEvolutionDisplay(week.evolution);
        
        // Créer un ID unique pour la semaine pour gérer le survol
        const weekId = `week-${week.weekDisplay.replace(/[^a-zA-Z0-9]/g, '-')}`;
        
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
              <td rowspan="3" class="text-center align-middle" style="vertical-align: middle !important;">${evolutionHtml}</td>`;
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
    const colorClass = isPositive ? 'text-success' : 'text-danger';
    const bgClass = isPositive ? 'bg-success' : 'bg-danger';
    const sign = isPositive ? '+' : '';
    
    return `
      <div class="d-flex align-items-center justify-content-center">
        <div class="evolution-indicator ${isPositive ? 'positive' : 'negative'}">
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
          checkbox.checked = response.data.countReadAsTreated || false;
        }
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

      const response = await window.electronAPI.invoke('api-settings-count-read-as-treated', {
        countReadAsTreated
      });

      if (response.success) {
        this.showNotification('Paramètres sauvegardés', 
          'Les paramètres du suivi hebdomadaire ont été mis à jour', 
          'success'
        );
        
        // Fermer la modal
        const modal = bootstrap.Modal.getInstance(document.getElementById('weeklySettingsModal'));
        if (modal) {
          modal.hide();
        }
        
        // Recharger les statistiques pour appliquer les nouveaux paramètres
        await this.loadCurrentWeekStats();
        
      } else {
        this.showNotification('Erreur', response.error, 'danger');
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
