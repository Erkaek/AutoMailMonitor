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
    
    this.updateInterval = null;
    this.charts = {};
    this.init();
  }

  // === INITIALISATION ===
  async init() {
    console.log('🚀 Initialisation de Mail Monitor...');
    
    try {
      this.setupEventListeners();
      await this.loadConfiguration();
      await this.checkConnection();
      await this.loadInitialData();
      this.startPeriodicUpdates();
      
      this.showNotification('Application prête', 'Mail Monitor est opérationnel', 'success');
      console.log('✅ MailMonitor app initialisée avec succès');
    } catch (error) {
      console.error('❌ Erreur lors de l\'initialisation:', error);
      this.showNotification('Erreur d\'initialisation', error.message, 'danger');
    }
  }

  /**
   * Configuration de l'actualisation automatique
   */
  setupAutoRefresh() {
    console.log('🔄 Configuration de l\'actualisation automatique...');
    
    // Actualisation des statistiques principales toutes les 10 secondes
    this.statsRefreshInterval = setInterval(() => {
      this.performStatsRefresh();
    }, 10000);
    
    // Actualisation des emails récents toutes les 15 secondes
    this.emailsRefreshInterval = setInterval(() => {
      this.performEmailsRefresh();
    }, 15000);
    
    // Actualisation complète toutes les 2 minutes
    this.fullRefreshInterval = setInterval(() => {
      this.performFullRefresh();
    }, 120000);
    
    // Actualisation des métriques VBA toutes les 20 secondes
    this.vbaRefreshInterval = setInterval(() => {
      this.loadVBAMetrics();
    }, 20000);
    
    console.log('✅ Auto-refresh configuré (10s stats, 15s emails, 2min complet, 20s VBA)');
  }

  /**
   * Configuration des listeners d'événements temps réel
   */
  setupRealtimeEventListeners() {
    console.log('🔔 Configuration des événements temps réel...');
    
    // Écouter les mises à jour de statistiques en temps réel
    if (window.electronAPI.onStatsUpdate) {
      window.electronAPI.onStatsUpdate((stats) => {
        console.log('📊 Événement stats temps réel reçu:', stats);
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
        console.log('🔄 Cycle de monitoring terminé:', cycleData);
        this.handleMonitoringCycleComplete(cycleData);
      });
    }
    
    console.log('✅ Événements temps réel configurés');
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
      console.log('🔄 Actualisation complète automatique...');
      await Promise.allSettled([
        this.checkConnection(),
        this.loadStats(),
        this.loadRecentEmails(),
        this.loadCategoryStats(),
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
    
    // Emails
    document.getElementById('refresh-emails')?.addEventListener('click', () => this.loadRecentEmails());
    
    // Monitoring
    document.getElementById('start-monitoring')?.addEventListener('click', () => this.startMonitoring());
    document.getElementById('stop-monitoring')?.addEventListener('click', () => this.stopMonitoring());
    document.getElementById('add-folder')?.addEventListener('click', () => this.showAddFolderModal());
    
    // Paramètres
    document.getElementById('settings-form')?.addEventListener('submit', (e) => this.saveSettings(e));
    document.getElementById('reset-settings')?.addEventListener('click', () => this.resetSettings());
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
      
      if (result.success) {
        this.state.folderCategories = result.folderCategories || {};
        console.log(`✅ Configuration rechargée: ${Object.keys(this.state.folderCategories).length} dossiers configurés`);
      } else {
        console.warn('⚠️ Aucune configuration trouvée lors du rechargement');
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
      this.loadVBAMetrics(), // Nouvelles métriques VBA
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

  async loadRecentEmails() {
    try {
      const emails = await window.electronAPI.getRecentEmails();
      this.state.recentEmails = emails || [];
      this.updateEmailsTable();
    } catch (error) {
      console.error('❌ Erreur chargement emails:', error);
      this.state.recentEmails = [];
      this.updateEmailsTable();
    }
  }

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

  async checkMonitoringStatus() {
    try {
      // Vérifier si des dossiers sont configurés pour le monitoring
      const foldersCount = Object.keys(this.state.folderCategories).length;
      
      if (foldersCount > 0) {
        console.log(`📁 ${foldersCount} dossier(s) configuré(s) - monitoring probablement actif`);
        
        // Assumer que le monitoring est actif si des dossiers sont configurés
        // (car il se démarre automatiquement au lancement)
        this.state.isMonitoring = true;
        this.updateMonitoringStatus(true);
        
        this.showNotification(
          'Monitoring automatique', 
          `Le monitoring automatique est actif sur ${foldersCount} dossier(s)`, 
          'info'
        );
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
    
    document.getElementById('emails-today').textContent = stats.emailsToday || '--';
    document.getElementById('emails-sent').textContent = stats.treatedToday || '--';  // Changé sentToday -> treatedToday
    document.getElementById('emails-unread').textContent = stats.unreadTotal || '--';
    
    if (this.state.lastUpdate) {
      document.getElementById('last-sync').textContent = this.state.lastUpdate.toLocaleTimeString();
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
    
    if (folders.length === 0) {
      container.innerHTML = `
        <div class="text-center text-muted py-4">
          <i class="bi bi-folder-plus fs-1 mb-3"></i>
          <p>Aucun dossier configuré</p>
          <p class="small">Cliquez sur "Ajouter un dossier" pour commencer</p>
        </div>
      `;
      return;
    }

    container.innerHTML = `
      <div class="table-responsive">
        <table class="table table-sm">
          <thead>
            <tr>
              <th>Dossier</th>
              <th>Catégorie</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${folders.map(([path, config]) => `
              <tr>
                <td>
                  <div class="fw-bold">${this.escapeHtml(config.name)}</div>
                  <small class="text-muted">${this.escapeHtml(path)}</small>
                </td>
                <td>
                  <span class="badge bg-primary">${this.escapeHtml(config.category)}</span>
                </td>
                <td>
                  <button class="btn btn-sm btn-outline-danger" onclick="app.removeFolderConfig('${this.escapeHtml(path)}')">
                    <i class="bi bi-trash"></i>
                  </button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  updateMonitoringStatus(isRunning) {
    const statusContainer = document.getElementById('monitoring-status');
    const startBtn = document.getElementById('start-monitoring');
    const stopBtn = document.getElementById('stop-monitoring');
    
    if (statusContainer) {
      const foldersCount = Object.keys(this.state.folderCategories).length;
      
      if (isRunning) {
        statusContainer.innerHTML = `
          <div class="d-flex align-items-center mb-3">
            <span class="status-indicator status-connected me-2"></span>
            <span class="fw-bold text-success">Actif</span>
          </div>
          <small class="text-muted">${foldersCount} dossier(s) surveillé(s)</small>
          <div class="mt-2">
            <small class="text-muted">Synchronisation toutes les secondes</small>
          </div>
        `;
      } else {
        statusContainer.innerHTML = `
          <div class="d-flex align-items-center mb-3">
            <span class="status-indicator status-disconnected me-2"></span>
            <span>Arrêté</span>
          </div>
          <small class="text-muted">${foldersCount} dossier(s) configuré(s)</small>
        `;
      }
    }
    
    if (startBtn && stopBtn) {
      startBtn.style.display = isRunning ? 'none' : 'inline-block';
      stopBtn.style.display = isRunning ? 'inline-block' : 'none';
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
    console.log('🎯 Mise à jour métriques VBA:', vbaMetrics);
    
    // Métriques quotidiennes (mise à jour des cartes principales)
    const daily = vbaMetrics.daily || {};
    document.getElementById('emails-today').textContent = daily.emailsReceived || '--';
    document.getElementById('emails-sent').textContent = daily.emailsProcessed || '--';
    document.getElementById('emails-unread').textContent = daily.emailsUnread || '--';
    
    // Métriques hebdomadaires
    const weekly = vbaMetrics.weekly || {};
    document.getElementById('stock-start').textContent = weekly.stockStart || '--';
    document.getElementById('stock-end').textContent = weekly.stockEnd || '--';
    document.getElementById('weekly-arrivals').textContent = weekly.arrivals || '--';
    document.getElementById('weekly-treatments').textContent = weekly.treatments || '--';
    
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
    // Rafraîchir l'affichage des dossiers en rechargeant la boîte mail courante
    const storeSelect = document.getElementById('store-select');
    if (storeSelect && storeSelect.value) {
      await this.loadFoldersForMailbox(storeSelect.value);
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
