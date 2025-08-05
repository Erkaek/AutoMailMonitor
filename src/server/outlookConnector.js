/**
 * Connecteur Outlook COM ultra robuste avec support FFI-NAPI
 * Gestion événements COM en temps réel, PowerShell de fallback, monitoring avancé
 */

const { exec, spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const { EventEmitter } = require('events');

// Importer le nouveau connecteur COM basé sur FFI-NAPI
const COMConnector = require('./comConnector');
let comConnector = null;
let comAvailable = false;

try {
  comConnector = new COMConnector();
  comAvailable = true;
  console.log('✅ FFI-NAPI COM disponible pour les événements');
} catch (error) {
  console.warn('⚠️ FFI-NAPI COM non disponible, utilisation PowerShell uniquement:', error.message);
  comAvailable = false;
}

class OutlookConnector extends EventEmitter {
  constructor() {
    super(); // Appel du constructeur EventEmitter
    
    // DIAGNOSTIC: Compter les instances
    if (!OutlookConnector.instanceCount) {
      OutlookConnector.instanceCount = 0;
    }
    OutlookConnector.instanceCount++;
    console.log(`🔍 DIAGNOSTIC: Création d'une nouvelle instance OutlookConnector #${OutlookConnector.instanceCount}`);
    
    // Initialize core properties
    this.isOutlookConnected = false;
    this.lastCheck = null;
    this.connectionAttempts = 0;
    this.maxRetries = 5;
    this.retryDelay = 2000;
    this.connectionState = 'disconnected';
    this.outlookVersion = null;
    this.lastError = null;
    this.checkInterval = null;
    this.isPowerShellRunning = false;
    
    // COM properties
    this.outlookApp = null;
    this.comConnection = null;
    this.eventHandlers = new Map(); // Stockage des handlers d'événements
    this.useComEvents = comAvailable; // Utiliser COM si FFI-NAPI disponible
    
    // Initialize configuration with defaults
    this.config = {
      timeout: 30000,
      cacheTTL: 300000, // 5 minutes
      healthCheckInterval: 30000, // 30 seconds
      autoReconnect: true,
      enableDetailedLogs: true,
      preferComEvents: true // Préférer les événements COM aux PowerShell
    };
    
    // Initialize data structures
    this.folders = new Map();
    this.stats = new Map();
    
    // Start auto-connection attempt
    this.autoConnect();
  }

  /**
   * Auto-connection attempt on instantiation
   */
  async autoConnect() {
    try {
      this.log('[AUTO-CONNECT] Tentative de connexion automatique...');
      
      // Éviter les double connexions
      if (this.isOutlookConnected || this.connectionState === 'connecting') {
        this.log('[AUTO-CONNECT] Connexion déjà en cours ou établie');
        return;
      }
      
      await this.establishConnection();
      this.log('[AUTO-CONNECT] Connexion automatique réussie');
    } catch (error) {
      this.log(`[AUTO-CONNECT] Échec connexion automatique: ${error.message}`);
      // Don't throw - allow manual connection later
    }
  }

  /**
   * Initialiser la connexion COM avec Outlook
   */
  async initializeComConnection() {
    try {
      if (!comConnector) {
        this.log('[COM] FFI-NAPI COM non disponible, utilisation PowerShell uniquement');
        this.useComEvents = false;
        return;
      }

      this.log('[COM] Initialisation de la connexion COM avec Outlook...');

      // Tenter de se connecter à Outlook via FFI-NAPI
      try {
        await comConnector.connectToOutlook();
        this.outlookApp = comConnector;
        this.log('[COM] Connexion à Outlook via FFI-NAPI réussie');
      } catch (error) {
        this.log(`[COM] Erreur connexion FFI-NAPI: ${error.message}`);
        // Utiliser le fallback PowerShell
        this.useComEvents = false;
        this.outlookApp = null;
        this.comConnection = null;
        return;
      }

      // Vérifier que la connexion fonctionne
      try {
        const nameSpace = await comConnector.getNamespace();
        if (!nameSpace) {
          throw new Error('Impossible d\'accéder au namespace MAPI');
        }

        this.comConnection = nameSpace;
        this.isOutlookConnected = true;
        this.connectionState = 'connected-com';
        this.useComEvents = true;

        this.log('[COM] ✅ Connexion COM établie avec succès');
      } catch (error) {
        this.log(`[COM] Erreur accès namespace: ${error.message}`);
        this.useComEvents = false;
        this.outlookApp = null;
        this.comConnection = null;
      }

    } catch (error) {
      this.log(`[COM] ❌ Erreur connexion COM: ${error.message}`);
      this.useComEvents = false;
      this.outlookApp = null;
      this.comConnection = null;
      
      // Ne pas lancer d'erreur, utiliser le fallback PowerShell
      this.log('[COM] Fallback vers PowerShell activé');
    }
  }

  /**
   * Configurer les événements COM pour un dossier spécifique
   */
  async setupFolderEvents(folderPath, callbacks) {
    try {
      if (!this.useComEvents || !this.comConnection) {
        throw new Error('Connexion COM non disponible');
      }

      this.log(`[COM-EVENTS] Configuration des événements pour: ${folderPath}`);

      // Récupérer le dossier Outlook
      const folder = this.getOutlookFolder(folderPath);
      if (!folder) {
        throw new Error(`Dossier non trouvé: ${folderPath}`);
      }

      // Configurer les événements du dossier
      const eventHandler = {
        folder: folder,
        callbacks: callbacks,
        
        // Événement ajout d'email
        onItemAdd: (item) => {
          try {
            this.log(`[COM-EVENT] Nouvel email détecté: ${item.Subject}`);
            const mailData = this.extractEmailDataFromComObject(item);
            if (callbacks.onNewMail) {
              callbacks.onNewMail(mailData);
            }
          } catch (error) {
            this.log(`[COM-EVENT] Erreur traitement nouvel email: ${error.message}`);
          }
        },

        // Événement modification d'email
        onItemChange: (item) => {
          try {
            this.log(`[COM-EVENT] Email modifié: ${item.Subject}`);
            const mailData = this.extractEmailDataFromComObject(item);
            if (callbacks.onMailChanged) {
              callbacks.onMailChanged(mailData);
            }
          } catch (error) {
            this.log(`[COM-EVENT] Erreur traitement email modifié: ${error.message}`);
          }
        },

        // Événement suppression d'email
        onItemRemove: () => {
          try {
            this.log(`[COM-EVENT] Email supprimé du dossier`);
            if (callbacks.onMailDeleted) {
              callbacks.onMailDeleted('unknown-id'); // COM ne fournit pas l'ID lors de la suppression
            }
          } catch (error) {
            this.log(`[COM-EVENT] Erreur traitement email supprimé: ${error.message}`);
          }
        }
      };

      // Attacher les événements
      folder.attachEvent('ItemAdd', eventHandler.onItemAdd);
      folder.attachEvent('ItemChange', eventHandler.onItemChange);
      folder.attachEvent('ItemRemove', eventHandler.onItemRemove);

      // Stocker le handler pour le nettoyage ultérieur
      this.eventHandlers.set(folderPath, eventHandler);

      this.log(`[COM-EVENTS] ✅ Événements configurés pour: ${folderPath}`);
      return eventHandler;

    } catch (error) {
      this.log(`[COM-EVENTS] ❌ Erreur configuration événements: ${error.message}`);
      throw error;
    }
  }

  /**
   * Supprimer les événements COM pour un dossier
   */
  async removeFolderEvents(folderPath, handler) {
    try {
      if (!handler || !handler.folder) {
        return;
      }

      this.log(`[COM-EVENTS] Suppression des événements pour: ${folderPath}`);

      // Détacher les événements
      handler.folder.detachEvent('ItemAdd', handler.onItemAdd);
      handler.folder.detachEvent('ItemChange', handler.onItemChange);
      handler.folder.detachEvent('ItemRemove', handler.onItemRemove);

      // Supprimer du stockage
      this.eventHandlers.delete(folderPath);

      this.log(`[COM-EVENTS] ✅ Événements supprimés pour: ${folderPath}`);

    } catch (error) {
      this.log(`[COM-EVENTS] ❌ Erreur suppression événements: ${error.message}`);
    }
  }

  /**
   * Récupérer un dossier Outlook par son chemin
   */
  getOutlookFolder(folderPath) {
    try {
      if (!this.comConnection) {
        throw new Error('Connexion COM non établie');
      }

      // Parser le chemin du dossier
      const pathParts = folderPath.split('\\').filter(part => part.length > 0);
      
      // Commencer par la racine
      let currentFolder = this.comConnection.GetDefaultFolder(6); // olFolderInbox
      
      // Si le chemin commence par une autre racine, l'ajuster
      if (pathParts[0] && pathParts[0].toLowerCase() !== 'inbox') {
        // Essayer de trouver le dossier racine
        const stores = this.comConnection.Stores;
        for (let i = 1; i <= stores.Count; i++) {
          const store = stores.Item(i);
          if (store.DisplayName.toLowerCase().includes(pathParts[0].toLowerCase())) {
            currentFolder = store.GetRootFolder();
            pathParts.shift(); // Supprimer la partie racine du chemin
            break;
          }
        }
      }

      // Naviguer dans le chemin
      for (const folderName of pathParts) {
        const folders = currentFolder.Folders;
        let found = false;
        
        for (let i = 1; i <= folders.Count; i++) {
          const folder = folders.Item(i);
          if (folder.Name.toLowerCase() === folderName.toLowerCase()) {
            currentFolder = folder;
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error(`Dossier non trouvé dans le chemin: ${folderName}`);
        }
      }

      return currentFolder;

    } catch (error) {
      this.log(`[COM] Erreur récupération dossier ${folderPath}: ${error.message}`);
      return null;
    }
  }

  /**
   * Extraire les données d'un email à partir d'un objet COM
   */
  extractEmailDataFromComObject(mailItem) {
    try {
      return {
        id: mailItem.EntryID || `temp-${Date.now()}`,
        subject: mailItem.Subject || 'Sans objet',
        sender: mailItem.SenderName || 'Expéditeur inconnu',
        senderEmail: mailItem.SenderEmailAddress || '',
        receivedDate: mailItem.ReceivedTime ? new Date(mailItem.ReceivedTime) : new Date(),
        body: mailItem.Body || '',
        isRead: mailItem.UnRead === false,
        importance: mailItem.Importance || 1,
        size: mailItem.Size || 0,
        attachmentCount: mailItem.Attachments ? mailItem.Attachments.Count : 0,
        categories: mailItem.Categories || '',
        messageClass: mailItem.MessageClass || ''
      };
    } catch (error) {
      this.log(`[COM] Erreur extraction données email: ${error.message}`);
      return {
        id: `error-${Date.now()}`,
        subject: 'Erreur extraction',
        sender: 'Erreur',
        senderEmail: '',
        receivedDate: new Date(),
        body: '',
        isRead: false,
        importance: 1,
        size: 0,
        attachmentCount: 0,
        categories: '',
        messageClass: ''
      };
    }
  }

  /**
   * Traiter les données d'un email (méthode commune pour COM et PowerShell)
   */
  async processEmailData(emailData) {
    // Cette méthode peut être utilisée pour normaliser les données
    // qu'elles proviennent de COM ou de PowerShell
    return {
      ...emailData,
      processedAt: new Date(),
      source: this.useComEvents ? 'com' : 'powershell'
    };
  }

  /**
   * Async initialization logic previously in constructor
   */
  async init(folderPath, maxEmails) {
    try {
      this.log('[INIT] Démarrage de l\'initialisation OutlookConnector...');
      
      // Ensure we're connected first
      if (!this.isOutlookConnected) {
        this.log('[INIT] Pas encore connecté, tentative de connexion...');
        await this.establishConnection();
      }
      
      this.log('[INIT] Connexion Outlook confirmée');
      
      // If folder path is provided, scan it
      if (folderPath) {
        this.log(`[INIT] Scan des emails du dossier: ${folderPath}`);
        const result = await this.getFolderEmailsWithPagination(folderPath, maxEmails);
        this.log(`[INIT] Scan terminé: ${result.EmailsRetrieved} emails trouvés`);
        return result;
      }
      
      this.log('[INIT] Initialisation terminée sans scan de dossier');
      return null;
      
    } catch (error) {
      this.log(`❌ [INIT] Erreur environnement: ${error.message}`);
      this.connectionState = 'unavailable';
      // Surface error for UI/debug
      console.error('[INIT] Erreur complète:', error);
      throw error;
    }
  }

  async checkCOMPermissions() {
    const testScript = `
      try {
        $ErrorActionPreference = "Stop"
        $comTest = New-Object -ComObject "Shell.Application"
        if ($comTest) {
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($comTest) | Out-Null
          Write-Output "COM_OK"
        }
      } catch {
        Write-Output "COM_ERROR: $($_.Exception.Message)"
      }
    `;
    
    const result = await this.executePowerShell(testScript, 5000);
    
    if (!result.includes('COM_OK')) {
      throw new Error('Permissions COM insuffisantes. Exécutez en tant qu\'administrateur.');
    }
    
    this.log('✅ Permissions COM vérifiées');
  }

  async detectOutlookVersion() {
    const versionScript = `
      try {
        $outlook = Get-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Office\\*\\Outlook\\InstallRoot" -ErrorAction SilentlyContinue
        if ($outlook) {
          $version = $outlook.Path | Split-Path -Parent | Split-Path -Leaf
          Write-Output "VERSION:$version"
        } else {
          Write-Output "VERSION:Unknown"
        }
      } catch {
        Write-Output "VERSION:Error"
      }
    `;
    
    try {
      const result = await this.executePowerShell(versionScript, 5000);
      if (result.includes('VERSION:')) {
        this.outlookVersion = result.split('VERSION:')[1].trim();
        this.log(`📧 Version Outlook détectée: ${this.outlookVersion}`);
      }
    } catch (error) {
      this.log(`⚠️ Impossible de détecter la version Outlook: ${error.message}`);
    }
  }

  async establishConnection() {
    // Protection contre les double connexions
    if (this.isOutlookConnected) {
      this.log('[CONNECT] Déjà connecté à Outlook');
      return;
    }
    
    if (this.connectionState === 'connecting') {
      this.log('[CONNECT] Connexion déjà en cours...');
      return;
    }
    
    this.log('[CONNECT] Établissement de la connexion Outlook...');
    this.connectionState = 'connecting';
    
    try {
      // Étape 1: Vérifier les permissions COM
      this.log('[CONNECT] Vérification des permissions COM...');
      await this.checkCOMPermissions();
      this.log('[CONNECT] Permissions COM OK');
      
      // Étape 2: Détecter la version d'Outlook installée
      this.log('[CONNECT] Détection de la version Outlook...');
      await this.detectOutlookVersion();
      this.log('[CONNECT] Version Outlook détectée');
      
      // Étape 3: Vérifier le processus Outlook
      const processRunning = await this.checkOutlookProcess();
      
      if (!processRunning) {
        this.log('⚠️ Processus Outlook non détecté, tentative de démarrage automatique...');
        await this.startOutlook();
        
        // Attendre que le processus démarre avec plus de patience
        await this.waitForProcess('OUTLOOK.EXE', 60000); // 60 secondes
      } else {
        this.log('✅ Processus Outlook déjà en cours d\'exécution');
        // Même si le processus tourne, attendre un peu pour s'assurer qu'il est prêt
        await this.sleep(2000);
      }
      
      // Étape 4: Tester la connexion COM avec retry
      await this.testCOMConnection();
      
      // Étape 5: Initialiser les données de base
      await this.loadInitialData();
      
      this.connectionState = 'connected';
      this.isOutlookConnected = true;
      this.connectionAttempts = 0;
      this.lastError = null;
      
      this.log('✅ Connexion Outlook établie avec succès');
      
      // Start health monitoring
      this.startHealthMonitoring();
      
    } catch (error) {
      this.handleConnectionError(error);
      throw error; // Re-throw for caller to handle
    }
  }

  async testCOMConnection() {
    const connectionScript = `
      try {
        $ErrorActionPreference = "Stop"
        $outlook = New-Object -ComObject Outlook.Application
        
        if ($outlook) {
          $namespace = $outlook.GetNamespace("MAPI")
          $defaultStore = $namespace.DefaultStore
          
          $result = @{
            Connected = $true
            Version = $outlook.Version
            ProfileName = $namespace.CurrentProfileName
            StoreCount = $namespace.Stores.Count
            DefaultStoreName = $defaultStore.DisplayName
          }
          
          # Test de connexion léger - pas d'accès aux données
          # (Évite l'accès non nécessaire à la boîte de réception)
          
          # Liberation des objets COM
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
          
          $result | ConvertTo-Json -Compress
        } else {
          throw "Impossible de creer l'objet Outlook.Application"
        }
      } catch {
        @{
          Connected = $false
          Error = $_.Exception.Message
          ErrorType = $_.Exception.GetType().Name
        } | ConvertTo-Json -Compress
      }
    `;
    
    for (let attempt = 1; attempt <= this.maxRetries; attempt++) {
      try {
        this.log(`🔄 Test connexion COM (tentative ${attempt}/${this.maxRetries})`);
        
        const result = await this.executePowerShell(connectionScript, this.config.timeout);
        const data = JSON.parse(result);
        
        if (data.Connected) {
          this.log(`✅ Connexion COM réussie - Profil: ${data.ProfileName}`);
          this.log(`📊 ${data.StoreCount} store(s) disponible(s), store par défaut: ${data.DefaultStoreName}`);
          return data;
        } else {
          throw new Error(`Connexion echouee: ${data.Error} (${data.ErrorType})`);
        }
        
      } catch (error) {
        this.log(`❌ Tentative ${attempt} echouee: ${error.message}`);
        
        if (attempt < this.maxRetries) {
          const delay = this.retryDelay * attempt; // Delai progressif
          this.log(`⏳ Nouvelle tentative dans ${delay}ms...`);
          await this.sleep(delay);
        } else {
          throw new Error(`Echec de la connexion COM apres ${this.maxRetries} tentatives: ${error.message}`);
        }
      }
    }
  }

  async loadInitialData() {
    this.log('📚 Chargement des donnees initiales...');
    
    try {
      // Charger les dossiers en parallele
      const [folders, stats] = await Promise.allSettled([
        this.loadFoldersRobust(),
        this.loadStatsRobust()
      ]);
      
      if (folders.status === 'fulfilled') {
        this.log(`📁 ${folders.value.length} dossiers charges`);
      } else {
        this.log(`⚠️ Erreur chargement dossiers: ${folders.reason.message}`);
      }
      
      if (stats.status === 'fulfilled') {
        this.log(`📊 Statistiques chargees`);
      } else {
        this.log(`⚠️ Erreur chargement stats: ${stats.reason.message}`);
      }
      
    } catch (error) {
              this.log(`⚠️ Erreur chargement donnees: ${error.message}`);
    }
  }

  async loadFoldersRobust() {
    const foldersScript = `
      try {
        $ErrorActionPreference = "Stop"
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        $folders = @()
        
        # Dossiers par defaut
        $defaultFolders = @{
          6 = "Inbox"           # olFolderInbox
          5 = "SentMail"        # olFolderSentMail
          16 = "Drafts"         # olFolderDrafts
          23 = "Junk"           # olFolderJunk
          3 = "DeletedItems"    # olFolderDeletedItems
        }
        
        foreach ($folderType in $defaultFolders.Keys) {
          try {
            $folder = $namespace.GetDefaultFolder($folderType)
            if ($folder) {
              $folders += @{
                Name = $folder.Name
                Path = $folder.FolderPath
                Type = $defaultFolders[$folderType]
                ItemCount = $folder.Items.Count
                UnreadCount = $folder.UnReadItemCount
                Size = $folder.Size
                LastModified = $folder.LastModificationTime.ToString("yyyy-MM-ddTHH:mm:ss")
              }
              [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folder) | Out-Null
            }
          } catch {
            # Continuer meme si un dossier specifique echoue
          }
        }
        
        # Dossiers personnalisés dans la racine (protection contre null)
        try {
          $store = $namespace.DefaultStore
          if ($store) {
            $rootFolder = $store.GetRootFolder()
            if ($rootFolder -and $rootFolder.Folders) {
              foreach ($subfolder in $rootFolder.Folders) {
                try {
                  if ($subfolder -and $subfolder.DefaultItemType -eq 0 -and $subfolder.Name) { # olMailItem
                    # Éviter les doublons avec les dossiers par défaut
                    $isDuplicate = $false
                    foreach ($existingFolder in $folders) {
                      if ($existingFolder.Name -eq $subfolder.Name) {
                        $isDuplicate = $true
                        break
                      }
                    }
                    
                    if (-not $isDuplicate) {
                      $folders += @{
                        Name = $subfolder.Name
                        Path = $subfolder.FolderPath
                        Type = "Custom"
                        ItemCount = $subfolder.Items.Count
                        UnreadCount = $subfolder.UnReadItemCount
                        Size = $subfolder.Size
                        LastModified = $subfolder.LastModificationTime.ToString("yyyy-MM-ddTHH:mm:ss")
                      }
                    }
                  }
                  if ($subfolder) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($subfolder) | Out-Null
                  }
                } catch {
                  # Ignorer les dossiers problématiques
                }
              }
            }
            if ($rootFolder) {
              [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
            }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($store) | Out-Null
          }
        } catch {
          # Ignorer l'erreur et continuer
        }
        
        # Libération des objets
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        
        # Assurer que la sortie est toujours un tableau JSON valide
        if ($folders.Count -eq 0) {
          Write-Output "[]"
        } elseif ($folders.Count -eq 1) {
          Write-Output "[$($folders[0] | ConvertTo-Json -Compress)]"
        } else {
          $folders | ConvertTo-Json -Compress
        }
        
      } catch {
        @{
          Error = $_.Exception.Message
          ErrorType = $_.Exception.GetType().Name
        } | ConvertTo-Json -Compress
      }
    `;
    
    try {
      const result = await this.executePowerShell(foldersScript, this.config.timeout);
      const data = JSON.parse(result);
      
      if (data.Error) {
        throw new Error(`Erreur PowerShell: ${data.Error} (${data.ErrorType})`);
      }
      
      if (Array.isArray(data)) {
        this.folders.clear();
        data.forEach(folder => {
          this.folders.set(folder.Path, folder);
        });
        this.stats.set('foldersLastUpdate', Date.now());
        this.log(`📁 ${data.length} dossiers chargés avec succès`);
        return data;
      } else {
        throw new Error('Format de données dossiers invalide');
      }
      
    } catch (error) {
      throw new Error(`Echec chargement dossiers: ${error.message}`);
    }
  }

  async loadStatsRobust() {
    const statsScript = `
      try {
        $ErrorActionPreference = "Stop"
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Statistiques générales sans accès aux données
        $stats = @{
          lastSync = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
          profileName = $namespace.CurrentProfileName
          serverVersion = $outlook.Version
          storeCount = $namespace.Stores.Count
          connected = $true
        }
        
        # Libération des objets
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        
        $stats | ConvertTo-Json -Compress
        
      } catch {
        @{
          Error = $_.Exception.Message
          ErrorType = $_.Exception.GetType().Name
        } | ConvertTo-Json -Compress
      }
    `;
    
    try {
      const result = await this.executePowerShell(statsScript, this.config.timeout);
      const data = JSON.parse(result);
      
      if (data.Error) {
        throw new Error(`Erreur PowerShell: ${data.Error} (${data.ErrorType})`);
      }
      
      this.stats.set('current', data);
      this.stats.set('lastUpdate', Date.now());
      return data;
      
    } catch (error) {
      throw new Error(`Echec chargement stats: ${error.message}`);
    }
  }

  async checkOutlookProcess() {
    try {
      const result = await this.executeCommand('tasklist', [
        '/fi', 'imagename eq OUTLOOK.EXE',
        '/fo', 'csv'
      ]);
      
      return result.includes('OUTLOOK.EXE');
    } catch (error) {
      this.log(`⚠️ Erreur verification processus: ${error.message}`);
      return false;
    }
  }

  async startOutlook() {
    this.log('🚀 Tentative de démarrage d\'Outlook...');
    
    try {
      // Vérifier d'abord si Outlook n'est pas déjà en cours de démarrage
      if (await this.checkOutlookProcess()) {
        this.log('✅ Outlook est déjà en cours d\'exécution');
        return true;
      }
      
      // Utiliser PowerShell pour démarrer Outlook de manière plus robuste
      const startScript = `
        try {
          $ErrorActionPreference = "Stop"
          
          # Chemins possibles pour Outlook - méthode simplifiée
          $outlookPaths = @(
            "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
            "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE",
            "C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE",
            "C:\\Program Files\\Microsoft Office\\Office15\\OUTLOOK.EXE",
            "C:\\Program Files (x86)\\Microsoft Office\\Office15\\OUTLOOK.EXE"
          )
          
          $outlookPath = $null
          foreach ($path in $outlookPaths) {
            if (Test-Path $path) {
              $outlookPath = $path
              Write-Output "FOUND_PATH:$path"
              break
            }
          }
          
          if (-not $outlookPath) {
            # Essayer de trouver Outlook via le registre
            try {
              $regPath = Get-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE" -ErrorAction SilentlyContinue
              if ($regPath -and $regPath.'(default)') {
                $outlookPath = $regPath.'(default)'
                Write-Output "FOUND_REGISTRY:$outlookPath"
              }
            } catch {
              # Ignorer les erreurs de registre
            }
          }
          
          if ($outlookPath -and (Test-Path $outlookPath)) {
            # Démarrer Outlook normalement (pas minimisé)
            Start-Process -FilePath $outlookPath
            Write-Output "STARTED:$outlookPath"
          } else {
            # Dernière tentative : essayer de démarrer Outlook via le PATH
            try {
              Start-Process -FilePath "outlook.exe" -ErrorAction Stop
              Write-Output "STARTED:outlook.exe"
            } catch {
              Write-Output "ERROR:Impossible de trouver ou démarrer Outlook"
            }
          }
        } catch {
          Write-Output "ERROR:$($_.Exception.Message)"
        }
      `;
      
      this.log('📧 Recherche et démarrage d\'Outlook...');
      const result = await this.executePowerShell(startScript, 10000);
      
      if (result.includes('ERROR:')) {
        const error = result.split('ERROR:')[1].trim();
        throw new Error(`Échec du démarrage d'Outlook: ${error}`);
      }
      
      if (result.includes('STARTED:')) {
        const path = result.split('STARTED:')[1].trim();
        this.log(`✅ Outlook démarré depuis: ${path}`);
        
        // Attendre que le processus soit vraiment démarré
        this.log('⏳ Attente de l\'initialisation d\'Outlook...');
        await this.sleep(3000); // Attente initiale
        
        return true;
      }
      
      if (result.includes('FOUND_PATH:') || result.includes('FOUND_REGISTRY:')) {
        this.log('✅ Outlook trouvé et démarrage en cours...');
        await this.sleep(3000);
        return true;
      }
      
      throw new Error('Impossible de démarrer Outlook - aucun chemin trouvé');
      
    } catch (error) {
      this.log(`❌ Échec démarrage Outlook: ${error.message}`);
      throw error;
    }
  }

  async waitForProcess(processName, timeout = 60000) { // Augmenté à 60 secondes
    const startTime = Date.now();
    let checkCount = 0;
    
    this.log(`⏳ Attente du processus ${processName} (timeout: ${timeout/1000}s)...`);
    
    while (Date.now() - startTime < timeout) {
      checkCount++;
      
      if (await this.checkOutlookProcess()) {
        this.log(`✅ Processus ${processName} détecté après ${checkCount} vérifications`);
        
        // Attendre encore plus longtemps pour l'initialisation complète d'Outlook
        this.log('⏳ Attente de l\'initialisation complète d\'Outlook...');
        await this.sleep(8000); // Augmenté à 8 secondes
        
        // Vérifier encore une fois que le processus est toujours là
        if (await this.checkOutlookProcess()) {
          this.log('✅ Outlook initialisé et prêt');
          return true;
        } else {
          this.log('⚠️ Outlook s\'est fermé pendant l\'initialisation, nouvelle tentative...');
          continue;
        }
      }
      
      // Afficher un message de progression
      if (checkCount % 10 === 0) {
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
        this.log(`⏳ Toujours en attente du processus ${processName} (${elapsed}s écoulées)...`);
      }
      
      await this.sleep(1000);
    }
    
    throw new Error(`Timeout: processus ${processName} non détecté après ${timeout/1000}s`);
  }

  startHealthMonitoring() {
    if (this.checkInterval) {
      clearInterval(this.checkInterval);
    }
    
    this.checkInterval = setInterval(async () => {
      try {
        await this.healthCheck();
      } catch (error) {
        this.log(`⚠️ Erreur health check: ${error.message}`);
      }
    }, this.config.healthCheckInterval);
    
    this.log('💓 Monitoring de santé démarré');
  }

  async healthCheck() {
    try {
      const processRunning = await this.checkOutlookProcess();
      
      if (!processRunning && this.isOutlookConnected) {
        this.log('⚠️ Processus Outlook fermé, mise à jour du statut');
        this.handleDisconnection();
        return;
      }
      
      if (processRunning && !this.isOutlookConnected && this.config.autoReconnect) {
        this.log('🔄 Processus détecté, tentative de reconnexion...');
        await this.establishConnection();
        return;
      }
      
      // Test de connexion COM léger
      if (this.isOutlookConnected) {
        const testResult = await this.quickCOMTest();
        if (!testResult) {
          this.log('⚠️ Test COM rapide échoué, reconnexion nécessaire');
          await this.establishConnection();
        }
      }
      
    } catch (error) {
      this.log(`❌ Erreur health check: ${error.message}`);
    }
  }

  async quickCOMTest() {
    const quickTest = `
      try {
        $outlook = New-Object -ComObject Outlook.Application
        $version = $outlook.Version
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        Write-Output "OK:$version"
      } catch {
        Write-Output "ERROR:$($_.Exception.Message)"
      }
    `;
    
    try {
      const result = await this.executePowerShell(quickTest, 5000);
      return result.startsWith('OK:');
    } catch (error) {
      return false;
    }
  }

  handleConnectionError(error) {
    this.connectionAttempts++;
    this.lastError = error;
    this.connectionState = 'failed';
    this.isOutlookConnected = false;
    
    this.log(`❌ Erreur connexion (tentative ${this.connectionAttempts}): ${error.message}`);
    
    if (this.connectionAttempts < this.maxRetries && this.config.autoReconnect) {
      const delay = this.retryDelay * this.connectionAttempts;
      this.log(`⏳ Nouvelle tentative dans ${delay}ms...`);
      
      setTimeout(() => {
        this.establishConnection();
      }, delay);
    } else {
      this.log(`💥 Arret des tentatives apres ${this.connectionAttempts} echecs`);
    }
  }

  handleDisconnection() {
    this.isOutlookConnected = false;
    this.connectionState = 'disconnected';
    this.folders.clear();
    this.stats.clear();
    this.log('🔌 Déconnexion Outlook détectée');
  }

  // === MÉTHODES PUBLIQUES ===

  async isConnected() {
    if (!this.isOutlookConnected) {
      return false;
    }
    
    // Vérification du cache
    const lastCheck = this.stats.get('lastConnectionCheck');
    if (lastCheck && Date.now() - lastCheck < 10000) {
      return this.isOutlookConnected;
    }
    
    // Test rapide
    const result = await this.quickCOMTest();
    this.stats.set('lastConnectionCheck', Date.now());
    
    if (!result && this.isOutlookConnected) {
      this.handleDisconnection();
    }
    
    return result;
  }

  async getStats() {
    if (!this.isOutlookConnected) {
      return null;
    }
    
    // Vérifier le cache
    const lastUpdate = this.stats.get('lastUpdate');
    if (lastUpdate && Date.now() - lastUpdate < this.config.cacheTTL) {
      return this.stats.get('current');
    }
    
    // Recharger les stats
    try {
      return await this.loadStatsRobust();
    } catch (error) {
      this.log(`⚠️ Erreur getStats: ${error.message}`);
      return this.stats.get('current') || null;
    }
  }

  getFolders() {
    return Array.from(this.folders.values());
  }

  // Récupérer toutes les boîtes mail connectées
  async getMailboxes() {
    const mailboxScript = `
      try {
        $ErrorActionPreference = "Stop"
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        $mailboxes = @()
        
        # Parcourir tous les stores (boites mail) avec plus de details
        foreach ($store in $namespace.Stores) {
          try {
            # Verifier que le store est accessible et valide
            $storeDisplayName = if ($store.DisplayName) { $store.DisplayName } else { "Store Sans Nom" }
            $storeType = "Inconnu"
            
            # Determiner le type de store plus precisement
            if ($store.ExchangeStoreType -ne $null) {
              switch ($store.ExchangeStoreType) {
                0 { $storeType = "Exchange Primaire" }
                1 { $storeType = "Exchange Public" }
                2 { $storeType = "Exchange Delegue" }
                default { $storeType = "Exchange (" + $store.ExchangeStoreType + ")" }
              }
            } elseif ($store.FilePath -ne $null -and $store.FilePath -ne "") {
              $storeType = "Fichier PST"
            } elseif ($store.StoreID -ne $null) {
              $storeType = "Store Systeme"
            }
            
            # Verifier si c'est un store valide (pas vide)
            $isValidStore = $true
            try {
              $rootFolder = $store.GetRootFolder()
              if ($rootFolder -eq $null) {
                $isValidStore = $false
              } else {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
              }
            } catch {
              $isValidStore = $false
            }
            
            if ($isValidStore) {
              $mailbox = @{
                Name = $storeDisplayName
                Type = $storeType
                Size = if ($store.Size -gt 0) { $store.Size } else { 0 }
                IsDefault = ($store.StoreID -eq $namespace.DefaultStore.StoreID)
                StoreID = $store.StoreID
                FilePath = if ($store.FilePath) { $store.FilePath } else { "" }
                ExchangeStoreType = if ($store.ExchangeStoreType -ne $null) { $store.ExchangeStoreType } else { -1 }
              }
              
              $mailboxes += $mailbox
            }
          } catch {
            # Ignorer les stores inaccessibles mais continuer
          }
        }
        
        # Liberation des objets
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        
        # S'assurer qu'on a au moins une boite mail
        if ($mailboxes.Count -eq 0) {
          throw "Aucune boite mail valide trouvee"
        }
        
        $mailboxes | ConvertTo-Json -Compress -Depth 3
        
      } catch {
        @{
          Error = $_.Exception.Message
          ErrorType = $_.Exception.GetType().Name
          Line = if ($_.InvocationInfo.ScriptLineNumber) { $_.InvocationInfo.ScriptLineNumber } else { 0 }
        } | ConvertTo-Json -Compress
      }
    `;
    
    try {
      const result = await this.executePowerShell(mailboxScript, this.config.timeout);
      const data = JSON.parse(result);
      
      if (data.Error) {
        throw new Error(`Erreur PowerShell: ${data.Error} (${data.ErrorType})`);
      }
      
      return Array.isArray(data) ? data : [data];
      
    } catch (error) {
      this.log(`⚠️ Erreur recuperation boites mail: ${error.message}`);
      return [];
    }
  }

  // Recuperer l'architecture complete des dossiers d'une boite mail
  async getFolderStructure(storeId = null) {
    this.log(`📁 [STRUCTURE] Debut getFolderStructure pour storeId: ${storeId || 'default'}`);
    
    const structureScript = `
      param([string]$TargetStoreId)
      
      try {
        $ErrorActionPreference = "Stop"
        
        # Configuration explicite de l'encodage pour éviter les problèmes de caractères
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
        
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Selectionner le store approprie
        $targetStore = $null
        if ($TargetStoreId -and $TargetStoreId -ne "" -and $TargetStoreId -ne "null") {
          foreach ($store in $namespace.Stores) {
            if ($store.StoreID -eq $TargetStoreId) {
              $targetStore = $store
              break
            }
          }
        } else {
          $targetStore = $namespace.DefaultStore
        }
        
        if (-not $targetStore) {
          throw "Store avec ID '$TargetStoreId' non trouve. Stores disponibles: $($namespace.Stores | ForEach-Object { $_.DisplayName + ' (' + $_.StoreID + ')' })"
        }
        
        # Fonction recursive pour explorer les dossiers
        function Get-FolderTree($folder, $level = 0) {
          try {
            # Utiliser ToString() explicite pour garantir l'encodage correct
            $folderName = $folder.Name.ToString()
            $folderPath = $folder.FolderPath.ToString()
            
            # Conversion explicite en UTF-8 pour les caractères spéciaux
            $folderName = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($folderName))
            $folderPath = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::UTF8.GetBytes($folderPath))
            
            $folderInfo = @{
              Name = $folderName
              Path = $folderPath
              Level = $level
              ItemCount = $folder.Items.Count
              UnreadCount = $folder.UnReadItemCount
              HasSubfolders = ($folder.Folders.Count -gt 0)
              DefaultItemType = $folder.DefaultItemType
              Subfolders = @()
            }
            
            # Explorer les sous-dossiers
            if ($folder.Folders.Count -gt 0) {
              $subfoldersList = @()
              foreach ($subfolder in $folder.Folders) {
                try {
                  $subfoldersList += Get-FolderTree $subfolder ($level + 1)
                  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($subfolder) | Out-Null
                } catch {
                  # Ignorer les dossiers inaccessibles
                }
              }
              # S'assurer que c'est toujours un tableau même avec un seul élément
              $folderInfo.Subfolders = @($subfoldersList)
            }
            
            return $folderInfo
          } catch {
            throw $_
          }
        }
        
        # Commencer par le dossier racine
        $rootFolder = $targetStore.GetRootFolder()
        $structure = Get-FolderTree $rootFolder
        
        # Liberation des objets
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        
        # Conversion JSON avec encodage UTF-8 explicite
        $jsonOutput = $structure | ConvertTo-Json -Depth 10 -Compress
        Write-Output $jsonOutput
        
      } catch {
        $errorInfo = @{
          Error = $_.Exception.Message
          ErrorType = $_.Exception.GetType().Name
          StackTrace = $_.ScriptStackTrace
          StoreId = $TargetStoreId
        }
        $errorJson = $errorInfo | ConvertTo-Json -Compress
        Write-Output $errorJson
      }
    `;
    
    try {
      this.log(`📁 [STRUCTURE] Execution du script PowerShell avec storeId: ${storeId || 'default'}...`);
      
      // Passer le storeId comme paramètre pour éviter l'interpolation et les problèmes d'encodage
      const scriptWithParams = `${structureScript}; ${structureScript.split('param')[0]} & { ${structureScript} } -TargetStoreId "${storeId || ''}"`;
      
      // Version simplifiée : écrire le script dans un fichier temporaire avec l'encodage correct
      const tempScript = `
        param([string]$TargetStoreId = "${storeId || ''}")
        ${structureScript.replace('param([string]$TargetStoreId)', '')}
      `;
      
      const result = await this.executePowerShell(tempScript, this.config.timeout * 2);
      
      this.log(`📁 [STRUCTURE] Resultat brut: ${result.substring(0, 200)}...`);
      
      const data = JSON.parse(result);
      
      if (data.Error) {
        this.log(`❌ [STRUCTURE] Erreur PowerShell: ${data.Error} (${data.ErrorType})`);
        if (data.StackTrace) {
          this.log(`📍 [STRUCTURE] Stack trace: ${data.StackTrace}`);
        }
        throw new Error(`Erreur PowerShell: ${data.Error} (${data.ErrorType})`);
      }
      
      this.log(`✅ [STRUCTURE] Structure recuperee avec succes pour ${storeId || 'store par defaut'}`);
      return data;
      
    } catch (error) {
      this.log(`⚠️ [STRUCTURE] Erreur structure dossiers: ${error.message}`);
      this.log(`⚠️ [STRUCTURE] Type: ${error.constructor.name}`);
      throw error;
    }
  }

  getConnectionInfo() {
    return {
      isConnected: this.isOutlookConnected,
      state: this.connectionState,
      attempts: this.connectionAttempts,
      lastError: this.lastError?.message,
      outlookVersion: this.outlookVersion,
      lastCheck: this.lastCheck,
      foldersCount: this.folders.size,
      hasStats: this.stats.has('current')
    };
  }

  // === MÉTHODES UTILITAIRES ===

  async executePowerShell(script, timeout = 10000, showProgress = false) {
    return new Promise((resolve, reject) => {
      const child = spawn('powershell', [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-OutputFormat', 'Text',
        '-InputFormat', 'Text',
        '-Command', `chcp 65001 > $null; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; ${script}`
      ], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
        encoding: 'utf8'
      });
      
      let stdout = '';
      let stderr = '';
      
      // Déduplication des messages PowerShell (PowerShell génère parfois des doublons)
      const seenProgressMessages = new Set();
      let lastProgressValue = -1;
      
      child.stdout.setEncoding('utf8');
      child.stderr.setEncoding('utf8');
      
      child.stdout.on('data', (data) => {
        stdout += data.toString('utf8');
      });
      
      child.stderr.on('data', (data) => {
        const stderrData = data.toString('utf8');
        stderr += stderrData;
        
        // Si showProgress est activé, afficher les logs de progression en temps réel
        if (showProgress && stderrData.trim()) {
          const lines = stderrData.trim().split('\n');
          lines.forEach(line => {
            line = line.trim();
            if (line) {
              // Ignorer les lignes de commande PowerShell et les messages d'erreur de format
              if (line.startsWith('Write-Error') || 
                  line.includes('+ CategoryInfo') || 
                  line.includes('+ FullyQualifiedErrorId') ||
                  line.startsWith('try {') ||
                  line.startsWith('$') ||
                  line.includes('exit 1') ||
                  line.includes('catch {')) {
                return;
              }
              
              // Extraire le pourcentage si présent pour déduplication
              const progressMatch = line.match(/\[(\d+(?:\.\d+)?)%\]/);
              let shouldEmit = true;
              
              if (progressMatch) {
                const currentProgress = parseFloat(progressMatch[1]);
                
                // Déduplication : ne pas réémettre le même pourcentage
                if (currentProgress <= lastProgressValue) {
                  shouldEmit = false;
                } else {
                  lastProgressValue = currentProgress;
                }
              }
              
              // Déduplication globale des messages
              const messageKey = line.replace(/\s+/g, ' ').trim();
              if (seenProgressMessages.has(messageKey)) {
                shouldEmit = false;
              } else {
                seenProgressMessages.add(messageKey);
              }
              
              if (shouldEmit) {
                console.log(`[PowerShell Progress] ${line}`);
                
                // Extraire et émettre les informations de progression pour l'UI
                if (line.includes('[PROGRESS]') || line.includes('[GLOBAL]') || line.includes('[SEMAINE]') || line.includes('[STATS]') || line.includes('[TERMINE]')) {
                  this.emit('powershell-progress', {
                    message: line,
                    timestamp: new Date().toISOString()
                  });
                }
              }
            }
          });
        }
      });
      
      const timeoutId = setTimeout(() => {
        child.kill('SIGKILL');
        reject(new Error(`Timeout PowerShell après ${timeout}ms`));
      }, timeout);
      
      child.on('close', (code) => {
        clearTimeout(timeoutId);
        
        if (code !== 0) {
          reject(new Error(`PowerShell exit code ${code}: ${stderr}`));
        } else {
          resolve(stdout.trim());
        }
      });
      
      child.on('error', (error) => {
        clearTimeout(timeoutId);
        reject(error);
      });
    });
  }

  async executeCommand(command, args = [], options = {}) {
    return new Promise((resolve, reject) => {
      const child = spawn(command, args, {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
        ...options
      });
      
      let stdout = '';
      let stderr = '';
      
      child.stdout.on('data', (data) => {
        stdout += data.toString();
      });
      
      child.stderr.on('data', (data) => {
        stderr += data.toString();
      });
      
      const timeout = options.timeout || 10000;
      const timeoutId = setTimeout(() => {
        child.kill('SIGKILL');
        reject(new Error(`Timeout commande après ${timeout}ms`));
      }, timeout);
      
      child.on('close', (code) => {
        clearTimeout(timeoutId);
        
        if (code !== 0 && !options.ignoreErrors) {
          reject(new Error(`Commande exit code ${code}: ${stderr}`));
        } else {
          resolve(stdout.trim());
        }
      });
      
      child.on('error', (error) => {
        clearTimeout(timeoutId);
        reject(error);
      });
    });
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  log(message) {
    if (this.config.enableDetailedLogs) {
      const timestamp = new Date().toISOString();
      // Remplacer les émojis et caractères accentués pour éviter les problèmes d'encodage Windows
      const cleanMessage = message
        .replace(/🚀/g, '[INIT]')
        .replace(/✅/g, '[OK]')
        .replace(/🔍/g, '[CHECK]')
        .replace(/📋/g, '[INFO]')
        .replace(/❌/g, '[ERROR]')
        .replace(/📧/g, '[OUTLOOK]')
        .replace(/⚠️/g, '[WARN]')
        .replace(/🔗/g, '[CONNECT]')
        .replace(/🔄/g, '[RETRY]')
        .replace(/📊/g, '[STATS]')
        .replace(/📚/g, '[LOAD]')
        .replace(/📁/g, '[FOLDER]')
        .replace(/⏳/g, '[WAIT]')
        .replace(/🩹/g, '[HEALTH]')
        .replace(/💓/g, '[HEALTH]')
        // Caractères accentués français
        .replace(/é/g, 'e')
        .replace(/è/g, 'e')
        .replace(/ê/g, 'e')
        .replace(/ë/g, 'e')
        .replace(/à/g, 'a')
        .replace(/â/g, 'a')
        .replace(/ä/g, 'a')
        .replace(/ç/g, 'c')
        .replace(/ù/g, 'u')
        .replace(/û/g, 'u')
        .replace(/ü/g, 'u')
        .replace(/ô/g, 'o')
        .replace(/ö/g, 'o')
        .replace(/î/g, 'i')
        .replace(/ï/g, 'i')
        .replace(/É/g, 'E')
        .replace(/È/g, 'E')
        .replace(/Ê/g, 'E')
        .replace(/Ë/g, 'E')
        .replace(/À/g, 'A')
        .replace(/Â/g, 'A')
        .replace(/Ä/g, 'A')
        .replace(/Ç/g, 'C')
        .replace(/Ù/g, 'U')
        .replace(/Û/g, 'U')
        .replace(/Ü/g, 'U')
        .replace(/Ô/g, 'O')
        .replace(/Ö/g, 'O')
        .replace(/Î/g, 'I')
        .replace(/Ï/g, 'I');
      
      console.log(`[${timestamp}] OutlookConnector: ${cleanMessage}`);
    }
  }

  // === RÉCUPÉRATION D'EMAILS ===

  /**
   * Version simplifiée - Pas de pagination nécessaire pour <500 emails
   */
  async getFolderEmailsWithPagination(folderPath, maxEmails = 500) {
    this.log(`� [SIMPLE] Récupération simple du dossier: ${folderPath} (max: ${maxEmails})`);
    
    // Utiliser directement getFolderEmails sans pagination
    return await this.getFolderEmails(folderPath, maxEmails);
  }

  /**
   * Récupère les emails d'un dossier spécifique (limité par COM à ~500)
   */
  async getFolderEmails(folderPath, limit = 5000, lastCheckDate = null, restrictFilter = null) {
    // Validation du chemin de dossier
    if (!folderPath || folderPath === 'folderCategories' || folderPath === 'undefined') {
      this.log(`❌ Chemin de dossier invalide: ${folderPath}`, 'ERROR');
      throw new Error(`Chemin de dossier invalide: ${folderPath}`);
    }

    // Protection contre les exécutions multiples - SYNCHRONE ET IMMÉDIATE
    if (this.isPowerShellRunning) {
      this.log(`⚠️ Analyse PowerShell déjà en cours - Requête REJETÉE pour: ${folderPath}`, 'WARNING');
      throw new Error('Une analyse PowerShell est déjà en cours. Veuillez attendre qu\'elle se termine.');
    }

    // DIAGNOSTIC: Qui appelle cette méthode ?
    const stack = new Error().stack;
    console.log(`🔍 DIAGNOSTIC getFolderEmails appelée pour: ${folderPath}`);
    console.log(`🔍 STACK TRACE:`);
    console.log(stack.split('\n').slice(1, 6).join('\n'));

    // Marquer IMMÉDIATEMENT comme en cours
    this.isPowerShellRunning = true;
    this.log(`🔒 Protection PowerShell activée pour: ${folderPath}`, 'INFO');
    
    try {
      this.log(`📧 Récupération des emails du dossier: ${folderPath}`);

      const emailsScript = `
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $OutputEncoding = [System.Text.Encoding]::UTF8
        $restrictFilter = ${restrictFilter ? '`' + restrictFilter + '`' : '$null'}
        try {
          $ErrorActionPreference = "Stop"
          $ProgressPreference = "SilentlyContinue"
          
          $outlook = New-Object -ComObject Outlook.Application
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Rechercher le dossier par son chemin
          $folder = $null
          
          # Cas spécial pour Inbox
          if ("${folderPath}" -eq "Inbox") {
            $folder = $namespace.GetDefaultFolder(6)  # olFolderInbox = 6
            Write-Error "[PROGRESS] [1%] Dossier Inbox trouvé directement" -ErrorAction Continue
          } else {
            $accounts = $namespace.Accounts
            
            foreach ($account in $accounts) {
              $rootFolder = $namespace.Folders($account.DisplayName)
              if ($rootFolder) {
                # Parcourir les dossiers pour trouver celui qui correspond
                $folderParts = "${folderPath}".Split("\\")
                $currentFolder = $rootFolder
                
                foreach ($part in $folderParts) {
                  if ($part -and $part -ne "" -and $part -ne $account.DisplayName) {
                    $found = $false
                    foreach ($subfolder in $currentFolder.Folders) {
                      if ($subfolder.Name -eq $part) {
                        $currentFolder = $subfolder
                        $found = $true
                        break
                      }
                    }
                    if (-not $found) {
                      break
                    }
                  }
                }
                
                if ($currentFolder -and $currentFolder.Name -eq ($folderParts[-1])) {
                  $folder = $currentFolder
                  break
                }
              }
            }
          }
          
          if (-not $folder) {
            Write-Output "FOLDER_NOT_FOUND"
            exit 1
          }
          
          # MÉTHODE ULTRA-RAPIDE: Récupération directe sans restrictions
          Write-Error "[PROGRESS] [0%] Initialisation de la récupération rapide..." -ErrorAction Continue
          
          $folderItems = $folder.Items
          if ($restrictFilter) {
            try {
              $folderItems = $folderItems.Restrict($restrictFilter)
              $totalItems = $folderItems.Count
              Write-Error "[PROGRESS] [3%] Restrict appliqué: $restrictFilter => $totalItems emails..." -ErrorAction Continue
            } catch {
              Write-Error "[PROGRESS] [3%] Erreur Restrict: $($_.Exception.Message)" -ErrorAction Continue
            }
          }
          
          # IMPORTANT: COM Outlook peut limiter l'accès aux items à 500 par défaut
          # Utilisation de la méthode Restrict pour contourner cette limitation
          $totalItems = $folderItems.Count
          Write-Error "[PROGRESS] [2%] Items détectés via COM: $totalItems..." -ErrorAction Continue
          
          # Si on a exactement 500 items, c'est probablement une limitation COM
          # Dans ce cas, utiliser une approche alternative avec restriction
          if ($totalItems -eq 500 -and ${limit} -gt 500) {
            Write-Error "[PROGRESS] [3%] Limitation COM détectée (500), utilisation de méthode alternative..." -ErrorAction Continue
            # Essayer d'obtenir tous les emails via restriction (sans limite de date)
            try {
              $allItems = $folderItems.Restrict("[MessageClass] = 'IPM.Note'")
              if ($allItems.Count -gt $totalItems) {
                $folderItems = $allItems
                $totalItems = $folderItems.Count
                Write-Error "[PROGRESS] [4%] Restriction appliquée: $totalItems emails accessibles..." -ErrorAction Continue
              } else {
                Write-Error "[PROGRESS] [4%] Restriction n'a pas permis d'accéder à plus d'emails. Limitation COM confirmée..." -ErrorAction Continue
                # La limitation COM est confirmée, on récupère ce qu'on peut
              }
            } catch {
              Write-Error "[PROGRESS] [4%] Erreur lors de la restriction: $($_.Exception.Message)" -ErrorAction Continue
            }
          }
          
          $limit = [Math]::Min($totalItems, ${limit})
          
          Write-Error "[PROGRESS] [5%] Tri par date (plus récents en premier)..." -ErrorAction Continue
          # Optimisation : si on ne prend que les premiers emails, pas besoin de trier TOUS
          if ($limit -lt $totalItems) {
            # Récupération directe des X premiers (déjà triés par défaut par date)
            $folderItems.Sort("[ReceivedTime]", $true)
          } else {
            $folderItems.Sort("[ReceivedTime]", $true)
          }
          
          Write-Error "[PROGRESS] [10%] Début de la récupération de $limit emails sur $totalItems..." -ErrorAction Continue
          
          # Mode EXPRESS pour gros volumes (> 2000 emails)
          $isExpressMode = $totalItems -gt 2000
          if ($isExpressMode) {
            Write-Error "[PROGRESS] [12%] Mode EXPRESS activé (volume important détecté)..." -ErrorAction Continue
          }
          
          $emails = @()
          $lastCheck = $null
          
          if ("${lastCheckDate}" -and "${lastCheckDate}" -ne "null") {
            $lastCheck = [DateTime]::Parse("${lastCheckDate}")
          }
          
          # Récupération par chunks avec progression linéaire OPTIMISÉE
          $chunkSize = if ($isExpressMode) { 50 } else { 75 }  # Chunks plus petits en mode express
          $totalChunks = [Math]::Ceiling($limit / $chunkSize)
          $processedCount = 0
          $currentProgress = 15  # Commencer après les étapes d'initialisation
          
          for ($chunk = 0; $chunk -lt $totalChunks; $chunk++) {
            $startIndex = $chunk * $chunkSize + 1
            $endIndex = [Math]::Min(($chunk + 1) * $chunkSize, $limit)
            
            # Progression linéaire de 15% à 90% pour le traitement des chunks
            $chunkProgress = [Math]::Round(15 + (($chunk / $totalChunks) * 75), 1)
            
            Write-Error "[PROGRESS] [$chunkProgress%] Traitement du chunk $($chunk + 1)/$totalChunks (emails $startIndex à $endIndex)..." -ErrorAction Continue
            
            for ($i = $startIndex; $i -le $endIndex -and $i -le $totalItems; $i++) {
              try {
                $item = $folderItems.Item($i)
                
                # Filtrer par date si spécifié
                if ($lastCheck -and $item.ReceivedTime -le $lastCheck) {
                  continue
                }
                
                # Mode EXPRESS : collecte de données optimisée
                if ($isExpressMode) {
                  $emailData = @{
                    EntryID = $item.EntryID
                    Subject = if($item.Subject -and $item.Subject.Length -gt 0) { 
                      $item.Subject.Substring(0, [Math]::Min(100, $item.Subject.Length)) 
                    } else { "(Sans objet)" }
                    SenderName = if($item.SenderName) { $item.SenderName } else { "Inconnu" }
                    ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss")
                    UnRead = $item.UnRead
                    HasAttachments = $item.Attachments.Count -gt 0
                    FolderName = $folder.Name
                  }
                } else {
                  # Mode STANDARD : collecte complète
                  $emailData = @{
                    EntryID = $item.EntryID
                    Subject = if($item.Subject) { $item.Subject } else { "(Sans objet)" }
                    SenderName = if($item.SenderName) { $item.SenderName } else { "Inconnu" }
                    SenderEmailAddress = if($item.SenderEmailAddress) { $item.SenderEmailAddress } else { "" }
                    ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss")
                    Size = $item.Size
                    UnRead = $item.UnRead
                    Importance = $item.Importance
                    Categories = if($item.Categories) { $item.Categories } else { "" }
                    FolderName = $folder.Name
                    HasAttachments = $item.Attachments.Count -gt 0
                    AttachmentCount = $item.Attachments.Count
                  }
                }
                
                $emails += $emailData
                $processedCount++
                
                # Feedback de progression plus régulier et linéaire (optimisé)
                if ($processedCount % 25 -eq 0) {  # Feedback toutes les 25 au lieu de 50
                  # Progression proportionnelle dans la plage du chunk actuel
                  $chunkStartProgress = 15 + (($chunk / $totalChunks) * 75)
                  $chunkEndProgress = 15 + ((($chunk + 1) / $totalChunks) * 75)
                  $itemsInChunk = $endIndex - $startIndex + 1
                  $itemsProcessedInChunk = ($processedCount % $chunkSize)
                  if ($itemsProcessedInChunk -eq 0 -and $processedCount -gt 0) {
                    $itemsProcessedInChunk = $chunkSize
                  }
                  
                  $itemProgress = [Math]::Round($chunkStartProgress + (($itemsProcessedInChunk / $itemsInChunk) * ($chunkEndProgress - $chunkStartProgress)), 1)
                  Write-Error "[PROGRESS] [$itemProgress%] $processedCount emails traités..." -ErrorAction Continue
                }
                
                # Pause adaptative pour éviter la surcharge
                if ($processedCount % 100 -eq 0) {
                  if ($isExpressMode) {
                    Start-Sleep -Milliseconds 5  # Pause minimale en mode express
                  } else {
                    Start-Sleep -Milliseconds 10  # Pause standard
                  }
                }
                
              } catch {
                # Ignorer les emails problématiques et continuer
                continue
              }
            }
          }
          
          Write-Error "[PROGRESS] [95%] Finalisation des données..." -ErrorAction Continue
          
          $result = @{
            FolderPath = "${folderPath}"
            FolderName = $folder.Name
            TotalItems = $totalItems
            UnreadItems = $folder.UnReadItemCount
            EmailsRetrieved = $emails.Count
            Emails = $emails
            Debug = @{
              Method = "Direct-Fast"
              LimitRequested = ${limit}
              TotalInFolder = $totalItems
              ProcessedCount = $processedCount
              ChunksUsed = $totalChunks
              ChunkSize = $chunkSize
            }
          }
          
          Write-Error "[PROGRESS] [100%] Récupération terminée: $($emails.Count) emails récupérés" -ErrorAction Continue
          Write-Error "[TERMINE] RAPIDE: $($emails.Count) emails récupérés en $totalChunks chunks de $chunkSize" -ErrorAction Continue
          
          # Libération des objets COM
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folderItems) | Out-Null
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folder) | Out-Null
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
          
          $json = $result | ConvertTo-Json -Depth 5 -Compress
          Write-Output $json
        }
        catch {
          Write-Output "ERROR: $($_.Exception.Message)"
          exit 1
        }
      `;

      const result = await this.executePowerShell(emailsScript, this.config.timeout * 5, true);
      
      if (result.includes('FOLDER_NOT_FOUND')) {
        throw new Error(`Dossier non trouvé: ${folderPath}`);
      }
      
      if (result.startsWith('ERROR:')) {
        throw new Error(result.replace('ERROR: ', ''));
      }

      const emailData = JSON.parse(result);
      
      // Afficher les informations de débogage si disponibles
      if (emailData.Debug) {
        this.log(`🐛 DEBUG: TotalItems=${emailData.Debug.TotalItemsInCollection}, Limit=${emailData.Debug.LimitRequested}, Processed=${emailData.Debug.ActualEmailsProcessed}`);
      }
      
      this.log(`📧 ${emailData.EmailsRetrieved} emails récupérés du dossier ${emailData.FolderName}`);
      
      return emailData;

    } catch (error) {
      this.log(`❌ Erreur récupération emails: ${error.message}`);
      throw error;
    } finally {
      // Libération du verrou PowerShell
      this.isPowerShellRunning = false;
      this.log(`🔓 Protection PowerShell désactivée`, 'INFO');
    }
  }

  /**
   * Méthode optimisée pour le monitoring - récupère seulement les métadonnées critiques
   */
  async getFolderEmailsForMonitoring(folderPath, limit = 2000) {
    // Protection contre les exécutions multiples - SYNCHRONE ET IMMÉDIATE
    if (this.isPowerShellRunning) {
      this.log(`⚠️ Analyse PowerShell déjà en cours - Requête monitoring REJETÉE pour: ${folderPath}`, 'WARNING');
      throw new Error('Une analyse PowerShell est déjà en cours. Veuillez attendre qu\'elle se termine.');
    }

    // Marquer IMMÉDIATEMENT comme en cours
    this.isPowerShellRunning = true;
    this.log(`🔒 Protection PowerShell activée pour monitoring: ${folderPath}`, 'INFO');
    
    try {
      this.log(`⚡ Monitoring rapide du dossier: ${folderPath}`);

      const monitoringScript = `
        try {
          $ErrorActionPreference = "Stop"
          $ProgressPreference = "SilentlyContinue"
          
          $outlook = New-Object -ComObject Outlook.Application
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Rechercher le dossier par son chemin
          $folder = $null
          $accounts = $namespace.Accounts
          
          foreach ($account in $accounts) {
            $rootFolder = $namespace.Folders($account.DisplayName)
            if ($rootFolder) {
              # Parcourir les dossiers pour trouver celui qui correspond
              $folderParts = "${folderPath}".Split("\\")
              $currentFolder = $rootFolder
              
              foreach ($part in $folderParts) {
                if ($part -and $part -ne "" -and $part -ne $account.DisplayName) {
                  $found = $false
                  foreach ($subfolder in $currentFolder.Folders) {
                    if ($subfolder.Name -eq $part) {
                      $currentFolder = $subfolder
                      $found = $true
                      break
                    }
                  }
                  if (-not $found) {
                    break
                  }
                }
              }
              
              if ($currentFolder -and $currentFolder.Name -eq ($folderParts[-1])) {
                $folder = $currentFolder
                break
              }
            }
          }
          
          if (-not $folder) {
            Write-Output "FOLDER_NOT_FOUND"
            exit 1
          }
          
          $emails = @()
          $items = $folder.Items
          $items.Sort("[ReceivedTime]", $true) # Tri par date décroissante
          
          $limit = ${limit}
          $count = 0
          
          # Optimisation : traitement par lots pour éviter les timeouts
          foreach ($item in $items) {
            if ($count -ge $limit) { break }
            
            try {
              # Récupération ultra-minimale pour la vitesse maximale
              $emailData = @{
                EntryID = $item.EntryID
                ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss")
                UnRead = $item.UnRead
                Subject = if($item.Subject -and $item.Subject.Length -gt 0) { 
                  $item.Subject.Substring(0, [Math]::Min(50, $item.Subject.Length)) 
                } else { 
                  "(Sans objet)" 
                }
                SenderName = if($item.SenderName) { $item.SenderName } else { "Inconnu" }
                HasAttachments = $item.Attachments.Count -gt 0
              }
              
              $emails += $emailData
              $count++
              
              # Petit break toutes les 100 itérations pour éviter les blocages
              if ($count % 100 -eq 0) {
                Start-Sleep -Milliseconds 10
              }
            }
            catch {
              # Ignorer les emails qui causent des erreurs et continuer
              continue
            }
          }
          
          $result = @{
            FolderPath = "${folderPath}"
            FolderName = $folder.Name
            TotalItems = $folder.Items.Count
            UnreadItems = $folder.UnReadItemCount
            EmailsRetrieved = $emails.Count
            Emails = $emails
            IsMonitoringMode = $true
          }
          
          $json = $result | ConvertTo-Json -Depth 3 -Compress
          Write-Output $json
        }
        catch {
          Write-Output "ERROR: $($_.Exception.Message)"
          exit 1
        }
      `;

      // Utiliser un timeout plus long pour le monitoring (30 secondes au lieu de 15)
      const result = await this.executePowerShell(monitoringScript, this.config.timeout * 2);
      
      if (result.includes('FOLDER_NOT_FOUND')) {
        throw new Error(`Dossier non trouvé: ${folderPath}`);
      }
      
      if (result.startsWith('ERROR:')) {
        throw new Error(result.replace('ERROR: ', ''));
      }

      const emailData = JSON.parse(result);
      
      this.log(`⚡ ${emailData.EmailsRetrieved} emails récupérés en mode monitoring (${emailData.FolderName})`);
      
      return emailData;

    } catch (error) {
      this.log(`❌ Erreur monitoring rapide: ${error.message}`);
      throw error;
    } finally {
      // Libération du verrou PowerShell
      this.isPowerShellRunning = false;
      this.log(`🔓 Protection PowerShell monitoring désactivée`, 'INFO');
    }
  }

  /**
   * Vérifie si un dossier existe dans Outlook
   */
  async folderExists(folderPath) {
    if (!folderPath) return false;

    try {
      const script = `
        try {
          $ErrorActionPreference = "Stop"
          $ProgressPreference = "SilentlyContinue"
          
          $outlook = New-Object -ComObject Outlook.Application
          $namespace = $outlook.GetNamespace("MAPI")
          
          # Rechercher le dossier par son chemin
          $folder = $null
          
          # Cas spécial pour Inbox
          if ("${folderPath}" -eq "Inbox") {
            $folder = $namespace.GetDefaultFolder(6)  # olFolderInbox = 6
          } else {
            $accounts = $namespace.Accounts
            
            foreach ($account in $accounts) {
              $rootFolder = $namespace.Folders($account.DisplayName)
              if ($rootFolder) {
                $folderParts = "${folderPath}".Split("\\")
                $currentFolder = $rootFolder
                
                foreach ($part in $folderParts) {
                  if ($part -and $part -ne "" -and $part -ne $account.DisplayName) {
                    $found = $false
                    foreach ($subfolder in $currentFolder.Folders) {
                      if ($subfolder.Name -eq $part) {
                        $currentFolder = $subfolder
                        $found = $true
                        break
                      }
                    }
                    if (-not $found) {
                      break
                    }
                  }
                }
                
                if ($currentFolder -and $currentFolder.Name -eq ($folderParts[-1])) {
                  $folder = $currentFolder
                  break
                }
              }
            }
          }
          
          if ($folder) {
            Write-Output "EXISTS"
          } else {
            Write-Output "NOT_FOUND"
          }
        }
        catch {
          Write-Output "ERROR: $($_.Exception.Message)"
        }
      `;

      const result = await this.executePowerShell(script, 10000);
      return result.trim() === "EXISTS";

    } catch (error) {
      this.log(`❌ Erreur vérification existence dossier ${folderPath}: ${error.message}`);
      return false;
    }
  }

  /**
   * Récupère la structure hiérarchique complète des dossiers Outlook
   */
  async getFolderStructure() {
    try {
      this.log('📁 Récupération de la structure hiérarchique des dossiers');

      const script = `
        try {
          $ErrorActionPreference = "Stop"
          $ProgressPreference = "SilentlyContinue"
          
          $outlook = New-Object -ComObject Outlook.Application
          $namespace = $outlook.GetNamespace("MAPI")
          
          function Get-FolderStructure($folder, $parentPath = "") {
            $folderPath = if ($parentPath) { "$parentPath\\$($folder.Name)" } else { $folder.Name }
            
            $folderInfo = @{
              Name = $folder.Name
              FolderPath = $folderPath
              Count = $folder.Items.Count
              UnreadCount = $folder.UnReadItemCount
              SubFolders = @()
            }
            
            # Parcourir les sous-dossiers
            foreach ($subfolder in $folder.Folders) {
              $folderInfo.SubFolders += Get-FolderStructure $subfolder $folderPath
            }
            
            return $folderInfo
          }
          
          $allFolders = @()
          
          # Parcourir tous les comptes
          foreach ($account in $namespace.Accounts) {
            try {
              $rootFolder = $namespace.Folders($account.DisplayName)
              if ($rootFolder) {
                $accountStructure = Get-FolderStructure $rootFolder
                $allFolders += $accountStructure
              }
            }
            catch {
              # Ignorer les comptes qui causent des erreurs
              continue
            }
          }
          
          $result = @{
            Folders = $allFolders
            Timestamp = [DateTime]::Now.ToString("yyyy-MM-ddTHH:mm:ss")
          }
          
          $json = $result | ConvertTo-Json -Depth 10 -Compress
          Write-Output $json
        }
        catch {
          Write-Output "ERROR: $($_.Exception.Message)"
        }
      `;

      const result = await this.executePowerShell(script, 30000);
      
      if (result.startsWith('ERROR:')) {
        throw new Error(result.replace('ERROR: ', ''));
      }

      const structureData = JSON.parse(result);
      this.log(`📁 Structure récupérée: ${structureData.Folders.length} comptes analysés`);
      
      return structureData.Folders;

    } catch (error) {
      this.log(`❌ Erreur récupération structure dossiers: ${error.message}`);
      throw error;
    }
  }

  // === NETTOYAGE ===

  async dispose() {
    this.log('🧹 Nettoyage du connecteur Outlook');
    
    if (this.checkInterval) {
      clearInterval(this.checkInterval);
      this.checkInterval = null;
    }
    
    this.folders.clear();
    this.stats.clear();
    this.requestQueue.length = 0;
    this.isOutlookConnected = false;
    this.connectionState = 'disposed';
    
    this.log('✅ Connecteur nettoyé');
  }
}

// Instance singleton
const outlookConnector = new OutlookConnector();

// Gestion propre de l'arrêt
process.on('SIGTERM', () => outlookConnector.dispose());
process.on('SIGINT', () => outlookConnector.dispose());

module.exports = outlookConnector;
