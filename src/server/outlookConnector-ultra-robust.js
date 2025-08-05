/**
 * Connecteur Outlook COM ultra robuste
 * Gestion avancée des erreurs, retry automatique, et monitoring en temps réel
 */

const { exec, spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

class OutlookConnector {
  constructor() {
    this.isOutlookConnected = false;
    this.lastCheck = null;
    this.connectionAttempts = 0;
    this.maxRetries = 5;
    this.retryDelay = 2000;
    this.checkInterval = null;
    this.folders = new Map();
    this.stats = new Map();
    this.connectionState = 'unknown'; // unknown, connecting, connected, failed, unavailable
    this.lastError = null;
    this.outlookVersion = null;
    this.comSecurityLevel = null;
    
    // Configuration avancée
    this.config = {
      timeout: 15000, // 15 secondes timeout
      maxConcurrentRequests: 3,
      enableDetailedLogs: true,
      autoReconnect: true,
      cacheTTL: 30000, // 30 secondes cache
      healthCheckInterval: 60000 // 1 minute
    };
    
    // File d'attente pour les requêtes
    this.requestQueue = [];
    this.activeRequests = 0;
    
    this.init();
  }

  async init() {
    this.log('🚀 Initialisation du connecteur Outlook COM ultra robuste');
    
    // Vérification de l'environnement
    await this.checkEnvironment();
    
    // Démarrage de la connexion initiale
    await this.establishConnection();
    
    // Démarrage du monitoring continu
    this.startHealthMonitoring();
    
    this.log('✅ Connecteur Outlook initialisé');
  }

  async checkEnvironment() {
    this.log('🔍 Vérification de l\'environnement Windows/Outlook...');
    
    try {
      // Vérifier si on est sur Windows
      if (process.platform !== 'win32') {
        throw new Error('Outlook COM n\'est disponible que sur Windows');
      }
      
      // Vérifier la version de PowerShell
      const psVersion = await this.executeCommand('powershell', ['$PSVersionTable.PSVersion.Major']);
      this.log(`📋 Version PowerShell: ${psVersion.trim()}`);
      
      // Vérifier les permissions COM
      await this.checkCOMPermissions();
      
      // Détecter la version d'Outlook installée
      await this.detectOutlookVersion();
      
    } catch (error) {
      this.log(`❌ Erreur environnement: ${error.message}`);
      this.connectionState = 'unavailable';
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
    this.log('🔗 Établissement de la connexion Outlook...');
    this.connectionState = 'connecting';
    
    try {
      // Étape 1: Vérifier le processus Outlook
      const processRunning = await this.checkOutlookProcess();
      
      if (!processRunning) {
        this.log('⚠️ Processus Outlook non détecté, tentative de démarrage...');
        await this.startOutlook();
        
        // Attendre que le processus démarre
        await this.waitForProcess('OUTLOOK.EXE', 30000);
      }
      
      // Étape 2: Tester la connexion COM avec retry
      await this.testCOMConnection();
      
      // Étape 3: Initialiser les données de base
      await this.loadInitialData();
      
      this.connectionState = 'connected';
      this.isOutlookConnected = true;
      this.connectionAttempts = 0;
      this.lastError = null;
      
      this.log('✅ Connexion Outlook établie avec succès');
      
    } catch (error) {
      this.handleConnectionError(error);
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
          
          # Test d'accès à la boîte de réception
          $inbox = $namespace.GetDefaultFolder(6)
          $result.InboxName = $inbox.Name
          $result.InboxItemCount = $inbox.Items.Count
          $result.UnreadCount = $inbox.UnReadItemCount
          
          # Libération des objets COM
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($inbox) | Out-Null
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
          
          $result | ConvertTo-Json -Compress
        } else {
          throw "Impossible de créer l'objet Outlook.Application"
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
          this.log(`📊 Boîte réception: ${data.InboxItemCount} éléments, ${data.UnreadCount} non lus`);
          return data;
        } else {
          throw new Error(`Connexion échouée: ${data.Error} (${data.ErrorType})`);
        }
        
      } catch (error) {
        this.log(`❌ Tentative ${attempt} échouée: ${error.message}`);
        
        if (attempt < this.maxRetries) {
          const delay = this.retryDelay * attempt; // Délai progressif
          this.log(`⏳ Nouvelle tentative dans ${delay}ms...`);
          await this.sleep(delay);
        } else {
          throw new Error(`Échec de la connexion COM après ${this.maxRetries} tentatives: ${error.message}`);
        }
      }
    }
  }

  async loadInitialData() {
    this.log('📚 Chargement des données initiales...');
    
    try {
      // Charger les dossiers en parallèle
      const [folders, stats] = await Promise.allSettled([
        this.loadFoldersRobust(),
        this.loadStatsRobust()
      ]);
      
      if (folders.status === 'fulfilled') {
        this.log(`📁 ${folders.value.length} dossiers chargés`);
      } else {
        this.log(`⚠️ Erreur chargement dossiers: ${folders.reason.message}`);
      }
      
      if (stats.status === 'fulfilled') {
        this.log(`📊 Statistiques chargées`);
      } else {
        this.log(`⚠️ Erreur chargement stats: ${stats.reason.message}`);
      }
      
    } catch (error) {
      this.log(`⚠️ Erreur chargement données: ${error.message}`);
    }
  }

  async loadFoldersRobust() {
    const foldersScript = `
      try {
        $ErrorActionPreference = "Stop"
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        $folders = @()
        
        # Dossiers par défaut
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
          } catch {
            # Continuer même si un dossier spécifique échoue
          }
        }
        
        # Dossiers personnalisés dans la racine
        $store = $namespace.DefaultStore
        $rootFolder = $store.GetRootFolder()
        
        foreach ($subfolder in $rootFolder.Folders) {
          if ($subfolder.DefaultItemType -eq 0) { # olMailItem
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
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($subfolder) | Out-Null
        }
        
        # Libération des objets
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($store) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        
        $folders | ConvertTo-Json -Compress
        
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
        return data;
      } else {
        throw new Error('Format de données dossiers invalide');
      }
      
    } catch (error) {
      throw new Error(`Échec chargement dossiers: ${error.message}`);
    }
  }

  async loadStatsRobust() {
    const statsScript = `
      try {
        $ErrorActionPreference = "Stop"
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Statistiques générales
        $inbox = $namespace.GetDefaultFolder(6)
        $sentItems = $namespace.GetDefaultFolder(5)
        $drafts = $namespace.GetDefaultFolder(16)
        
        # Compter les emails d'aujourd'hui
        $today = (Get-Date).Date
        $todayFilter = "[ReceivedTime] >= '$($today.ToString("MM/dd/yyyy"))'"
        $sentTodayFilter = "[SentOn] >= '$($today.ToString("MM/dd/yyyy"))'"
        
        $emailsToday = ($inbox.Items.Restrict($todayFilter)).Count
        $sentToday = ($sentItems.Items.Restrict($sentTodayFilter)).Count
        
        $stats = @{
          emailsToday = $emailsToday
          sentToday = $sentToday
          unreadTotal = $inbox.UnReadItemCount
          totalInbox = $inbox.Items.Count
          totalSent = $sentItems.Items.Count
          totalDrafts = $drafts.Items.Count
          lastSync = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
          profileName = $namespace.CurrentProfileName
          serverVersion = $outlook.Version
        }
        
        # Libération des objets
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($inbox) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sentItems) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($drafts) | Out-Null
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
      throw new Error(`Échec chargement stats: ${error.message}`);
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
      this.log(`⚠️ Erreur vérification processus: ${error.message}`);
      return false;
    }
  }

  async startOutlook() {
    this.log('🚀 Tentative de démarrage d\'Outlook...');
    
    try {
      // Tenter de démarrer Outlook
      const outlookPaths = [
        'outlook.exe',
        'C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE',
        'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE',
        'C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE',
        'C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE'
      ];
      
      for (const outlookPath of outlookPaths) {
        try {
          await this.executeCommand(outlookPath, [], { timeout: 5000 });
          this.log(`✅ Outlook démarré depuis: ${outlookPath}`);
          return true;
        } catch (error) {
          // Continuer avec le chemin suivant
        }
      }
      
      throw new Error('Impossible de trouver l\'exécutable Outlook');
      
    } catch (error) {
      this.log(`❌ Échec démarrage Outlook: ${error.message}`);
      throw error;
    }
  }

  async waitForProcess(processName, timeout = 30000) {
    const startTime = Date.now();
    
    while (Date.now() - startTime < timeout) {
      if (await this.checkOutlookProcess()) {
        this.log(`✅ Processus ${processName} détecté`);
        
        // Attendre encore 5 secondes pour l'initialisation complète
        await this.sleep(5000);
        return true;
      }
      
      await this.sleep(1000);
    }
    
    throw new Error(`Timeout: processus ${processName} non détecté après ${timeout}ms`);
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
      this.log(`💥 Arrêt des tentatives après ${this.connectionAttempts} échecs`);
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

  async executePowerShell(script, timeout = 10000) {
    return new Promise((resolve, reject) => {
      const child = spawn('powershell', [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-Command', script
      ], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true
      });
      
      let stdout = '';
      let stderr = '';
      
      child.stdout.on('data', (data) => {
        stdout += data.toString();
      });
      
      child.stderr.on('data', (data) => {
        stderr += data.toString();
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
      console.log(`[${timestamp}] OutlookConnector: ${message}`);
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
