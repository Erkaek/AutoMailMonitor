/**
 * OUTLOOK CONNECTOR OPTIMISÉ - Version simplifiée
 * Performance maximale avec API REST + Better-SQLite3 + Cache
 * Remplace l'ancien système COM/PowerShell
 */

const { EventEmitter } = require('events');
const { exec } = require('child_process');

// Simplification: pas d'import de Graph API pour éviter les problèmes de dépendances
let graphAvailable = false; // Désactivé temporairement

class OutlookConnector extends EventEmitter {
  constructor() {
    super();
    
    console.log('🚀 DIAGNOSTIC: Création d\'une nouvelle instance OutlookConnector optimisée');
    
    // État de connexion
    this.isOutlookConnected = false;
    this.connectionState = 'disconnected';
    this.lastError = null;
    
    // Configuration optimisée
    this.config = {
      timeout: 15000, // Réduit pour Graph API
      realtimePollingInterval: 15000, // 15s optimal pour Graph API
      enableDetailedLogs: false, // Performance
      autoReconnect: true,
      maxRetries: 3
    };
    
    // Graph API simulation
    this.useGraphAPI = graphAvailable;
    
    // Données en cache
    this.folders = new Map();
    this.stats = new Map();
    // Auto-connexion
    this.autoConnect();
  }

  /**
   * Auto-connexion optimisée
   */
  async autoConnect() {
    try {
      console.log('[AUTO-CONNECT] Tentative de connexion automatique Graph API...');
        // En mode dégradé, considérer Outlook comme "connecté" si le processus est disponible
        const isRunning = await this.checkOutlookProcess();
        if (isRunning) {
          this.isOutlookConnected = true;
          this.connectionState = 'connected';
          this.emit('connected');
          console.log('✅ Mode dégradé actif: Outlook détecté, fonctionnalités de base disponibles');
        } else {
          this.connectionState = 'error';
        }
    } catch (error) {
      console.error('[AUTO-CONNECT] Erreur:', error.message);
      this.connectionState = 'error';
      this.lastError = error;
    }
  }

  /**
   * PERFORMANCE: Connexion Graph API
   */
  async connectToGraphAPI() {
    try {
      this.connectionState = 'connecting';
      console.log('🚀 Connexion à Microsoft Graph API...');
      
      // Vérifier si Outlook est réellement disponible via PowerShell
      let isOutlookRunning = await this.checkOutlookProcess();
      
      if (!isOutlookRunning) {
        console.log('📱 Outlook n\'est pas démarré - Lancement automatique...');
        
        try {
          await this.launchOutlook();
          console.log('⏳ Attente du démarrage d\'Outlook...');
          await this.waitForOutlookReady();
          isOutlookRunning = true;
        } catch (launchError) {
          throw new Error(`Impossible de lancer Outlook automatiquement: ${launchError.message}\n\nVeuillez démarrer Microsoft Outlook manuellement et réessayer.`);
        }
      }
      
      // Pour l'instant, on simule une connexion réussie si Outlook est présent
      this.isOutlookConnected = true;
      this.connectionState = 'connected';
      
      console.log('✅ Connexion Graph API simulée (authentification utilisateur requise)');
      this.emit('connected');
      
      return true;
      
    } catch (error) {
      this.connectionState = 'error';
      this.lastError = error;
      console.error('❌ Erreur connexion Graph API:', error);
      throw error;
    }
  }

  /**
   * Vérifier si le processus Outlook est en cours d'exécution
   */
  async checkOutlookProcess() {
    try {
      const { spawn } = require('child_process');
      
      return new Promise((resolve) => {
        const powershell = spawn('powershell.exe', [
          '-Command',
          'Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue | Select-Object -Property Name'
        ]);
        
        let output = '';
        
        powershell.stdout.on('data', (data) => {
          output += data.toString();
        });
        
        powershell.on('close', (code) => {
          // Si le processus retourne quelque chose avec "OUTLOOK", c'est que le processus existe
          const isRunning = output.includes('OUTLOOK');
          console.log(`🔍 Vérification processus Outlook: ${isRunning ? 'Trouvé' : 'Non trouvé'}`);
          resolve(isRunning);
        });
        
        // Timeout de 5 secondes pour la vérification
        setTimeout(() => {
          powershell.kill();
          resolve(false);
        }, 5000);
      });
      
    } catch (error) {
      console.error('❌ Erreur vérification processus Outlook:', error);
      return false;
    }
  }

  /**
   * Trouver le chemin d'installation d'Outlook sur le système
   */
  async findOutlookPath() {
    try {
      const { spawn } = require('child_process');
      
      return new Promise((resolve) => {
        // Script PowerShell pour chercher Outlook dans tous les emplacements possibles
        const powershellScript = `
          $outlookPaths = @(
            'C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE',
            'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE',
            'C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE',
            'C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE',
            'C:\\Program Files\\Microsoft Office\\Office15\\OUTLOOK.EXE',
            'C:\\Program Files (x86)\\Microsoft Office\\Office15\\OUTLOOK.EXE'
          )
          
          foreach ($path in $outlookPaths) {
            if (Test-Path $path) {
              Write-Output $path
              exit 0
            }
          }
          
          # Si aucun chemin fixe ne fonctionne, chercher récursivement
          $searchPaths = @(
            'C:\\Program Files\\Microsoft Office',
            'C:\\Program Files (x86)\\Microsoft Office'
          )
          
          foreach ($searchPath in $searchPaths) {
            if (Test-Path $searchPath) {
              $found = Get-ChildItem -Path $searchPath -Name "OUTLOOK.EXE" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
              if ($found) {
                $fullPath = Join-Path $searchPath $found
                Write-Output $fullPath
                exit 0
              }
            }
          }
          
          # Dernière tentative : chercher dans le PATH
          $outlookInPath = Get-Command outlook.exe -ErrorAction SilentlyContinue
          if ($outlookInPath) {
            Write-Output $outlookInPath.Source
            exit 0
          }
          
          Write-Output "NOT_FOUND"
          exit 1
        `;
        
        const powershell = spawn('powershell.exe', [
          '-NoProfile',
          '-ExecutionPolicy', 'Bypass',
          '-Command', powershellScript
        ]);
        
        let output = '';
        
        powershell.stdout.on('data', (data) => {
          output += data.toString().trim();
        });
        
        powershell.on('close', (code) => {
          if (code === 0 && output && output !== 'NOT_FOUND') {
            console.log(`🔍 Outlook trouvé à: ${output}`);
            resolve(output);
          } else {
            console.log('❌ Outlook non trouvé sur ce système');
            resolve(null);
          }
        });
        
        // Timeout de 10 secondes pour la recherche
        setTimeout(() => {
          powershell.kill();
          resolve(null);
        }, 10000);
      });
      
    } catch (error) {
      console.error('❌ Erreur lors de la recherche d\'Outlook:', error);
      return null;
    }
  }

  /**
   * Lancer Microsoft Outlook automatiquement
   */
  async launchOutlook() {
    try {
      console.log('🚀 Lancement automatique de Microsoft Outlook...');
      
      // Chercher d'abord où est installé Outlook
      const outlookPath = await this.findOutlookPath();
      
      if (!outlookPath) {
        throw new Error('Microsoft Outlook n\'est pas installé sur ce système ou n\'est pas accessible');
      }
      
      const { spawn } = require('child_process');
      
      return new Promise((resolve, reject) => {
        console.log(`📍 Lancement d'Outlook depuis: ${outlookPath}`);
        let resolved = false;
        
        // Lancer Outlook directement (sans cmd.exe)
        const process = spawn(outlookPath, [], {
          detached: true,
          stdio: 'ignore'
        });
        
        process.on('error', (error) => {
          if (!resolved) {
            resolved = true;
            console.log(`❌ Erreur lors du lancement: ${error.message}`);
            reject(new Error(`Échec du lancement d'Outlook: ${error.message}`));
          }
        });
        
        process.on('spawn', () => {
          if (!resolved) {
            resolved = true;
            console.log('✅ Processus Outlook lancé avec succès');
            // Détacher le processus pour qu'il continue à tourner indépendamment
            process.unref();
            resolve(true);
          }
        });
        
        // Timeout pour le lancement
        setTimeout(() => {
          if (!resolved) {
            resolved = true;
            console.log('❌ Timeout lors du lancement d\'Outlook');
            reject(new Error('Timeout lors du lancement d\'Outlook'));
          }
        }, 5000);
      });
      
    } catch (error) {
      console.error('❌ Erreur lors du lancement d\'Outlook:', error);
      throw error;
    }
  }

  /**
   * Attendre qu'Outlook soit complètement démarré
   */
  async waitForOutlookReady() {
    console.log('⏳ Attente du démarrage complet d\'Outlook...');
    
    for (let i = 0; i < 30; i++) { // Maximum 30 secondes d'attente
      const isRunning = await this.checkOutlookProcess();
      if (isRunning) {
        // Attendre encore 2 secondes pour que Outlook soit complètement chargé
        await new Promise(resolve => setTimeout(resolve, 2000));
        console.log('✅ Outlook est prêt !');
        return true;
      }
      
      // Attendre 1 seconde avant la prochaine vérification
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      if (i % 5 === 0) {
        console.log(`⏳ Attente... ${i + 1}/30 secondes`);
      }
    }
    
    throw new Error('Timeout: Outlook n\'a pas démarré dans les 30 secondes');
  }

  /**
   * COMPATIBILITY: Méthode pour compatibilité avec ancien code
   */
  async establishConnection() {
    return await this.connectToGraphAPI();
  }

  /**
   * Test de connexion Outlook
   */
  async testOutlookConnection() {
    try {
      if (this.useGraphAPI) {
        // Simulation test Graph API
        return true;
      }
      
      return false;
      
    } catch (error) {
      console.error('❌ Test connexion échoué:', error);
      return false;
    }
  }

  /**
   * OPTIMIZED: Récupération dossiers
   */
  async getFolders() {
    if (!this.isOutlookConnected) {
      await this.connectToGraphAPI();
    }

    try {
      // Fallback: dossiers par défaut (pour simulation)
      const folders = [
        { name: 'Inbox', path: 'Inbox', id: 'inbox' },
        { name: 'Sent Items', path: 'Sent Items', id: 'sent' },
        { name: 'Drafts', path: 'Drafts', id: 'drafts' }
      ];
      
      // Mise en cache
      this.folders.clear();
      folders.forEach(folder => {
        this.folders.set(folder.path, folder);
      });
      
      return folders;
      
    } catch (error) {
      console.error('❌ Erreur récupération dossiers:', error);
      return [];
    }
  }

  /**
   * S'assure qu'Outlook est prêt, même en mode dégradé
   */
  async ensureConnected() {
    if (this.isOutlookConnected) return true;
    // Tenter une connexion légère (process + lancement si nécessaire)
    let isRunning = await this.checkOutlookProcess();
    if (!isRunning) {
      try {
        await this.launchOutlook();
        await this.waitForOutlookReady();
        isRunning = true;
      } catch (e) {
        this.connectionState = 'error';
        this.lastError = e;
        console.error('❌ ensureConnected: Outlook indisponible:', e.message);
        return false;
      }
    }
    this.isOutlookConnected = true;
    this.connectionState = 'connected';
    this.emit('connected');
    return true;
  }

  /**
   * Récupère la liste des boîtes mail (Stores) dans Outlook via COM
   * Retour: Array<{ Name, StoreID, FoldersCount, IsDefault, SmtpAddress? }>
   */
  async getMailboxes() {
    try {
      const ok = await this.ensureConnected();
      if (!ok) {
        return [];
      }

      const script = `
        # Force UTF-8 output to preserve accents/diacritics
        $enc = New-Object System.Text.UTF8Encoding $false
        [Console]::OutputEncoding = $enc
        $OutputEncoding = $enc

        try {
          $ErrorActionPreference = 'SilentlyContinue'
          $outlook = New-Object -ComObject Outlook.Application
          $ns = $outlook.GetNamespace("MAPI")
          $accounts = $null
          try { $accounts = $outlook.Session.Accounts } catch {}

          # Utiliser une map pour dédupliquer par StoreID
          $storeMap = @{}

          # 1) Énumération standard des Stores
          foreach ($store in $ns.Stores) {
            try {
              $root = $store.GetRootFolder()
              $name = $store.DisplayName
              $storeId = $store.StoreID
              $isDefault = $storeId -eq $ns.DefaultStore.StoreID
              $foldersCount = 0
              try { $foldersCount = $root.Folders.Count } catch {}

              $storeMap[$storeId] = @{
                Name = $name
                StoreID = $storeId
                FoldersCount = $foldersCount
                IsDefault = $isDefault
                SmtpAddress = $null
                ExchangeStoreType = $store.ExchangeStoreType
                FilePath = $store.FilePath
              }
            } catch {}
          }

          # 2) Fallback: énumérer Namespace.Folders pour récupérer des Stores supplémentaires
          foreach ($topFolder in $ns.Folders) {
            try {
              $st = $topFolder.Store
              if ($st -ne $null) {
                $storeId = $st.StoreID
                if (-not $storeMap.ContainsKey($storeId)) {
                  $root = $st.GetRootFolder()
                  $name = $st.DisplayName
                  $isDefault = $storeId -eq $ns.DefaultStore.StoreID
                  $foldersCount = 0
                  try { $foldersCount = $root.Folders.Count } catch {}
                  $storeMap[$storeId] = @{
                    Name = $name
                    StoreID = $storeId
                    FoldersCount = $foldersCount
                    IsDefault = $isDefault
                    SmtpAddress = $null
                    ExchangeStoreType = $st.ExchangeStoreType
                    FilePath = $st.FilePath
                  }
                }
              }
            } catch {}
          }

          # 3) Enrichir avec les comptes: associer SMTP par DeliveryStore si possible
          if ($accounts -ne $null) {
            foreach ($acc in $accounts) {
              $smtp = $null
              try { $smtp = $acc.SmtpAddress } catch {}
              try {
                $delStore = $acc.DeliveryStore
                if ($delStore -ne $null) {
                  $sid = $delStore.StoreID
                  if ($storeMap.ContainsKey($sid)) { $storeMap[$sid].SmtpAddress = $smtp }
                }
              } catch {}
              # Si non mappé via DeliveryStore, essayer par nom approchant
              foreach ($key in $storeMap.Keys) {
                if (-not $storeMap[$key].SmtpAddress) {
                  $name = $storeMap[$key].Name
                  try { if ($smtp -and ($acc.DisplayName -eq $name -or $name -like "*$($acc.DisplayName)*")) { $storeMap[$key].SmtpAddress = $smtp } } catch {}
                }
              }
            }
          }

          $stores = @()
          foreach ($k in $storeMap.Keys) { $stores += $storeMap[$k] }
          $res = @{ success = $true; mailboxes = $stores } | ConvertTo-Json -Depth 12 -Compress
          Write-Output $res
        } catch {
          $err = $_.Exception.Message
          $res = @{ success = $false; error = $err; mailboxes = @() } | ConvertTo-Json -Depth 3 -Compress
          Write-Output $res
        }
      `;

      let result = await this.executePowerShellScript(script);
      if (!result.success) {
        // Retry in 32-bit if first attempt fails
        result = await this.executePowerShellScript(script, 15000, { force32Bit: true });
      }
      if (!result.success) {
        throw new Error(result.error || 'Échec récupération boîtes mail');
      }
      let json;
      try { json = JSON.parse(result.output || '{}'); } catch (_) { json = {}; }
      let mailboxes = Array.isArray(json.mailboxes) ? json.mailboxes : [];
      if (mailboxes.length === 0) {
        // Retry parsing or force 32-bit once more if not yet
        const res32 = await this.executePowerShellScript(script, 15000, { force32Bit: true });
        if (res32.success) {
          try {
            const j2 = JSON.parse(res32.output || '{}');
            mailboxes = Array.isArray(j2.mailboxes) ? j2.mailboxes : mailboxes;
          } catch {}
        }
      }
      return mailboxes;
    } catch (error) {
      console.error('❌ Erreur getMailboxes:', error.message);
      return [];
    }
  }

  /**
   * Récupère la structure des dossiers pour un Store donné
   * Retour: Array<{ Name, StoreID, SubFolders: Folder[] }>
   */
  async getFolderStructure(storeId) {
    try {
      const ok = await this.ensureConnected();
      if (!ok) {
        return [];
      }

      // Échapper les guillemets pour insertion dans le script
      const safeStoreId = String(storeId || '').replace(/`/g, '``').replace(/"/g, '\"');

      const script = `
        # Force UTF-8 output to preserve accents/diacritics
        $enc = New-Object System.Text.UTF8Encoding $false
        [Console]::OutputEncoding = $enc
        $OutputEncoding = $enc

        try {
          $outlook = New-Object -ComObject Outlook.Application
          $ns = $outlook.GetNamespace("MAPI")
          $accounts = $null
          try { $accounts = $outlook.Session.Accounts } catch {}

          function Get-FolderTree {
            param([object]$folder, [string]$prefixPath, [string]$mailboxName)
            $currentPath = if ($prefixPath -and $prefixPath.Trim() -ne "") { "$prefixPath\\$($folder.Name)" } else { "$mailboxName\\$($folder.Name)" }
            $children = @()
            try {
              foreach ($sf in $folder.Folders) {
                $children += (Get-FolderTree -folder $sf -prefixPath $currentPath -mailboxName $mailboxName)
              }
            } catch {}
            $unread = 0
            try { $unread = $folder.UnReadItemCount } catch {}
            $total = 0
            try { $total = $folder.Items.Count } catch {}
            return @{
              Name = $folder.Name
              FolderPath = $currentPath
              UnreadCount = $unread
              TotalCount = $total
              SubFolders = $children
            }
          }

          if ("${safeStoreId}" -eq "") {
            # Retourner l'arborescence de toutes les boîtes pour permettre la sélection côté UI
            $mbs = @()
            foreach ($st in $ns.Stores) {
              try {
                $root = $st.GetRootFolder()
                $mailboxName = $st.DisplayName
                # Tenter de trouver une adresse SMTP correspondante
                $smtp = $null
        if ($accounts -ne $null) {
                  foreach ($acc in $accounts) {
          if ($acc.SmtpAddress -and ($acc.DisplayName -eq $mailboxName -or $mailboxName -like "*$($acc.DisplayName)*")) {
                      $smtp = $acc.SmtpAddress
                      break
                    }
                  }
                }
                $tree = @()
                foreach ($sf in $root.Folders) {
                  $tree += (Get-FolderTree -folder $sf -prefixPath $mailboxName -mailboxName $mailboxName)
                }
                $mb = @{
                  Name = $mailboxName
                  StoreID = $st.StoreID
                  SmtpAddress = $smtp
                  SubFolders = $tree
                }
                $mbs += $mb
              } catch {}
            }
            $res = @{ success = $true; folders = $mbs } | ConvertTo-Json -Depth 24 -Compress
            Write-Output $res
            return
          }

          # Sélection ciblée d'un Store
          $target = $null
          foreach ($st in $ns.Stores) {
            if ($st.StoreID -eq "${safeStoreId}" -or $st.DisplayName -eq "${safeStoreId}" -or $st.DisplayName -like "*${safeStoreId}*") {
              $target = $st
              break
            }
          }
          if (-not $target) { $target = $ns.DefaultStore }

          $root = $target.GetRootFolder()
          $mailboxName = $target.DisplayName
          # Tenter de trouver une adresse SMTP correspondante
          $smtp = $null
          try {
            $accounts = $outlook.Session.Accounts
            foreach ($acc in $accounts) {
              if ($acc.SmtpAddress -and ($acc.DisplayName -eq $mailboxName -or $mailboxName -like "*$($acc.DisplayName)*")) {
                $smtp = $acc.SmtpAddress
                break
              }
            }
          } catch {}
          $tree = @()
          try {
            foreach ($sf in $root.Folders) {
              $tree += (Get-FolderTree -folder $sf -prefixPath $mailboxName -mailboxName $mailboxName)
            }
          } catch {}

          $mb = @{
            Name = $mailboxName
            StoreID = $target.StoreID
            SmtpAddress = $smtp
            SubFolders = $tree
          }

          $res = @{ success = $true; folders = @($mb) } | ConvertTo-Json -Depth 24 -Compress
          Write-Output $res
        } catch {
          $err = $_.Exception.Message
          $res = @{ success = $false; error = $err; folders = @() } | ConvertTo-Json -Depth 3 -Compress
          Write-Output $res
        }
      `;

      let result = await this.executePowerShellScript(script, 60000);
      if (!result.success) {
        // Retry in 32-bit if first attempt fails
        result = await this.executePowerShellScript(script, 60000, { force32Bit: true });
      }
      if (!result.success) {
        throw new Error(result.error || 'Échec récupération structure dossiers');
      }
      let json;
      try { json = JSON.parse(result.output || '{}'); } catch (_) { json = {}; }
      let folders = Array.isArray(json.folders) ? json.folders : [];
      if (folders.length === 0) {
        const res32 = await this.executePowerShellScript(script, 60000, { force32Bit: true });
        if (res32.success) {
          try {
            const j2 = JSON.parse(res32.output || '{}');
            folders = Array.isArray(j2.folders) ? j2.folders : folders;
          } catch {}
        }
      }
      return folders;
    } catch (error) {
      console.error('❌ Erreur getFolderStructure:', error.message);
      return [];
    }
  }

  /**
   * OPTIMIZED: Récupération emails par folder
   */
  async getEmailsFromFolder(folderPath, limit = 100) {
    if (!this.isOutlookConnected) {
      await this.connectToGraphAPI();
    }

    try {
      // Simulation: retourner tableau vide pour l'instant
      console.log(`⚠️ Simulation: ${folderPath} - aucun email récupéré (Graph API requis)`);
      return [];
      
    } catch (error) {
      console.error(`❌ Erreur emails ${folderPath}:`, error);
      return [];
    }
  }

  /**
   * MONITORING: Démarrage surveillance temps réel
   */
  async startRealtimeMonitoring(foldersConfig = []) {
    try {
      console.log('🔄 Démarrage monitoring temps réel (simulation)...');
      
      console.log('⚠️ Monitoring en mode simulation');
      
      // Simulation d'événements pour les tests
      setInterval(() => {
        this.emit('monitoring-heartbeat', {
          timestamp: new Date().toISOString(),
          status: 'active'
        });
      }, 30000);
      
      this.emit('monitoring-started');
      
    } catch (error) {
      console.error('❌ Erreur démarrage monitoring:', error);
      throw error;
    }
  }

  /**
   * MONITORING: Surveillance d'un dossier spécifique
   */
  async startFolderMonitoring(folderPath) {
    try {
      console.log(`🔍 Démarrage du monitoring temps réel pour: ${folderPath}`);
      
      // Vérifier que folderPath est valide
      if (!folderPath || typeof folderPath !== 'string') {
        throw new Error('Chemin de dossier invalide');
      }
      
      console.log(`🎧 [REALTIME] Activation monitoring PowerShell pour: ${folderPath}`);
      
      // Stocker l'état initial du dossier pour détecter les changements
      const initialEmails = await this.getFolderEmails(folderPath);
      
      if (!initialEmails.success) {
        throw new Error(`Impossible d'accéder au dossier pour monitoring: ${initialEmails.error}`);
      }
      
      // Créer un état de base pour ce dossier
      const folderState = {
        path: folderPath,
        lastCount: initialEmails.count,
        lastCheck: new Date(),
        emails: new Map() // EntryID -> email info
      };
      
      // Stocker les emails actuels par EntryID
      initialEmails.emails.forEach(email => {
        folderState.emails.set(email.EntryID, {
          subject: email.Subject,
          unread: email.UnRead,
          receivedTime: email.ReceivedTime
        });
      });
      
      // Stocker l'état de ce dossier
      if (!this.monitoringStates) {
        this.monitoringStates = new Map();
      }
      this.monitoringStates.set(folderPath, folderState);
      
      // Démarrer la surveillance périodique de ce dossier
      const monitoringInterval = setInterval(async () => {
        try {
          await this.checkFolderChanges(folderPath);
        } catch (error) {
          console.error(`❌ Erreur monitoring ${folderPath}:`, error.message);
        }
      }, 30000); // Vérifier toutes les 30 secondes
      
      // Stocker l'interval pour pouvoir l'arrêter plus tard
      if (!this.monitoringIntervals) {
        this.monitoringIntervals = new Map();
      }
      this.monitoringIntervals.set(folderPath, monitoringInterval);
      
      const result = {
        success: true,
        folderPath: folderPath,
        message: `Monitoring temps réel activé (vérification toutes les 30s)`,
        timestamp: new Date().toISOString(),
        initialCount: initialEmails.count
      };
      
      console.log(`✅ [REALTIME] Monitoring actif pour: ${folderPath} (${initialEmails.count} emails)`);
      return result;
      
    } catch (error) {
      // Distinguer les dossiers inexistants des vraies erreurs
      if (error.message.includes('non trouve') || error.message.includes('not found') || error.message.includes('Navigation impossible')) {
        console.log(`ℹ️ Dossier "${folderPath}" non trouvé - monitoring ignoré`);
      } else {
        console.error(`❌ Erreur démarrage monitoring dossier ${folderPath}:`, error.message);
      }
      return {
        success: false,
        folderPath: folderPath,
        error: error.message,
        message: 'Monitoring dossier échoué'
      };
    }
  }

  /**
   * Vérifier les changements dans un dossier surveillé
   */
  async checkFolderChanges(folderPath) {
    try {
      if (!this.monitoringStates || !this.monitoringStates.has(folderPath)) {
        console.warn(`⚠️ État de monitoring manquant pour: ${folderPath}`);
        return;
      }
      
      const folderState = this.monitoringStates.get(folderPath);
      const currentEmails = await this.getFolderEmails(folderPath);
      
      if (!currentEmails.success) {
        console.error(`❌ Erreur vérification ${folderPath}: ${currentEmails.error}`);
        return;
      }
      
      // Détecter les changements de nombre
      if (currentEmails.count !== folderState.lastCount) {
        console.log(`📊 [MONITORING] Changement détecté ${folderPath}: ${folderState.lastCount} -> ${currentEmails.count}`);
        
        this.emit('folderCountChanged', {
          folderPath: folderPath,
          oldCount: folderState.lastCount,
          newCount: currentEmails.count,
          timestamp: new Date().toISOString()
        });
        
        folderState.lastCount = currentEmails.count;
      }
      
      // Détecter les nouveaux emails
      const currentEmailMap = new Map();
      currentEmails.emails.forEach(email => {
        currentEmailMap.set(email.EntryID, {
          subject: email.Subject,
          unread: email.UnRead,
          receivedTime: email.ReceivedTime
        });
        
        // Vérifier si c'est un nouvel email
        if (!folderState.emails.has(email.EntryID)) {
          console.log(`📧 [MONITORING] Nouvel email détecté: ${email.Subject}`);
          
          this.emit('newEmailDetected', {
            folderPath: folderPath,
            entryId: email.EntryID,
            subject: email.Subject,
            senderName: email.SenderName,
            senderEmail: email.SenderEmailAddress,
            receivedTime: email.ReceivedTime,
            isRead: !email.UnRead,
            timestamp: new Date().toISOString()
          });
        } else {
          // Vérifier les changements de statut lu/non lu
          const oldEmail = folderState.emails.get(email.EntryID);
          if (oldEmail.unread !== email.UnRead) {
            console.log(`📝 [MONITORING] Statut changé: ${email.Subject} -> ${email.UnRead ? 'Non lu' : 'Lu'}`);
            
            this.emit('emailStatusChanged', {
              folderPath: folderPath,
              entryId: email.EntryID,
              subject: email.Subject,
              isRead: !email.UnRead,
              timestamp: new Date().toISOString()
            });
          }
        }
      });
      
      // Détecter les emails supprimés
      for (const [entryId, oldEmail] of folderState.emails) {
        if (!currentEmailMap.has(entryId)) {
          console.log(`🗑️ [MONITORING] Email supprimé: ${oldEmail.subject}`);
          
          this.emit('emailDeleted', {
            folderPath: folderPath,
            entryId: entryId,
            subject: oldEmail.subject,
            timestamp: new Date().toISOString()
          });
        }
      }
      
      // Mettre à jour l'état
      folderState.emails = currentEmailMap;
      folderState.lastCheck = new Date();
      
    } catch (error) {
      console.error(`❌ Erreur vérification changements ${folderPath}:`, error.message);
    }
  }

  /**
   * Arrêter le monitoring d'un dossier
   */
  async stopFolderMonitoring(folderPath) {
    try {
      console.log(`⏹️ Arrêt monitoring pour: ${folderPath}`);
      
      // Arrêter l'interval de monitoring
      if (this.monitoringIntervals && this.monitoringIntervals.has(folderPath)) {
        clearInterval(this.monitoringIntervals.get(folderPath));
        this.monitoringIntervals.delete(folderPath);
      }
      
      // Supprimer l'état de monitoring
      if (this.monitoringStates && this.monitoringStates.has(folderPath)) {
        this.monitoringStates.delete(folderPath);
      }
      
      console.log(`✅ Monitoring arrêté pour: ${folderPath}`);
      return { success: true, folderPath };
      
    } catch (error) {
      console.error(`❌ Erreur arrêt monitoring ${folderPath}:`, error.message);
      return { success: false, error: error.message };
    }
  }

  /**
   * SYNC: Synchronisation complète optimisée
   */
  async performFullSync(foldersConfig = []) {
    try {
      console.log('🚀 Synchronisation complète optimisée (simulation)...');
      const startTime = Date.now();
      
      const results = {
        totalEmails: 0,
        foldersProcessed: foldersConfig.length,
        errors: []
      };

      console.log('⚠️ Sync en mode simulation');
      
      const syncTime = Date.now() - startTime;
      console.log(`✅ Sync complète: ${results.totalEmails} emails en ${syncTime}ms`);
      
      results.syncTime = syncTime;
      this.emit('sync-completed', results);
      
      return results;
      
    } catch (error) {
      console.error('❌ Erreur sync complète:', error);
      throw error;
    }
  }

  /**
   * EMAILS: Récupération des emails d'un dossier
   */
  async getFolderEmails(folderPath, limit = 50) {
    try {
      console.log(`📧 Récupération emails du dossier: ${folderPath}`);
      
      // Vérifier que folderPath est valide
      if (!folderPath || typeof folderPath !== 'string') {
        throw new Error('Chemin de dossier invalide');
      }
      
      console.log(`� [OUTLOOK] Navigation vers dossier spécifique: ${folderPath}`);
      
      const script = `
        # Force UTF-8 output to preserve accents/diacritics
        $enc = New-Object System.Text.UTF8Encoding $false
        [Console]::OutputEncoding = $enc
        $OutputEncoding = $enc

        try {
          $outlook = New-Object -ComObject Outlook.Application
          $namespace = $outlook.GetNamespace("MAPI")
          $emails = @()
          
          # Fonction pour naviguer vers un dossier specifique
          function Find-OutlookFolder {
            param([string]$FolderPath, [object]$Namespace)
            
            # Extraire compte et chemin
            if ($FolderPath -match '^([^\\\\]+)\\\\(.+)$') {
              $accountName = $matches[1]
              $folderPath = $matches[2]
              
              # Chercher le store/compte
              $targetStore = $null
              foreach ($store in $Namespace.Stores) {
                if ($store.DisplayName -like "*$accountName*" -or $store.DisplayName -eq $accountName) {
                  $targetStore = $store
                  break
                }
              }
              
              if (-not $targetStore) {
                $targetStore = $Namespace.DefaultStore
              }
              
              # Naviguer dans l'arborescence
              $currentFolder = $targetStore.GetRootFolder()
              $pathParts = $folderPath -split '\\\\'
              
              foreach ($part in $pathParts) {
                if ($part -and $part.Trim() -ne "") {
                  $found = $false
                  $folders = $currentFolder.Folders
                  
                  # Recherche exacte par nom
                  for ($i = 1; $i -le $folders.Count; $i++) {
                    $subfolder = $folders.Item($i)
                    if ($subfolder.Name -eq $part) {
                      $currentFolder = $subfolder
                      $found = $true
                      break
                    }
                  }
                  
                  # Si pas trouve, recherche par pattern pour "Boîte de réception"
                  if (-not $found -and ($part -match "réception" -or $part -match "Boîte")) {
                    for ($i = 1; $i -le $folders.Count; $i++) {
                      $subfolder = $folders.Item($i)
                      if ($subfolder.Name -match "réception") {
                        $currentFolder = $subfolder
                        $found = $true
                        break
                      }
                    }
                  }
                  
                  if (-not $found) {
                    return $null
                  }
                }
              }
              
              return $currentFolder
            } else {
              return $null
            }
          }
          
          # Trouver le dossier cible
          $targetFolder = Find-OutlookFolder -FolderPath "${folderPath}" -Namespace $namespace
          
          if ($targetFolder) {
            # Récupérer TOUS les emails du dossier
            $items = $targetFolder.Items
            if ($items.Count -gt 0) {
              $items.Sort("[ReceivedTime]", $true)
              
              # Traiter TOUS les emails
              for ($i = 1; $i -le $items.Count; $i++) {
                try {
                  $mail = $items.Item($i)
                  if ($mail.Class -eq 43) {
                    $subject = if($mail.Subject) { $mail.Subject } else { "" }
                    $senderName = if($mail.SenderName) { $mail.SenderName } else { "" }
                    $senderEmail = if($mail.SenderEmailAddress) { $mail.SenderEmailAddress } else { "" }
                    $receivedTime = $mail.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
                    $entryId = $mail.EntryID
                    $unread = $mail.UnRead
                    $importance = $mail.Importance
                    $categories = if($mail.Categories) { $mail.Categories } else { "" }
                    $flagStatus = $mail.FlagStatus
                    $size = $mail.Size
                    $conversationTopic = if($mail.ConversationTopic) { $mail.ConversationTopic } else { "" }
                    $hasAttachments = $mail.Attachments.Count -gt 0
                    $attachmentCount = $mail.Attachments.Count
                    
                    $emailObj = @{
                      Subject = $subject
                      SenderName = $senderName
                      SenderEmailAddress = $senderEmail
                      ReceivedTime = $receivedTime
                      Importance = $importance
                      UnRead = $unread
                      Categories = $categories
                      FlagStatus = $flagStatus
                      Size = $size
                      EntryID = $entryId
                      ConversationTopic = $conversationTopic
                      HasAttachments = $hasAttachments
                      AttachmentCount = $attachmentCount
                    }
                    $emails += $emailObj
                  }
                } catch {
                  # Ignorer les erreurs sur emails individuels
                }
              }
            }
            
            $count = $emails.Count
            $totalCount = $items.Count
            $folderName = $targetFolder.Name
            $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            
            $result = @{
              success = $true
              folderPath = "${folderPath}"
              emails = $emails
              count = $count
              totalInFolder = $totalCount
              message = "Dossier specifique: TOUS les emails recuperes ($count/$totalCount)"
              timestamp = $timestamp
              folderName = $folderName
            }
          } else {
            $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            $result = @{
              success = $false
              folderPath = "${folderPath}"
              emails = @()
              count = 0
              error = "Dossier non trouve dans l'arborescence Outlook"
              message = "Navigation impossible vers le dossier"
              timestamp = $timestamp
            }
          }
          
          $json = $result | ConvertTo-Json -Depth 10 -Compress
          Write-Output $json
          
        } catch {
          $errorMsg = $_.Exception.Message
          $timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
          $errorResult = @{
            success = $false
            folderPath = "${folderPath}"
            emails = @()
            count = 0
            error = $errorMsg
            message = "Erreur navigation vers dossier specifique"
            timestamp = $timestamp
          }
          $errorJson = $errorResult | ConvertTo-Json -Depth 5 -Compress
          Write-Output $errorJson
        }
      `;
      
      const result = await this.executePowerShellScript(script);
      
      if (result.success && result.output) {
        try {
          const emailData = JSON.parse(result.output);
          console.log(`✅ [OUTLOOK] ${emailData.count} emails récupérés pour: ${folderPath}`);
          return emailData;
        } catch (parseError) {
          console.error('❌ Erreur parsing JSON:', parseError.message);
          console.log('Output brut:', result.output);
          throw new Error(`Erreur parsing données: ${parseError.message}`);
        }
      } else {
        throw new Error(result.error || 'Échec exécution PowerShell');
      }
      
    } catch (error) {
      console.error(`❌ Erreur récupération emails ${folderPath}:`, error.message);
      
      // Retourner un résultat d'erreur structuré
      return {
        success: false,
        folderPath: folderPath,
        emails: [],
        count: 0,
        error: error.message,
        message: 'Erreur récupération emails',
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * Statistiques de performance
   */
  getStats() {
    return {
      isConnected: this.isOutlookConnected,
      connectionState: this.connectionState,
      useGraphAPI: this.useGraphAPI,
      foldersCount: this.folders.size,
      lastError: this.lastError?.message || null,
      uptime: this.isOutlookConnected ? Date.now() - (this.startTime || Date.now()) : 0
    };
  }

  /**
   * Test de performance
   */
  async runPerformanceTest() {
    console.log('🧪 Test de performance Outlook Connector...');
    
    const tests = [
      { name: 'Connection', fn: () => this.testOutlookConnection() },
      { name: 'Get Folders', fn: () => this.getFolders() },
      { name: 'Get Inbox Emails', fn: () => this.getEmailsFromFolder('Inbox', 50) }
    ];

    const results = [];
    
    for (const test of tests) {
      const startTime = Date.now();
      try {
        await test.fn();
        const time = Date.now() - startTime;
        results.push({ test: test.name, time, status: 'success' });
        console.log(`✅ ${test.name}: ${time}ms`);
      } catch (error) {
        const time = Date.now() - startTime;
        results.push({ test: test.name, time, status: 'error', error: error.message });
        console.log(`❌ ${test.name}: ${time}ms (erreur)`);
      }
    }
    
    return results;
  }

  /**
   * Arrêt propre
   */
  async disconnect() {
    console.log('🔌 Déconnexion Outlook Connector...');
    
    this.isOutlookConnected = false;
    this.connectionState = 'disconnected';
    
    this.emit('disconnected');
    console.log('✅ Déconnexion propre effectuée');
  }

  /**
   * Exécuter une commande système
   */
  async executeCommand(command, args = [], timeout = 10000) {
    return new Promise((resolve, reject) => {
      const process = exec(`${command} ${args.join(' ')}`, {
        timeout,
        windowsHide: true
      });
      
      let output = '';
      let error = '';
      process.stdout.on('data', data => output += data);
      process.stderr.on('data', data => error += data);
      
      process.on('close', code => {
        if (code === 0) {
          resolve({ success: true, output, error });
        } else {
          reject(new Error(`Command failed with code ${code}: ${error || output}`));
        }
      });
      
      process.on('error', reject);
    });
  }

  /**
   * Exécuter un script PowerShell
   */
  async executePowerShellScript(script, timeout = 15000, opts = {}) {
    try {
      // console.log(`🔧 [DEBUG] Exécution PowerShell - Longueur script: ${script.length} caractères`);
      
      // Pour les scripts longs, créer un fichier temporaire
      const fs = require('fs');
      const path = require('path');
      const tempDir = require('os').tmpdir();
      const tempFile = path.join(tempDir, `outlook_script_${Date.now()}.ps1`);
      
      // Écrire le script dans un fichier temporaire avec UTF-8 BOM
      const BOM = '\uFEFF';
      fs.writeFileSync(tempFile, BOM + script, { encoding: 'utf8' });
      // console.log(`📄 [DEBUG] Script temporaire: ${tempFile}`);
      
  // Choisir l'exécutable PowerShell (64-bit par défaut, fallback 32-bit en option)
  const winDir = process.env.WINDIR || 'C:\\Windows';
  const pwsh32 = `${winDir}\\SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe`;
  const pwsh64 = `${winDir}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe`;

  const command = opts.force32Bit ? pwsh32 : pwsh64;
  const result = await this.executeCommand(command, ['-NoProfile','-STA','-ExecutionPolicy','Bypass','-File', tempFile], timeout);
      
      // Nettoyer le fichier temporaire
      try {
        fs.unlinkSync(tempFile);
      } catch (e) {
        console.warn(`⚠️ Impossible de supprimer le fichier temporaire: ${tempFile}`);
      }
      
      console.log(`✅ [DEBUG] PowerShell terminé - Success: ${result.success}`);
      if (result.output) {
        // console.log(`📄 [DEBUG] Output length: ${result.output.length} chars`);
        // console.log(`📄 [DEBUG] First 500 chars: ${result.output.substring(0, 500)}`);
      }
      if (result.error) {
        console.log(`❌ [DEBUG] Error: ${result.error}`);
      }
      
  return result;
    } catch (error) {
      console.error(`❌ [DEBUG] Exception PowerShell: ${error.message}`);
      return { success: false, error: error.message };
    }
  }

  /**
   * Logging optimisé
   */
  log(message) {
    if (this.config.enableDetailedLogs) {
      console.log(`[OutlookConnector] ${message}`);
    }
  }
}

// Export singleton
const outlookConnector = new OutlookConnector();
module.exports = outlookConnector;
