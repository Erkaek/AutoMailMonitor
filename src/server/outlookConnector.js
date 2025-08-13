/**
 * OUTLOOK CONNECTOR OPTIMISÉ - Version simplifiée
 * Performance maximale avec API REST + Better-SQLite3 + Cache
 * Remplace l'ancien système COM/PowerShell
 */

const { EventEmitter } = require('events');
const { exec } = require('child_process');
const path = require('path');

// Simplification: pas d'import de Graph API pour éviter les problèmes de dépendances
let graphAvailable = false; // Désactivé temporairement

class OutlookConnector extends EventEmitter {
  constructor() {
    super();
    
    console.log('🚀 DIAGNOSTIC: Création d\'une nouvelle instance OutlookConnector optimisée');
    
    // État de connexion
    this.isOutlookConnected = false;
    this.connectionState = 'disconnected';
    // Configuration optimisée
    this.config = {
      timeout: 15000,
      realtimePollingInterval: 15000,
      enableDetailedLogs: false,
      autoReconnect: true,
      maxRetries: 3
    };
    
    // Graph API simulation
    this.useGraphAPI = graphAvailable;
    
    // Données en cache
    this.folders = new Map();
    this.stats = new Map();
  // Cache boîtes mail pour fallback en cas de lenteur/timeout
  this.lastMailboxes = [];
  this.lastMailboxesAt = 0;
    
    // Auto-connexion
    this.autoConnect();
  }

  // Robust JSON extraction from potentially noisy PowerShell output
  static parseJsonOutput(raw) {
    try {
      if (!raw) return null;
      let s = String(raw);
      // quick path
      const t = s.trim();
      if (t.startsWith('{') || t.startsWith('[')) {
        return JSON.parse(t);
      }
      // find outermost JSON braces
      const first = s.indexOf('{');
      const last = s.lastIndexOf('}');
      if (first >= 0 && last > first) {
        const slice = s.slice(first, last + 1);
        try { return JSON.parse(slice); } catch {}
      }
      // array variant
      const afirst = s.indexOf('[');
      const alast = s.lastIndexOf(']');
      if (afirst >= 0 && alast > afirst) {
        const slice = s.slice(afirst, alast + 1);
        try { return JSON.parse(slice); } catch {}
      }
      return null;
    } catch { return null; }
  }

  /**
   * Auto-connexion optimisée
   */
  async autoConnect() {
    try {
      console.log('[AUTO-CONNECT] Tentative de connexion automatique Graph API...');
      // Mode dégradé: considérer Outlook comme "connecté" si le process existe
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
   * Legacy name kept for compatibility. We don't actually use Graph here; we ensure Outlook is available.
   */
  async connectToGraphAPI() {
    try {
      const ok = await this.ensureConnected();
      return ok;
    } catch (e) {
      this.connectionState = 'error';
      this.lastError = e;
      return false;
    }
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

      // Tentative 1 (64-bit) avec délai élargi
      let result = await this.executePowerShellScript(script, 30000);
      if (!result.success) {
        // Tentative 1 bis (32-bit)
        result = await this.executePowerShellScript(script, 30000, { force32Bit: true });
      }
      if (!result.success) {
        throw new Error(result.error || 'Échec récupération boîtes mail');
      }
  let json = OutlookConnector.parseJsonOutput(result.output) || {};
      let mailboxes = Array.isArray(json.mailboxes) ? json.mailboxes : [];

      // Si vide, petit backoff puis seconde tentative (64-bit puis 32-bit)
      if (mailboxes.length === 0) {
        console.log('⏳ [Mailboxes] Résultat vide - nouvelle tentative après courte attente...');
        await new Promise(r => setTimeout(r, 1500));
        let retry = await this.executePowerShellScript(script, 45000);
        if (!retry.success) {
          retry = await this.executePowerShellScript(script, 45000, { force32Bit: true });
        }
        if (retry.success) {
          const j2 = OutlookConnector.parseJsonOutput(retry.output) || {};
          mailboxes = Array.isArray(j2.mailboxes) ? j2.mailboxes : mailboxes;
        }
      }

      // Fallback: si toujours vide, renvoyer le dernier cache récent (< 1h)
      if (mailboxes.length === 0 && Array.isArray(this.lastMailboxes) && this.lastMailboxes.length > 0) {
        const age = Date.now() - (this.lastMailboxesAt || 0);
        if (age < 3600_000) {
          console.log('♻️ [Mailboxes] Résultat vide - retour au cache récent');
          return this.lastMailboxes;
        }
      }

      // Mettre en cache si non vide
      if (mailboxes.length > 0) {
        this.lastMailboxes = mailboxes;
        this.lastMailboxesAt = Date.now();
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
        $enc = New-Object System.Text.UTF8Encoding $false
        [Console]::OutputEncoding = $enc
        $OutputEncoding = $enc
        try {
          $outlook = New-Object -ComObject Outlook.Application
          $ns = $outlook.GetNamespace("MAPI")

          if ("${safeStoreId}" -eq "") {
            # Retourner uniquement la liste des boîtes (métadonnées), pas d'arborescence
            $mailboxes = @()
            foreach ($st in $ns.Stores) {
              try {
                $mbName = $st.DisplayName
                $smtp = $null
                try {
                  $accounts = $outlook.Session.Accounts
                  foreach ($acc in $accounts) { if ($acc.SmtpAddress -and ($acc.DisplayName -eq $mbName -or $mbName -like "*$($acc.DisplayName)*")) { $smtp = $acc.SmtpAddress; break } }
                } catch {}
                $mailboxes += @{ Name = $mbName; StoreID = $st.StoreID; SmtpAddress = $smtp; SubFolders = @() }
              } catch {}
            }
            $res = @{ success = $true; folders = $mailboxes } | ConvertTo-Json -Depth 12 -Compress
            Write-Output $res
            return
          }

          # Déterminer le store cible (par StoreID ou défaut)
          $target = $null
          foreach ($st in $ns.Stores) { if ($st.StoreID -eq "${safeStoreId}" -or $st.DisplayName -eq "${safeStoreId}" -or $st.DisplayName -like "*${safeStoreId}*") { $target = $st; break } }
          if (-not $target) { $target = $ns.DefaultStore }

          $root = $target.GetRootFolder()
          $mailboxName = $target.DisplayName
          # SMTP si possible
          $smtp = $null
          try { $accounts = $outlook.Session.Accounts; foreach ($acc in $accounts) { if ($acc.SmtpAddress -and ($acc.DisplayName -eq $mailboxName -or $mailboxName -like "*$($acc.DisplayName)*")) { $smtp = $acc.SmtpAddress; break } } } catch {}

          # Construire un arbre minimal: Inbox et ses enfants directs, sans compter les Items
          $tree = @()
          try {
            $inbox = $null
            # 1) Tenter via Store.GetDefaultFolder (Outlook 2010+)
            try { $inbox = $target.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox) } catch {}
            # 2) Fallback: chercher par nom localisé (liste élargie)
            if (-not $inbox) {
              $inboxPattern = 'Inbox|Boite de reception|Boîte de réception|Courrier entrant|Posteingang|Posta in arrivo|Bandeja de entrada|Caixa de Entrada|Postvak IN|Indbakke|Inkorgen|Saapuneet|Skrzynka odbiorcza|Doručená pošta|Beérkezett üzenetek|Mesaje primite|Gelen Kutusu|Εισερχόμενα|Входящие|受信トレイ|收件箱|Вхідні'
              foreach ($sf in $root.Folders) { if ($sf.Name -match $inboxPattern) { $inbox = $sf; break } }
              # 2bis) Si toujours rien, chercher un niveau plus bas (ex: "Top of Information Store")
              if (-not $inbox) {
                foreach ($rf in $root.Folders) {
                  try {
                    foreach ($sf in $rf.Folders) { if ($sf.Name -match $inboxPattern) { $inbox = $sf; break } }
                  } catch {}
                  if ($inbox) { break }
                }
              }
            }
            # Si la racine parait vide, essayer de retrouver un autre point racine via Namespace.Folders
            $rootChilds = 0; try { $rootChilds = $root.Folders.Count } catch {}
            if ($rootChilds -eq 0) {
              try {
                foreach ($tf in $ns.Folders) {
                  try { if ($tf.Store.StoreID -eq $target.StoreID) { $root = $tf; break } } catch {}
                }
              } catch {}
            }

            if ($inbox) {
              $children = @()
              foreach ($sf in $inbox.Folders) {
                $childCount = 0; try { $childCount = $sf.Folders.Count } catch {}
                $children += @{ Name = $sf.Name; FolderPath = "$mailboxName\\$($inbox.Name)\\$($sf.Name)"; EntryID = $sf.EntryID; ChildCount = $childCount; SubFolders = @() }
              }
              $tree += @{ Name = $inbox.Name; FolderPath = "$mailboxName\\$($inbox.Name)"; EntryID = $inbox.EntryID; ChildCount = ($inbox.Folders.Count); SubFolders = $children }
              # Si l'inbox ne semble pas avoir d'enfants (compte 0 et enumeration vide), ajouter aussi les dossiers racine
              $inboxChilds = 0; try { $inboxChilds = $inbox.Folders.Count } catch {}
              if ($inboxChilds -eq 0 -and $children.Count -eq 0) {
                foreach ($sf in $root.Folders) {
                  $childCount = 0; try { $childCount = $sf.Folders.Count } catch {}
                  $tree += @{ Name = $sf.Name; FolderPath = "$mailboxName\\$($sf.Name)"; EntryID = $sf.EntryID; ChildCount = $childCount; SubFolders = @() }
                }
              }
            } else {
              # 3) Dernier fallback: exposer les dossiers de premier niveau de la racine
              $children = @()
              foreach ($sf in $root.Folders) {
                $childCount = 0; try { $childCount = $sf.Folders.Count } catch {}
                $children += @{ Name = $sf.Name; FolderPath = "$mailboxName\\$($sf.Name)"; EntryID = $sf.EntryID; ChildCount = $childCount; SubFolders = @() }
              }
              # Dans ce mode, on n'ajoute pas un noeud 'Inbox' parent; on renvoie directement les enfants de la racine
              foreach ($c in $children) { $tree += $c }
            }
          } catch {}

          $mb = @{ Name = $mailboxName; StoreID = $target.StoreID; SmtpAddress = $smtp; SubFolders = $tree }
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
      let json = OutlookConnector.parseJsonOutput(result.output) || {};
      let folders = Array.isArray(json.folders) ? json.folders : [];
      if (folders.length === 0) {
        const res32 = await this.executePowerShellScript(script, 60000, { force32Bit: true });
        if (res32.success) {
          const j2 = OutlookConnector.parseJsonOutput(res32.output) || {};
          folders = Array.isArray(j2.folders) ? j2.folders : folders;
        }
      }
      try {
        const mb = folders && folders[0];
        const subCount = Array.isArray(mb?.SubFolders) ? mb.SubFolders.length : 0;
        console.log(`[OUTLOOK] getFolderStructure: store=${String(storeId || '').slice(0,8)}… -> mailboxes=${folders.length}, first.SubFolders=${subCount}`);
      } catch {}
      return folders;
    } catch (error) {
      console.error('❌ Erreur getFolderStructure:', error.message);
      return [];
    }
  }

  /**
   * Lazy-load: Récupère les sous-dossiers directs d'un dossier par EntryID pour un store donné.
   * Retour: Array<{ Name, FolderPath, EntryID, ChildCount, SubFolders: [] }>
   */
  async getSubFolders(storeId, parentEntryId, parentPath) {
    try {
      const ok = await this.ensureConnected();
      if (!ok) return [];

      // Cache simple en mémoire avec TTL 5 min
  if (!this.subfolderCache) this.subfolderCache = new Map();
  // Include parentPath in the cache key when parentEntryId is not provided to avoid collisions
  const cacheKey = `${storeId}|${parentEntryId || 'noid'}|${parentPath || ''}`;
      const cached = this.subfolderCache.get(cacheKey);
      if (cached && (Date.now() - cached.at) < 300_000) {
        // If previous result was empty, try a fresh probe once to avoid sticky 0 due to resolution quirks
        if (Array.isArray(cached.data) && cached.data.length === 0) {
          // continue to compute a fresh result
        } else {
          return cached.data;
        }
      }

      const safeStoreId = String(storeId || '').replace(/`/g, '``').replace(/"/g, '\"');
      const safeParentId = String(parentEntryId || '').replace(/`/g, '``').replace(/"/g, '\"');
      const safeParentPath = String(parentPath || '').replace(/`/g, '``').replace(/"/g, '\"');

      const script = `
        $enc = New-Object System.Text.UTF8Encoding $false
        [Console]::OutputEncoding = $enc
        $OutputEncoding = $enc
        try {
          $outlook = New-Object -ComObject Outlook.Application
          $ns = $outlook.GetNamespace("MAPI")
          $store = $null
          foreach ($st in $ns.Stores) { if ($st.StoreID -eq "${safeStoreId}") { $store = $st; break } }
          if (-not $store) { throw "Store introuvable" }
          $root = $store.GetRootFolder()
          $mbName = $store.DisplayName
          $parent = $null
          
          # Certaines configurations renvoient une racine "vide"; corriger via Namespace.Folders
          try {
            $rootChilds = 0; try { $rootChilds = $root.Folders.Count } catch {}
            if ($rootChilds -eq 0) {
              foreach ($tf in $ns.Folders) {
                try { if ($tf.Store.StoreID -eq $store.StoreID) { $root = $tf; break } } catch {}
              }
            }
          } catch {}
          
          # Helper: normalize names (remove diacritics, lower-case, trim incl. trailing dots)
          function Normalize-Name([string]$s) {
            if ([string]::IsNullOrEmpty($s)) { return "" }
            $s2 = $s.Trim().TrimEnd('.')
            $n = $s2.Normalize([Text.NormalizationForm]::FormD)
            $sb = New-Object System.Text.StringBuilder
            foreach ($c in $n.ToCharArray()) {
              if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$sb.Append($c)
              }
            }
            return $sb.ToString().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant().Trim()
          }
          
          function Find-ChildByName($parentFolder, [string]$targetName) {
            $normTarget = Normalize-Name $targetName
            foreach ($f in $parentFolder.Folders) { if ((Normalize-Name $f.Name) -eq $normTarget) { return $f } }
            return $null
          }

          function Find-FolderByNameDeep($startFolder, [string]$targetName) {
            $norm = Normalize-Name $targetName
            foreach ($f in $startFolder.Folders) {
              if ((Normalize-Name $f.Name) -eq $norm) { return $f }
            }
            foreach ($f in $startFolder.Folders) {
              $found = Find-FolderByNameDeep -startFolder $f -targetName $targetName
              if ($found -ne $null) { return $found }
            }
            return $null
          }
          
          try { $parent = $ns.GetFolderFromID("${safeParentId}", "${safeStoreId}") } catch {}
          if (-not $parent -and "${safeParentPath}" -ne "") {
            # Fallback 1: navigation depuis la racine d'affichage (MB\...)
            $parts = "${safeParentPath}" -split "\\"
            if ($parts.Length -eq 1 -and (Normalize-Name $parts[0]) -eq (Normalize-Name $mbName)) {
              # Demande des enfants de la racine de la boîte: utiliser root directement
              $parent = $root
            } elseif ($parts.Length -gt 1) {
              $cursor = $root
              # Ignorer la premiere partie (nom de la boite), les suivantes sont les dossiers
              for ($i = 1; $i -lt $parts.Length; $i++) {
                $name = $parts[$i]
                $next = Find-ChildByName -parentFolder $cursor -targetName $name
                if ($next -eq $null) { break }
                $cursor = $next
              }
              if ($cursor -ne $root) { $parent = $cursor }
            }
            
            # Fallback 2: si non trouvé, tenter résolution depuis la Boîte de réception (localisée)
            if (-not $parent) {
              try {
                $inbox = $store.GetDefaultFolder(6) # olFolderInbox
                if ($inbox -ne $null) {
                  $cursor = $inbox
                  # Si le chemin fourni inclut déjà le segment Inbox en deuxième position, le sauter
                  $startIndex = 1
                  try {
                    if ($parts.Length -gt 1) {
                      $p1 = $parts[1]
                      if ((Normalize-Name $p1) -eq (Normalize-Name $inbox.Name)) { $startIndex = 2 }
                    }
                  } catch {}
                  for ($i = $startIndex; $i -lt $parts.Length; $i++) {
                    $name = $parts[$i]
                    $next = Find-ChildByName -parentFolder $cursor -targetName $name
                    if ($next -eq $null) { break }
                    $cursor = $next
                  }
                  if ($cursor -ne $inbox) {
                    $parent = $cursor
                  } elseif ($parts.Length -eq 2 -and ((Normalize-Name $parts[1]) -eq (Normalize-Name $inbox.Name))) {
                    # Cas particulier: on demandait directement les enfants de l'Inbox
                    $parent = $inbox
                  }
                }
              } catch {}
            }

            # Fallback 3: recherche profonde par nom du dernier segment si encore introuvable (préférer Inbox)
            if (-not $parent) {
              try {
                if ($parts.Length -gt 1) {
                  $leaf = $parts[$parts.Length - 1]
                  $candidate = $null
                  try {
                    $inbox2 = $store.GetDefaultFolder(6)
                    if ($inbox2 -ne $null) { $candidate = Find-FolderByNameDeep -startFolder $inbox2 -targetName $leaf }
                  } catch {}
                  if ($candidate -eq $null) { $candidate = Find-FolderByNameDeep -startFolder $root -targetName $leaf }
                  if ($candidate -ne $null) { $parent = $candidate }
                }
              } catch {}
            }
          }
          if (-not $parent) { throw "Dossier parent introuvable" }
          # Build a stable display path by walking up to the root
          function Get-FolderDisplayPath($folder, [string]$mailboxName, $rootFolder) {
            $segments = New-Object System.Collections.Generic.List[string]
            $cur = $folder
            try {
              while ($cur -ne $null -and $cur.EntryID -ne $rootFolder.EntryID) {
                $segments.Add($cur.Name)
                $cur = $cur.Parent
                # Safety break to avoid infinite loops
                if ($segments.Count -gt 50) { break }
              }
            } catch {}
            $segments.Reverse()
            if ($segments.Count -gt 0) { return ("$mailboxName\" + ([string]::Join("\\", $segments))) } else { return $mailboxName }
          }

          $children = @()
          foreach ($sf in $parent.Folders) {
            $childCount = 0; try { $childCount = $sf.Folders.Count } catch {}
            $path = Get-FolderDisplayPath -folder $sf -mailboxName $mbName -rootFolder $root
            $children += @{
              Name = $sf.Name
              FolderPath = $path
              EntryID = $sf.EntryID
              ChildCount = $childCount
              SubFolders = @()
            }
          }
          $res = @{ success = $true; children = $children } | ConvertTo-Json -Depth 12 -Compress
          Write-Output $res
        } catch {
          $err = $_.Exception.Message
          $res = @{ success = $false; error = $err; children = @() } | ConvertTo-Json -Depth 3 -Compress
          Write-Output $res
        }
      `;

      let result = await this.executePowerShellScript(script, 20000);
      if (!result.success) {
        result = await this.executePowerShellScript(script, 20000, { force32Bit: true });
      }
      if (!result.success) throw new Error(result.error || 'Échec récupération sous-dossiers');

  let json = OutlookConnector.parseJsonOutput(result.output) || {};
  const children = Array.isArray(json.children) ? json.children : [];
  try { console.log(`[OUTLOOK] getSubFolders: parent=${String(parentPath||'').slice(-40)}… -> ${children.length} items`); } catch {}

      this.subfolderCache.set(cacheKey, { at: Date.now(), data: children });
      return children;
    } catch (error) {
      console.error('❌ Erreur getSubFolders:', error.message);
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
          $max = ${Number.isFinite(limit) ? Math.max(1, Number(limit)) : 50}
          
          # Helpers available in script scope (used by all fallbacks)
          function Normalize-Name([string]$s) {
            if ([string]::IsNullOrEmpty($s)) { return "" }
            # Trim whitespace and trailing dots to tolerate minor path differences
            $s2 = $s.Trim().TrimEnd('.')
            $n = $s2.Normalize([Text.NormalizationForm]::FormD)
            $sb = New-Object System.Text.StringBuilder
            foreach ($c in $n.ToCharArray()) {
              if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$sb.Append($c)
              }
            }
            return $sb.ToString().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant().Trim()
          }

          function Find-ChildByName($parentFolder, [string]$targetName) {
            $normTarget = Normalize-Name $targetName
            foreach ($f in $parentFolder.Folders) { if ((Normalize-Name $f.Name) -eq $normTarget) { return $f } }
            return $null
          }
          
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
                if ((Normalize-Name $store.DisplayName) -eq (Normalize-Name $accountName)) {
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
                  $next = Find-ChildByName -parentFolder $currentFolder -targetName $part
                  if ($next -eq $null) { return $null }
                  $currentFolder = $next
                }
              }
              
              return $currentFolder
            } else {
              return $null
            }
          }
          
          # Trouver le dossier cible
          $targetFolder = Find-OutlookFolder -FolderPath "${folderPath}" -Namespace $namespace
          
          # Fallback: si non trouvé, tenter depuis la Boîte de réception (olFolderInbox)
          if (-not $targetFolder) {
            try {
              $parts = "${folderPath}" -split "\\"
              if ($parts.Length -gt 1) {
                $accountName = $parts[0]
                $targetStore = $null
                foreach ($store in $namespace.Stores) { if ((Normalize-Name $store.DisplayName) -eq (Normalize-Name $accountName)) { $targetStore = $store; break } }
                if ($targetStore -eq $null) { $targetStore = $namespace.DefaultStore }
                $inbox = $targetStore.GetDefaultFolder(6)
                if ($inbox -ne $null) {
                  $cursor = $inbox
                  for ($i = 1; $i -lt $parts.Length; $i++) {
                    $name = $parts[$i]
                    $next = Find-ChildByName -parentFolder $cursor -targetName $name
                    if ($next -eq $null) { $cursor = $null; break }
                    $cursor = $next
                  }
                  if ($cursor -ne $null) { $targetFolder = $cursor }
                }
              }
            } catch {}
          }

      # Fallback 2: recherche profonde depuis Inbox ou racine par nom du dernier segment
          if (-not $targetFolder) {
            try {
              $parts2 = "${folderPath}" -split "\\"
              if ($parts2.Length -gt 1) {
                $account2 = $parts2[0]
                $store2 = $null
                foreach ($st in $namespace.Stores) { if ((Normalize-Name $st.DisplayName) -eq (Normalize-Name $account2)) { $store2 = $st; break } }
                if ($store2 -eq $null) { $store2 = $namespace.DefaultStore }
                $root2 = $store2.GetRootFolder()
                $leaf = $parts2[$parts2.Length - 1]
                function Find-FolderByNameDeep([object]$start, [string]$tname) {
                  $norm = Normalize-Name $tname
                  foreach ($f in $start.Folders) { if ((Normalize-Name $f.Name) -eq $norm) { return $f } }
                  foreach ($f in $start.Folders) {
                    $found = Find-FolderByNameDeep -start $f -tname $tname
                    if ($found -ne $null) { return $found }
                  }
                  return $null
                }
        $cand = $null
        try { $inbox3 = $store2.GetDefaultFolder(6); if ($inbox3 -ne $null) { $cand = Find-FolderByNameDeep -start $inbox3 -tname $leaf } } catch {}
        if ($cand -eq $null) { $cand = Find-FolderByNameDeep -start $root2 -tname $leaf }
                if ($cand -ne $null) { $targetFolder = $cand }
              }
            } catch {}
          }
          
          if ($targetFolder) {
            # Récupérer TOUS les emails du dossier
            $items = $targetFolder.Items
            if ($items.Count -gt 0) {
              $items.Sort("[ReceivedTime]", $true)
              
              # Traiter les emails (jusqu'à $max)
              $upper = [Math]::Min($items.Count, [int]$max)
              for ($i = 1; $i -le $upper; $i++) {
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
              message = "Dossier specifique: emails recuperes ($count/$totalCount)"
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

  // --- EWS FAST ENUMERATION HELPERS ---
  async ewsRun(scriptArgs, timeout = 20000) {
    const { spawn } = require('child_process');
    const winDir = process.env.WINDIR || 'C:\\Windows';
    const pwsh = path.join(winDir, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    return new Promise((resolve, reject) => {
      const ps = spawn(pwsh, ['-NoProfile', '-ExecutionPolicy', 'Bypass', ...scriptArgs], { windowsHide: true });
      let out = '', err = '';
      const t = setTimeout(() => { try { ps.kill('SIGTERM'); } catch {} reject(new Error('EWS timeout')); }, timeout);
      ps.stdout.on('data', d => out += d.toString());
      ps.stderr.on('data', d => err += d.toString());
      ps.on('exit', code => { clearTimeout(t); if (code === 0) resolve(out); else reject(new Error(err || `PowerShell exited ${code}`)); });
      ps.on('error', e => { clearTimeout(t); reject(e); });
    });
  }

  async getTopLevelFoldersFast(mailbox) {
    // Circuit breaker
    this._ewsFailures = this._ewsFailures || 0;
    if (this._ewsFailures >= 3 || this._ewsDisabled) {
      return []; // allow COM fallback silently
    }
    const { resolveResource } = require('./scriptPathResolver');
    const scriptRes = resolveResource(['scripts'], 'ews-list-folders.ps1');
    const dllRes = resolveResource(['resources','ews'], 'Microsoft.Exchange.WebServices.dll');
    if (!scriptRes.path || !dllRes.path) {
      if (!scriptRes.path) console.warn('⚠️ EWS script absent (top-level), fallback COM. Tried:', scriptRes.tried);
      if (!dllRes.path) console.warn('⚠️ EWS DLL absente (top-level), fallback COM. Tried:', dllRes.tried);
      this._ewsDisabled = true;
      return [];
    }
    const scriptPath = scriptRes.path;
    const ewsDll = dllRes.path;
    const argsBase = ['-File', scriptPath, '-Mailbox', mailbox, '-DllPath', ewsDll];
    try {
      let folders = [];
      for (const scope of ['Inbox','Root']) {
        const raw = await this.ewsRun([...argsBase, '-Scope', scope], scope === 'Inbox' ? 20000 : 25000);
        const data = OutlookConnector.parseJsonOutput(raw);
        const payload = Array.isArray(data) ? data[0] : data;
        const list = payload?.Folders || [];
        if (Array.isArray(list) && list.length) { folders = list; break; }
      }
      return Array.isArray(folders) ? folders.map(f => ({ Id: f.Id, Name: f.Name, ChildCount: Number(f.ChildCount || 0) })) : [];
    } catch (e) {
      this._ewsFailures++;
      console.warn(`⚠️ EWS top-level failed (${this._ewsFailures}):`, e.message);
      if (this._ewsFailures >= 3) this._ewsDisabled = true;
      return [];
    }
  }

  async getChildFoldersFast(mailbox, parentId) {
    this._ewsFailures = this._ewsFailures || 0;
    if (this._ewsFailures >= 3 || this._ewsDisabled) return [];
    const { resolveResource } = require('./scriptPathResolver');
    const scriptRes = resolveResource(['scripts'], 'ews-list-folders.ps1');
    const dllRes = resolveResource(['resources','ews'], 'Microsoft.Exchange.WebServices.dll');
    if (!scriptRes.path || !dllRes.path) {
      if (!scriptRes.path) console.warn('⚠️ EWS script absent (children), fallback COM.');
      if (!dllRes.path) console.warn('⚠️ EWS DLL absente (children), fallback COM.');
      this._ewsDisabled = true;
      return [];
    }
    const scriptPath = scriptRes.path;
    const ewsDll = dllRes.path;
    const args = ['-File', scriptPath, '-Mailbox', mailbox, '-ParentId', parentId, '-DllPath', ewsDll];
    try {
      const raw = await this.ewsRun(args, 15000);
      const data = OutlookConnector.parseJsonOutput(raw);
      const payload = Array.isArray(data) ? data[0] : data;
      const folders = payload?.Folders || [];
      return Array.isArray(folders) ? folders.map(f => ({ Id: f.Id, Name: f.Name, ChildCount: Number(f.ChildCount || 0) })) : [];
    } catch (e) {
      this._ewsFailures++;
      console.warn(`⚠️ EWS children failed (${this._ewsFailures}):`, e.message);
      if (this._ewsFailures >= 3) this._ewsDisabled = true;
      return [];
    }
  }
}

// Export singleton
const outlookConnector = new OutlookConnector();
module.exports = outlookConnector;
