/**
 * SERVICE D'ÉCOUTE DES ÉVÉNEMENTS COM OUTLOOK
 * Gère l'écoute en temps réel des événements Outlook (ItemAdd, ItemChange, NewMail)
 * Optimisé pour une surveillance efficace sans polling
 */

const { EventEmitter } = require('events');
const { spawn } = require('child_process');
const path = require('path');

class OutlookEventsService extends EventEmitter {
  constructor() {
    super();
    this.isListening = false;
    this.outlookProcess = null;
    this.monitoredFolders = new Map();
    this.lastEventTime = null;
    this.eventBuffer = [];
    this.bufferTimeout = null;
    
    // Configuration
    this.config = {
      bufferDelay: 500, // Délai pour grouper les événements (ms)
      maxBufferSize: 50, // Taille max du buffer d'événements
      reconnectDelay: 5000, // Délai avant reconnexion (ms)
      maxReconnectAttempts: 10
    };
    
    this.reconnectAttempts = 0;
  }

  /**
   * Démarre l'écoute des événements COM Outlook avec polling intelligent
   */
  async startListening(folders = []) {
    try {
      if (this.isListening) {
        console.log('⚠️ [OutlookEvents] Écoute déjà active');
        return { success: false, message: 'Écoute déjà active' };
      }

      console.log('🔔 [OutlookEvents] Démarrage de l\'écoute des événements COM...');
      
      // Mettre à jour la liste des dossiers surveillés
      this.updateMonitoredFolders(folders);
      
      // Démarrer le processus PowerShell d'écoute COM avec méthode hybride
      await this.startHybridListener();
      
      this.isListening = true;
      this.reconnectAttempts = 0;
      
      console.log(`✅ [OutlookEvents] Écoute COM active sur ${this.monitoredFolders.size} dossiers`);
      this.emit('listening-started', { folders: this.monitoredFolders.size });
      
      return { success: true, message: `Écoute COM démarrée sur ${this.monitoredFolders.size} dossiers` };
      
    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur démarrage écoute:', error);
      this.isListening = false;
      throw error;
    }
  }

  /**
   * Arrête l'écoute des événements COM
   */
  async stopListening() {
    try {
      if (!this.isListening) {
        console.log('⚠️ [OutlookEvents] Écoute non active');
        return { success: false, message: 'Écoute non active' };
      }

      console.log('🛑 [OutlookEvents] Arrêt de l\'écoute des événements COM...');
      
      this.isListening = false;
      
      // Arrêter le processus PowerShell
      if (this.outlookProcess) {
        this.outlookProcess.kill('SIGTERM');
        this.outlookProcess = null;
      }
      
      // Nettoyer les buffers
      this.clearEventBuffer();
      
      console.log('✅ [OutlookEvents] Écoute COM arrêtée');
      this.emit('listening-stopped');
      
      return { success: true, message: 'Écoute COM arrêtée' };
      
    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur arrêt écoute:', error);
      throw error;
    }
  }

  /**
   * Met à jour la liste des dossiers surveillés
   */
  updateMonitoredFolders(folders) {
    this.monitoredFolders.clear();
    
    if (Array.isArray(folders)) {
      folders.forEach(item => {
        const isObj = item && typeof item === 'object';
        const path = isObj ? (item.path || item.folderPath || '') : item;
        const entryId = isObj ? (item.entryId || item.EntryID || item.EntryId || '') : '';
        const storeId = isObj ? (item.storeId || item.StoreID || item.StoreId || '') : '';
        const storeName = isObj ? (item.storeName || item.StoreName || '') : '';
        const key = entryId || path;
        if (!key) return;
        this.monitoredFolders.set(key, {
          path,
          entryId,
          storeId,
          storeName,
          lastEventTime: null,
          eventCount: 0
        });
      });
    }
    
    console.log(`📁 [OutlookEvents] ${this.monitoredFolders.size} dossiers configurés pour l'écoute`);
  }

  /**
   * Démarre le processus PowerShell d'écoute COM
   */
  async startCOMListener() {
    return new Promise((resolve, reject) => {
      try {
        // Créer le script PowerShell d'écoute COM
        const scriptPath = this.createCOMListenerScript();
        
        // Démarrer le processus PowerShell
        this.outlookProcess = spawn('powershell.exe', [
          '-ExecutionPolicy', 'Bypass',
          '-NoProfile',
          '-WindowStyle', 'Hidden',
          '-File', scriptPath
        ], {
          stdio: ['pipe', 'pipe', 'pipe'],
          windowsHide: true
        });

        // Gérer la sortie du processus
        this.outlookProcess.stdout.on('data', (data) => {
          this.handleCOMEvent(data.toString());
        });

        this.outlookProcess.stderr.on('data', (data) => {
          console.error('[ERROR] [OutlookEvents] Erreur PowerShell:', data.toString());
        });

        this.outlookProcess.on('close', (code) => {
          console.log(`[INFO] [OutlookEvents] Processus PowerShell ferme (code: ${code})`);
          this.handleProcessClose(code);
        });

        this.outlookProcess.on('error', (error) => {
          console.error('❌ [OutlookEvents] Erreur processus PowerShell:', error);
          reject(error);
        });

        // Attendre que le processus soit prêt
        setTimeout(() => {
          if (this.outlookProcess && !this.outlookProcess.killed) {
            resolve();
          } else {
            reject(new Error('Impossible de démarrer le processus d\'écoute COM'));
          }
        }, 2000);

      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Démarre l'écoute hybride (COM + polling intelligent)
   */
  async startHybridListener() {
    try {
      console.log('🔔 [OutlookEvents] Démarrage du listener hybride...');
      
      // Créer le script PowerShell de polling intelligent
      const scriptPath = this.createIntelligentPollingScript();
      
      // Démarrer le processus PowerShell
      this.outlookProcess = spawn('powershell', [
        '-NoProfile',
        '-NonInteractive', 
        '-ExecutionPolicy', 'Bypass',
        '-File', scriptPath
      ], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
        encoding: 'utf8'
      });

      // Gestion des données de sortie
      this.outlookProcess.stdout.on('data', (data) => {
        const output = data.toString('utf8');
        this.handlePollingEvent(output);
      });

      this.outlookProcess.stderr.on('data', (data) => {
        const error = data.toString('utf8');
        if (!error.includes('VERBOSE') && error.trim()) {
          console.error('⚠️ [OutlookEvents] PowerShell stderr:', error);
        }
      });

      this.outlookProcess.on('close', (code) => {
        console.log(`🛑 [OutlookEvents] Processus PowerShell fermé avec code: ${code}`);
        if (this.isListening && code !== 0) {
          this.handleProcessCrash();
        }
      });

      this.outlookProcess.on('error', (error) => {
        console.error('❌ [OutlookEvents] Erreur processus PowerShell:', error);
        this.handleProcessCrash();
      });

      // Attendre que le processus démarre
      await this.waitForProcessStart();
      
      console.log('✅ [OutlookEvents] Listener hybride démarré');

    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur démarrage listener hybride:', error);
      throw error;
    }
  }

  /**
   * Crée le script PowerShell d'écoute des événements COM
   */
  createCOMListenerScript() {
    // Récupérer la liste des dossiers surveillés
      const monitoredFolders = Array.from(this.monitoredFolders.values());
    
    const scriptContent = `
# Script d'écoute des événements COM Outlook - SURVEILLANCE DES DOSSIERS SPÉCIFIQUES
# Surveille les événements ItemAdd, ItemChange dans les dossiers configurés

$ErrorActionPreference = "SilentlyContinue"

try {
    Write-Host "[INFO] [COM] Demarrage ecoute evenements Outlook..."
    
    # Connexion à Outlook via COM
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Host "[OK] [COM] Connexion Outlook etablie"
    
    # Liste des dossiers à surveiller - RÉCUPÉRÉS DEPUIS LA CONFIGURATION
      $foldersToMonitor = @(
        ${monitoredFolders.map(f => `@{ Path="${(f.path||'').replace(/"/g,'\"')}"; EntryId="${(f.entryId||'').replace(/"/g,'\"')}"; StoreId="${(f.storeId||'').replace(/"/g,'\"')}"; StoreName="${(f.storeName||'').replace(/"/g,'\"')}" }`).join(',\n      ')}
      )
    
    Write-Host "[FOLDER] [COM] Configuration surveillance pour $($foldersToMonitor.Length) dossiers"
    
    # Configuration des événements pour chaque dossier spécifique
    foreach ($folderInfo in $foldersToMonitor) {
      try {
        $folderPath = $folderInfo.Path
        $entryId = $folderInfo.EntryId
        $storeId = $folderInfo.StoreId
        Write-Host "[COM] Configuration evenements pour: $($folderPath) (EntryId=$entryId)"

        # Navigation prioritaire via EntryId + StoreId
        $folder = $null
        if ($entryId -and $entryId -ne "") {
          try {
            if ($storeId -and $storeId -ne "") {
              $folder = $namespace.GetFolderFromID($entryId, $storeId)
            } else {
              $folder = $namespace.GetFolderFromID($entryId)
            }
          } catch {}
        }

        # Fallback: navigation via chemin affiché
        if (-not $folder -and $folderPath) {
          $pathParts = $folderPath -split "\\"
          
          if ($pathParts.Length -ge 2) {
            $storeName = $pathParts[0]
            $folderNames = $pathParts[1..($pathParts.Length-1)]
            
            # Trouver le store
            $targetStore = $null
            foreach ($store in $namespace.Stores) {
              if ($store.DisplayName -eq $storeName -or $store.StoreID -eq $storeId) {
                /**
                 * Traite les événements COM reçus
                 */
                handleCOMEvent(data) {
                ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                FolderPath = $item.Parent.FolderPath
                IsRead = $item.UnRead -eq $false
                Size = $item.Size
                Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
              }
              
              $json = $eventData | ConvertTo-Json -Compress
              Write-Host "EVENT_DATA:$json"
            }
          }
          
          Register-ObjectEvent -InputObject $folder.Items -EventName "ItemChange" -Action {
            param($item)
            if ($item.Class -eq 43) {  # olMail
              $eventData = @{
                Type = "ItemChange"
                EntryID = $item.EntryID
                Subject = $item.Subject
                SenderName = $item.SenderName
                ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                FolderPath = $item.Parent.FolderPath
                IsRead = $item.UnRead -eq $false
                Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
              }
              
              $json = $eventData | ConvertTo-Json -Compress
              Write-Host "EVENT_DATA:$json"
            }
          }
          
          Write-Host "[OK] [COM] Evenements configures pour: $($folder.Name)"
        } else {
          Write-Host "[ERROR] [COM] Impossible de trouver le dossier: $folderPath"
        }
        
      } catch {
        Write-Host "[ERROR] [COM] Erreur configuration evenements pour $folderPath : $($_.Exception.Message)"
      }
    }
    
    Write-Host "[OK] [COM] Listeners evenements configures"
    Write-Host "[LISTEN] [COM] Ecoute en cours... (Ctrl+C pour arreter)"
    
    # Boucle infinie pour maintenir l'écoute
    while ($true) {
        Start-Sleep -Seconds 1
        
        # Ping de vie toutes les 30 secondes
        if ((Get-Date).Second % 30 -eq 0) {
            Write-Host "HEARTBEAT:$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        }
    }
    
} catch {
    Write-Host "[ERROR] [COM] Erreur: $($_.Exception.Message)"
    exit 1
}
`;

    const scriptPath = path.join(__dirname, '..', '..', 'temp', 'outlook-events-listener.ps1');
    const fs = require('fs');
    
    // Créer le dossier temp s'il n'existe pas
    const tempDir = path.dirname(scriptPath);
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    // Écrire le script avec encodage UTF-8 BOM pour PowerShell
    const scriptBuffer = Buffer.concat([
      Buffer.from([0xEF, 0xBB, 0xBF]), // UTF-8 BOM
      Buffer.from(scriptContent, 'utf8')
    ]);
    fs.writeFileSync(scriptPath, scriptBuffer);
    
    return scriptPath;
  }

  /**
    foreach ($folderInfo in $foldersToMonitor) {
   * Traite les événements COM reçus
        $folderPath = $folderInfo.Path
        $entryId = $folderInfo.EntryId
        $storeId = $folderInfo.StoreId
        Write-Host "[COM] Configuration evenements pour: $($folderPath) (EntryId=$entryId)"

        # Navigation prioritaire via EntryId + StoreId
        $folder = $null
        if ($entryId -and $entryId -ne "") {
          try {
            if ($storeId -and $storeId -ne "") {
              $folder = $namespace.GetFolderFromID($entryId, $storeId)
            } else {
              $folder = $namespace.GetFolderFromID($entryId)
            }
          } catch {}
        }

        # Fallback: navigation via chemin affiché
        if (-not $folder -and $folderPath) {
          $pathParts = $folderPath -split "\\\\"
          if ($pathParts.Length -ge 2) {
            $storeName = $pathParts[0]
            $folderNames = $pathParts[1..($pathParts.Length-1)]
            $targetStore = $null
            foreach ($store in $namespace.Stores) {
              if ($store.DisplayName -eq $storeName -or $store.StoreID -eq $storeId) {
                $targetStore = $store
                break
              }
            }
            if ($targetStore) {
              if ($folderNames[0] -like "*reception*" -or $folderNames[0] -like "*Inbox*" -or $folderNames[0] -eq "Boite de reception") {
                $currentFolder = $namespace.GetDefaultFolder(6)
                $remainingFolders = $folderNames[1..($folderNames.Length-1)]
                foreach ($folderName in $remainingFolders) {
                  $found = $false
                  foreach ($subfolder in $currentFolder.Folders) {
                    if ($subfolder.Name -eq $folderName) {
                      $currentFolder = $subfolder
                      $found = $true
                      break
                    }
                  }
                  if (-not $found) { throw "Dossier non trouve: $folderName" }
                }
              } else {
                $currentFolder = $targetStore.GetRootFolder()
                foreach ($folderName in $folderNames) {
                  $found = $false
                  foreach ($subfolder in $currentFolder.Folders) {
                    if ($subfolder.Name -eq $folderName) { $currentFolder = $subfolder; $found = $true; break }
                  }
                  if (-not $found) { throw "Dossier non trouve: $folderName" }
                }
              }
              $folder = $currentFolder
            }
          }
        }
  handleCOMEvent(data) {
    try {
      const lines = data.split('\n').filter(line => line.trim());
      
      for (const line of lines) {
        if (line.startsWith('EVENT_DATA:')) {
          const jsonData = line.replace('EVENT_DATA:', '');
          const eventData = JSON.parse(jsonData);
          
          // Vérifier si l'événement concerne un dossier surveillé
          if (this.isMonitoredFolder(eventData.FolderPath)) {
            this.processOutlookEvent(eventData);
          }
        } else if (line.startsWith('HEARTBEAT:')) {
          this.lastEventTime = new Date();
          // Heartbeat silencieux pour vérifier que l'écoute fonctionne
        } else if (line.includes('[COM]')) {
          console.log(`🔔 [OutlookEvents] ${line}`);
        }
      }
    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur traitement événement:', error);
    }
  }

  /**
   * Vérifie si un dossier est surveillé
   */
  isMonitoredFolder(folderPath) {
    if (!folderPath) return false;
    
    // Vérification exacte
    if (this.monitoredFolders.has(folderPath)) {
      return true;
    }
    
    // Vérification par inclusion (pour gérer les variations de format)
    for (const [monitoredPath] of this.monitoredFolders) {
      if (folderPath.includes(monitoredPath) || monitoredPath.includes(folderPath)) {
        return true;
      }
    }
    
    return false;
  }

  /**
   * Traite un événement Outlook spécifique
   */
  processOutlookEvent(eventData) {
    try {
      console.log(`📧 [OutlookEvents] Événement ${eventData.Type}: ${eventData.Subject} (${eventData.FolderPath})`);
      
      // Mettre à jour les statistiques du dossier
      if (this.monitoredFolders.has(eventData.FolderPath)) {
        const folder = this.monitoredFolders.get(eventData.FolderPath);
        folder.lastEventTime = new Date();
        folder.eventCount++;
      }
      
      // Ajouter au buffer d'événements
      this.addToEventBuffer(eventData);
      
      // Émettre immédiatement l'événement pour mise à jour temps réel
      if (eventData.Type === 'ItemAdd') {
        this.emit('newEmail', {
          entryId: eventData.EntryID,
          subject: eventData.Subject,
          senderName: eventData.SenderName,
          senderEmail: eventData.SenderEmailAddress,
          receivedTime: eventData.ReceivedTime,
          folderPath: eventData.FolderPath,
          isRead: eventData.IsRead,
          size: eventData.Size
        });
      } else if (eventData.Type === 'ItemChange') {
        this.emit('emailChanged', {
          entryId: eventData.EntryID,
          subject: eventData.Subject,
          folderPath: eventData.FolderPath,
          isRead: eventData.IsRead,
          changeTime: eventData.Timestamp
        });
      }
      
    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur traitement événement Outlook:', error);
    }
  }

  /**
   * Ajoute un événement au buffer pour traitement groupé
   */
  addToEventBuffer(eventData) {
    this.eventBuffer.push({
      ...eventData,
      processedTime: new Date()
    });
    
    // Limiter la taille du buffer
    if (this.eventBuffer.length > this.config.maxBufferSize) {
      this.eventBuffer = this.eventBuffer.slice(-this.config.maxBufferSize);
    }
    
    // Programmer le traitement du buffer
    this.scheduleBufferProcessing();
  }

  /**
   * Programme le traitement groupé des événements
   */
  scheduleBufferProcessing() {
    if (this.bufferTimeout) {
      clearTimeout(this.bufferTimeout);
    }
    
    this.bufferTimeout = setTimeout(() => {
      this.processEventBuffer();
    }, this.config.bufferDelay);
  }

  /**
   * Traite le buffer d'événements groupés
   */
  processEventBuffer() {
    if (this.eventBuffer.length === 0) return;
    
    try {
      console.log(`📊 [OutlookEvents] Traitement de ${this.eventBuffer.length} événements groupés`);
      
      // Grouper les événements par type et dossier
      const groupedEvents = this.groupEventsByFolder(this.eventBuffer);
      
      // Émettre les événements groupés
      this.emit('eventsProcessed', {
        totalEvents: this.eventBuffer.length,
        groupedEvents: groupedEvents,
        timestamp: new Date()
      });
      
      // Vider le buffer
      this.eventBuffer = [];
      
    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur traitement buffer:', error);
    }
  }

  /**
   * Crée le script PowerShell de polling intelligent
   */
  createIntelligentPollingScript() {
    const fs = require('fs');
    const scriptPath = path.join(__dirname, '../../temp/outlook-polling-listener.ps1');
    
    const monitoredFolders = Array.from(this.monitoredFolders.values());
    
  const scriptContent = `
# Script de polling intelligent Outlook - DÉTECTION CHANGEMENTS EN TEMPS RÉEL
# Surveille les modifications d'emails avec vérification des lastModificationTime

$ErrorActionPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"

# Force UTF-8 output to preserve accents/diacritics
$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc

# Configuration
$PollIntervalSeconds = 2  # Vérification toutes les 2 secondes
$DeepScanIntervalMinutes = 1  # Scan profond toutes les 1 minute

try {
    Write-Host "[INFO] [POLLING] Démarrage monitoring intelligent Outlook..."
    
    # Connexion à Outlook via COM
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Host "[OK] [POLLING] Connexion Outlook établie"
    
    # Liste des dossiers à surveiller
    $foldersToMonitor = @(
      ${monitoredFolders
        .map(f => `@{ Path="${(f.path||'').replace(/"/g,'""')}"; EntryId="${(f.entryId||'').replace(/"/g,'""')}"; StoreId="${(f.storeId||'').replace(/"/g,'""')}"; StoreName="${(f.storeName||'').replace(/"/g,'""')}" }`)
        .join('\n      ')}
    )
    
    Write-Host "[FOLDER] [POLLING] Configuration surveillance pour $($foldersToMonitor.Length) dossiers"
    
    # Cache des états des emails (pour détecter les changements)
    $emailCache = @{}
    $lastDeepScan = Get-Date
    
    # Fonction pour obtenir les informations d'un email
    function Get-EmailInfo($item) {
        $parent = $null
        try { $parent = $item.Parent } catch {}
        $store = $null
        try { if ($parent -ne $null) { $store = $parent.Store } } catch {}
        return @{
            EntryID = $item.EntryID
            InternetMessageId = if($item.InternetMessageId) { $item.InternetMessageId } elseif($item.InternetMessageID) { $item.InternetMessageID } else { "" }
            Subject = if($item.Subject) { $item.Subject } else { "(Sans objet)" }
            SenderName = if($item.SenderName) { $item.SenderName } else { "Inconnu" }
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
            UnRead = $item.UnRead
            Size = $item.Size
            LastModificationTime = $item.LastModificationTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Importance = $item.Importance
            Categories = if($item.Categories) { $item.Categories } else { "" }
            FolderPath = if($parent -ne $null) { $parent.FolderPath } else { "" }
            StoreId = if($store -ne $null) { $store.StoreID } else { "" }
            StoreName = if($store -ne $null) { $store.DisplayName } else { "" }
        }
    }
    
    # Fonction pour traiter un dossier
    function Process-Folder($folderInfo) {
        try {
        $folderPath = $folderInfo.Path
        $entryId = $folderInfo.EntryId
        $storeId = $folderInfo.StoreId
            
        # Navigation prioritaire via EntryId + StoreId
        $folder = $null
        if ($entryId -and $entryId -ne "") {
          try {
            if ($storeId -and $storeId -ne "") {
              $folder = $namespace.GetFolderFromID($entryId, $storeId)
            } else {
              $folder = $namespace.GetFolderFromID($entryId)
            }
          } catch {}
        }
            
        # Fallback: navigation via chemin affiché
        if (-not $folder -and $folderPath) {
          $pathParts = $folderPath -split "\\"
                
          if ($pathParts.Length -ge 2) {
            $storeName = $pathParts[0]
            $folderNames = $pathParts[1..($pathParts.Length-1)]
                    
            # Trouver le store
            $targetStore = $null
            foreach ($store in $namespace.Stores) {
              if ($store.DisplayName -eq $storeName -or $store.StoreID -eq $storeId) {
                $targetStore = $store
                break
              }
            }
                    
            if ($targetStore) {
              # Utiliser GetDefaultFolder pour éviter les problèmes d'encodage
              if ($folderNames[0] -like "*reception*" -or $folderNames[0] -like "*Inbox*" -or $folderNames[0] -eq "Boite de reception") {
                $currentFolder = $namespace.GetDefaultFolder(6)  # olFolderInbox
                $remainingFolders = $folderNames[1..($folderNames.Length-1)]
                foreach ($folderName in $remainingFolders) {
                  $found = $false
                  foreach ($subfolder in $currentFolder.Folders) {
                    if ($subfolder.Name -eq $folderName) {
                      $currentFolder = $subfolder
                      $found = $true
                      break
                    }
                  }
                  if (-not $found) {
                    throw "Dossier non trouvé: $folderName"
                  }
                }
              } else {
                $currentFolder = $targetStore.GetRootFolder()
                foreach ($folderName in $folderNames) {
                  $found = $false
                  foreach ($subfolder in $currentFolder.Folders) {
                    if ($subfolder.Name -eq $folderName) {
                      $currentFolder = $subfolder
                      $found = $true
                      break
                    }
                  }
                  if (-not $found) {
                    throw "Dossier non trouvé: $folderName"
                  }
                }
              }
                        
              $folder = $currentFolder
            }
          }
        }
            
            if ($folder) {
              $folderName = $folder.Name
              $currentTime = Get-Date
                
                # Vérifier chaque email dans le dossier
                foreach ($item in $folder.Items) {
                    if ($item.Class -eq 43) {  # olMail
                        $emailInfo = Get-EmailInfo $item
                        $entryId = $emailInfo.EntryID
                        
                        # Vérifier si cet email a changé
                        if ($emailCache.ContainsKey($entryId)) {
                            $cachedInfo = $emailCache[$entryId]
                            
                            # Comparer les propriétés critiques
                            if ($cachedInfo.UnRead -ne $emailInfo.UnRead -or 
                                $cachedInfo.LastModificationTime -ne $emailInfo.LastModificationTime -or
                                $cachedInfo.Categories -ne $emailInfo.Categories) {
                                
                                # Changement détecté !
                                $eventData = @{
                                    Type = "ItemChange"
                                    EntryID = $entryId
                                    Subject = $emailInfo.Subject
                                    SenderName = $emailInfo.SenderName
                                    ReceivedTime = $emailInfo.ReceivedTime
                                    FolderPath = $folderPath
                                    FolderName = $folderName
                                    StoreId = $folderInfo.StoreId
                                    StoreName = $folderInfo.StoreName
                                    IsRead = $emailInfo.UnRead -eq $false
                                    Size = $emailInfo.Size
                                    Importance = $emailInfo.Importance
                                    Categories = $emailInfo.Categories
                                    InternetMessageId = $emailInfo.InternetMessageId
                                    ChangeType = "Modified"
                                    LastModificationTime = $emailInfo.LastModificationTime
                                    DetectedAt = $currentTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
                                    Changes = @()
                                }
                                
                                # Identifier les types de changements
                                if ($cachedInfo.UnRead -ne $emailInfo.UnRead) {
                                    $eventData.Changes += if ($emailInfo.UnRead) { "MarkedUnread" } else { "MarkedRead" }
                                }
                                if ($cachedInfo.Categories -ne $emailInfo.Categories) {
                                    $eventData.Changes += "CategoriesChanged"
                                }
                                if ($cachedInfo.LastModificationTime -ne $emailInfo.LastModificationTime) {
                                    $eventData.Changes += "ContentModified"
                                }
                                
                                $json = $eventData | ConvertTo-Json -Compress
                                Write-Host "EVENT_DATA:$json"
                                
                                # Mettre à jour le cache
                                $emailCache[$entryId] = $emailInfo
                            }
                        } else {
                            # Nouvel email détecté (seulement lors du premier scan, pas pour les anciens)
                            if ($emailCache.Count -gt 0) {  # Pas le premier scan global
                                $eventData = @{
                                    Type = "ItemAdd"
                                    EntryID = $entryId
                                    Subject = $emailInfo.Subject
                                    SenderName = $emailInfo.SenderName
                                    ReceivedTime = $emailInfo.ReceivedTime
                                    FolderPath = $folderPath
                                    FolderName = $folderName
                                    StoreId = $folderInfo.StoreId
                                    StoreName = $folderInfo.StoreName
                                    IsRead = $emailInfo.UnRead -eq $false
                                    Size = $emailInfo.Size
                                    Importance = $emailInfo.Importance
                                    Categories = $emailInfo.Categories
                                    InternetMessageId = $emailInfo.InternetMessageId
                                    ChangeType = "Added"
                                    DetectedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
                                    Changes = @()
                                }
                                
                                # NOUVEAU: Détecter les changements de statut même pour les "nouveaux" emails
                                # (peut être un email existant non en cache à cause du redémarrage du script)
                                if ($emailInfo.UnRead -eq $false) {
                                    $eventData.Changes += "MarkedRead"
                                } else {
                                    $eventData.Changes += "MarkedUnread"
                                }
                                
                                $json = $eventData | ConvertTo-Json -Compress
                                Write-Host "EVENT_DATA:$json"
                            }
                            
                            # Ajouter au cache
                            $emailCache[$entryId] = $emailInfo
                        }
                        
                        # Libération COM
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($item) | Out-Null
                    }
                }
                
                # Libération COM du dossier
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folder) | Out-Null
                
            } else {
              Write-Host "[ERROR] [POLLING] Impossible de trouver le dossier: $folderPath"
            }
            
        } catch {
            Write-Host "[ERROR] [POLLING] Erreur traitement dossier $folderPath : $($_.Exception.Message)"
        }
    }
    
    Write-Host "[OK] [POLLING] Démarrage de la boucle de monitoring..."
    
    # Premier scan pour initialiser le cache (sans émettre d'événements)
    Write-Host "[INIT] [POLLING] Initialisation du cache..."
    foreach ($folderInfo in $foldersToMonitor) {
      Process-Folder $folderInfo
    }
    Write-Host "[INIT] [POLLING] Cache initialisé avec $($emailCache.Count) emails"
    
    # Boucle principale de monitoring
    $iterationCount = 0
    while ($true) {
        try {
            $iterationCount++
            $currentTime = Get-Date
            
            # Traiter chaque dossier surveillé
            foreach ($folderInfo in $foldersToMonitor) {
              Process-Folder $folderInfo
            }
            
            # Heartbeat périodique
            if ($iterationCount % 15 -eq 0) {  # Toutes les 30 secondes (15 * 2s)
                Write-Host "HEARTBEAT:$($currentTime.ToString('yyyy-MM-dd HH:mm:ss'))"
            }
            
            # Nettoyage du cache périodique (scan profond)
            if (($currentTime - $lastDeepScan).TotalMinutes -ge $DeepScanIntervalMinutes) {
                
                # Garder seulement les emails récents dans le cache (optimisation mémoire)
                $recentThreshold = $currentTime.AddHours(-24)
                $emailsToRemove = @()
                foreach ($entryId in $emailCache.Keys) {
                    $emailInfo = $emailCache[$entryId]
                    try {
                        if ([DateTime]::ParseExact($emailInfo.ReceivedTime, "yyyy-MM-dd HH:mm:ss", $null) -lt $recentThreshold) {
                            $emailsToRemove += $entryId
                        }
                    } catch {
                        # Si on ne peut pas parser la date, garder l'email
                    }
                }
                foreach ($entryId in $emailsToRemove) {
                    $emailCache.Remove($entryId)
                }
                
                $lastDeepScan = $currentTime
                if ($emailsToRemove.Count -gt 0) {
                    Write-Host "[DEEP] [POLLING] Cache nettoyé : $($emailsToRemove.Count) emails anciens supprimés"
                }
            }
            
            # Attendre avant la prochaine itération
            Start-Sleep -Seconds $PollIntervalSeconds
            
        } catch {
            Write-Host "[ERROR] [POLLING] Erreur dans la boucle principale: $($_.Exception.Message)"
            Start-Sleep -Seconds 5
        }
    }
    
} catch {
    Write-Host "[ERROR] [POLLING] Erreur fatale: $($_.Exception.Message)"
    exit 1
} finally {
    # Libération des objets COM
    if ($namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null }
}
`;
    
    // Écrire le script avec encodage UTF-8 BOM
    const scriptBuffer = Buffer.concat([
      Buffer.from([0xEF, 0xBB, 0xBF]), // UTF-8 BOM
      Buffer.from(scriptContent, 'utf8')
    ]);
    fs.writeFileSync(scriptPath, scriptBuffer);
    
    return scriptPath;
  }

  /**
   * Traite les événements de polling intelligent
   */
  handlePollingEvent(data) {
    try {
      const lines = data.split('\n').filter(line => line.trim());
      
      for (const line of lines) {
        if (line.startsWith('EVENT_DATA:')) {
          const jsonData = line.replace('EVENT_DATA:', '');
          const eventData = JSON.parse(jsonData);
          
          console.log(`🔔 [OutlookEvents] Changement détecté: ${eventData.ChangeType} - ${eventData.Subject} (${eventData.Changes ? eventData.Changes.join(', ') : 'N/A'})`);
          
          // Émettre l'événement pour le système
          this.emit('email-changed', eventData);
          
        } else if (line.startsWith('HEARTBEAT:')) {
          this.lastEventTime = new Date();
          // Heartbeat silencieux
        } else if (line.includes('[POLLING]')) {
          console.log(`🔔 [OutlookEvents] ${line}`);
        }
      }
    } catch (error) {
      console.error('❌ [OutlookEvents] Erreur traitement événement polling:', error);
    }
  }

  /**
   * Attendre que le processus démarre
   */
  async waitForProcessStart() {
    return new Promise((resolve, reject) => {
      setTimeout(() => {
        if (this.outlookProcess && !this.outlookProcess.killed) {
          resolve();
        } else {
          reject(new Error('Impossible de démarrer le processus d\'écoute'));
        }
      }, 2000);
    });
  }

  /**
   * Gère les crashes de processus
   */
  handleProcessCrash() {
    if (this.reconnectAttempts < this.config.maxReconnectAttempts) {
      this.reconnectAttempts++;
      console.log(`⚠️ [OutlookEvents] Tentative de reconnexion ${this.reconnectAttempts}/${this.config.maxReconnectAttempts}...`);
      
      setTimeout(async () => {
        try {
          await this.startHybridListener();
        } catch (error) {
          console.error('❌ [OutlookEvents] Échec reconnexion:', error);
        }
      }, this.config.reconnectDelay);
    } else {
      console.error('❌ [OutlookEvents] Nombre maximum de tentatives de reconnexion atteint');
      this.isListening = false;
      this.emit('listening-failed');
    }
  }

  /**
   * Groupe les événements par dossier
   */
  groupEventsByFolder(events) {
    const grouped = {};
    
    events.forEach(event => {
      const folderPath = event.FolderPath;
      if (!grouped[folderPath]) {
        grouped[folderPath] = {
          newEmails: 0,
          changedEmails: 0,
          events: []
        };
      }
      
      if (event.Type === 'ItemAdd') {
        grouped[folderPath].newEmails++;
      } else if (event.Type === 'ItemChange') {
        grouped[folderPath].changedEmails++;
      }
      
      grouped[folderPath].events.push(event);
    });
    
    return grouped;
  }

  /**
   * Nettoie le buffer d'événements
   */
  clearEventBuffer() {
    this.eventBuffer = [];
    if (this.bufferTimeout) {
      clearTimeout(this.bufferTimeout);
      this.bufferTimeout = null;
    }
  }

  /**
   * Gère la fermeture du processus PowerShell
   */
  handleProcessClose(code) {
    this.isListening = false;
    
    if (code !== 0 && this.reconnectAttempts < this.config.maxReconnectAttempts) {
      console.log(`🔄 [OutlookEvents] Tentative de reconnexion (${this.reconnectAttempts + 1}/${this.config.maxReconnectAttempts})...`);
      
      setTimeout(() => {
        this.reconnectAttempts++;
        this.startCOMListener().catch(error => {
          console.error('❌ [OutlookEvents] Échec reconnexion:', error);
        });
      }, this.config.reconnectDelay);
    } else {
      console.log('❌ [OutlookEvents] Écoute COM terminée définitivement');
      this.emit('listening-failed', { code, attempts: this.reconnectAttempts });
    }
  }

  /**
   * Récupère les statistiques de l'écoute
   */
  getListeningStats() {
    return {
      isListening: this.isListening,
      monitoredFolders: this.monitoredFolders.size,
      lastEventTime: this.lastEventTime,
      bufferSize: this.eventBuffer.length,
      reconnectAttempts: this.reconnectAttempts
    };
  }
}

module.exports = OutlookEventsService;
