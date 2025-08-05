/**
 * Connecteur Outlook avec gestion d'événements en temps réel
 * Évite le polling en utilisant les événements COM natifs
 */

const { exec, spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const EventEmitter = require('events');

class OutlookEventConnector extends EventEmitter {
  constructor() {
    super();
    this.isOutlookConnected = false;
    this.eventListenersActive = false;
    this.lastError = null;
    this.outlookVersion = null;
    this.connectionState = 'unknown';
    this.monitoredFolders = new Map();
    this.eventHandlers = new Map();
    
    // Configuration
    this.config = {
      timeout: 30000,
      enableDetailedLogs: true,
      autoReconnect: true,
      initialSyncEnabled: true
    };
    
    this.init();
  }

  async init() {
    this.log('🚀 Initialisation du connecteur Outlook avec événements temps réel');
    
    // Vérification de l'environnement
    await this.checkEnvironment();
    
    // Établissement de la connexion
    await this.establishConnection();
    
    // Configuration des événements temps réel
    await this.setupEventListeners();
    
    this.log('✅ Connecteur Outlook avec événements initialisé');
  }

  async checkEnvironment() {
    this.log('🔍 Vérification de l\'environnement Windows/Outlook...');
    
    try {
      if (process.platform !== 'win32') {
        throw new Error('Outlook COM n\'est disponible que sur Windows');
      }
      
      // Vérifier PowerShell
      const psVersion = await this.executeCommand('powershell', ['$PSVersionTable.PSVersion.Major']);
      this.log(`📋 Version PowerShell: ${psVersion.trim()}`);
      
      // Vérifier COM
      await this.checkCOMPermissions();
      
      // Détecter Outlook
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
        # Test plus robuste pour COM
        $outlook = New-Object -ComObject Outlook.Application
        if ($outlook) {
          try {
            $namespace = $outlook.GetNamespace("MAPI")
            if ($namespace) {
              Write-Output "COM_OK"
            } else {
              Write-Output "COM_ERROR: Namespace inaccessible"
            }
          } catch {
            Write-Output "COM_OK_LIMITED: Outlook accessible mais namespace limité"
          }
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
      } catch {
        # Test alternatif avec Shell.Application
        try {
          $comTest = New-Object -ComObject "Shell.Application"
          if ($comTest) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($comTest) | Out-Null
            Write-Output "COM_OK_SHELL"
          }
        } catch {
          Write-Output "COM_ERROR: $($_.Exception.Message)"
        }
      }
    `;
    
    const result = await this.executePowerShell(testScript, 10000);
    
    if (result.includes('COM_OK') || result.includes('COM_OK_LIMITED') || result.includes('COM_OK_SHELL')) {
      this.log('✅ Permissions COM vérifiées (ou test alternatif réussi)');
    } else {
      this.log('⚠️ Permissions COM limitées, tentative de contournement...');
      // Ne pas lever d'erreur, laisser le connecteur essayer quand même
    }
  }

  async detectOutlookVersion() {
    const versionScript = `
      try {
        $outlook = Get-ItemProperty "HKLM:\\SOFTWARE\\Microsoft\\Office\\*\\Outlook\\InstallRoot" -ErrorAction SilentlyContinue
        if ($outlook) {
          $version = $outlook.Path.Split('\\')[-2]
          Write-Output $version
        } else {
          Write-Output "unknown"
        }
      } catch {
        Write-Output "error"
      }
    `;
    
    this.outlookVersion = await this.executePowerShell(versionScript, 5000);
    this.log(`📧 Version Outlook détectée: ${this.outlookVersion}`);
  }

  async establishConnection() {
    this.log('🔌 Établissement de la connexion Outlook...');
    
    try {
      // Vérifier si Outlook est démarré
      const outlookRunning = await this.isOutlookRunning();
      
      if (!outlookRunning) {
        this.log('⚠️ Processus Outlook non détecté, tentative de démarrage automatique...');
        await this.startOutlook();
      }
      
      // Établir la connexion COM
      await this.connectCOM();
      
      this.isOutlookConnected = true;
      this.connectionState = 'connected';
      this.log('✅ Connexion Outlook établie avec succès');
      
    } catch (error) {
      this.log(`❌ Échec connexion: ${error.message}`);
      this.connectionState = 'failed';
      this.lastError = error;
      throw error;
    }
  }

  async setupEventListeners() {
    this.log('🎧 Configuration des événements temps réel...');
    
    try {
      const eventScript = `
        $ErrorActionPreference = "Stop"
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
        
        # Connexion à Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Configuration des événements
        Register-ObjectEvent -InputObject $outlook -EventName "NewMail" -Action {
          $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
          Write-Host "EVENT_NEWMAIL:$timestamp"
        }
        
        Register-ObjectEvent -InputObject $outlook -EventName "ItemSend" -Action {
          param($item)
          $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
          Write-Host "EVENT_ITEMSEND:$timestamp"
        }
        
        # Événements sur les dossiers
        $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
        
        Register-ObjectEvent -InputObject $inbox.Items -EventName "ItemAdd" -Action {
          param($item)
          $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
          $entryId = ""
          try { $entryId = $item.EntryID } catch { }
          Write-Host "EVENT_ITEMADD:$timestamp:$entryId"
        }
        
        Register-ObjectEvent -InputObject $inbox.Items -EventName "ItemChange" -Action {
          param($item)
          $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
          $entryId = ""
          $unread = $false
          try { 
            $entryId = $item.EntryID 
            $unread = $item.UnRead
          } catch { }
          Write-Host "EVENT_ITEMCHANGE:$timestamp:$entryId:$unread"
        }
        
        Register-ObjectEvent -InputObject $inbox.Items -EventName "ItemRemove" -Action {
          $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
          Write-Host "EVENT_ITEMREMOVE:$timestamp"
        }
        
        Write-Host "EVENT_LISTENERS_READY"
        
        # Maintenir le script en vie
        while ($true) {
          Start-Sleep -Seconds 1
        }
      `;
      
      // Démarrer le script PowerShell d'écoute des événements
      this.eventProcess = spawn('powershell', ['-Command', eventScript], {
        stdio: ['pipe', 'pipe', 'pipe']
      });
      
      // Écouter les événements
      this.eventProcess.stdout.on('data', (data) => {
        const output = data.toString().trim();
        this.handleOutlookEvent(output);
      });
      
      this.eventProcess.stderr.on('data', (data) => {
        this.log(`❌ Erreur événements: ${data.toString()}`);
      });
      
      this.eventProcess.on('close', (code) => {
        this.log(`⚠️ Processus événements fermé (code: ${code})`);
        this.eventListenersActive = false;
        
        // Auto-restart si nécessaire
        if (this.config.autoReconnect && this.isOutlookConnected) {
          setTimeout(() => this.setupEventListeners(), 5000);
        }
      });
      
      this.eventListenersActive = true;
      this.log('✅ Événements temps réel configurés');
      
    } catch (error) {
      this.log(`❌ Erreur configuration événements: ${error.message}`);
      throw error;
    }
  }

  handleOutlookEvent(eventData) {
    if (!eventData || eventData.trim() === '') return;
    
    const lines = eventData.split('\n');
    
    for (const line of lines) {
      if (line.startsWith('EVENT_')) {
        const parts = line.split(':');
        const eventType = parts[0];
        const timestamp = parts[1];
        
        switch (eventType) {
          case 'EVENT_NEWMAIL':
            this.log(`📬 Nouveau mail détecté à ${timestamp}`);
            this.emit('newMail', { timestamp });
            break;
            
          case 'EVENT_ITEMADD':
            const entryId = parts[2] || '';
            this.log(`➕ Email ajouté: ${entryId}`);
            this.emit('itemAdd', { timestamp, entryId });
            break;
            
          case 'EVENT_ITEMCHANGE':
            const itemId = parts[2] || '';
            const unread = parts[3] === 'True';
            this.log(`📝 Email modifié: ${itemId} (non-lu: ${unread})`);
            this.emit('itemChange', { timestamp, entryId: itemId, unread });
            break;
            
          case 'EVENT_ITEMREMOVE':
            this.log(`🗑️ Email supprimé à ${timestamp}`);
            this.emit('itemRemove', { timestamp });
            break;
            
          case 'EVENT_LISTENERS_READY':
            this.log('🎧 Écouteurs d\'événements prêts');
            this.emit('listenersReady');
            break;
        }
      }
    }
  }

  // Synchronisation initiale complète (appelée une seule fois au démarrage)
  async performInitialSync(folderPath) {
    this.log(`🔄 Synchronisation initiale du dossier: ${folderPath}`);
    
    const syncScript = `
      $ErrorActionPreference = "Stop"
      
      try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Récupérer le dossier
        $folder = $namespace.Folders("${folderPath.split('\\')[1]}").Folders("${folderPath.split('\\')[2]}")
        
        $emails = @()
        $items = $folder.Items
        $count = $items.Count
        
        Write-Host "SYNC_FOLDER_INFO:$count emails trouvés"
        
        for ($i = 1; $i -le $count; $i++) {
          try {
            $item = $items.Item($i)
            
            if ($item.Class -eq 43) { # MailItem
              $emailData = @{
                EntryID = $item.EntryID
                Subject = $item.Subject
                SenderEmailAddress = ""
                ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                UnRead = $item.UnRead
                Size = $item.Size
                FolderPath = "${folderPath}"
              }
              
              # Gestion sécurisée de l'email expéditeur
              try {
                if ($item.SenderEmailType -eq "EX") {
                  $emailData.SenderEmailAddress = $item.Sender.GetExchangeUser().PrimarySmtpAddress
                } else {
                  $emailData.SenderEmailAddress = $item.SenderEmailAddress
                }
              } catch {
                $emailData.SenderEmailAddress = "unknown@unknown.com"
              }
              
              # Convertir en JSON et émettre
              $json = $emailData | ConvertTo-Json -Compress
              Write-Host "SYNC_EMAIL_DATA:$json"
            }
          } catch {
            Write-Host "SYNC_ERROR_ITEM:$i - $($_.Exception.Message)"
          }
        }
        
        Write-Host "SYNC_COMPLETE:$count"
        
      } catch {
        Write-Host "SYNC_ERROR:$($_.Exception.Message)"
      }
    `;
    
    return new Promise((resolve, reject) => {
      const process = spawn('powershell', ['-Command', syncScript]);
      const emails = [];
      let folderInfo = null;
      
      process.stdout.on('data', (data) => {
        const output = data.toString().trim();
        const lines = output.split('\n');
        
        for (const line of lines) {
          if (line.startsWith('SYNC_FOLDER_INFO:')) {
            folderInfo = line.split(':')[1];
          } else if (line.startsWith('SYNC_EMAIL_DATA:')) {
            try {
              const jsonData = line.substring('SYNC_EMAIL_DATA:'.length);
              const emailData = JSON.parse(jsonData);
              emails.push(emailData);
            } catch (error) {
              this.log(`❌ Erreur parsing email: ${error.message}`);
            }
          } else if (line.startsWith('SYNC_COMPLETE:')) {
            const totalCount = line.split(':')[1];
            this.log(`✅ Sync terminée: ${emails.length}/${totalCount} emails traités`);
            resolve({ emails, totalCount: parseInt(totalCount) });
          } else if (line.startsWith('SYNC_ERROR:')) {
            reject(new Error(line.split(':')[1]));
          }
        }
      });
      
      process.stderr.on('data', (data) => {
        this.log(`❌ Erreur sync: ${data.toString()}`);
      });
      
      process.on('close', (code) => {
        if (code !== 0) {
          reject(new Error(`Processus sync fermé avec code ${code}`));
        }
      });
    });
  }

  // Récupération d'un email spécifique par EntryID (pour les événements)
  async getEmailByEntryId(entryId) {
    if (!entryId) return null;
    
    const getEmailScript = `
      $ErrorActionPreference = "Stop"
      
      try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        $item = $namespace.GetItemFromID("${entryId}")
        
        if ($item -and $item.Class -eq 43) {
          $emailData = @{
            EntryID = $item.EntryID
            Subject = $item.Subject
            SenderEmailAddress = ""
            ReceivedTime = $item.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            UnRead = $item.UnRead
            Size = $item.Size
            FolderPath = $item.Parent.FolderPath
          }
          
          try {
            if ($item.SenderEmailType -eq "EX") {
              $emailData.SenderEmailAddress = $item.Sender.GetExchangeUser().PrimarySmtpAddress
            } else {
              $emailData.SenderEmailAddress = $item.SenderEmailAddress
            }
          } catch {
            $emailData.SenderEmailAddress = "unknown@unknown.com"
          }
          
          $json = $emailData | ConvertTo-Json -Compress
          Write-Host "EMAIL_DATA:$json"
        } else {
          Write-Host "EMAIL_NOT_FOUND"
        }
        
      } catch {
        Write-Host "EMAIL_ERROR:$($_.Exception.Message)"
      }
    `;
    
    return new Promise((resolve, reject) => {
      const process = spawn('powershell', ['-Command', getEmailScript]);
      
      process.stdout.on('data', (data) => {
        const output = data.toString().trim();
        
        if (output.startsWith('EMAIL_DATA:')) {
          try {
            const jsonData = output.substring('EMAIL_DATA:'.length);
            const emailData = JSON.parse(jsonData);
            resolve(emailData);
          } catch (error) {
            reject(error);
          }
        } else if (output.startsWith('EMAIL_NOT_FOUND')) {
          resolve(null);
        } else if (output.startsWith('EMAIL_ERROR:')) {
          reject(new Error(output.split(':')[1]));
        }
      });
      
      process.on('close', (code) => {
        if (code !== 0) {
          reject(new Error(`Processus fermé avec code ${code}`));
        }
      });
    });
  }

  async isOutlookRunning() {
    try {
      const result = await this.executeCommand('tasklist', ['/FI', 'IMAGENAME eq OUTLOOK.EXE']);
      return result.includes('OUTLOOK.EXE');
    } catch {
      return false;
    }
  }

  async startOutlook() {
    this.log('🔥 Tentative de démarrage d\'Outlook...');
    
    const startScript = `
      try {
        $outlookPath = Get-ItemProperty "HKLM:\\SOFTWARE\\Microsoft\\Office\\*\\Outlook\\InstallRoot" | Select-Object -First 1 -ExpandProperty Path
        if ($outlookPath) {
          $exePath = Join-Path $outlookPath "OUTLOOK.EXE"
          if (Test-Path $exePath) {
            Start-Process $exePath
            Write-Output "OUTLOOK_STARTED:$exePath"
          } else {
            Write-Output "OUTLOOK_NOT_FOUND"
          }
        } else {
          Write-Output "OUTLOOK_NOT_INSTALLED"
        }
      } catch {
        Write-Output "OUTLOOK_ERROR:$($_.Exception.Message)"
      }
    `;
    
    const result = await this.executePowerShell(startScript);
    
    if (result.includes('OUTLOOK_STARTED:')) {
      const exePath = result.split(':')[1];
      this.log(`✅ Outlook démarré depuis: ${exePath}`);
      
      // Attendre que le processus soit complètement initialisé
      await this.waitForOutlookReady();
    } else {
      throw new Error('Impossible de démarrer Outlook automatiquement');
    }
  }

  async waitForOutlookReady() {
    this.log('⏳ Attente de l\'initialisation d\'Outlook...');
    
    const maxWait = 60; // 60 secondes max
    let attempts = 0;
    
    while (attempts < maxWait) {
      try {
        const isReady = await this.testOutlookCOMConnection();
        if (isReady) {
          this.log('✅ Outlook initialisé et prêt');
          return;
        }
      } catch {
        // Continuer à attendre
      }
      
      await new Promise(resolve => setTimeout(resolve, 1000));
      attempts++;
    }
    
    throw new Error('Timeout: Outlook n\'a pas pu être initialisé dans les temps');
  }

  async testOutlookCOMConnection() {
    const testScript = `
      try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)
        Write-Output "COM_READY"
      } catch {
        Write-Output "COM_NOT_READY"
      }
    `;
    
    const result = await this.executePowerShell(testScript, 5000);
    return result.includes('COM_READY');
  }

  async connectCOM() {
    const connectionScript = `
      $ErrorActionPreference = "Stop"
      
      try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Test de connexion
        $profile = $namespace.CurrentProfileName
        Write-Host "COM_CONNECTED:$profile"
        
        # Statistiques de base
        $inbox = $namespace.GetDefaultFolder(6)
        $totalItems = $inbox.Items.Count
        $unreadItems = $inbox.UnReadItemCount
        
        Write-Host "COM_STATS:$totalItems,$unreadItems"
        
      } catch {
        Write-Host "COM_ERROR:$($_.Exception.Message)"
      }
    `;
    
    const result = await this.executePowerShell(connectionScript);
    
    if (result.includes('COM_CONNECTED:')) {
      const profile = result.split('COM_CONNECTED:')[1].split('\n')[0];
      this.log(`✅ Connexion COM réussie - Profil: ${profile}`);
      
      if (result.includes('COM_STATS:')) {
        const stats = result.split('COM_STATS:')[1].split('\n')[0].split(',');
        this.log(`📊 Boîte réception: ${stats[0]} éléments, ${stats[1]} non lus`);
      }
    } else {
      throw new Error('Échec de la connexion COM');
    }
  }

  // Arrêt propre
  async cleanup() {
    this.log('🧹 Arrêt du connecteur événements...');
    
    if (this.eventProcess) {
      this.eventProcess.kill();
      this.eventProcess = null;
    }
    
    this.eventListenersActive = false;
    this.isOutlookConnected = false;
    this.removeAllListeners();
    
    this.log('✅ Connecteur événements arrêté');
  }

  // Utilitaires
  async executeCommand(command, args = [], timeout = 10000) {
    return new Promise((resolve, reject) => {
      const process = exec(`${command} ${args.join(' ')}`, {
        timeout,
        windowsHide: true
      });
      
      let output = '';
      process.stdout.on('data', data => output += data);
      process.stderr.on('data', data => output += data);
      
      process.on('close', code => {
        if (code === 0) resolve(output);
        else reject(new Error(`Command failed with code ${code}: ${output}`));
      });
      
      process.on('error', reject);
    });
  }

  async executePowerShell(script, timeout = 15000) {
    const escapedScript = script.replace(/"/g, '`"');
    return this.executeCommand('powershell', ['-Command', `"${escapedScript}"`], timeout);
  }

  log(message) {
    if (this.config.enableDetailedLogs) {
      const timestamp = new Date().toISOString();
      console.log(`[${timestamp}] OutlookEventConnector: ${message}`);
    }
  }

  // Getters pour compatibilité
  get connected() {
    return this.isOutlookConnected;
  }

  get hasEventListeners() {
    return this.eventListenersActive;
  }
}

module.exports = OutlookEventConnector;
