/**
 * Service de gestion des données - Mail Monitor
 * Couche d'abstraction pour la persistance et la synchronisation des données
 */

class DataService {
  constructor() {
    this.cache = new Map();
    this.syncQueue = [];
    this.lastSync = null;
    this.isOnline = true;
    
    this.initializeStorage();
    this.setupSyncScheduler();
  }

  // === INITIALISATION ===
  initializeStorage() {
    // Initialisation des données en cache
    const cachedData = this.loadFromLocalStorage();
    if (cachedData) {
      this.cache = new Map(Object.entries(cachedData));
      console.log('✅ Données chargées depuis le stockage local');
    }
  }

  setupSyncScheduler() {
    // Synchronisation automatique toutes les 5 minutes
    setInterval(() => {
      if (this.isOnline && this.syncQueue.length > 0) {
        this.processSyncQueue();
      }
    }, 5 * 60 * 1000);
  }

  // === GESTION DU CACHE ===
  set(key, value, options = {}) {
    const entry = {
      value,
      timestamp: Date.now(),
      ttl: options.ttl || 3600000, // 1 heure par défaut
      persistent: options.persistent !== false
    };
    
    this.cache.set(key, entry);
    
    if (entry.persistent) {
      this.saveToLocalStorage();
    }
    
    // Ajouter à la queue de sync si nécessaire
    if (options.sync) {
      this.addToSyncQueue('SET', key, value);
    }
    
    return true;
  }

  get(key, defaultValue = null) {
    const entry = this.cache.get(key);
    
    if (!entry) {
      return defaultValue;
    }
    
    // Vérifier l'expiration
    if (Date.now() - entry.timestamp > entry.ttl) {
      this.cache.delete(key);
      this.saveToLocalStorage();
      return defaultValue;
    }
    
    return entry.value;
  }

  has(key) {
    return this.cache.has(key) && !this.isExpired(key);
  }

  delete(key) {
    const deleted = this.cache.delete(key);
    if (deleted) {
      this.saveToLocalStorage();
      this.addToSyncQueue('DELETE', key);
    }
    return deleted;
  }

  clear() {
    this.cache.clear();
    this.saveToLocalStorage();
    this.addToSyncQueue('CLEAR');
  }

  // === PERSISTANCE LOCALE ===
  saveToLocalStorage() {
    try {
      const persistentData = {};
      for (const [key, entry] of this.cache.entries()) {
        if (entry.persistent && !this.isExpired(key)) {
          persistentData[key] = entry;
        }
      }
      localStorage.setItem('automail-data', JSON.stringify(persistentData));
    } catch (error) {
      console.error('Erreur sauvegarde locale:', error);
    }
  }

  loadFromLocalStorage() {
    try {
      const data = localStorage.getItem('automail-data');
      return data ? JSON.parse(data) : null;
    } catch (error) {
      console.error('Erreur chargement local:', error);
      return null;
    }
  }

  // === SYNCHRONISATION ===
  addToSyncQueue(operation, key, value = null) {
    this.syncQueue.push({
      id: `sync_${Date.now()}_${Math.random()}`,
      operation,
      key,
      value,
      timestamp: Date.now(),
      retries: 0
    });
  }

  async processSyncQueue() {
    if (this.syncQueue.length === 0) return;
    
    console.log(`🔄 Traitement de ${this.syncQueue.length} opérations de sync`);
    
    const operations = [...this.syncQueue];
    this.syncQueue = [];
    
    for (const operation of operations) {
      try {
        await this.executeSyncOperation(operation);
        console.log(`✅ Sync réussie: ${operation.operation} ${operation.key}`);
      } catch (error) {
        console.error(`❌ Échec sync: ${operation.operation} ${operation.key}`, error);
        
        // Retry logic
        if (operation.retries < 3) {
          operation.retries++;
          this.syncQueue.push(operation);
        }
      }
    }
    
    this.lastSync = new Date();
  }

  async executeSyncOperation(operation) {
    // Ici on enverrait les données vers un serveur distant
    // Pour l'instant, simulation
    return new Promise((resolve) => {
      setTimeout(() => {
        console.log(`Simulation sync: ${operation.operation}`, operation.key);
        resolve();
      }, 100);
    });
  }

  // === MÉTHODES SPÉCIALISÉES ===
  
  // Emails
  getEmails(folder = 'inbox', limit = 50) {
    const cacheKey = `emails_${folder}_${limit}`;
    return this.get(cacheKey, []);
  }

  setEmails(folder, emails, options = {}) {
    const cacheKey = `emails_${folder}_${options.limit || 50}`;
    return this.set(cacheKey, emails, { 
      ttl: 10 * 60 * 1000, // 10 minutes
      ...options 
    });
  }

  // Statistiques
  getStats(period = 'today') {
    const cacheKey = `stats_${period}`;
    return this.get(cacheKey, {
      emailsReceived: 0,
      emailsSent: 0,
      unreadCount: 0,
      avgResponseTime: 0
    });
  }

  setStats(period, stats) {
    const cacheKey = `stats_${period}`;
    return this.set(cacheKey, stats, { 
      ttl: 5 * 60 * 1000, // 5 minutes
      sync: true 
    });
  }

  // Configuration
  getSettings() {
    return this.get('user_settings', {
      syncInterval: 60,
      notifications: true,
      autoStart: false,
      theme: 'light'
    });
  }

  setSettings(settings) {
    return this.set('user_settings', settings, { 
      persistent: true,
      sync: true 
    });
  }

  // Analytics
  getAnalytics(type, period = '7d') {
    const cacheKey = `analytics_${type}_${period}`;
    return this.get(cacheKey, null);
  }

  setAnalytics(type, period, data) {
    const cacheKey = `analytics_${type}_${period}`;
    return this.set(cacheKey, data, { 
      ttl: 30 * 60 * 1000, // 30 minutes
      sync: true 
    });
  }

  // === MÉTHODES UTILITAIRES ===
  
  isExpired(key) {
    const entry = this.cache.get(key);
    if (!entry) return true;
    
    return Date.now() - entry.timestamp > entry.ttl;
  }

  getCacheInfo() {
    let totalSize = 0;
    let expiredCount = 0;
    
    for (const [key, entry] of this.cache.entries()) {
      totalSize += JSON.stringify(entry).length;
      if (this.isExpired(key)) {
        expiredCount++;
      }
    }
    
    return {
      totalEntries: this.cache.size,
      expiredEntries: expiredCount,
      estimatedSize: `${(totalSize / 1024).toFixed(2)} KB`,
      lastSync: this.lastSync,
      syncQueueSize: this.syncQueue.length
    };
  }

  cleanup() {
    console.log('🧹 Nettoyage du cache...');
    
    const before = this.cache.size;
    
    for (const key of this.cache.keys()) {
      if (this.isExpired(key)) {
        this.cache.delete(key);
      }
    }
    
    const after = this.cache.size;
    const cleaned = before - after;
    
    if (cleaned > 0) {
      console.log(`✅ ${cleaned} entrées expirées supprimées`);
      this.saveToLocalStorage();
    }
    
    return cleaned;
  }

  // === EXPORT/IMPORT ===
  
  exportData() {
    const exportData = {
      version: '1.0.0',
      timestamp: new Date().toISOString(),
      cache: Object.fromEntries(this.cache.entries()),
      settings: this.getSettings(),
      cacheInfo: this.getCacheInfo()
    };
    
    return JSON.stringify(exportData, null, 2);
  }

  importData(jsonData) {
    try {
      const data = JSON.parse(jsonData);
      
      if (data.version && data.cache) {
        this.clear();
        this.cache = new Map(Object.entries(data.cache));
        this.saveToLocalStorage();
        
        console.log('✅ Données importées avec succès');
        return true;
      }
      
      throw new Error('Format de données invalide');
    } catch (error) {
      console.error('❌ Erreur import:', error);
      return false;
    }
  }

  // === MONITORING ===
  
  startMonitoring() {
    console.log('📊 Démarrage du monitoring des données');
    
    // Nettoyage automatique toutes les heures
    setInterval(() => {
      this.cleanup();
    }, 60 * 60 * 1000);
    
    // Log des statistiques toutes les 10 minutes
    setInterval(() => {
      const info = this.getCacheInfo();
      console.log('📈 Cache info:', info);
    }, 10 * 60 * 1000);
  }

  setOnlineStatus(isOnline) {
    const wasOnline = this.isOnline;
    this.isOnline = isOnline;
    
    if (!wasOnline && isOnline) {
      console.log('🌐 Connexion rétablie - traitement de la queue de sync');
      this.processSyncQueue();
    }
  }
}

// Instance singleton
const dataService = new DataService();

// Auto-démarrage du monitoring
dataService.startMonitoring();

module.exports = dataService;
