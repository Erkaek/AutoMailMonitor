/**
 * Service de logging centralisé avec filtres et catégories
 */

class LogService {
  constructor() {
    this.listeners = new Set();
    this.logHistory = [];
    // Mode debug constant: conserver davantage d'historique en mémoire.
    this.maxHistorySize = 10000;
    
    // Niveaux de log
    this.levels = {
      DEBUG: { value: 0, label: 'DEBUG', emoji: '🔍', color: '#6c757d' },
      INFO: { value: 1, label: 'INFO', emoji: 'ℹ️', color: '#0dcaf0' },
      WARN: { value: 2, label: 'WARN', emoji: '⚠️', color: '#ffc107' },
      ERROR: { value: 3, label: 'ERROR', emoji: '❌', color: '#dc3545' },
      SUCCESS: { value: 1, label: 'SUCCESS', emoji: '✅', color: '#198754' }
    };
    
    // Catégories
    this.categories = {
      INIT: '🚀 Init',
      SYNC: '🔄 Sync',
      COM: '📡 COM',
      DB: '💾 DB',
      PS: '⚙️ PowerShell',
      IPC: '📱 IPC',
      CONFIG: '⚙️ Config',
      WEEKLY: '📅 Weekly',
      EMAIL: '📧 Email',
      PERF: '⚡ Performance',
      SECURITY: '🔒 Security',
      CACHE: '💨 Cache',
      START: '▶️ Start',
      STOP: '⏹️ Stop',
      AUTO: '🤖 Auto'
    };
    
    // Mode très verbeux: DEBUG toujours actif.
    this.currentLevel = this.levels.DEBUG.value;
    this.stats = { DEBUG: 0, INFO: 0, WARN: 0, ERROR: 0, SUCCESS: 0 };
  }

  setLevel(levelName) {
    const level = this.levels[levelName];
    if (level) {
      this.currentLevel = level.value;
    }
  }

  log(level, category, message, data = null) {
    const levelObj = this.levels[level];
    if (!levelObj || levelObj.value < this.currentLevel) {
      return; // Niveau trop bas, on ignore
    }

    const timestamp = new Date().toISOString();
    const categoryLabel = this.categories[category] || category;
    
    const logEntry = {
      timestamp,
      level: level,
      levelValue: levelObj.value,
      levelEmoji: levelObj.emoji,
      levelColor: levelObj.color,
      category,
      categoryLabel,
      message,
      data: data ? (typeof data === 'string' ? data : JSON.stringify(data, null, 2)) : null
    };

    // Ajouter à l'historique
    this.logHistory.push(logEntry);
    if (this.logHistory.length > this.maxHistorySize) {
      this.logHistory.shift();
    }

    // Stats
    this.stats[level] = (this.stats[level] || 0) + 1;

    // Console classique (pour compatibilité)
    const consoleMsg = `[${timestamp}] [${level}] [${category}] ${message}`;
    switch (level) {
      case 'ERROR':
        console.error(consoleMsg, data || '');
        break;
      case 'WARN':
        console.warn(consoleMsg, data || '');
        break;
      default:
        console.log(consoleMsg, data || '');
    }

    // Notifier les listeners (interface)
    this.notifyListeners(logEntry);

    return logEntry;
  }

  debug(category, message, data) {
    return this.log('DEBUG', category, message, data);
  }

  info(category, message, data) {
    return this.log('INFO', category, message, data);
  }

  warn(category, message, data) {
    return this.log('WARN', category, message, data);
  }

  error(category, message, data) {
    return this.log('ERROR', category, message, data);
  }

  success(category, message, data) {
    return this.log('SUCCESS', category, message, data);
  }

  addListener(callback) {
    this.listeners.add(callback);
  }

  removeListener(callback) {
    this.listeners.delete(callback);
  }

  notifyListeners(logEntry) {
    this.listeners.forEach(listener => {
      try {
        listener(logEntry);
      } catch (err) {
        console.error('Erreur notification listener:', err);
      }
    });
  }

  getHistory(filters = {}) {
    const limit = Number.isFinite(filters.limit) ? Math.max(1, Number(filters.limit)) : 0;
    const levelValue = (filters.level && filters.level !== 'ALL') ? this.levels[filters.level]?.value : undefined;
    const category = (filters.category && filters.category !== 'ALL') ? filters.category : null;
    const searchLower = (filters.search && String(filters.search).trim()) ? String(filters.search).toLowerCase() : null;

    const out = [];
    // Parcours inverse pour récupérer d'abord les logs les plus récents et
    // pouvoir s'arrêter tôt dès que la limite est atteinte.
    for (let i = this.logHistory.length - 1; i >= 0; i--) {
      const log = this.logHistory[i];
      if (!log) continue;

      if (levelValue !== undefined && log.levelValue < levelValue) continue;
      if (category && log.category !== category) continue;
      if (searchLower) {
        const msg = String(log.message || '').toLowerCase();
        const data = log.data ? String(log.data).toLowerCase() : '';
        if (!msg.includes(searchLower) && !data.includes(searchLower)) continue;
      }

      out.push(log);
      if (limit > 0 && out.length >= limit) break;
    }

    // Repasser dans l'ordre chronologique attendu par l'UI.
    out.reverse();
    return out;
  }

  getStats() {
    return { ...this.stats };
  }

  clear() {
    this.logHistory = [];
    this.stats = { DEBUG: 0, INFO: 0, WARN: 0, ERROR: 0, SUCCESS: 0 };
    this.notifyListeners({ type: 'clear' });
  }
}

// Singleton
const logService = new LogService();

module.exports = logService;
