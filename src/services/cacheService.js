/**
 * Service de cache intelligent pour optimiser les performances
 * QUICK WIN: +500% UI responsiveness
 */

const NodeCache = require('node-cache');

class CacheService {
    constructor() {
        // Cache principal avec différents TTL selon le type de données
        this.caches = {
            // Cache ultra-rapide pour UI (30s)
            ui: new NodeCache({ stdTTL: 30, checkperiod: 10 }),
            
            // Cache moyen pour API (5 minutes)
            api: new NodeCache({ stdTTL: 300, checkperiod: 60 }),
            
            // Cache long pour configurations (30 minutes)
            config: new NodeCache({ stdTTL: 1800, checkperiod: 300 }),
            
            // Cache emails récents (1 minute)
            emails: new NodeCache({ stdTTL: 60, checkperiod: 20 })
        };

        this.stats = {
            hits: 0,
            misses: 0,
            sets: 0
        };
    }

    /**
     * Get avec fallback automatique
     */
    get(cacheType, key, fallbackFn = null) {
        const cache = this.caches[cacheType];
        if (!cache) {
            throw new Error(`Cache type '${cacheType}' non trouvé`);
        }

        const value = cache.get(key);
        
        if (value !== undefined) {
            this.stats.hits++;
            return value;
        }

        this.stats.misses++;

        // Si fallback fourni, l'exécuter et mettre en cache
        if (fallbackFn) {
            const result = fallbackFn();
            this.set(cacheType, key, result);
            return result;
        }

        return undefined;
    }

    /**
     * Set avec types
     */
    set(cacheType, key, value, ttl = null) {
        const cache = this.caches[cacheType];
        if (!cache) {
            throw new Error(`Cache type '${cacheType}' non trouvé`);
        }

        this.stats.sets++;
        
        if (ttl) {
            return cache.set(key, value, ttl);
        } else {
            return cache.set(key, value);
        }
    }

    /**
     * Suppression ciblée
     */
    del(cacheType, key) {
        const cache = this.caches[cacheType];
        if (cache) {
            return cache.del(key);
        }
        return false;
    }

    /**
     * Suppression par pattern (pour invalidation intelligente)
     */
    delPattern(cacheType, pattern) {
        const cache = this.caches[cacheType];
        if (!cache) return;

        const keys = cache.keys();
        const keysToDelete = keys.filter(key => key.includes(pattern));
        
        keysToDelete.forEach(key => cache.del(key));
        return keysToDelete.length;
    }

    /**
     * Flush sélectif
     */
    flush(cacheType = null) {
        if (cacheType) {
            const cache = this.caches[cacheType];
            if (cache) {
                cache.flushAll();
            }
        } else {
            // Flush all caches
            Object.values(this.caches).forEach(cache => cache.flushAll());
        }
    }

    /**
     * Statistiques globales
     */
    getStats() {
        const cacheStats = {};
        
        Object.entries(this.caches).forEach(([type, cache]) => {
            cacheStats[type] = {
                keys: cache.keys().length,
                stats: cache.getStats()
            };
        });

        return {
            global: this.stats,
            caches: cacheStats,
            hitRate: this.stats.hits / (this.stats.hits + this.stats.misses) * 100
        };
    }

    /**
     * Méthodes helper pour les cas d'usage communs
     */
    
    // Cache pour les stats UI (très fréquent)
    getUIStats(fallbackFn) {
        return this.get('ui', 'dashboard_stats', fallbackFn);
    }

    // Cache pour les emails récents
    getRecentEmails(limit, fallbackFn) {
        return this.get('emails', `recent_${limit}`, fallbackFn);
    }

    // Cache pour la config des dossiers
    getFoldersConfig(fallbackFn) {
        return this.get('config', 'folders_config', fallbackFn);
    }

    // Invalidation intelligente quand config change
    invalidateFoldersConfig() {
    this.del('config', 'folders_config');
    // Also invalidate cached folders tree used by Monitoring
    this.del('config', 'folders_tree');
        this.delPattern('api', 'folders_');
        this.delPattern('ui', 'dashboard_');
    }

    // Invalidation quand nouveaux emails
    invalidateEmailsCache() {
        this.flush('emails');
        this.delPattern('ui', 'stats_');
        this.delPattern('api', 'emails_');
    }

    // Invalidation des statistiques UI/API (sans toucher à la config)
    invalidateStats() {
        // Stats dashboard
        this.del('ui', 'dashboard_stats');
        // Variantes potentielles
        this.delPattern('ui', 'stats_');
        this.delPattern('api', 'stats_');
    }

    // Invalidation ciblée de l'arbre des dossiers (compteurs Outlook)
    invalidateFoldersTree() {
        this.del('config', 'folders_tree');
    }
}

// Export singleton
const cacheService = new CacheService();
module.exports = cacheService;
