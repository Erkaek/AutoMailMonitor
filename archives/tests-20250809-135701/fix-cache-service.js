const fs = require('fs');

// Lire le fichier
const filePath = './src/services/unifiedMonitoringService.js';
let content = fs.readFileSync(filePath, 'utf8');

// Remplacer toutes les occurrences
const oldPattern = /this\.cacheService\.invalidateStats\(\);/g;
const newPattern = `if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        this.cacheService.invalidateStats();
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }`;

content = content.replace(oldPattern, newPattern);

// Écrire le fichier
fs.writeFileSync(filePath, content, 'utf8');
console.log('✅ Toutes les occurrences de cacheService.invalidateStats ont été corrigées');
