# Système de Logs Filtrable

## Vue d'ensemble

Le nouveau système de logs offre une interface moderne avec filtres, catégories et recherche en temps réel.

## Caractéristiques

### Niveaux de logs
- **DEBUG** 🔍 - Informations de débogage détaillées
- **INFO** ℹ️ - Informations générales
- **SUCCESS** ✅ - Opérations réussies
- **WARN** ⚠️ - Avertissements
- **ERROR** ❌ - Erreurs

### Catégories
- **INIT** 🚀 - Initialisation
- **SYNC** 🔄 - Synchronisation
- **COM** 📡 - Communications COM
- **DB** 💾 - Base de données
- **PS** ⚙️ - PowerShell
- **IPC** 📱 - Inter-Process Communication
- **CONFIG** ⚙️ - Configuration
- **WEEKLY** 📅 - Statistiques hebdomadaires
- **EMAIL** 📧 - Emails
- **PERF** ⚡ - Performance
- **START** ▶️ - Démarrage
- **STOP** ⏹️ - Arrêt

## Utilisation

### Dans le code

```javascript
const logService = require('./services/logService');

// Différents niveaux
logService.debug('SYNC', 'Détails de synchronisation', { folder: 'Test' });
logService.info('INIT', 'Service initialisé');
logService.success('DB', 'Base de données connectée');
logService.warn('CONFIG', 'Configuration manquante');
logService.error('COM', 'Erreur de connexion', error);

// Ou avec la méthode générique
logService.log('INFO', 'SYNC', 'Message', data);
```

### Dans l'interface

1. Cliquer sur **Logs** dans le menu
2. Utiliser les filtres :
   - **Niveau minimum** : Afficher uniquement les logs d'un niveau donné et supérieur
   - **Catégorie** : Filtrer par catégorie spécifique
   - **Recherche** : Recherche textuelle dans les messages et données
3. Options :
   - **Défilement automatique** : Suit automatiquement les nouveaux logs
   - **Pause** : Met en pause la réception des nouveaux logs
   - **Export** : Exporte les logs filtrés en fichier texte
   - **Effacer** : Supprime tous les logs

## Migration depuis l'ancien système

Les services qui utilisent `this.log()` dans `UnifiedMonitoringService` ont été automatiquement migrés pour utiliser le nouveau système en parallèle.

Pour migrer d'autres services :

```javascript
// Avant
console.log('[INFO] Message');

// Après
const logService = require('./services/logService');
logService.info('CATEGORY', 'Message');
```

## API IPC

### Handlers disponibles

- `api-get-log-history` : Récupère l'historique avec filtres
- `api-clear-logs` : Efface tous les logs
- `api-get-log-stats` : Récupère les statistiques

### Événements

- `log-entry` : Nouveau log reçu en temps réel
- `logs-cleared` : Logs effacés

## Performance

- Historique limité à 2000 entrées
- Affichage limité à 500 entrées visibles
- Recherche avec debounce de 300ms
- Filtres optimisés côté serveur

## Fichiers créés

- `src/services/logService.js` - Service de logging centralisé
- `public/logs.html` - Interface de visualisation
- `public/js/logs.js` - Logique frontend
