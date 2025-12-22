# Documentation AutoMailMonitor

Cette documentation couvre les systèmes critiques de l'application.

## Documents Disponibles

### 📋 [LOGS_SYSTEM.md](./LOGS_SYSTEM.md)
Documentation technique du système de logs avec filtres.

**Contenu :**
- Architecture du logService
- Niveaux et catégories de logs
- API et événements IPC
- Intégration dans les services

**Pour :** Développeurs

---

### 📖 [LOGS_USER_GUIDE.md](./LOGS_USER_GUIDE.md)
Guide utilisateur pour l'interface de logs.

**Contenu :**
- Utilisation de l'interface logs.html
- Filtrage par niveau et catégorie
- Recherche et export
- Conseils de dépannage

**Pour :** Utilisateurs finaux

---

### 🔄 [AUTO_UPDATE_SYSTEM.md](./AUTO_UPDATE_SYSTEM.md)
Documentation complète du système de mise à jour automatique.

**Contenu :**
- Architecture de l'updateManager
- Flux de mise à jour
- Configuration et paramètres
- Événements IPC
- Gestion des erreurs et retry
- Support dépôts privés (token GitHub)
- Sécurité et validation
- Troubleshooting

**Pour :** Développeurs et DevOps

---

### 🧪 [AUTO_UPDATE_TESTING.md](./AUTO_UPDATE_TESTING.md)
Guide pratique pour tester les mises à jour.

**Contenu :**
- Test local avec serveur HTTP
- Test avec GitHub Releases réelles
- Scénarios de test détaillés
- Checklist avant release
- Procédures de rollback
- Monitoring post-release
- FAQ et support

**Pour :** QA et Développeurs

---

## Quick Start

### Pour les Développeurs

1. **Logs** : Lire [LOGS_SYSTEM.md](./LOGS_SYSTEM.md) pour intégrer le logging dans vos services
   ```javascript
   const logService = require('../services/logService');
   logService.info('CATEGORY', 'Message', 'Details');
   ```

2. **Mises à jour** : Lire [AUTO_UPDATE_SYSTEM.md](./AUTO_UPDATE_SYSTEM.md) pour comprendre le flux
   ```javascript
   const updateManager = require('../services/updateManager');
   updateManager.initialize();
   ```

### Pour les Testeurs

1. **Logs** : Suivre [LOGS_USER_GUIDE.md](./LOGS_USER_GUIDE.md) pour analyser les problèmes
2. **Updates** : Suivre [AUTO_UPDATE_TESTING.md](./AUTO_UPDATE_TESTING.md) pour tester les releases

### Pour les Utilisateurs

Consultez uniquement [LOGS_USER_GUIDE.md](./LOGS_USER_GUIDE.md) pour :
- Voir l'activité de l'application
- Filtrer les erreurs
- Exporter les logs pour support

---

## Structure du Code

```
src/
├── services/
│   ├── logService.js          # ← Système de logs centralisé
│   ├── updateManager.js       # ← Gestionnaire de mises à jour
│   ├── optimizedDatabaseService.js
│   ├── unifiedMonitoringService.js
│   └── outlookEventsService.js
├── main/
│   ├── index.js               # ← Point d'entrée (utilise logService & updateManager)
│   ├── preload.js             # ← Bridge IPC
│   └── logger.js              # ← Ancien système (legacy)
└── server/
    └── outlookConnector.js    # ← COM Outlook

public/
├── index.html                 # ← Interface principale
├── logs.html                  # ← Interface de logs filtrables
└── js/
    ├── app.js                 # ← Logique UI (listeners update events)
    └── logs.js                # ← Gestion interface logs

docs/
├── README.md                  # ← Ce fichier
├── LOGS_SYSTEM.md
├── LOGS_USER_GUIDE.md
├── AUTO_UPDATE_SYSTEM.md
└── AUTO_UPDATE_TESTING.md
```

---

## Workflows Importants

### 1. Ajouter du Logging dans un Service

```javascript
// En haut du fichier
const logService = require('../services/logService');

// Dans votre code
logService.debug('DB', 'Query executed', query);
logService.info('SYNC', 'Synchronisation démarrée');
logService.success('SYNC', 'Synchronisation terminée', { count: 42 });
logService.warn('COM', 'Connexion Outlook instable');
logService.error('DB', 'Erreur requête', error.message);
```

**Catégories disponibles :**
INIT, SYNC, COM, DB, PS, IPC, CONFIG, WEEKLY, EMAIL, PERF, SECURITY, CACHE, START, STOP, AUTO

### 2. Tester une Mise à Jour

```bash
# 1. Setup serveur local
mkdir -p /tmp/update-test && cd /tmp/update-test
echo "version: 1.0.1" > latest.yml
http-server -p 8080

# 2. Modifier version app
# package.json → "version": "1.0.0"

# 3. Lancer
npm start

# 4. Observer logs
# Onglet Logs → Catégorie: INIT → Rechercher "MAJ"
```

### 3. Publier une Release

```bash
# 1. Bump version
npm version minor  # 1.0.0 → 1.1.0

# 2. Build et publish
npm run build

# 3. Créer release GitHub
gh release create v1.1.0 \
  --title "Release v1.1.0" \
  --notes "$(git log --oneline $(git describe --tags --abbrev=0)..HEAD)"

# 4. Upload assets
gh release upload v1.1.0 dist/*.exe dist/*.yml
```

---

## Maintenance

### Nettoyage des Logs

Les logs sont automatiquement limités à 2000 entrées en mémoire. Pour nettoyer manuellement :

```javascript
// Via IPC
electronAPI.clearLogs();

// Ou dans l'interface
Onglet Logs → Bouton "Effacer"
```

### Monitoring des Mises à Jour

```bash
# Vérifier les downloads d'une release
gh release view v1.0.1 --json assets \
  --jq '.assets[] | "\(.name): \(.downloadCount) downloads"'
```

---

## Contribution

Pour ajouter de la documentation :

1. Créer un nouveau fichier `.md` dans `docs/`
2. Ajouter une section dans ce README
3. Suivre le format existant (titre, contenu, pour qui)
4. Mettre à jour la table des matières

---

## Versions

| Document | Dernière MAJ | Version |
|----------|--------------|---------|
| LOGS_SYSTEM.md | 2024-01-15 | 1.0 |
| LOGS_USER_GUIDE.md | 2024-01-15 | 1.0 |
| AUTO_UPDATE_SYSTEM.md | 2024-01-15 | 1.0 |
| AUTO_UPDATE_TESTING.md | 2024-01-15 | 1.0 |

---

## Support

En cas de question :

1. **Développement** : Consulter les docs techniques (LOGS_SYSTEM, AUTO_UPDATE_SYSTEM)
2. **Tests** : Suivre les guides de test (AUTO_UPDATE_TESTING)
3. **Utilisation** : Lire le guide utilisateur (LOGS_USER_GUIDE)
4. **Problèmes** : Vérifier les logs dans l'application (Onglet Logs)

---

**Navigation :**
- ← [Retour au projet](../README.md)
- → [Logs System](./LOGS_SYSTEM.md)
- → [Auto-Update System](./AUTO_UPDATE_SYSTEM.md)
