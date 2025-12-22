# Système de Mise à Jour Automatique

## Vue d'ensemble

L'application utilise `electron-updater` pour gérer les mises à jour automatiques via GitHub Releases. Le système est complètement revu et robustifié avec retry, timeout, et notifications détaillées.

## Architecture

### Composants

1. **updateManager.js** (`src/services/updateManager.js`)
   - Service singleton gérant toute la logique de mise à jour
   - Configuration centralisée (timeouts, retry, intervalles)
   - Gestion des événements auto-updater
   - Retry automatique en cas d'erreur
   - Logging détaillé via logService

2. **main/index.js**
   - Initialise updateManager au démarrage
   - Configure la vérification périodique
   - Expose IPC handler pour vérifications manuelles
   - Connecte updateManager à la fenêtre principale

3. **preload.js**
   - Expose les événements de mise à jour au renderer
   - Bridge sécurisé via contextBridge

4. **app.js**
   - Écoute les événements de mise à jour
   - Affiche des toasts pour informer l'utilisateur
   - Gestion de la progression du téléchargement

## Configuration

### Paramètres (updateManager.config)

```javascript
{
  requestTimeout: 30000,           // 30 secondes max par requête
  maxRetries: 3,                   // 3 tentatives en cas d'échec
  retryDelay: 5000,                // 5 secondes entre chaque retry
  checkInterval: 2 * 60 * 60 * 1000, // 2 heures entre vérifications auto
  minCheckInterval: 5 * 60 * 1000  // 5 minutes minimum entre checks
}
```

### Comportement

- **Téléchargement automatique** : Activé (`autoDownload: true`)
- **Installation à la fermeture** : Activée (`autoInstallOnAppQuit: true`)
- **Pre-releases** : Supportées (configurable via `ALLOW_PRERELEASE` env)
- **Cache** : Désactivé (`Cache-Control: no-cache`)

## Flux de Mise à Jour

### 1. Vérification au démarrage

```
App ready
  ↓
updateManager.initialize()
  ↓
Vérification initiale (runInitialUpdateCheck)
  ↓
Affichage dans la fenêtre de chargement
```

### 2. Vérification périodique

```
Toutes les 2 heures
  ↓
updateManager.checkForUpdates()
  ↓
Respect du minCheckInterval (5 min)
  ↓
Retry automatique si échec (max 3x)
```

### 3. Vérification manuelle

```
Utilisateur clique "Vérifier les MAJ"
  ↓
IPC: app-check-updates-now
  ↓
updateManager.checkManually()
  ↓
Retour: { success, updateInfo, currentVersion }
```

### 4. Téléchargement disponible

```
Mise à jour détectée
  ↓
Toast: "Version X disponible"
  ↓
Téléchargement automatique
  ↓
Notifications de progression (10%, 20%...)
  ↓
Téléchargement terminé
```

### 5. Installation

```
update-downloaded
  ↓
Dialog: "Redémarrer maintenant / Plus tard ?"
  ↓
Si OUI: quitAndInstall()
Si NON: Installation au prochain démarrage
```

## Événements IPC

### Envoyés au Renderer

| Événement | Payload | Description |
|-----------|---------|-------------|
| `update-checking` | `{}` | Vérification démarrée |
| `update-available` | `{ version, releaseDate, releaseNotes }` | Mise à jour trouvée |
| `update-not-available` | `{}` | Aucune mise à jour |
| `update-error` | `{ error }` | Erreur rencontrée |
| `update-download-progress` | `{ percent, transferred, total, bytesPerSecond }` | Progression |
| `update-pending-restart` | `{ version }` | Installation reportée |

### Reçus du Renderer

| Handler | Paramètres | Retour |
|---------|-----------|--------|
| `app-check-updates-now` | Aucun | `{ success, updateInfo, currentVersion, error? }` |

## Gestion des Erreurs

### Retry Automatique

En cas d'erreur réseau ou timeout :

1. Incrémente `updateCheckAttempts`
2. Calcule délai : `retryDelay * attempts` (5s, 10s, 15s)
3. Relance `checkForUpdates()` après délai
4. Abandonne après `maxRetries` (3)

### Timeout Protection

```javascript
Promise.race([
  autoUpdater.checkForUpdates(),
  timeout(30000)
])
```

Évite de bloquer indéfiniment l'application.

### Logging

Tous les événements sont loggés via `logService` :

- Catégorie : `INIT` ou `IPC`
- Niveaux : DEBUG, INFO, SUCCESS, WARN, ERROR
- Visible dans l'onglet Logs de l'application

## Dépôts Privés

### Configuration du Token GitHub

Le système supporte les dépôts privés via token d'authentification.

**Méthodes (par ordre de priorité) :**

1. **Fichier bundlé** : `src/main/updaterToken.js`
   ```javascript
   module.exports = 'ghp_your_token_here';
   ```
   ⚠️ **Ne JAMAIS committer ce fichier** (dans `.gitignore`)

2. **Variable d'environnement** :
   - `GH_TOKEN`
   - `UPDATER_TOKEN`
   - `ELECTRON_UPDATER_TOKEN`

Le token est ajouté aux headers :
```javascript
Authorization: token ghp_xxxxx
```

## Publication de Releases

### Format attendu

GitHub Releases avec assets :

```
v1.2.3
  ├─ MonApp-Setup-1.2.3.exe (Windows)
  ├─ MonApp-1.2.3.dmg (macOS)
  ├─ MonApp-1.2.3.AppImage (Linux)
  └─ latest.yml / latest-mac.yml / latest-linux.yml
```

### Configuration package.json

```json
{
  "build": {
    "appId": "com.example.mailmonitor",
    "publish": [{
      "provider": "github",
      "owner": "Erkaek",
      "repo": "AutoMailMonitor",
      "private": true
    }]
  }
}
```

### Publication

```bash
# Build + publish
npm run build

# Ou avec electron-builder
electron-builder --publish always
```

## Interface Utilisateur

### Toasts de Notification

- **Vérification** : Silencieux (console seulement)
- **Disponible** : Toast vert avec version
- **Téléchargement** : Toast bleu avec progression
- **Erreur** : Toast rouge avec message
- **En attente** : Toast info (installation au prochain démarrage)

### Dialogues

**Installation prête :**
```
┌─────────────────────────────────────┐
│ Mise à jour prête                   │
├─────────────────────────────────────┤
│ La version X a été téléchargée.     │
│                                     │
│ Nouveautés:                         │
│ - Feature 1                         │
│ - Bug fix 2                         │
│                                     │
│ Voulez-vous redémarrer ?            │
├─────────────────────────────────────┤
│ [Redémarrer]  [Plus tard]          │
└─────────────────────────────────────┘
```

## Tests

### Test Local

1. **Mock Server** :
   ```bash
   npm install -g http-server
   cd dist
   http-server -p 8080
   ```

2. **dev-app-update.yml** (déjà présent) :
   ```yaml
   provider: generic
   url: http://localhost:8080
   ```

3. **Tester** :
   - Changer version dans `package.json` (ex: 1.0.0 → 0.9.0)
   - Lancer l'app : `npm start`
   - Créer fake release dans `dist/` avec version supérieure

### Test Production

1. Créer une release GitHub avec tag `v1.0.1`
2. Builder l'app avec version `1.0.0`
3. Lancer et vérifier la détection

## Troubleshooting

### Problème : Pas de vérification

**Vérifier :**
- `autoUpdater.updateConfigPath` dans logs
- Token GitHub si dépôt privé
- Connexion internet
- Firewall/Proxy

**Logs à consulter :**
```
Onglet Logs → Catégorie: INIT → Rechercher "MAJ" ou "update"
```

### Problème : Timeout

**Causes possibles :**
- Réseau lent
- GitHub API rate limit
- Fichiers trop volumineux

**Solutions :**
- Augmenter `requestTimeout` dans updateManager.config
- Vérifier les limites API GitHub
- Réduire la taille des assets

### Problème : Erreur 404

**Causes :**
- Dépôt privé sans token
- Owner/Repo incorrect dans package.json
- Release non publiée (draft)

**Solution :**
```json
{
  "build": {
    "publish": {
      "owner": "VotreUsername",  // ← Vérifier
      "repo": "VotreRepo"        // ← Vérifier
    }
  }
}
```

## Sécurité

### Headers de Sécurité

- Authorization header seulement si token présent
- Token jamais loggé en clair
- Cache désactivé pour éviter versions obsolètes

### Validation

`electron-updater` vérifie automatiquement :
- Signature des fichiers (Windows: Authenticode)
- Checksums (SHA512)
- latest.yml integrity

### Recommandations

1. **Toujours signer les binaires** (Windows/macOS)
2. **Utiliser HTTPS** pour le provider
3. **Tester sur release pre-release d'abord**
4. **Monitorer les logs** après déploiement

## Performance

### Optimisations

- **Coalescing** : minCheckInterval évite checks trop fréquents
- **Background** : Téléchargement non-bloquant
- **Timeout** : Évite freeze de l'app
- **Retry backoff** : Évite spam en cas d'erreur

### Impact

| Opération | Impact | Durée |
|-----------|--------|-------|
| Vérification | Minimal | ~1-3s |
| Téléchargement | Variable | Dépend de la taille |
| Installation | Bloquant | ~5-10s |

## Roadmap

### Améliorations futures

- [ ] **Delta updates** : Télécharger seulement les différences
- [ ] **Channel selection** : Beta/Stable/Canary
- [ ] **Rollback** : Revenir à version précédente
- [ ] **Silent updates** : Installation sans dialog
- [ ] **Progressive rollout** : 10% → 50% → 100%
- [ ] **Metrics** : Statistiques d'adoption des versions

## Support

Pour plus d'informations :
- [electron-updater docs](https://www.electron.build/auto-update)
- [GitHub Releases API](https://docs.github.com/en/rest/releases)
- Logs internes : Onglet Logs de l'application
