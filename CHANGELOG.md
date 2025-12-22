# Changelog

Toutes les modifications notables de ce projet seront documentées dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/),
et ce projet adhère au [Semantic Versioning](https://semver.org/lang/fr/).

## [Non publié]

### Ajouté

#### Système de Logs Amélioré (2024-01-15)
- **logService.js** : Service centralisé de logging avec niveaux et catégories
  - Niveaux : DEBUG, INFO, SUCCESS, WARN, ERROR
  - Catégories : INIT, SYNC, COM, DB, PS, IPC, CONFIG, WEEKLY, EMAIL, PERF, SECURITY, CACHE, START, STOP, AUTO
  - Limitation automatique à 2000 entrées en mémoire
  - Streaming temps réel vers l'interface
  
- **logs.html** : Interface moderne de visualisation des logs
  - Filtrage par niveau (ALL, DEBUG, INFO, WARN, ERROR)
  - Filtrage par catégorie avec sélection multiple
  - Recherche en temps réel (debounced 300ms)
  - Statistiques par niveau (cartes colorées)
  - Export en fichier texte
  - Auto-scroll avec pause/resume
  - Navigation depuis l'interface principale
  
- **logs.js** : Logique frontend pour la gestion des logs
  - Chargement paginé de l'historique
  - Application de filtres côté client
  - Mise à jour temps réel via IPC
  - Gestion de l'auto-scroll intelligent

- **IPC Handlers** pour les logs :
  - `api-get-log-history` : Récupère l'historique avec filtres
  - `api-clear-logs` : Efface tous les logs
  - `api-get-log-stats` : Obtient les statistiques
  - Stream temps réel : `log-realtime`

- **Documentation** :
  - `docs/LOGS_SYSTEM.md` : Documentation technique complète
  - `docs/LOGS_USER_GUIDE.md` : Guide utilisateur pour l'interface

#### Système de Mise à Jour Automatique Robuste (2024-01-15)
- **updateManager.js** : Gestionnaire centralisé des mises à jour
  - Configuration complète (timeouts, retry, intervalles)
  - Retry automatique avec backoff exponentiel (3 tentatives max)
  - Timeout protection (30 secondes par requête)
  - Logging détaillé via logService
  - Support dépôts GitHub privés (token authentication)
  - Événements IPC détaillés pour l'UI
  
- **Événements de Mise à Jour** :
  - `update-checking` : Vérification démarrée
  - `update-available` : Mise à jour trouvée (version, releaseNotes)
  - `update-not-available` : Aucune mise à jour
  - `update-error` : Erreur rencontrée
  - `update-download-progress` : Progression téléchargement (%)
  - `update-pending-restart` : Installation reportée

- **Notifications UI** :
  - Toasts Bootstrap pour informer l'utilisateur
  - Progression de téléchargement en temps réel
  - Dialogues améliorés avec release notes
  - Gestion de "Plus tard" (installation au prochain démarrage)

- **Configuration** :
  - Vérification au démarrage (non-bloquante)
  - Vérification périodique (2 heures, configurable)
  - Vérification manuelle (bouton dans Paramètres)
  - Respect du délai minimum entre checks (5 min)

- **Documentation** :
  - `docs/AUTO_UPDATE_SYSTEM.md` : Documentation technique complète
  - `docs/AUTO_UPDATE_TESTING.md` : Guide de test détaillé
  - `docs/README.md` : Index de toute la documentation

### Modifié

#### Services
- **unifiedMonitoringService.js** :
  - Intégration avec logService (dual logging ancien/nouveau)
  - Ajout de `forceFullResync()` pour resynchronisation complète
  - Ajout de `getFolderSyncIdentity()` pour extraction identifiants dossiers
  - Ajout de `safeGetFolderSyncState()` avec gestion d'erreur
  - Correction de `performCompleteSync()` (logique de pagination)

- **outlookConnector.js** :
  - Migration vers logService pour tous les logs
  - Catégories appropriées (COM, INIT, PS)
  - Conservation du système legacy en parallèle

#### Interface Principale
- **index.html** :
  - Lien vers logs.html dans la sidebar
  - Icône journal-text pour l'onglet Logs
  - Bouton "Resync complète" dans l'onglet Monitoring

- **app.js** :
  - Ajout de `setupUpdateListeners()` pour événements MAJ
  - Fonction `forceFullResync()` avec confirmation utilisateur
  - Système de toasts pour notifications
  - Gestion de la progression de téléchargement

- **preload.js** :
  - Exposition des événements de mise à jour au renderer
  - Bridge sécurisé via contextBridge

#### Processus Principal
- **main/index.js** :
  - Import de updateManager et logService
  - Remplacement de l'ancien code auto-update par updateManager
  - Simplification de la vérification périodique
  - Connexion updateManager à mainWindow pour notifications IPC
  - Ajout IPC handler `api-force-full-resync`
  - Streaming logs vers renderer

### Corrigé

#### Bugs Critiques
- **Historique des emails manquant** :
  - `safeGetFolderSyncState()` était appelée mais jamais définie
  - Ajout de 3 méthodes manquantes dans unifiedMonitoringService
  - Correction de la logique de sync baseline/incrémentale

- **Erreur JavaScript main/index.js** :
  - Bloc de code orphelin causant SyntaxError
  - Suppression du return isolé

- **Erreur PowerShell** :
  - `Try-ParseDate` appelée avant définition
  - Déplacement de la fonction en haut du script
  - Ordre de définition respecté

### Sécurité

- **Token GitHub** :
  - Support fichier bundlé : `src/main/updaterToken.js`
  - Support variables d'env : GH_TOKEN, UPDATER_TOKEN, ELECTRON_UPDATER_TOKEN
  - Token jamais loggé en clair
  - Header Authorization seulement si token présent

- **Validation des Mises à Jour** :
  - Vérification automatique des signatures (Authenticode Windows)
  - Checksums SHA512 validés par electron-updater
  - HTTPS forcé pour dépôts GitHub

### Performance

- **Optimisations Logs** :
  - Limitation mémoire (2000 entrées max)
  - Debounce search (300ms)
  - Filtrage côté client pour réactivité
  - Streaming au lieu de polling

- **Optimisations Updates** :
  - Coalescing des checks (minCheckInterval 5min)
  - Téléchargement non-bloquant
  - Timeout pour éviter freeze
  - Backoff exponentiel sur retry

### Documentation

Création de 4 nouveaux documents dans `docs/` :

1. **LOGS_SYSTEM.md** (technique) :
   - Architecture du logService
   - API complète
   - Intégration dans les services
   - Événements IPC

2. **LOGS_USER_GUIDE.md** (utilisateur) :
   - Guide d'utilisation de l'interface
   - Filtrage et recherche
   - Export et partage
   - Troubleshooting

3. **AUTO_UPDATE_SYSTEM.md** (technique) :
   - Architecture updateManager
   - Configuration et paramètres
   - Flux de mise à jour
   - Gestion erreurs et retry
   - Sécurité et validation
   - Troubleshooting avancé

4. **AUTO_UPDATE_TESTING.md** (QA) :
   - Test local avec http-server
   - Test avec GitHub Releases
   - Scénarios de test complets
   - Checklist avant release
   - Procédures de rollback
   - Monitoring post-release

5. **docs/README.md** (index) :
   - Vue d'ensemble de la documentation
   - Quick start par rôle (dev/QA/user)
   - Structure du code
   - Workflows importants

---

## [1.0.0] - Date non définie

### Ajouté
- Application Electron de surveillance Outlook
- Synchronisation automatique des emails
- Base de données Better-SQLite3
- Interface Bootstrap 5
- Scripts PowerShell pour COM Outlook
- Onglet suivi hebdomadaire
- Import activité XLSB
- Performances personnelles
- Monitoring temps réel

### Infrastructures
- electron-updater pour mises à jour auto
- Better-SQLite3 pour performance
- PowerShell COM pour Outlook
- IPC main/renderer pour communication

---

## Format des Entrées

Chaque changement doit être catégorisé sous :
- **Ajouté** : Nouvelles fonctionnalités
- **Modifié** : Changements de fonctionnalités existantes
- **Déprécié** : Fonctionnalités qui seront supprimées
- **Retiré** : Fonctionnalités supprimées
- **Corrigé** : Corrections de bugs
- **Sécurité** : Changements liés à la sécurité

---

## Liens

- [Dépôt GitHub](https://github.com/Erkaek/AutoMailMonitor)
- [Documentation](./docs/README.md)
- [Releases](https://github.com/Erkaek/AutoMailMonitor/releases)
