# AutoMailMonitor

Application de surveillance des emails Outlook avec classification automatique basée sur la localisation des dossiers.

## 🎯 Vue d'ensemble

AutoMailMonitor est une application de bureau basée sur Electron qui surveille les emails dans Outlook et les classe automatiquement selon leur localisation dans l'arborescence des dossiers, tout en respectant strictement la confidentialité des utilisateurs.

### Fonctionnalités Principales

- 🔒 **Protection de la confidentialité** : Aucun accès automatique au contenu des emails
- 📁 **Classification par localisation** : Classification basée sur la position des dossiers, pas sur le contenu
- 🏷️ **3 catégories simplifiées** : Déclarations, Règlements, Mails simples
- 📊 **Dashboard en temps réel** : Interface web intégrée avec métriques et analytics
- 🗄️ **Base SQLite** : Persistance des données et historique complet
- 🔄 **Monitoring intelligent** : Surveillance continue avec polling optimisé

## 🏗️ Architecture

### Structure du Projet

```
AutoMailMonitor/
├── data/                          # 📁 Données centralisées
│   ├── emails.db                  # Base SQLite principale
│   ├── folders-config.json        # Mapping dossier → catégorie
│   ├── settings.json              # Configuration utilisateur
│   └── vba-categories.json        # Définition des 3 catégories
├── src/
│   ├── main/                      # 🚀 Process principal Electron
│   │   ├── index.js               # Point d'entrée et IPC handlers
│   │   └── preload.js             # Script de sécurité Electron
│   ├── server/                    # 🔗 Interface Outlook
│   │   ├── outlookConnector.js    # Connecteur PowerShell/COM
│   │   ├── server.js              # Serveur Express local
│   │   └── htmlTemplate.js        # Templates HTML
│   ├── services/                  # ⚙️ Services métier
│   │   ├── monitoringService.js   # Service principal de monitoring
│   │   ├── vbaMetricsService.js   # Métriques et analytics
│   │   └── databaseService.js     # Gestion base de données
│   └── utils/                     # 🛠️ Utilitaires
├── public/                        # 🌐 Interface web
│   ├── index.html                 # Interface principale
│   ├── css/style.css              # Styles Bootstrap
│   └── js/app.js                  # Logique frontend
├── docs/                          # 📚 Documentation
│   └── ARCHITECTURE.md            # Architecture détaillée
└── resources/                     # 📋 Ressources et exemples
```

### Services Principaux

#### 1. MonitoringService (`src/services/monitoringService.js`)
- Service principal de surveillance des emails
- Gestion du polling intelligent et des cycles de monitoring
- Synchronisation avec la base de données
- Export singleton pour usage global

#### 2. VBAMetricsService (`src/services/vbaMetricsService.js`)
- Collecte et calcul des métriques email
- Classification automatique par localisation de dossier
- Analytics et rapports de performance
- Compatible avec les macros VBA existantes

#### 3. DatabaseService (`src/services/databaseService.js`)
- Interface SQLite pour persistance des données
- Gestion des schémas et migrations
- Optimisations de performance pour gros volumes

#### 4. OutlookConnector (`src/server/outlookConnector.js`)
- Interface PowerShell vers COM Outlook
- Gestion des connexions et de la robustesse
- Support UTF-8 pour caractères français
- Validation et sécurisation des accès

## 🚀 Installation et Configuration

### Prérequis

- Windows 10/11
- Microsoft Outlook installé et configuré
- Node.js 16+ et npm
- PowerShell avec droits d'exécution

### Installation

1. **Cloner le projet**
   ```bash
   git clone <repository-url>
   cd AutoMailMonitor
   ```

2. **Installer les dépendances**
   ```bash
   npm install
   ```

3. **Démarrer l'application**
   ```bash
   npm start
   ```

### Configuration Initiale

1. **Lancer l'application** : L'interface web s'ouvre automatiquement
2. **Configurer les dossiers** : Aller dans l'onglet "Configuration"
3. **Sélectionner les dossiers Outlook** à surveiller
4. **Assigner les catégories** : Déclarations, Règlements, ou Mails simples
5. **Sauvegarder** : Le monitoring démarre automatiquement

## 📊 Système de Classification

### 3 Catégories Uniques

| Catégorie | Couleur | Icône | Description |
|-----------|---------|-------|-------------|
| **Déclarations** | `#ff6b6b` | 📋 | Documents officiels et déclarations |
| **Règlements** | `#4ecdc4` | ⚖️ | Factures, paiements, règlements |
| **Mails simples** | `#45b7d1` | 📧 | Correspondance générale |

### Logique de Classification

- **Basée sur la localisation** : Classification selon le dossier Outlook où se trouve l'email
- **Configuration flexible** : Mapping personnalisable dans `data/folders-config.json`
- **Pas d'analyse de contenu** : Respect total de la confidentialité

Exemple de configuration :
```json
{
  "\\FlotteAuto\\Boîte de réception\\Déclarations": {
    "category": "declarations",
    "name": "Dossier Déclarations"
  },
  "\\FlotteAuto\\Boîte de réception\\Factures": {
    "category": "reglements", 
    "name": "Dossier Factures"
  }
}
```

## 🔒 Confidentialité et Sécurité

### Principes de Protection

- ✅ **Surveillance des structures de dossiers uniquement**
- ✅ **Aucune lecture automatique du contenu des emails**
- ✅ **Classification basée sur la localisation géographique**
- ✅ **Validation et sanitisation de toutes les entrées**
- ❌ **Aucun accès aux pièces jointes**
- ❌ **Aucune indexation du contenu textuel**

### Données Collectées

1. **Métadonnées uniquement** : Expéditeur, destinataire, sujet, date
2. **Statuts de lecture** : Lu/non-lu, traité/non-traité
3. **Localisation** : Chemin du dossier Outlook
4. **Identifiants** : EntryID Outlook pour synchronisation

## 📈 Métriques et Analytics

### Dashboard Principal

- **Statistiques temps réel** : Emails reçus, traités, non-lus
- **Évolution hebdomadaire** : Tendances et performances
- **Répartition par catégorie** : Distribution des emails
- **Alertes et notifications** : Statut de monitoring

### Métriques Avancées

- **Temps de traitement moyen** par catégorie
- **Volume d'emails par dossier** et période
- **Détection des pics d'activité** et tendances
- **Rapports d'efficacité** et de performance

## 🛠️ Développement

### Architecture Technique

- **Frontend** : HTML5, Bootstrap 5, JavaScript ES6+
- **Backend** : Electron, Node.js, Express
- **Base de données** : SQLite3 avec optimisations
- **Interface Outlook** : PowerShell + COM Automation
- **Communication** : IPC Electron pour sécurité

### Scripts de Développement

```bash
# Démarrage en mode développement
npm run dev

# Tests unitaires
npm test

# Build pour production
npm run build

# Nettoyage des dépendances
npm run clean
```

### Structure des Logs

```
[TIMESTAMP] [SERVICE] [LEVEL] Message
[2024-03-15T10:30:45.123Z] [MonitoringService] [INFO] 📁 3 dossiers configurés
[2024-03-15T10:30:46.456Z] [OutlookConnector] [SUCCESS] ✅ Connexion établie
```

## 🔧 Configuration Avancée

### Variables d'Environnement

```bash
# Mode de logging détaillé
NODE_ENV=development

# Niveau de log (debug, info, warn, error)
LOG_LEVEL=info

# Intervalle de monitoring (ms)
MONITORING_INTERVAL=30000

# Limite d'emails par scan
EMAIL_SCAN_LIMIT=1000
```

### Paramètres de Performance

Dans `data/settings.json` :
```json
{
  "monitoring": {
    "scanInterval": 30000,
    "treatReadEmailsAsProcessed": false,
    "autoStart": true
  },
  "database": {
    "purgeOldDataAfterDays": 365,
    "enableEventLogging": true
  },
  "ui": {
    "emailsLimit": 50,
    "theme": "default",
    "language": "fr"
  }
}
```

## 🐛 Dépannage

### Problèmes Courants

1. **Erreur connexion Outlook**
   ```
   Solution : Vérifier qu'Outlook est ouvert et configuré
   ```

2. **Caractères français corrompus**
   ```
   Solution : Vérifier l'encoding UTF-8 (chcp 65001)
   ```

3. **Monitoring ne démarre pas**
   ```
   Solution : Configurer au moins un dossier dans l'interface
   ```

4. **Base de données corrompue**
   ```
   Solution : Supprimer data/emails.db (reconstruction automatique)
   ```

### Logs de Diagnostic

Les logs sont disponibles dans :
- **Console développeur** : F12 dans l'interface
- **Console Electron** : Terminal de lancement
- **Logs PowerShell** : Sortie des scripts COM

## 📋 Changelog

### Version Actuelle (2024-03)

- ✅ **Architecture nettoyée** : Suppression des services redondants
- ✅ **Centralisation des données** : Migration vers `data/`
- ✅ **Classification simplifiée** : 3 catégories uniquement
- ✅ **Protection confidentialité** : Aucun accès automatique emails
- ✅ **Performance optimisée** : Polling intelligent et cache
- ✅ **Documentation complète** : Architecture et guide utilisateur

### Roadmap

- 🔄 **API REST** : Interface pour intégration externe
- 📱 **Notifications système** : Alertes desktop natives
- 🌐 **Support multi-comptes** : Gestion plusieurs boîtes Outlook
- 📤 **Export rapports** : PDF et Excel
- 🎨 **Thèmes personnalisés** : Interface adaptable

## 📄 Licence

Ce projet est développé pour un usage interne et respecte les politiques de confidentialité en vigueur.

## 🤝 Support

Pour toute question ou problème :
1. Consulter la documentation dans `docs/`
2. Vérifier les logs de diagnostic
3. Contacter l'équipe de développement
