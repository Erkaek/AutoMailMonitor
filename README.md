# AutoMailMonitor

Surveillance professionnelle d’emails Outlook (Electron) avec dashboard, suivi hebdo et mises à jour automatiques.

## 🚀 Vue d’ensemble

AutoMailMonitor surveille Outlook et agrège des métriques sans lire le contenu des emails. L’app fournit:

- Dashboard en temps réel (KPIs, graphiques Chart.js)
- Répartition par dossiers, analyse temporelle, performance système
- Suivi hebdomadaire avec commentaires (CRUD) et sélection de semaine
- Monitoring des dossiers Outlook (arbre, catégories Déclarations/Règlements/Mails simples)
- Thèmes clair/sombre, sidebar repliable, UI responsive
- SQLite (better-sqlite3, WAL) + cache
- Mise à jour automatique (builds) + vérification Git au démarrage (dev)

Confidentialité: pas d’analyse du contenu des emails; classification par localisation de dossiers.

## 🧩 Prérequis

- Windows 10/11
- Microsoft Outlook installé et configuré
- Node.js 16+ et Git
- PowerShell (par défaut sous Windows)

## ⚙️ Installation & exécution

```powershell
# Cloner et installer
git clone https://github.com/Erkaek/AutoMailMonitor.git
cd AutoMailMonitor
npm install

# Démarrer en dev
npm start
```

Au premier lancement:

- Configurez les dossiers surveillés depuis l’onglet Monitoring/Paramètres
- Choisissez un thème (clair/sombre) dans la sidebar


## 📦 Build (Electron)

Les builds utilisent electron-builder et publient sur GitHub Releases (déjà configuré dans package.json → build.publish).

```powershell
# Build Windows (NSIS)
npx electron-builder -w

# Publier (si GH_TOKEN configuré)
npx electron-builder -w --publish always
```

Paramètres clés (package.json):

- appId/productName: com.tanguy.mailmonitor / “Mail Monitor”
- publish: GitHub (owner: Erkaek, repo: AutoMailMonitor)
- win.target: nsis

### Signature Windows (certificat)

- Par défaut sur CI, si aucun certificat n'est fourni, un certificat d'authenticode auto-signé est généré et utilisé pour signer l'installeur et l'exécutable.
  - Cela évite les alertes "Éditeur inconnu" (binaire non signé), mais RESTE non approuvé par SmartScreen/Defender (avertissement possible à l'installation).
  - Pour une confiance complète, fournissez un certificat d'éditeur émis par une AC :
    - Secrets requis dans le repo: `CODE_SIGNING_CERT_BASE64` (PFX en base64) et `CODE_SIGNING_CERT_PASSWORD`.
    - Le workflow CI utilise automatiquement ces secrets pour signer avec l'AC.
  - Aucun changement de code n'est nécessaire; electron-builder consomme `CSC_LINK`/`CSC_KEY_PASSWORD` injectés par le workflow.

## 🔄 Mises à jour

- Application packagée: auto-update via GitHub Releases.
  - Vérification au démarrage, puis toutes les 30 minutes.
  - Téléchargement + prompt de redémarrage.
- Mode développement: vérification Git au démarrage.
  - fetch + comparaison HEAD..origin/BRANCHE
  - prompt “pull & redémarrer” si des commits distants existent.

## 🧾 Versionnement (source unique)

- Changez la version UNIQUEMENT dans `package.json` (clé `version`).
- L’UI (footer, À propos) lit dynamiquement `app.getVersion()` via IPC — aucune autre mise à jour nécessaire.

## 🗂️ Structure du projet (simplifiée)

```plaintext
AutoMailMonitor/
├─ public/                 # UI (Bootstrap/Chart.js)
│  ├─ index.html
│  ├─ css/style.css, css/themes.css
│  └─ js/app.js
├─ src/
│  ├─ main/                # Processus principal Electron
│  │  ├─ index.js          # Fenêtres, IPC, auto-update, check Git dev
│  │  ├─ preload.js        # API sécurisée vers le renderer
│  │  └─ preload-loading.js
│  ├─ server/
│  │  └─ outlookConnector.js   # Bridge PowerShell/COM Outlook
│  └─ services/
│     ├─ optimizedDatabaseService.js
│     ├─ cacheService.js
│     └─ (autres services)
├─ resources/new logo/
├─ config/app-settings.json
├─ data/                   # Données locales (ignorées par Git)
└─ package.json
```

## 🧪 Scripts npm utiles

```json
{
  "start": "npx electron .",
  "import-activity": "node src/cli/import-activity.js",
  "postinstall": "electron-builder install-app-deps"
}
```

## 🔧 Dépannage (Windows/Outlook/Git)

- Outlook COM: assurez-vous qu’Outlook est ouvert et configuré. Les chemins PowerShell 32/64 bits sont gérés dans `outlookConnector.js`.
- SQLite verrouillée: fermez les instances de l’app; l’utilisation de WAL réduit les conflits.
- Fins de ligne (LF/CRLF): `.gitattributes` est fourni; vous pouvez normaliser avec:

```powershell
# Stratégie recommandée (LF global, CRLF pour .bat/.ps1)
git config --global core.autocrlf input
git config --global core.eol lf
# Normalisation
git add --renormalize .
git commit -m "chore: normalize line endings" ; git push
```

- .gitignore: si des fichiers ignorés sont déjà suivis, “désuivez-les” sans les supprimer:

```powershell
git rm -r --cached -- data temp logs dist out release archives node_modules
git rm -r --cached -- *.db *.db-wal *.db-shm *.log
git commit -m "chore: stop tracking ignored files" ; git push
```

## 🔐 Confidentialité

- Pas d’accès automatique au contenu des emails
- Classification par localisation (dossier) uniquement
- Métadonnées minimales pour les métriques

## 📣 Astuces UI

````markdown
# Mail Monitor

Surveille des dossiers Outlook et affiche des statistiques simples.

## Installation (dev)

```powershell
git clone https://github.com/Erkaek/AutoMailMonitor.git
cd AutoMailMonitor
npm install
npm start
```

## Utilisation

- Dans l’onglet Monitoring, sélectionnez les dossiers Outlook à surveiller.
- L’app collecte les nouveaux mails et calcule des stats (jour/semaine).

## Import activité (.xlsb)

- Menu Importer: choisissez un fichier .xlsb (feuilles S1..S52).
- Sécurité: priorité à PowerShell/Excel COM; sinon lecteur isolé.
- Pour forcer COM uniquement: définir `XLSB_IMPORT_DISABLE_JS=1`.

## Données & build

- Base locale: `data/emails.db` (SQLite, WAL).
- Build Windows (optionnel): `npx electron-builder -w`.

## Licence & support

Usage interne. © 2025 Tanguy Raingeard.
Issues/PR bienvenues.
````
