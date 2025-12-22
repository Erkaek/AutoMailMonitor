# Guide de Test - Système de Mise à Jour

## Test Rapide en Local

### 1. Préparation

Créez un serveur local pour simuler GitHub Releases :

```bash
# Installer http-server globalement (une seule fois)
npm install -g http-server

# Créer un dossier de test
mkdir -p /tmp/update-test
cd /tmp/update-test
```

### 2. Créer des Faux Fichiers de Release

```bash
# Créer un fichier latest.yml simulé
cat > latest.yml << 'EOF'
version: 1.0.1
releaseDate: 2024-01-15T10:00:00.000Z
files:
  - url: AutoMailMonitor-Setup-1.0.1.exe
    sha512: fake-sha512-hash
    size: 50000000
path: AutoMailMonitor-Setup-1.0.1.exe
sha512: fake-sha512-hash
releaseNotes: |
  ## Nouveautés v1.0.1
  - Amélioration du système de mise à jour
  - Corrections de bugs
  - Performance optimisée
EOF

# Créer un fichier exe vide (juste pour le test)
touch AutoMailMonitor-Setup-1.0.1.exe
```

### 3. Démarrer le Serveur Local

```bash
# Dans /tmp/update-test
http-server -p 8080 --cors

# Le serveur est maintenant accessible sur http://localhost:8080
```

### 4. Configuration de l'App

Le fichier `dev-app-update.yml` est déjà configuré pour pointer sur localhost.

Vérifiez son contenu :
```yaml
provider: generic
url: http://localhost:8080
```

### 5. Modifier la Version de l'App

Dans `package.json`, changez temporairement la version :

```json
{
  "version": "1.0.0"  // ← Plus petit que 1.0.1 du serveur
}
```

### 6. Lancer l'Application

```bash
cd /workspaces/AutoMailMonitor
npm start
```

### 7. Observer les Logs

1. **Console** : Vérifiez les messages de vérification
   ```
   🔍 Vérification des mises à jour...
   🎉 Mise à jour disponible: v1.0.1
   ```

2. **Onglet Logs** dans l'app :
   - Filtrer par catégorie `INIT`
   - Rechercher "MAJ" ou "update"

3. **Toast** : Une notification doit apparaître en haut à droite

### 8. Vérification Manuelle

1. Ouvrir l'onglet **Paramètres**
2. Cliquer sur **"Vérifier les mises à jour"**
3. Observer le résultat dans la console/toast

## Test avec GitHub Release Réelle

### 1. Créer une Release de Test

```bash
# Via GitHub CLI (gh)
gh release create v1.0.1-test \
  --title "Test Release v1.0.1" \
  --notes "Release de test pour auto-update" \
  --prerelease

# Ou manuellement sur https://github.com/Erkaek/AutoMailMonitor/releases/new
```

### 2. Builder et Publier

```bash
# Build de l'application
npm run build

# Ou avec publication automatique
npm run build -- --publish always
```

### 3. Tester avec Version Inférieure

```json
// package.json
{
  "version": "1.0.0"  // ← Inférieur à 1.0.1-test
}
```

```bash
npm start
```

### 4. Vérifier le Téléchargement

Observez dans les logs :
```
⬇️ Téléchargement: 10%
⬇️ Téléchargement: 20%
...
✅ Mise à jour v1.0.1-test téléchargée
```

## Scénarios de Test

### ✅ Scénario 1 : Mise à jour disponible

**Setup :**
- App version : 1.0.0
- Release version : 1.0.1

**Résultat attendu :**
1. Toast vert : "Mise à jour v1.0.1 disponible !"
2. Téléchargement automatique
3. Toast bleu : progression (10%, 20%...)
4. Dialog : "Redémarrer maintenant ?"

### ✅ Scénario 2 : Aucune mise à jour

**Setup :**
- App version : 1.0.1
- Release version : 1.0.1

**Résultat attendu :**
1. Log : "Aucune mise à jour disponible"
2. Pas de toast (silencieux)

### ✅ Scénario 3 : Erreur réseau

**Setup :**
- Serveur arrêté ou URL invalide

**Résultat attendu :**
1. Log : "Erreur lors de la vérification"
2. Retry automatique (3x avec délai)
3. Toast rouge après échec final

### ✅ Scénario 4 : Timeout

**Setup :**
- Serveur très lent (simuler avec `tc` ou throttling)

**Résultat attendu :**
1. Timeout après 30 secondes
2. Retry automatique
3. Log détaillé de l'erreur

### ✅ Scénario 5 : Installation reportée

**Setup :**
- Mise à jour téléchargée
- Cliquer "Plus tard" dans le dialog

**Résultat attendu :**
1. Toast info : "Sera installée au prochain démarrage"
2. Event : `update-pending-restart`
3. Installation automatique au prochain lancement

## Vérifications de Sécurité

### ✅ Token GitHub non exposé

```bash
# Rechercher le token dans les logs
grep -r "ghp_" /path/to/app/logs

# Résultat attendu : RIEN (token jamais loggé)
```

### ✅ HTTPS uniquement en production

```javascript
// package.json build config
{
  "publish": {
    "provider": "github",  // ← Force HTTPS
    "private": true
  }
}
```

### ✅ Signature des binaires

```bash
# Windows : Vérifier Authenticode
signtool verify /pa dist/MonApp-Setup.exe

# macOS : Vérifier signature
codesign --verify --verbose dist/MonApp.dmg
```

## Debug

### Activer les Logs Détaillés

```bash
# Variable d'environnement
export DEBUG=electron-updater

npm start
```

### Inspecter latest.yml

```bash
# Télécharger manuellement
curl -v https://github.com/Erkaek/AutoMailMonitor/releases/latest/download/latest.yml

# ou depuis localhost
curl http://localhost:8080/latest.yml
```

### Vérifier la Configuration

Dans la console de l'app :

```javascript
// Afficher la config de l'auto-updater
console.log(require('electron-updater').autoUpdater.updateConfigPath);
```

## Checklist Avant Release

- [ ] Version correcte dans `package.json`
- [ ] Changelog/Release notes rédigées
- [ ] Binaires signés (Windows/macOS)
- [ ] Test sur version inférieure fonctionnel
- [ ] Logs vérifiés (pas d'erreurs)
- [ ] Token GitHub configuré si repo privé
- [ ] `dev-app-update.yml` dans `.gitignore` (ne pas publish)
- [ ] Tests sur tous les OS cibles (Win/Mac/Linux)

## Rollback d'Urgence

Si une release pose problème :

### Option 1 : Dépublier la Release

```bash
# Via GitHub CLI
gh release delete v1.0.1 --yes

# Ou sur https://github.com/Erkaek/AutoMailMonitor/releases
```

### Option 2 : Release Corrective Immédiate

```bash
# Créer v1.0.2 rapidement avec fix
gh release create v1.0.2 \
  --title "Hotfix v1.0.2" \
  --notes "Correction urgente du bug X"
```

Les utilisateurs téléchargeront directement v1.0.2.

## Monitoring Post-Release

### Vérifier l'Adoption

```bash
# Analyser les téléchargements
gh release view v1.0.1 --json assets

# Outputs
{
  "assets": [
    {
      "name": "AutoMailMonitor-Setup-1.0.1.exe",
      "downloadCount": 42  // ← Nombre de téléchargements
    }
  ]
}
```

### Surveiller les Erreurs

1. Demander aux utilisateurs de partager leurs logs
2. Chercher pattern d'erreur commun
3. Préparer hotfix si nécessaire

## Questions Fréquentes

### Q: Pourquoi ma mise à jour n'est pas détectée ?

**Réponses possibles :**
- Version dans `package.json` >= version release
- Repo privé sans token GitHub
- Release est un draft (non publiée)
- Firewall bloque GitHub API

### Q: Comment forcer une vérification ?

```javascript
// Dans la console DevTools
electronAPI.checkUpdatesNow().then(console.log)
```

Ou via l'onglet Paramètres → Bouton "Vérifier les MAJ".

### Q: L'installation échoue, pourquoi ?

**Causes fréquentes :**
- App déjà ouverte (fermer toutes les instances)
- Droits admin requis (Windows)
- Antivirus bloque (whitelist l'app)
- Fichier corrompu (re-télécharger)

### Q: Comment tester sans publier sur GitHub ?

Utilisez le serveur local (voir section Test Rapide).

## Support

En cas de problème :

1. **Consulter les logs** : Onglet Logs → Catégorie INIT
2. **Activer debug** : `DEBUG=electron-updater npm start`
3. **Vérifier latest.yml** : Download manuel pour inspecter
4. **Tester version locale** : http-server + dev-app-update.yml

Pour plus d'aide : voir [AUTO_UPDATE_SYSTEM.md](./AUTO_UPDATE_SYSTEM.md)
