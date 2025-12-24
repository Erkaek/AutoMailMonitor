# 🗂️ Validation des Chemins de Dossiers - Résumé des Changements

## 🎯 Problème Résolu
Les chemins incomplets dans `folder_configurations` (ex: `"11- Tanguy"` sans préfixe `"FlotteAuto\Boîte de réception\"`) causaient l'erreur:
```
❌ Erreur synchronisation dossier 11- Tanguy: Dossier introuvable (EntryID ou chemin)
```

Ces chemins **ne peuvent jamais fonctionner** avec Outlook car:
- Outlook a besoin du chemin complet (avec hiérarchie boîte→dossier→sous-dossier)
- OU des IDs uniques (storeId + entryId)

## ✅ Corrections Appliquées

### 1. **Validation au Chargement** (`unifiedMonitoringService.js`)
- Les chemins trop courts (sans antislash ET sans IDs) sont **automatiquement filtrés** au démarrage
- Chaque chemin invalide supprimé est **loggé** et **nettoyé de la BDD en arrière-plan**
- Log visible lors du démarrage:
  ```
  📁 8 dossiers configurés pour le monitoring (2 orphelins supprimés: "11- Tanguy", "Mails simples"...)
  ```

### 2. **Validation à l'Ajout** (`main/index.js`)
- Impossible d'ajouter un dossier avec un chemin invalide via l'UI
- Message d'erreur clair indiquant le chemin correct à utiliser:
  ```
  ❌ Dossier invalide: "11- Tanguy" (chemin trop court, pas de boîte-mère). 
  Sélectionnez le chemin complet (ex: "FlotteAuto\Boîte de réception\11- Tanguy") 
  ou fournissez les IDs Outlook (storeId + entryId).
  ```

## 🧹 Nettoyage Manuel (Optionnel)

Deux options:

### Option A: Nettoyage Automatique (Recommandé)
Le nettoyage se fait **automatiquement** lors du démarrage:
1. Redémarrez l'application
2. Observez les logs: `🗑️ [CLEANUP] Dossier orpheline...`
3. Les chemins invalides sont supprimés de la BDD

### Option B: Nettoyage Via Script
Si vous voulez nettoyer manuellement:
```bash
node scripts/cleanup-invalid-folder-paths.js
```

Affiche:
```
🔍 Scanning folder_configurations pour les chemins invalides...
⚠️ Trouvé 2 chemins invalides (orphelins):
  1. "11- Tanguy"
  2. "Mails simples"
🗑️ Suppression des 2 chemins invalides...
✅ 2 lignes supprimées.
ℹ️ Dossiers monitorés restants: 8
```

## 📋 Chemins Valides vs Invalides

### ❌ INVALIDES (seront rejetés/supprimés)
```
"11- Tanguy"                    # Pas de boîte-mère
"Mails simples"                 # Idem
"dossier seul"                  # Idem
```

### ✅ VALIDES (acceptés)
```
"FlotteAuto\Boîte de réception\11- Tanguy"           # Chemin complet
"FlotteAuto\Boîte de réception\11- Tanguy\1- DECLARATION"  # Sous-dossier
"storeId123|entryIdABC"         # Résolution par IDs (si utilisé)
```

## 🔄 Comportement Post-Correction

1. **Au Démarrage**: Scan des chemins, filtrage automatique des invalides
2. **Pendant le Sync**: Seuls les chemins valides sont synchronisés
3. **Logs**: Messages clairs indiquant les chemins nettoyés
4. **Redémarrage**: Les dossiers monitorés restants fonctionnent normalement

## ⚙️ Détails Techniques

### Validation Effectuée
```javascript
// Un chemin est VALIDE si:
const hasBackslash = folderPath.includes('\\') || folderPath.includes('/');
const hasIds = storeId && entryId;
const isValid = hasBackslash || hasIds;
```

### Point d'Entrée du Nettoyage
- `src/services/unifiedMonitoringService.js` ligne ~310
- `loadMonitoredFolders()` filtre + nettoie à chaque chargement

### Suppression BDD
```javascript
dbService.deleteFolderConfiguration(folderPath)  // Non-bloquant
```

## 📌 Résumé des Fichiers Modifiés

| Fichier | Changement |
|---------|-----------|
| `unifiedMonitoringService.js` | Filtre paths invalides + log nettoyage |
| `index.js` (main) | Validation à l'ajout via UI |
| `cleanup-invalid-folder-paths.js` | Script de nettoyage manuel |

## 🚀 Prochaines Étapes
Après correction:
1. Redémarrez l'application → logs de nettoyage visibles
2. Vérifiez que les bons dossiers restent monitorés
3. Tentez une nouvelle synchronisation → pas d'erreur "Dossier introuvable"
