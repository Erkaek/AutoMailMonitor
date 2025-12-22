# 📊 Guide d'utilisation des Logs

## Accès rapide

Cliquez sur **Logs** dans le menu latéral pour accéder à la nouvelle interface de logs.

## Interface

### En-tête
- **Statistiques en temps réel** : Compteurs par niveau (DEBUG, INFO, SUCCESS, WARN, ERROR)
- **Total** : Nombre total de logs

### Filtres

#### 1. Niveau minimum
Sélectionnez le niveau minimum à afficher :
- **Tous les niveaux** : Affiche tous les logs
- **DEBUG et plus** : DEBUG, INFO, SUCCESS, WARN, ERROR
- **INFO et plus** : INFO, SUCCESS, WARN, ERROR
- **WARN et plus** : WARN, ERROR
- **ERROR seulement** : Uniquement les erreurs

#### 2. Catégorie
Filtrez par type d'opération :
- **Toutes** : Tous les types
- **Init** 🚀 : Initialisation de l'application
- **Sync** 🔄 : Synchronisation des emails
- **COM** 📡 : Communications COM avec Outlook
- **DB** 💾 : Opérations base de données
- **PowerShell** ⚙️ : Scripts PowerShell
- **IPC** 📱 : Communication inter-processus
- **Config** ⚙️ : Configuration
- **Weekly** 📅 : Statistiques hebdomadaires
- **Start/Stop** ▶️⏹️ : Démarrage/arrêt

#### 3. Recherche
Recherche textuelle instantanée dans :
- Messages de logs
- Données associées (JSON, erreurs, etc.)

### Options

- **Défilement automatique** ✅ : Active/désactive le scroll automatique vers les nouveaux logs
- **Pause** ⏸️ : Met en pause la réception des nouveaux logs (utile pour analyser)
- **Export** 💾 : Exporte les logs filtrés en fichier texte
- **Effacer** 🗑️ : Supprime tous les logs de la mémoire

## Codes couleur

- 🔍 **Gris** : DEBUG - Informations de débogage
- ℹ️ **Bleu** : INFO - Informations générales
- ✅ **Vert** : SUCCESS - Opérations réussies
- ⚠️ **Jaune** : WARN - Avertissements
- ❌ **Rouge** : ERROR - Erreurs

## Astuces

1. **Diagnostic de problèmes** :
   - Filtrer sur ERROR pour voir uniquement les erreurs
   - Utiliser la recherche pour trouver un dossier ou email spécifique

2. **Monitoring de performance** :
   - Filtrer sur PERF pour voir les métriques de performance
   - Surveiller les catégories SYNC et DB

3. **Débogage** :
   - Passer en mode DEBUG pour voir tous les détails
   - Mettre en pause pour analyser sans nouveaux logs qui arrivent

4. **Export pour support** :
   - Filtrer les logs pertinents
   - Exporter en fichier texte
   - Envoyer au support pour diagnostic

## Raccourcis

- **Filtre rapide** : Changez la catégorie pour focus sur un composant
- **Recherche en temps réel** : Tapez pendant que les logs arrivent
- **Pause/Reprendre** : Un clic pour figer l'affichage

## Limitations

- **Historique** : 2000 derniers logs conservés
- **Affichage** : 500 logs maximum visibles simultanément
- **Rafraîchissement** : Automatique en temps réel
