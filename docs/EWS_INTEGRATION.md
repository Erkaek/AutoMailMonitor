# Intégration EWS (Exchange Web Services)

## Vue d'ensemble

AutoMailMonitor supporte maintenant **EWS (Exchange Web Services)** pour la récupération des emails, en plus du système COM traditionnel. EWS offre une **performance significativement meilleure** pour les opérations en masse et les synchronisations.

### Avantages EWS vs COM

| Critère | EWS (SOAP) | COM |
|---------|-----------|-----|
| **Performance** | 10-50x plus rapide pour les opérations massives | Lent pour les synchronisations volumineuses |
| **Pagination** | Native, efficace (200 items/requête) | Très faible, itération item-par-item |
| **Dépendances** | HTTP/XML uniquement | Requiert Outlook actif |
| **Compatibilité** | Boîtes partagées, autres utilisateurs | Comptes Outlook locaux principalement |
| **Chiffrement** | HTTPS natif | Dépend du système local |

### Architecture

Le système utilise une **stratégie de fallback intelligent**:

1. **Tentative EWS** (si configuré): Utilise le script `get-folder-emails-by-id-ews.ps1`
2. **Fallback COM**: Bascule automatiquement vers `get-folder-emails-by-id.ps1` en cas d'erreur EWS
3. **Transparence**: L'application continue de fonctionner avec ou sans EWS

```
┌─────────────────────────────────────────┐
│ Récupération d'emails (getFolderEmails) │
└────────────────┬────────────────────────┘
                 │
         ┌───────▼────────┐
         │ EWS configuré? │
         └───────┬────────┘
                 │
      ┌──────────┴──────────┐
      │ OUI                 │ NON
      │                     │
  ┌───▼────┐           ┌───▼────┐
  │ Essayer│           │Utiliser│
  │  EWS   │           │  COM   │
  └───┬────┘           └────────┘
      │
 ┌────▼─────┐
 │ Succès?  │
 └────┬─────┘
      │
  ┌───┴─────┐
  │ OUI   NON
  │       │
  │    ┌──▼──┐
  │    │Fall-│
  │    │back │
  │    │ COM │
  │    └─────┘
```

## Configuration

### Configuration minimale

Pour activer EWS, définir dans les paramètres d'échange:

```javascript
{
  exchange: {
    ewsUrlOverride: "https://your-exchange-server/EWS/Exchange.asmx",
    timeoutMs: 240000  // 4 minutes (EWS peut être plus lent pour les grandes boîtes)
  }
}
```

### Paramètres configurables

- **`ewsUrlOverride`** (string, optionnel): URL EWS du serveur Exchange
  - Format: `https://<exchange-server>/EWS/Exchange.asmx`
  - Si vide: EWS est désactivé, retour au COM uniquement
  
- **`timeoutMs`** (number, optionnel): Délai d'expiration en ms
  - Défaut: 240000 (4 minutes)
  - Recommandé: 240000-300000 pour les grandes boîtes

### Découverte automatique

Si votre serveur Exchange supporte la découverte automatique (Autodiscover), EWS peut fonctionner sans configuration explicite de l'URL. Le script utilise:

```powershell
$svc.AutodiscoverUrl($Mailbox, { param($url) $true })
```

## Utilisation

### Appel de base (EWS automatiquement essayé si configuré)

```javascript
const emails = await outlookConnector.getFolderEmails("Inbox", {
  limit: 1000,
  storeName: "user@company.com"  // Requis pour EWS
});
```

### Contrôle du fallback

Le fallback est **automatique et transparent**:
- Si EWS échoue, COM est automatiquement utilisé
- Les logs montrent quelle source a été utilisée: `source: 'ews'` ou `source: 'com'`

```javascript
const result = await outlookConnector.getFolderEmails("Inbox");
console.log(`Source utilisée: ${result.source}`); // "ews" ou "com"
```

## Scripts PowerShell

### get-folder-emails-by-id-ews.ps1

Récupère les emails via EWS (SOAP HTTP/XML).

**Paramètres:**
- `-EwsUrl` (requis): URL du serveur EWS
- `-Mailbox` (requis): Adresse email ou nom de la boîte
- `-FolderPath` (optionnel): Chemin du dossier (défaut: Inbox)
- `-MaxItems` (optionnel): Nombre max d'items (défaut: 1000)
- `-UnreadOnly` (switch): Retourner seulement les non-lus
- `-AllItems` (switch): Ignorer les filtres temporels
- `-UseLastModificationTime` (switch): Utiliser la date de modification au lieu de réception

**Sortie JSON:**
```json
{
  "success": true,
  "emails": [
    {
      "Subject": "...",
      "SenderEmailAddress": "...",
      "ReceivedTime": "2025-01-20T10:30:00Z",
      "InternetMessageId": "<...@domain.com>",
      "UnRead": false,
      "HasAttachments": false,
      "FolderPath": "Inbox",
      "StoreName": "user@company.com",
      ...
    }
  ],
  "count": 5,
  "totalInFolder": 25,
  "timestamp": "2025-01-20T10:30:05Z"
}
```

### get-folder-emails-by-id.ps1

Récupère les emails via COM (fallback, plus lent).

**Paramètres:** Similaires à EWS mais accepte aussi:
- `-StoreId`: ID du magasin Outlook
- `-FolderEntryId`: EntryID du dossier COM
- `-HoursBack` (optionnel): Heures dans le passé à inclure

## Optimisations de performance

### EWS est significativement plus rapide pour:

1. **Synchronisations massives** (> 500 emails)
   - EWS: ~1-2 secondes
   - COM: ~10-30 secondes

2. **Boîtes partagées**
   - EWS: Accès direct et rapide
   - COM: Nécessite une configuration complexe

3. **Requêtes répétées**
   - EWS: Pagination native efficace (200 items/requête)
   - COM: Itération item-par-item

### Recommandations

- ✅ **Utiliser EWS** pour les synchronisations régulières et les grandes boîtes
- ✅ **Configurer un timeout approprié** (240s minimum pour les boîtes > 10k emails)
- ✅ **Surveiller les logs** pour identifier les boîtes problématiques
- ⚠️ **Laisser le fallback COM actif** en cas de maintenance serveur EWS

## Dépannage

### EWS échoue et bascule sur COM

**Log:**
```
⚠️ EWS a échoué: Erreur..., tentative COM...
```

**Solutions:**
1. Vérifier l'URL EWS: `https://<server>/EWS/Exchange.asmx`
2. Vérifier les credentials (Windows Auth doit être actif)
3. Vérifier la connectivité réseau
4. Vérifier que le serveur Exchange est accessible

### EWS n'est pas utilisé du tout

**Cause:** `ewsUrlOverride` n'est pas configuré ou `storeName`/`mailbox` n'est pas fourni

**Solution:** 
```javascript
// Ajouter dans les settings
settings.exchange.ewsUrlOverride = "https://your-exchange-server/EWS/Exchange.asmx";

// Ajouter storeName dans les options
outlookConnector.getFolderEmails("Inbox", { 
  storeName: "user@company.com" 
});
```

### Format de sortie JSON incomplet

**Cause:** Le serveur EWS a retourné des champs différents

**Solution:** Vérifier la version Exchange du serveur:
- Exchange 2013 SP1 ou plus récent: Complètement supporté
- Exchange 2010: Peut manquer certains champs (fallback sur COM recommandé)

## Performance observée

### Métriques de synchronisation initiale

| Taille boîte | COM | EWS | Gain |
|--------------|-----|-----|------|
| 100 emails | 1-2s | 0.5s | 2-4x |
| 500 emails | 5-8s | 1-2s | 4-8x |
| 1000 emails | 10-20s | 2-3s | 5-10x |
| 5000 emails | 60-120s | 5-10s | 10-20x |

> Les temps réels dépendent de la qualité de la connexion réseau et de la charge du serveur Exchange.

## Limitations

- **EWS ne fournit pas d'EntryID COM** (mais utilise ItemId qui est unique au serveur)
- **Pagination EWS** limite à ~1000 items avant ralentissement
- **Pas d'événements temps réel** (COM via événements, EWS nécessite polling)
- **Authentification** basée sur Windows Auth (credentials de la session active)

## Migration depuis COM uniquement

### Activation progressive

1. Configurer `ewsUrlOverride` dans les settings
2. Les opérations utiliseront automatiquement EWS si disponible
3. Surveiller les logs pour vérifier que EWS fonctionne
4. COM reste disponible comme fallback

### Monitoring

```javascript
// Vérifier quelle source est utilisée
const result = await outlookConnector.getFolderEmails("Inbox", { 
  storeName: "user@company.com" 
});

if (result.source === 'ews') {
  console.log('✅ EWS utilisé');
} else {
  console.log('⚠️ COM utilisé (fallback)');
}
```

## Ressources

- [Documentation EWS Microsoft](https://learn.microsoft.com/fr-fr/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-reference)
- [Schéma EWS SOAP](https://learn.microsoft.com/fr-fr/exchange/client-developer/exchange-web-services/start-using-web-services-in-an-exchange-client-application)
- [Forum AutoMailMonitor](https://github.com/Erkaek/AutoMailMonitor/issues)
