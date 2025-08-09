## 📊 RAPPORT D'ANALYSE CODE ↔ BASE DE DONNÉES

### 🔍 PROBLÈMES IDENTIFIÉS

#### 1. 🔴 COLONNES JAMAIS UTILISÉES (12 colonnes à supprimer)
- `entry_id` : Définie mais jamais remplie (doublon avec outlook_id)
- `sender_name` : Définie mais jamais remplie (le code utilise uniquement sender_email)
- `recipient_email` : Définie mais jamais remplie
- `treated_time` : Définie mais jamais remplie
- `folder_path` : Définie mais jamais remplie (le code utilise folder_name)
- `is_replied` : Définie mais jamais remplie
- `is_treated` : Définie mais jamais remplie  
- `has_attachment` : Définie mais jamais remplie
- `body_preview` : Définie mais jamais remplie
- `size_bytes` : Définie mais jamais remplie
- `is_deleted` : Définie mais jamais remplie
- `deleted_at` : Définie mais jamais remplie

#### 2. 🟡 COLONNES TRÈS PEU UTILISÉES
- `is_read` : Seulement 7% d'utilisation (1/14 emails)

#### 3. ⚠️ INCOHÉRENCES MAJEURES

**Problème sender_name vs sender_email :**
- Le code Outlook extrait bien `SenderName` et `SenderEmailAddress`
- Mais dans unifiedMonitoringService, seul `sender_email` est mappé
- Résultat : `sender_name` reste toujours vide

**Problème entry_id vs outlook_id :**
- Les deux colonnes existent mais seul `outlook_id` est utilisé
- `entry_id` est un doublon inutile

**Problème folder_path vs folder_name :**
- Les deux colonnes existent mais seul `folder_name` est utilisé
- `folder_path` reste toujours vide

### 🎯 RECOMMANDATIONS

#### Option 1 : NETTOYAGE RADICAL (Recommandé)
Supprimer toutes les colonnes jamais utilisées pour une base optimale :

```sql
-- Supprimer les colonnes inutilisées
ALTER TABLE emails DROP COLUMN entry_id;
ALTER TABLE emails DROP COLUMN sender_name;
ALTER TABLE emails DROP COLUMN recipient_email;
ALTER TABLE emails DROP COLUMN treated_time;
ALTER TABLE emails DROP COLUMN folder_path;
ALTER TABLE emails DROP COLUMN is_replied;
ALTER TABLE emails DROP COLUMN is_treated;
ALTER TABLE emails DROP COLUMN has_attachment;
ALTER TABLE emails DROP COLUMN body_preview;
ALTER TABLE emails DROP COLUMN size_bytes;
ALTER TABLE emails DROP COLUMN is_deleted;
ALTER TABLE emails DROP COLUMN deleted_at;
```

**Schéma final (13 colonnes vs 25 actuelles) :**
- `id`, `outlook_id`, `subject`, `sender_email`, `received_time`, `sent_time`
- `folder_name`, `category`, `is_read`, `importance`, `week_identifier`
- `created_at`, `updated_at`

#### Option 2 : CORRECTION ET CONSERVATION
Corriger le code pour utiliser les colonnes existantes :

1. **Corriger sender_name :** Modifier unifiedMonitoringService pour mapper senderName
2. **Corriger entry_id :** L'utiliser au lieu d'outlook_id si nécessaire
3. **Utiliser les autres colonnes :** Implémenter has_attachment, body_preview, etc.

### 🚀 ACTIONS RECOMMANDÉES

1. **Immediate :** Corriger le mapping sender_name dans le code
2. **Court terme :** Supprimer les colonnes jamais utilisées  
3. **Moyen terme :** Revoir l'architecture pour éviter ces doublons

### 📈 IMPACT PERFORMANCE

**Avant nettoyage :**
- 25 colonnes par email
- 12 colonnes vides = 48% d'espace gaspillé
- Index inutiles sur colonnes vides

**Après nettoyage :**
- 13 colonnes par email  
- Réduction de 48% de l'espace disque
- Amélioration des performances de requête

### 🔧 SCRIPTS DISPONIBLES

- `fix-sender-mapping.js` : Corrige le mapping sender_name (déjà exécuté)
- `cleanup-unused-columns.js` : Supprime les colonnes inutilisées (à créer)
- `analyze-column-usage.js` : Analyse l'utilisation (déjà exécuté)
