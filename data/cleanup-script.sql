-- Script de nettoyage de la base de données AutoMailMonitor
-- Généré automatiquement le 2025-08-03T20:36:16.030Z

-- Sauvegarde des données importantes avant suppression
CREATE TABLE IF NOT EXISTS cleanup_backup AS SELECT 'backup_created' as status, datetime('now') as timestamp;

-- Suppression des tables redondantes/inutiles
DROP TABLE IF EXISTS "email_tracking";
DROP TABLE IF EXISTS "user_settings";
DROP TABLE IF EXISTS "vba_categories";
DROP TABLE IF EXISTS "daily_metrics";
DROP TABLE IF EXISTS "configurations";
DROP TABLE IF EXISTS "settings";

-- Optimisation de la base
VACUUM;
ANALYZE;

-- Vérification finale
SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;
