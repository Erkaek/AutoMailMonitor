-- Script SQL pour configurer le monitoring du dossier Inbox
-- À exécuter manuellement dans un outil SQLite

-- Créer la table si elle n'existe pas
CREATE TABLE IF NOT EXISTS folder_configurations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_path TEXT UNIQUE NOT NULL,
    category TEXT NOT NULL,
    folder_name TEXT,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Ajouter le dossier Inbox pour monitoring
INSERT OR REPLACE INTO folder_configurations (folder_path, category, folder_name, updated_at)
VALUES ('\\Inbox', 'inbox', 'Boîte de réception', CURRENT_TIMESTAMP);

-- Vérifier la configuration
SELECT * FROM folder_configurations;
