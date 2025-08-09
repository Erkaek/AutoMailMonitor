-- Script pour vérifier le schéma de la table emails
.mode table
.headers on

-- Afficher la structure de la table emails
PRAGMA table_info(emails);

-- Vérifier s'il y a encore une colonne event_type
SELECT name FROM pragma_table_info('emails') WHERE name = 'event_type';
