-- Script SQL pour corriger les chemins de dossiers

-- Afficher les chemins actuels
SELECT 'AVANT CORRECTION:' as info;
SELECT folder_path, folder_name, category FROM folder_config ORDER BY folder_path;

-- Corriger les chemins avec les vrais caractères français
UPDATE folder_config 
SET folder_path = 'erkaekanon@outlook.com\Boîte de réception\testA' 
WHERE folder_path = 'erkaekanon@outlook.com\Bo├«te de r├®ception\testA';

UPDATE folder_config 
SET folder_path = 'erkaekanon@outlook.com\Boîte de réception\test' 
WHERE folder_path = 'erkaekanon@outlook.com\Bo├«te de r├®ception\test';

UPDATE folder_config 
SET folder_path = 'erkaekanon@outlook.com\Boîte de réception\test\test-1' 
WHERE folder_path = 'erkaekanon@outlook.com\Bo├«te de r├®ception\test\test-1';

UPDATE folder_config 
SET folder_path = 'erkaekanon@outlook.com\Boîte de réception\testA\test-c' 
WHERE folder_path = 'erkaekanon@outlook.com\Bo├«te de r├®ception\testA\test-c';

-- Afficher les chemins après correction
SELECT 'APRES CORRECTION:' as info;
SELECT folder_path, folder_name, category FROM folder_config ORDER BY folder_path;
