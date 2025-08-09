@echo off
echo =====================================================
echo     Fix automatique better-sqlite3 - Mail Monitor
echo =====================================================
echo.

:: Vérifier si on est dans le bon répertoire
if not exist "package.json" (
    echo ❌ Erreur: package.json non trouvé
    echo Veuillez exécuter ce script depuis le dossier du projet
    pause
    exit /b 1
)

:: Rebuild better-sqlite3
echo 🔨 Rebuild de better-sqlite3 en cours...
call npm rebuild better-sqlite3

if %ERRORLEVEL% EQU 0 (
    echo ✅ Rebuild terminé avec succès
    echo.
    echo 🧪 Test du module...
    node -e "try { require('better-sqlite3'); console.log('✅ better-sqlite3 fonctionne !'); } catch(e) { console.log('❌ Erreur:', e.message); process.exit(1); }"
    
    if %ERRORLEVEL% EQU 0 (
        echo.
        echo 🎉 better-sqlite3 est maintenant opérationnel !
        echo Vous pouvez lancer l'application avec: npm start
    ) else (
        echo.
        echo ❌ Le module ne fonctionne toujours pas
        echo Essayez: npm install --force
    )
) else (
    echo ❌ Échec du rebuild
    echo Essayez: npm install --force
)

echo.
pause
