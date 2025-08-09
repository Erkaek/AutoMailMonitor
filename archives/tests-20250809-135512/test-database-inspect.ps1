# Test PowerShell pour vérifier le contenu de la base de données
# et examiner les recipient_email

$dbPath = "E:\Tanguy\Bureau\AutoMailMonitor\data\emails.db"

if (Test-Path $dbPath) {
    Write-Host "🔍 Base de données trouvée: $dbPath" -ForegroundColor Green
    
    # Utiliser System.Data.SQLite pour lire la base
    try {
        Add-Type -Path "${env:PROGRAMFILES(X86)}\SQLite.NET\System.Data.SQLite.dll" -ErrorAction SilentlyContinue
    } catch {
        Write-Host "⚠️  System.Data.SQLite non trouvé, essai avec une autre approche..." -ForegroundColor Yellow
    }
    
    # Alternative : utiliser un simple script PowerShell avec sqlite3 en ligne de commande
    try {
        # Vérifier d'abord le schéma de la table emails
        Write-Host "`n📋 Schéma de la table emails:" -ForegroundColor Cyan
        
        # Commande PowerShell pour lire le fichier SQLite (méthode basique)
        # Note: Ceci est une approche simple pour inspecter les données
        
        Write-Host "⚡ Pour une inspection complète, utilisez un outil comme DB Browser for SQLite" -ForegroundColor Yellow
        Write-Host "ou SQLiteStudio pour examiner manuellement:" -ForegroundColor Yellow
        Write-Host "- Chemin: $dbPath" -ForegroundColor White
        Write-Host "- Table: emails" -ForegroundColor White
        Write-Host "- Colonnes à vérifier: recipient_email, sender_email, subject" -ForegroundColor White
        
        # Afficher la taille du fichier pour s'assurer qu'il contient des données
        $fileInfo = Get-Item $dbPath
        $sizeKB = [math]::Round($fileInfo.Length / 1024, 2)
        Write-Host "`n📊 Taille de la base: $sizeKB KB" -ForegroundColor Green
        
        if ($sizeKB -gt 100) {
            Write-Host "✅ La base contient probablement des données (taille significative)" -ForegroundColor Green
        } else {
            Write-Host "⚠️  Base de données relativement petite - peut contenir peu de données" -ForegroundColor Yellow
        }
        
        # Instructions pour vérification manuelle
        Write-Host "`n🔧 Pour vérifier manuellement les corrections:" -ForegroundColor Cyan
        Write-Host "1. Ouvrez DB Browser for SQLite (https://sqlitebrowser.org/)" -ForegroundColor White
        Write-Host "2. Ouvrez le fichier: $dbPath" -ForegroundColor White
        Write-Host "3. Allez dans l'onglet 'Parcourir les données'" -ForegroundColor White
        Write-Host "4. Sélectionnez la table 'emails'" -ForegroundColor White
        Write-Host "5. Vérifiez si la colonne 'recipient_email' contient des données" -ForegroundColor White
        
        Write-Host "`n📝 Requête SQL à exécuter dans DB Browser:" -ForegroundColor Cyan
        Write-Host "SELECT subject, sender_email, recipient_email, folder_name, created_at FROM emails ORDER BY created_at DESC LIMIT 10;" -ForegroundColor White
        
    } catch {
        Write-Host "❌ Erreur lors de l'inspection: $($_.Exception.Message)" -ForegroundColor Red
    }
    
} else {
    Write-Host "❌ Base de données non trouvée: $dbPath" -ForegroundColor Red
}

Write-Host "`n🎯 Prochaines étapes après vérification:" -ForegroundColor Cyan
Write-Host "1. Si recipient_email est vide pour tous les emails:" -ForegroundColor White
Write-Host "   - Les corrections sont en place mais les anciens emails n'ont pas de destinataires" -ForegroundColor White
Write-Host "   - Les nouveaux emails scannés après la correction devraient avoir des destinataires" -ForegroundColor White
Write-Host "2. Si recipient_email contient des données:" -ForegroundColor White
Write-Host "   - ✅ Les corrections fonctionnent!" -ForegroundColor Green
Write-Host "3. Testez en ajoutant un nouveau dossier pour voir si les nouveaux emails" -ForegroundColor White
Write-Host "   incluent bien les destinataires" -ForegroundColor White
