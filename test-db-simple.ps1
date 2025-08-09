# Test PowerShell pour verifier le contenu de la base de donnees

$dbPath = "E:\Tanguy\Bureau\AutoMailMonitor\data\emails.db"

Write-Host "Verification de la base de donnees..." -ForegroundColor Green

if (Test-Path $dbPath) {
    Write-Host "Base trouvee: $dbPath" -ForegroundColor Green
    
    $fileInfo = Get-Item $dbPath
    $sizeKB = [math]::Round($fileInfo.Length / 1024, 2)
    Write-Host "Taille de la base: $sizeKB KB" -ForegroundColor Green
    
    if ($sizeKB -gt 100) {
        Write-Host "La base contient probablement des donnees" -ForegroundColor Green
    } else {
        Write-Host "Base relativement petite" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Pour verifier les corrections:" -ForegroundColor Cyan
    Write-Host "1. Ouvrez DB Browser for SQLite" -ForegroundColor White
    Write-Host "2. Ouvrez le fichier: $dbPath" -ForegroundColor White
    Write-Host "3. Utilisez cette requete SQL:" -ForegroundColor White
    Write-Host "SELECT subject, sender_email, recipient_email, folder_name, created_at FROM emails ORDER BY created_at DESC LIMIT 10;" -ForegroundColor Yellow
    
} else {
    Write-Host "Base de donnees non trouvee: $dbPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "Les corrections apportees:" -ForegroundColor Cyan
Write-Host "1. Scripts PowerShell modifies pour recuperer Recipients" -ForegroundColor White
Write-Host "2. Methodes de base de donnees mises a jour pour inclure recipient_email" -ForegroundColor White
Write-Host "3. Nouveaux emails scannes devraient avoir des destinataires" -ForegroundColor White
