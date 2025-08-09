try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    Write-Output "✅ Outlook COM test OK"
    Write-Output "Nom de la boîte: $($inbox.Name)"
    Write-Output "Nombre d'éléments: $($inbox.Items.Count)"
    
    # Test navigation vers un sous-dossier (testA)
    try {
        $testFolder = $inbox.Folders.Item("testA")
        Write-Output "✅ Dossier testA trouvé"
        Write-Output "Emails dans testA: $($testFolder.Items.Count)"
    } catch {
        Write-Output "❌ Dossier testA non trouvé: $($_.Exception.Message)"
    }
    
} catch {
    Write-Output "❌ Erreur COM Outlook: $($_.Exception.Message)"
}
