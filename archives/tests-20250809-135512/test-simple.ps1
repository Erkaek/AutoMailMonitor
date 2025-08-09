try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Output "=== Test de navigation dans Outlook ==="
    
    # Lister tous les comptes
    $accounts = $namespace.Accounts
    Write-Output "Comptes disponibles:"
    foreach ($account in $accounts) {
        Write-Output "  - $($account.DisplayName)"
    }
    
    # Essayer de trouver le compte erkaekanon
    $targetAccount = $null
    foreach ($account in $accounts) {
        if ($account.DisplayName -like "*erkaekanon*") {
            $targetAccount = $account
            Write-Output "Compte trouvé: $($account.DisplayName)"
            break
        }
    }
    
    if ($targetAccount) {
        $rootFolder = $namespace.Folders($targetAccount.DisplayName)
        Write-Output "Dossier racine: $($rootFolder.Name)"
        
        # Lister les sous-dossiers principaux
        Write-Output "Sous-dossiers:"
        foreach ($subfolder in $rootFolder.Folders) {
            Write-Output "  - $($subfolder.Name) ($($subfolder.Items.Count) éléments)"
            
            if ($subfolder.Name -eq "Boîte de réception") {
                Write-Output "    Sous-dossiers de la boîte de réception:"
                foreach ($inboxSubfolder in $subfolder.Folders) {
                    Write-Output "      - $($inboxSubfolder.Name) ($($inboxSubfolder.Items.Count) éléments)"
                }
            }
        }
    } else {
        Write-Output "Compte erkaekanon non trouvé, utilisation de la boîte par défaut"
        $inbox = $namespace.GetDefaultFolder(6)
        Write-Output "Boîte par défaut: $($inbox.Name) ($($inbox.Items.Count) éléments)"
        
        Write-Output "Sous-dossiers de la boîte par défaut:"
        foreach ($subfolder in $inbox.Folders) {
            Write-Output "  - $($subfolder.Name) ($($subfolder.Items.Count) éléments)"
        }
    }
    
} catch {
    Write-Output "Erreur: $($_.Exception.Message)"
}
