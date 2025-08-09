# Test de navigation PowerShell vers un dossier spécifique
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8

try {
    # Test avec le dossier "test" qui contient 9 emails
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Host "=== TEST NAVIGATION DOSSIER ==="
    
    # 1. Trouver le store erkaekanon@outlook.com
    $targetStore = $null
    foreach ($store in $namespace.Stores) {
        Write-Host "Store disponible: '$($store.DisplayName)'"
        if ($store.DisplayName -eq "erkaekanon@outlook.com") {
            $targetStore = $store
            Write-Host "✓ Store cible trouvé: $($store.DisplayName)"
            break
        }
    }
    
    if ($targetStore) {
        # 2. Naviguer vers la racine
        $rootFolder = $targetStore.GetRootFolder()
        Write-Host "Racine: $($rootFolder.Name)"
        
        # 3. Lister les dossiers de premier niveau
        Write-Host "Dossiers niveau 1:"
        foreach ($folder in $rootFolder.Folders) {
            Write-Host "  - $($folder.Name) ($($folder.Items.Count) items)"
        }
        
        # 4. Aller vers "Boîte de réception"
        $inboxFolder = $null
        foreach ($folder in $rootFolder.Folders) {
            if ($folder.Name -eq "Boîte de réception") {
                $inboxFolder = $folder
                Write-Host "✓ Boîte de réception trouvée: $($folder.Items.Count) items"
                break
            }
        }
        
        if ($inboxFolder) {
            # 5. Lister les sous-dossiers de la boîte de réception
            Write-Host "Sous-dossiers de la boîte de réception:"
            foreach ($subfolder in $inboxFolder.Folders) {
                Write-Host "  - $($subfolder.Name) ($($subfolder.Items.Count) items)"
            }
            
            # 6. Test navigation vers "test"
            $testFolder = $null
            foreach ($subfolder in $inboxFolder.Folders) {
                if ($subfolder.Name -eq "test") {
                    $testFolder = $subfolder
                    Write-Host "✓ Dossier 'test' trouvé: $($subfolder.Items.Count) items"
                    break
                }
            }
            
            if ($testFolder) {
                Write-Host "=== SUCCÈS: Navigation vers test réussie ==="
                Write-Host "Emails dans le dossier test: $($testFolder.Items.Count)"
                
                # Lister quelques emails
                if ($testFolder.Items.Count -gt 0) {
                    Write-Host "Premiers emails:"
                    for ($i = 1; $i -le [Math]::Min(3, $testFolder.Items.Count); $i++) {
                        $email = $testFolder.Items.Item($i)
                        if ($email.Class -eq 43) {
                            Write-Host "  - $($email.Subject) (de: $($email.SenderName))"
                        }
                    }
                }
            } else {
                Write-Host "❌ Dossier 'test' non trouvé"
            }
        } else {
            Write-Host "❌ Boîte de réception non trouvée"
        }
    } else {
        Write-Host "❌ Store erkaekanon@outlook.com non trouvé"
    }
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
}

Read-Host "Appuyez sur Entrée pour fermer..."
