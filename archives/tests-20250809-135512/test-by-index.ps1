# Navigation par index pour éviter les problèmes d'encodage
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== NAVIGATION PAR INDEX ==="
    
    $store = $namespace.Stores | Where-Object { $_.DisplayName -eq "erkaekanon@outlook.com" }
    if ($store) {
        Write-Host "Store trouve: $($store.DisplayName)"
        $rootFolder = $store.GetRootFolder()
        $folders = $rootFolder.Folders
        
        Write-Host "`nDossiers disponibles:"
        for ($i = 1; $i -le $folders.Count; $i++) {
            $folder = $folders.Item($i)
            Write-Host "  $i. '$($folder.Name)' ($($folder.Items.Count) items)"
            
            # Chercher celui qui ressemble à "Boîte de réception"
            if ($folder.Name -match "reception" -or $folder.Name -match "réception" -or $folder.Items.Count -eq 477) {
                Write-Host "    --> CECI EST PROBABLEMENT LA BOITE DE RECEPTION"
                
                # Explorer ce dossier
                $subFolders = $folder.Folders
                Write-Host "    Sous-dossiers:"
                for ($j = 1; $j -le $subFolders.Count; $j++) {
                    $subFolder = $subFolders.Item($j)
                    Write-Host "      $j. '$($subFolder.Name)' ($($subFolder.Items.Count) items)"
                    
                    # Test pour "test" et "testA"
                    if ($subFolder.Name -eq "test") {
                        Write-Host "        ✅ DOSSIER TEST TROUVE ! $($subFolder.Items.Count) emails"
                        
                        # Récupérer quelques emails
                        if ($subFolder.Items.Count -gt 0) {
                            Write-Host "        Emails:"
                            $items = $subFolder.Items
                            for ($k = 1; $k -le [Math]::Min(3, $subFolder.Items.Count); $k++) {
                                $email = $items.Item($k)
                                Write-Host "          $k. $($email.Subject) - $($email.SenderName)"
                            }
                        }
                    }
                    
                    if ($subFolder.Name -eq "testA") {
                        Write-Host "        ✅ DOSSIER TESTA TROUVE ! $($subFolder.Items.Count) emails"
                    }
                }
            }
        }
    }
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
