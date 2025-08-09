# Test avec les VRAIS noms directement
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== TEST AVEC VRAIS NOMS ==="
    
    # Tester directement avec les vrais noms
    $store = $namespace.Stores | Where-Object { $_.DisplayName -eq "erkaekanon@outlook.com" }
    if ($store) {
        Write-Host "Store trouve: $($store.DisplayName)"
        $rootFolder = $store.GetRootFolder()
        
        # Chercher "Boîte de réception"
        $inboxFolder = $rootFolder.Folders | Where-Object { $_.Name -eq "Boîte de réception" }
        if ($inboxFolder) {
            Write-Host "Boite de reception trouvee: $($inboxFolder.Items.Count) items"
            
            # Chercher "test"
            $testFolder = $inboxFolder.Folders | Where-Object { $_.Name -eq "test" }
            if ($testFolder) {
                Write-Host "Dossier test trouve: $($testFolder.Items.Count) emails"
                
                # Lister les emails
                if ($testFolder.Items.Count -gt 0) {
                    Write-Host "`nEmails dans test:"
                    $items = $testFolder.Items
                    for ($i = 1; $i -le [Math]::Min(5, $testFolder.Items.Count); $i++) {
                        $email = $items.Item($i)
                        Write-Host "  $i. $($email.Subject) - $($email.SenderName)"
                    }
                }
            } else {
                Write-Host "Dossier test NON trouve"
                Write-Host "Dossiers disponibles dans Boite de reception:"
                foreach ($folder in $inboxFolder.Folders) {
                    Write-Host "  - $($folder.Name) ($($folder.Items.Count) items)"
                }
            }
            
            # Chercher "testA"
            $testAFolder = $inboxFolder.Folders | Where-Object { $_.Name -eq "testA" }
            if ($testAFolder) {
                Write-Host "`nDossier testA trouve: $($testAFolder.Items.Count) emails"
            } else {
                Write-Host "`nDossier testA NON trouve"
            }
            
        } else {
            Write-Host "Boite de reception NON trouvee"
            Write-Host "Dossiers disponibles dans root:"
            foreach ($folder in $rootFolder.Folders) {
                Write-Host "  - $($folder.Name) ($($folder.Items.Count) items)"
            }
        }
    } else {
        Write-Host "Store erkaekanon@outlook.com NON trouve"
        Write-Host "Stores disponibles:"
        foreach ($s in $namespace.Stores) {
            Write-Host "  - $($s.DisplayName)"
        }
    }
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
