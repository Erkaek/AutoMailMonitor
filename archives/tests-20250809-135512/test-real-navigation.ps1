# Test navigation avec les vrais noms de dossiers
try {
    $outlook = New-Object -ComObject Outlook.Application
    $nameSpace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== TEST NAVIGATION AVEC VRAIS NOMS ==="
    
    # Test 1: Navigation vers "Boîte de réception"
    Write-Host "`nTest 1: Navigation Boite de reception..."
    $inboxStore = $nameSpace.Stores | Where-Object { $_.DisplayName -eq "erkaekanon@outlook.com" }
    if ($inboxStore) {
        $rootFolder = $inboxStore.GetRootFolder()
        $inboxFolder = $rootFolder.Folders | Where-Object { $_.Name -eq "Boîte de réception" }
        if ($inboxFolder) {
            Write-Host "SUCCESS: Boite de reception trouvee - $($inboxFolder.Items.Count) items"
            
            # Test 2: Navigation vers sous-dossier "test"
            Write-Host "`nTest 2: Navigation sous-dossier test..."
            $testFolder = $inboxFolder.Folders | Where-Object { $_.Name -eq "test" }
            if ($testFolder) {
                Write-Host "SUCCESS: Dossier test trouve - $($testFolder.Items.Count) items"
                
                # Lister quelques emails
                if ($testFolder.Items.Count -gt 0) {
                    Write-Host "`nEmails dans test:"
                    for ($i = 1; $i -le [Math]::Min(3, $testFolder.Items.Count); $i++) {
                        $email = $testFolder.Items.Item($i)
                        Write-Host "- $($email.Subject) (from: $($email.SenderName))"
                    }
                }
                
                # Test 3: Navigation vers sous-sous-dossier "test-1"
                Write-Host "`nTest 3: Navigation sous-sous-dossier test-1..."
                $test1Folder = $testFolder.Folders | Where-Object { $_.Name -eq "test-1" }
                if ($test1Folder) {
                    Write-Host "SUCCESS: Dossier test-1 trouve - $($test1Folder.Items.Count) items"
                } else {
                    Write-Host "ECHEC: Dossier test-1 non trouve"
                }
            } else {
                Write-Host "ECHEC: Dossier test non trouve"
            }
            
            # Test 4: Navigation vers sous-dossier "testA"
            Write-Host "`nTest 4: Navigation sous-dossier testA..."
            $testAFolder = $inboxFolder.Folders | Where-Object { $_.Name -eq "testA" }
            if ($testAFolder) {
                Write-Host "SUCCESS: Dossier testA trouve - $($testAFolder.Items.Count) items"
            } else {
                Write-Host "ECHEC: Dossier testA non trouve"
            }
        } else {
            Write-Host "ECHEC: Boite de reception non trouvee"
        }
    } else {
        Write-Host "ECHEC: Store erkaekanon@outlook.com non trouve"
    }
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
