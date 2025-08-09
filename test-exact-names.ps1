# Test avec le VRAI nom exact detecte
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== TEST AVEC NOM EXACT ==="
    
    $store = $namespace.Stores | Where-Object { $_.DisplayName -eq "erkaekanon@outlook.com" }
    if ($store) {
        Write-Host "Store trouve: $($store.DisplayName)"
        $rootFolder = $store.GetRootFolder()
        
        # Utiliser le vrai nom detecte: "Boîte de réception"
        $inboxFolder = $rootFolder.Folders | Where-Object { $_.Name -eq "Boîte de réception" }
        if ($inboxFolder) {
            Write-Host "SUCCESS! Boite de reception trouvee: $($inboxFolder.Items.Count) items"
            
            # Chercher "test"
            $testFolder = $inboxFolder.Folders | Where-Object { $_.Name -eq "test" }
            if ($testFolder) {
                Write-Host "SUCCESS! Dossier test trouve: $($testFolder.Items.Count) emails"
                
                # Récupérer et afficher TOUS les emails
                if ($testFolder.Items.Count -gt 0) {
                    Write-Host "`n=== EMAILS DANS LE DOSSIER TEST ==="
                    $items = $testFolder.Items
                    $items.Sort("[ReceivedTime]", $true)
                    
                    for ($i = 1; $i -le $testFolder.Items.Count; $i++) {
                        try {
                            $email = $items.Item($i)
                            if ($email.Class -eq 43) { # Email
                                $subject = if($email.Subject) { $email.Subject } else { "(Pas de sujet)" }
                                $from = if($email.SenderName) { $email.SenderName } else { "(Expéditeur inconnu)" }
                                $received = $email.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                                $isRead = -not $email.UnRead
                                
                                Write-Host "  $i. [$received] $subject"
                                Write-Host "      De: $from - Lu: $isRead"
                            }
                        } catch {
                            Write-Host "  $i. [Erreur lecture email: $($_.Exception.Message)]"
                        }
                    }
                    
                    Write-Host "`n✅ TOTAL: $($testFolder.Items.Count) emails trouvés dans le dossier 'test'"
                }
            } else {
                Write-Host "Dossier test NON trouve"
                Write-Host "Sous-dossiers disponibles:"
                foreach ($folder in $inboxFolder.Folders) {
                    Write-Host "  - '$($folder.Name)' ($($folder.Items.Count) items)"
                }
            }
            
            # Aussi tester testA
            $testAFolder = $inboxFolder.Folders | Where-Object { $_.Name -eq "testA" }
            if ($testAFolder) {
                Write-Host "`nSUCCESS! Dossier testA trouve: $($testAFolder.Items.Count) emails"
            } else {
                Write-Host "`nDossier testA NON trouve"
            }
            
        } else {
            Write-Host "ECHEC: Boite de reception non trouvee"
        }
    } else {
        Write-Host "ECHEC: Store non trouve"
    }
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
    Write-Host "Stack: $($_.ScriptStackTrace)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
