# Analyser les caracteres exacts des noms de dossiers
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== ANALYSE DES CARACTERES ==="
    
    $store = $namespace.Stores | Where-Object { $_.DisplayName -eq "erkaekanon@outlook.com" }
    if ($store) {
        $rootFolder = $store.GetRootFolder()
        
        foreach ($folder in $rootFolder.Folders) {
            $name = $folder.Name
            if ($name -like "*réception*" -or $name -like "*reception*") {
                Write-Host "`nDossier trouve: '$name'"
                Write-Host "Longueur: $($name.Length)"
                Write-Host "Caracteres:"
                for ($i = 0; $i -lt $name.Length; $i++) {
                    $char = $name[$i]
                    $unicode = [int][char]$char
                    Write-Host "  $i : '$char' = Unicode $unicode"
                }
                
                # Tester la correspondance
                if ($name -eq "Boîte de réception") {
                    Write-Host "MATCH avec 'Boîte de réception'"
                } else {
                    Write-Host "PAS de match avec 'Boîte de réception'"
                }
                
                # Test direct de navigation
                Write-Host "`nTest de navigation dans ce dossier:"
                $subFolders = $folder.Folders
                foreach ($subFolder in $subFolders) {
                    Write-Host "  - '$($subFolder.Name)' ($($subFolder.Items.Count) items)"
                    
                    if ($subFolder.Name -eq "test") {
                        Write-Host "    SUCCESS! Dossier 'test' trouve avec $($subFolder.Items.Count) emails"
                    }
                    if ($subFolder.Name -eq "testA") {
                        Write-Host "    SUCCESS! Dossier 'testA' trouve avec $($subFolder.Items.Count) emails"
                    }
                }
                break
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
