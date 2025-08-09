# Explorer la structure exacte Outlook
try {
    $outlook = New-Object -ComObject Outlook.Application
    $nameSpace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== EXPLORATION STRUCTURE EXACTE ==="
    
    # Explorer tous les stores
    Write-Host "`nStores disponibles:"
    $stores = $nameSpace.Stores
    for ($i = 1; $i -le $stores.Count; $i++) {
        $store = $stores.Item($i)
        Write-Host "Store $($i): '$($store.DisplayName)' (Type: $($store.StoreType))"
        
        if ($store.DisplayName -like "*erkaekanon*") {
            Write-Host "`n=== EXPLORATION DU STORE erkaekanon ==="
            try {
                $rootFolder = $store.GetRootFolder()
                Write-Host "Root folder: '$($rootFolder.Name)'"
                
                Write-Host "`nDossiers dans le root:"
                $folders = $rootFolder.Folders
                for ($j = 1; $j -le $folders.Count; $j++) {
                    $folder = $folders.Item($j)
                    Write-Host "  - '$($folder.Name)' ($($folder.Items.Count) items)"
                    
                    # Si c'est la boite de reception, explorer dedans
                    if ($folder.Name -like "*reception*" -or $folder.Name -like "*Inbox*" -or $folder.Name -like "*Bo*te*") {
                        Write-Host "`n    EXPLORATION DE: '$($folder.Name)'"
                        $subFolders = $folder.Folders
                        for ($k = 1; $k -le $subFolders.Count; $k++) {
                            $subFolder = $subFolders.Item($k)
                            Write-Host "      - '$($subFolder.Name)' ($($subFolder.Items.Count) items)"
                            
                            # Explorer les sous-sous-dossiers si c'est "test"
                            if ($subFolder.Name -eq "test") {
                                Write-Host "`n        EXPLORATION DE test:"
                                $subSubFolders = $subFolder.Folders
                                for ($l = 1; $l -le $subSubFolders.Count; $l++) {
                                    $subSubFolder = $subSubFolders.Item($l)
                                    Write-Host "          - '$($subSubFolder.Name)' ($($subSubFolder.Items.Count) items)"
                                }
                            }
                        }
                    }
                }
            } catch {
                Write-Host "Erreur exploration store: $($_.Exception.Message)"
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
