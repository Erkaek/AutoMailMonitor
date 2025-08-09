# Test de navigation Outlook simplifie
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== TEST NAVIGATION OUTLOOK ==="
    
    # Fonction de navigation avec corrections d'encodage
    function Find-OutlookFolder {
        param([string]$FolderPath, [object]$Namespace)
        
        Write-Host "Navigation vers: $FolderPath"
        
        # Nettoyer le chemin
        $cleanPath = $FolderPath -replace '\\\\', '\'
        
        # Corriger les problèmes d'encodage
        $cleanPath = $cleanPath -replace 'BoÃ®te de rÃ©ception', 'Boîte de réception'
        $cleanPath = $cleanPath -replace 'Ã®', 'î'
        $cleanPath = $cleanPath -replace 'Ã©', 'é'
        
        Write-Host "Chemin nettoye: $cleanPath"
        
        # Extraire compte et chemin
        if ($cleanPath -match '^([^\\]+)\\(.+)$') {
            $accountName = $matches[1]
            $folderPath = $matches[2]
            
            Write-Host "  Compte: $accountName"
            Write-Host "  Chemin: $folderPath"
            
            # Chercher le store
            $targetStore = $null
            foreach ($store in $Namespace.Stores) {
                if ($store.DisplayName -eq $accountName) {
                    $targetStore = $store
                    Write-Host "  Store trouve: $($store.DisplayName)"
                    break
                }
            }
            
            if (-not $targetStore) {
                Write-Host "  Store non trouve"
                return $null
            }
            
            # Naviguer
            $currentFolder = $targetStore.GetRootFolder()
            $pathParts = $folderPath -split '\\'
            
            foreach ($part in $pathParts) {
                if ($part -and $part.Trim() -ne "") {
                    Write-Host "    Recherche: $part"
                    $found = $false
                    
                    foreach ($subfolder in $currentFolder.Folders) {
                        if ($subfolder.Name -eq $part) {
                            $currentFolder = $subfolder
                            $found = $true
                            Write-Host "    Trouve: $part"
                            break
                        }
                    }
                    
                    if (-not $found) {
                        Write-Host "    Non trouve: $part"
                        return $null
                    }
                }
            }
            
            return $currentFolder
        }
        return $null
    }
    
    # Tester les chemins
    $paths = @(
        "erkaekanon@outlook.com\Boîte de réception\testA",
        "erkaekanon@outlook.com\Boîte de réception\test"
    )
    
    foreach ($path in $paths) {
        Write-Host "`n" + "="*50
        Write-Host "TEST: $path"
        
        $folder = Find-OutlookFolder -FolderPath $path -Namespace $namespace
        
        if ($folder) {
            Write-Host "SUCCES ! Dossier: $($folder.Name)"
            Write-Host "Emails: $($folder.Items.Count)"
            
            # Lister quelques emails
            if ($folder.Items.Count -gt 0) {
                Write-Host "`nPremiers emails:"
                $items = $folder.Items
                for ($i = 1; $i -le [Math]::Min(3, $folder.Items.Count); $i++) {
                    try {
                        $email = $items.Item($i)
                        if ($email.Class -eq 43) {
                            Write-Host "  $i. $($email.Subject) - $($email.SenderName)"
                        }
                    } catch {
                        Write-Host "  $i. [Erreur lecture email]"
                    }
                }
            }
        } else {
            Write-Host "ECHEC: Navigation impossible"
        }
    }
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
