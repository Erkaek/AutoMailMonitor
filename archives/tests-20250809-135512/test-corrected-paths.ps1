# Test avec les VRAIS chemins de la BDD
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    Write-Host "=== TEST NAVIGATION AVEC CHEMINS CORRIGES ==="
    
    # Liste des chemins corrigés de la BDD
    $pathsToTest = @(
        "erkaekanon@outlook.com\Boîte de réception\testA",
        "erkaekanon@outlook.com\Boîte de réception\test",
        "erkaekanon@outlook.com\Boîte de réception\test\test-1",
        "erkaekanon@outlook.com\Boîte de réception\testA\test-c"
    )
    
    # Fonction pour naviguer vers un dossier spécifique
    function Find-OutlookFolder {
        param([string]$FolderPath, [object]$Namespace)
        
        Write-Host "`nTentative navigation vers: '$FolderPath'"
        
        # Extraire compte et chemin
        if ($FolderPath -match '^([^\\]+)\\(.+)$') {
            $accountName = $matches[1]
            $folderPath = $matches[2]
            
            Write-Host "  Compte: '$accountName'"
            Write-Host "  Chemin: '$folderPath'"
            
            # Chercher le store/compte
            $targetStore = $null
            foreach ($store in $Namespace.Stores) {
                Write-Host "  Checking store: '$($store.DisplayName)'"
                if ($store.DisplayName -like "*$accountName*" -or $store.DisplayName -eq $accountName) {
                    $targetStore = $store
                    Write-Host "  ✅ Store trouvé: '$($store.DisplayName)'"
                    break
                }
            }
            
            if (-not $targetStore) {
                Write-Host "  ❌ Store non trouvé - utilisation du store par défaut"
                $targetStore = $Namespace.DefaultStore
            }
            
            # Naviguer dans l'arborescence
            $currentFolder = $targetStore.GetRootFolder()
            Write-Host "  Root folder: '$($currentFolder.Name)'"
            
            $pathParts = $folderPath -split '\\'
            Write-Host "  Parties du chemin: $($pathParts -join ' -> ')"
            
            foreach ($part in $pathParts) {
                if ($part -and $part.Trim() -ne "") {
                    Write-Host "    Recherche de: '$part'"
                    $found = $false
                    
                    # Lister tous les sous-dossiers pour debug
                    Write-Host "    Dossiers disponibles:"
                    foreach ($subfolder in $currentFolder.Folders) {
                        Write-Host "      - '$($subfolder.Name)'"
                        if ($subfolder.Name -eq $part) {
                            $currentFolder = $subfolder
                            $found = $true
                            Write-Host "    ✅ Trouvé: '$part' - Navigation réussie"
                            break
                        }
                    }
                    
                    if (-not $found) {
                        Write-Host "    ❌ '$part' non trouvé dans le dossier actuel"
                        return $null
                    }
                }
            }
            
            Write-Host "  🎯 Navigation complète réussie vers: '$($currentFolder.Name)'"
            Write-Host "     Items dans le dossier: $($currentFolder.Items.Count)"
            return $currentFolder
        } else {
            Write-Host "  ❌ Format de chemin invalide"
            return $null
        }
    }
    
    # Tester chaque chemin
    foreach ($testPath in $pathsToTest) {
        Write-Host "`n" + "="*50
        $result = Find-OutlookFolder -FolderPath $testPath -Namespace $namespace
        
        if ($result) {
            Write-Host "✅ SUCCÈS: Dossier '$($result.Name)' trouvé avec $($result.Items.Count) emails"
            
            # Lister quelques emails si disponibles
            if ($result.Items.Count -gt 0) {
                Write-Host "   Premiers emails:"
                for ($i = 1; $i -le [Math]::Min(3, $result.Items.Count); $i++) {
                    $email = $result.Items.Item($i)
                    Write-Host "   - $($email.Subject) (de: $($email.SenderName))"
                }
            }
        } else {
            Write-Host "❌ ÉCHEC: Impossible de naviguer vers '$testPath'"
        }
    }
    
} catch {
    Write-Host "ERREUR GLOBALE: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
