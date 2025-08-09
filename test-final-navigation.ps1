# Test complet de navigation et récupération d'emails
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    
    # Fonction de navigation améliorée avec corrections d'encodage
    function Find-OutlookFolder {
        param([string]$FolderPath, [object]$Namespace)
        
        Write-Host "Navigation vers: '$FolderPath'"
        
        # Nettoyer le chemin (supprimer doubles backslashes, etc.)
        $cleanPath = $FolderPath -replace '\\\\\\\\', '\\' -replace '\\\\', '\'
        
        # Corriger les problèmes d'encodage courants - TOUTES les variantes
        $cleanPath = $cleanPath -replace 'Bo├«te de r├®ception', 'Boîte de réception'
        $cleanPath = $cleanPath -replace 'BoÃ®te de rÃ©ception', 'Boîte de réception'
        $cleanPath = $cleanPath -replace 'BoÃ®te de rÃ©ception', 'Boîte de réception'
        
        # Corrections d'encodage UTF-8 supplémentaires
        $cleanPath = $cleanPath -replace 'Ã®', 'î'
        $cleanPath = $cleanPath -replace 'Ã©', 'é'
        $cleanPath = $cleanPath -replace 'Ã¨', 'è'
        $cleanPath = $cleanPath -replace 'Ã ', 'à'
        
        Write-Host "Chemin nettoyé: '$cleanPath'"
        
        # Extraire compte et chemin
        if ($cleanPath -match '^([^\\]+)\\(.+)$') {
            $accountName = $matches[1]
            $folderPath = $matches[2]
            
            Write-Host "  Compte: '$accountName'"
            Write-Host "  Chemin: '$folderPath'"
            
            # Chercher le store/compte
            $targetStore = $null
            foreach ($store in $Namespace.Stores) {
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
                    Write-Host "    🔍 Recherche de: '$part'"
                    $found = $false
                    
                    foreach ($subfolder in $currentFolder.Folders) {
                        if ($subfolder.Name -eq $part) {
                            $currentFolder = $subfolder
                            $found = $true
                            Write-Host "    ✅ Trouvé: '$part' - Navigation réussie"
                            break
                        }
                    }
                    
                    if (-not $found) {
                        Write-Host "    ❌ '$part' non trouvé"
                        return $null
                    }
                }
            }
            
            Write-Host "  🎯 Navigation complète réussie !"
            return $currentFolder
        } else {
            Write-Host "  ❌ Format de chemin invalide"
            return $null
        }
    }
    
    # Chemins à tester (les vrais chemins de la BDD)
    $pathsToTest = @(
        "erkaekanon@outlook.com\Boîte de réception\testA",
        "erkaekanon@outlook.com\Boîte de réception\test"
    )
    
    $results = @()
    
    foreach ($testPath in $pathsToTest) {
        Write-Host "`n" + "="*60
        Write-Host "TEST: $testPath"
        Write-Host "="*60
        
        $folder = Find-OutlookFolder -FolderPath $testPath -Namespace $namespace
        
        if ($folder) {
            Write-Host "✅ SUCCÈS ! Dossier trouvé: '$($folder.Name)'"
            Write-Host "📧 Emails dans le dossier: $($folder.Items.Count)"
            
            $result = @{
                Path = $testPath
                Success = $true
                FolderName = $folder.Name
                EmailCount = $folder.Items.Count
                Emails = @()
            }
            
            # Récupérer quelques emails pour test
            if ($folder.Items.Count -gt 0) {
                Write-Host "`n📋 Premiers emails:"
                $items = $folder.Items
                $items.Sort("[ReceivedTime]", $true)
                
                for ($i = 1; $i -le [Math]::Min(5, $folder.Items.Count); $i++) {
                    try {
                        $email = $items.Item($i)
                        if ($email.Class -eq 43) { # Email
                            $emailInfo = @{
                                Subject = if($email.Subject) { $email.Subject } else { "(Pas de sujet)" }
                                From = if($email.SenderName) { $email.SenderName } else { "(Expéditeur inconnu)" }
                                ReceivedTime = $email.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                                IsRead = -not $email.UnRead
                            }
                            $result.Emails += $emailInfo
                            Write-Host "  $i. $($emailInfo.Subject) - $($emailInfo.From) ($($emailInfo.ReceivedTime))"
                        }
                    } catch {
                        Write-Host "  $i. [Erreur lecture email]"
                    }
                }
            }
            
            $results += $result
            
        } else {
            Write-Host "❌ ÉCHEC: Navigation impossible"
            $results += @{
                Path = $testPath
                Success = $false
                Error = "Navigation impossible"
            }
        }
    }
    
    # Sauvegarder les résultats
    $jsonResults = $results | ConvertTo-Json -Depth 3
    $jsonResults | Out-File -FilePath "test-navigation-results.json" -Encoding UTF8
    
    Write-Host "`n" + "="*60
    Write-Host "🎯 RÉSUMÉ FINAL"
    Write-Host "="*60
    
    foreach ($result in $results) {
        if ($result.Success) {
            Write-Host "✅ $($result.Path): $($result.EmailCount) emails trouvés"
        } else {
            Write-Host "❌ $($result.Path): Échec"
        }
    }
    
    Write-Host "`n📄 Résultats détaillés sauvegardés dans: test-navigation-results.json"
    
} catch {
    Write-Host "❌ ERREUR GLOBALE: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.ScriptStackTrace)"
} finally {
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
