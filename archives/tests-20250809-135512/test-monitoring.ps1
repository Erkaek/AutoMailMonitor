# Script de test pour monitoring en temps réel
param([string]$FolderPath = "erkaekanon@outlook.com\Boîte de réception\test")

try {
    $ErrorActionPreference = "Stop"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # Rechercher le dossier
    $folder = $null
    $accounts = $namespace.Accounts
    foreach ($account in $accounts) {
        if ($account.DisplayName -like "*erkaekanon*") {
            $rootFolder = $namespace.Folders($account.DisplayName)
            if ($rootFolder) {
                $inbox = $rootFolder.Folders("Boîte de réception")
                if ($inbox) {
                    $testFolder = $inbox.Folders("test")
                    if ($testFolder) {
                        $folder = $testFolder
                        break
                    }
                }
            }
        }
    }
    
    if (-not $folder) {
        throw "Dossier non trouvé: $FolderPath"
    }
    
    Write-Output "MONITORING_STARTED:$($folder.Items.Count)"
    Write-Output "ADVANCED_MONITORING:Surveillance complète activée"
    
    # Cache initial des états
    $lastCount = $folder.Items.Count
    $lastEmailStates = @{}
    
    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    foreach ($item in $items) {
        try {
            $lastEmailStates[$item.EntryID] = @{
                Subject = if($item.Subject) { $item.Subject } else { "(Sans objet)" }
                IsRead = (-not $item.UnRead)
                LastModificationTime = if ($item.LastModificationTime) { $item.LastModificationTime } else { $item.ReceivedTime }
            }
        } catch {
            # Ignorer les erreurs d'accès
        }
    }
    
    Write-Output "INITIAL_STATE:$lastCount emails cachés"
    
    # Boucle de monitoring (seulement 3 cycles pour test)
    for ($cycle = 1; $cycle -le 3; $cycle++) {
        Start-Sleep -Seconds 3
        Write-Output "CYCLE:$cycle - Vérification..."
        
        try {
            $currentCount = $folder.Items.Count
            $currentEmailStates = @{}
            $items = $folder.Items
            $items.Sort("[ReceivedTime]", $true)
            
            foreach ($item in $items) {
                try {
                    $entryId = $item.EntryID
                    $currentState = @{
                        Subject = if($item.Subject) { $item.Subject } else { "(Sans objet)" }
                        IsRead = (-not $item.UnRead)
                        LastModificationTime = if ($item.LastModificationTime) { $item.LastModificationTime } else { $item.ReceivedTime }
                    }
                    
                    $currentEmailStates[$entryId] = $currentState
                    
                    # Détecter les changements
                    if ($lastEmailStates.ContainsKey($entryId)) {
                        $lastState = $lastEmailStates[$entryId]
                        
                        # CHANGEMENT DE STATUT LU/NON LU
                        if ($lastState.IsRead -ne $currentState.IsRead) {
                            $statusChange = if ($currentState.IsRead) { "MARKED_READ" } else { "MARKED_UNREAD" }
                            $statusSubject = $currentState.Subject
                            Write-Output "STATUS_CHANGE:${entryId}:${statusChange}:${statusSubject}:${FolderPath}"
                        }
                    }
                } catch {
                    continue
                }
            }
            
            # Mettre à jour le cache
            $lastEmailStates = $currentEmailStates
            $lastCount = $currentCount
            
        } catch {
            Write-Output "ERROR_MONITORING:$($_.Exception.Message)"
        }
    }
    
    Write-Output "TEST_COMPLETED:Monitoring test terminé"
    
} catch {
    Write-Output "ERROR_SETUP:$($_.Exception.Message)"
}
