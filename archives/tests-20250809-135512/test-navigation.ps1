try {
  $ErrorActionPreference = "Stop"
  $outlook = New-Object -ComObject Outlook.Application
  $namespace = $outlook.GetNamespace("MAPI")
  
  # Analyser le chemin du dossier
  $folderPath = "erkaekanon@outlook.com\Boîte de réception\testA"
  $pathParts = $folderPath.Split("\") | Where-Object { $_ -ne "" }
  
  Write-Output "Debug: Chemin analysé: $($pathParts -join ' -> ')"
  
  # Rechercher le bon compte et dossier
  $folder = $null
  
  if ($pathParts.Count -ge 2) {
    $accountName = $pathParts[0]
    $folderNames = $pathParts[1..($pathParts.Count-1)]
    
    Write-Output "Debug: Compte recherché: $accountName"
    Write-Output "Debug: Dossiers: $($folderNames -join ' -> ')"
    
    # Rechercher le compte correspondant
    $accounts = $namespace.Accounts
    Write-Output "Debug: Comptes disponibles:"
    foreach ($account in $accounts) {
      Write-Output "  - $($account.DisplayName)"
      if ($account.DisplayName -like "*" + $accountName.Split("@")[0] + "*") {
        Write-Output "  ✓ Compte correspondant trouvé!"
        try {
          $rootFolder = $namespace.Folders($account.DisplayName)
          if ($rootFolder) {
            $currentFolder = $rootFolder
            Write-Output "  - Dossier racine: $($rootFolder.Name)"
            
            # Naviguer dans la hiérarchie des dossiers
            foreach ($folderName in $folderNames) {
              Write-Output "  - Navigation vers: $folderName"
              if ($folderName -ne "Boîte de réception" -and $folderName -ne "Inbox") {
                $currentFolder = $currentFolder.Folders($folderName)
                Write-Output "    ✓ Trouvé: $($currentFolder.Name)"
              } elseif ($folderName -eq "Boîte de réception" -or $folderName -eq "Inbox") {
                $currentFolder = $currentFolder.Folders("Boîte de réception")
                Write-Output "    ✓ Boîte de réception: $($currentFolder.Name)"
              }
            }
            
            $folder = $currentFolder
            Write-Output "✅ Dossier final trouvé: $($folder.Name)"
            break
          }
        } catch {
          Write-Output "  ❌ Erreur navigation: $($_.Exception.Message)"
        }
      }
    }
  }
  
  if (-not $folder) {
    throw "Dossier non trouvé: $folderPath"
  }
  
  $items = $folder.Items
  Write-Output "📧 Nombre d'emails dans le dossier: $($items.Count)"
  
  # Test de récupération de quelques emails
  $maxEmails = [Math]::Min(3, $items.Count)
  for ($i = 1; $i -le $maxEmails; $i++) {
    try {
      $mail = $items.Item($i)
      if ($mail.Class -eq 43) {
        $subject = if($mail.Subject) { $mail.Subject } else { "(Sans objet)" }
        $isRead = -not $mail.UnRead
        Write-Output "  Email ${i}: $subject (Lu: $isRead)"
      }
    } catch {
      Write-Output "  Erreur email ${i}: $($_.Exception.Message)"
    }
  }
  
} catch {
  Write-Output "❌ Erreur: $($_.Exception.Message)"
}
