try {
  $ErrorActionPreference = "Stop"
  $outlook = New-Object -ComObject Outlook.Application
  $namespace = $outlook.GetNamespace("MAPI")
  
  Write-Host "RECHERCHE du dossier avec 14 emails..."
  Write-Host ""
  
  # Fonction pour explorer récursivement tous les dossiers
  function Search-FolderWithCount($folder, $targetCount = 14, $level = 0) {
    $indent = "  " * $level
    try {
      $itemCount = $folder.Items.Count
      $unreadCount = $folder.UnReadItemCount
      $folderPath = $folder.FolderPath
      
      Write-Host "${indent}$($folder.Name): $itemCount emails ($unreadCount non lus) - $folderPath"
      
      # Si c'est le dossier recherché
      if ($itemCount -eq $targetCount) {
        Write-Host "*** TROUVE: $folderPath ***" -ForegroundColor Green
        
        # Afficher quelques emails pour confirmer
        Write-Host "Premiers emails:"
        $count = 0
        foreach ($item in $folder.Items) {
          if ($count -ge 5) { break }
          $status = if ($item.UnRead -eq $false) { "Lu" } else { "Non lu" }
          Write-Host "  - $($item.Subject) [$status]"
          $count++
        }
      }
      
      # Explorer les sous-dossiers
      if ($folder.Folders.Count -gt 0) {
        foreach ($subfolder in $folder.Folders) {
          try {
            Search-FolderWithCount $subfolder $targetCount ($level + 1)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($subfolder) | Out-Null
          } catch {
            # Ignorer les dossiers inaccessibles
          }
        }
      }
    } catch {
      Write-Host "${indent}[ERREUR] $($folder.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
  }
  
  # Explorer toutes les boîtes mail
  foreach ($store in $namespace.Stores) {
    try {
      Write-Host "=== STORE: $($store.DisplayName) ===" -ForegroundColor Yellow
      $rootFolder = $store.GetRootFolder()
      Search-FolderWithCount $rootFolder 14 0
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
      Write-Host ""
    } catch {
      Write-Host "Erreur store $($store.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
    }
  }
  
  # Libération des objets
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
  
} catch {
  Write-Host "ERREUR:" $_.Exception.Message -ForegroundColor Red
}
