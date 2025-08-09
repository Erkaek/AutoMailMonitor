try {
  $outlook = New-Object -ComObject Outlook.Application
  $namespace = $outlook.GetNamespace("MAPI")
  
  Write-Host "=== TOUTES LES BOITES MAIL DETECTEES ==="
  Write-Host ""
  
  $storeIndex = 1
  foreach ($store in $namespace.Stores) {
    try {
      Write-Host "Boite $storeIndex : $($store.DisplayName)"
      Write-Host "  Type: $($store.ExchangeStoreType)"
      Write-Host "  Fichier: $($store.FilePath)"
      
      # Essayer d'accéder au dossier racine
      try {
        $rootFolder = $store.GetRootFolder()
        Write-Host "  Dossiers principaux:"
        
        foreach ($folder in $rootFolder.Folders) {
          if ($folder.DefaultItemType -eq 0) { # Email folders
            Write-Host "    - $($folder.Name) ($($folder.Items.Count) items, $($folder.UnReadItemCount) non lus)"
          }
        }
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rootFolder) | Out-Null
      } catch {
        Write-Host "    Erreur accès dossiers: $($_.Exception.Message)"
      }
      
      Write-Host ""
      $storeIndex++
    } catch {
      Write-Host "Erreur lecture boite $storeIndex : $($_.Exception.Message)"
      $storeIndex++
    }
  }
  
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
  
} catch {
  Write-Host "ERREUR:" $_.Exception.Message
}
