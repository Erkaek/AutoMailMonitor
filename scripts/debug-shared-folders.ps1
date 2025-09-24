# Script de debug pour diagnostiquer les dossiers partagés
param([string]$StoreId)

try {
  # Connexion COM
  $outlook = New-Object -ComObject Outlook.Application
  $ns = $outlook.GetNamespace("MAPI")
  
  # Récupérer le store par ID
  $store = $null
  foreach ($st in $ns.Stores) {
    if ($st.StoreID -eq $StoreId) {
      $store = $st
      break
    }
  }
  
  if (!$store) {
    Write-Output "Store non trouvé avec ID: $StoreId"
    exit 1
  }
  
  Write-Output "=== DIAGNOSTIC STORE ==="
  Write-Output "Store trouvé: $($store.DisplayName)"
  Write-Output "Type: $($store.ExchangeStoreType)"
  Write-Output "=== DOSSIERS RACINE ==="
  
  $root = $store.GetRootFolder()
  
  foreach ($folder in $root.Folders) {
    $folderName = $folder.Name
    Write-Output "Dossier: '$folderName'"
    Write-Output "  ChildCount: $($folder.Folders.Count)"
    
    # Analyse complète du nom pour "Boîte de réception"
    Write-Output "  Longueur: $($folderName.Length)"
    Write-Output "  Bytes: $([System.Text.Encoding]::UTF8.GetBytes($folderName) -join ',')"
    Write-Output "  Chars: $($folderName.ToCharArray() -join ',')"
    
    # Tests multiples
    $isInbox1 = ($folderName -eq "Boîte de réception")
    $isInbox2 = ($folderName -like "*Boîte*")
    $isInbox3 = ($folderName -contains "réception")
    $isInbox4 = ($folderName -match "Boîte")
    
    Write-Output "  Tests: eq=$isInbox1, like=$isInbox2, contains=$isInbox3, match=$isInbox4"
    
    # Force le test pour le dossier qui semble être la boîte de réception
    if ($folderName.Length -gt 10 -and ($folderName -like "*o*t*" -or $folderName -like "*éception*" -or $folderName.Contains("Bo"))) {
      Write-Output "  ==> TENTATIVE FORCEE DE DETECTION!"
      
      try {
        Write-Output "  Sous-dossiers via .Folders:"
        foreach ($subFolder in $folder.Folders) {
          Write-Output "    - $($subFolder.Name) (ID: $($subFolder.EntryID))"
        }
      } catch {
        Write-Output "    Erreur .Folders: $($_.Exception.Message)"
      }
      
      # Test alternatif via Items (au cas où)
      try {
        Write-Output "  Items count: $($folder.Items.Count)"
      } catch {
        Write-Output "    Erreur .Items: $($_.Exception.Message)"
      }
    }
    Write-Output "---"
  }
  
} catch {
  Write-Output "Erreur: $($_.Exception.Message)"
  exit 1
}
