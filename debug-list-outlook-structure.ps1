# Script pour lister toute l'arborescence Outlook
# Configuration d'encodage
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8

try {
    # Connexion à Outlook
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # Fonction pour lister récursivement les dossiers
    function List-FolderStructure {
        param($folder, $level = 0)
        
        $indent = "  " * $level
        $fullPath = ""
        
        # Construire le chemin complet
        $currentFolder = $folder
        $pathParts = @()
        while ($currentFolder -ne $null -and $currentFolder.Parent -ne $null) {
            $pathParts = @($currentFolder.Name) + $pathParts
            $currentFolder = $currentFolder.Parent
        }
        
        if ($pathParts.Count -gt 0) {
            $fullPath = $pathParts -join "\"
        }
        
        $folderInfo = @{
            name = $folder.Name
            fullPath = $fullPath
            level = $level
            itemCount = $folder.Items.Count
            folderType = $folder.DefaultItemType
        }
        
        Write-Host "$indent- $($folder.Name) ($($folder.Items.Count) items) - Chemin: $fullPath"
        
        # Lister les sous-dossiers
        foreach ($subfolder in $folder.Folders) {
            List-FolderStructure $subfolder ($level + 1)
        }
    }
    
    # Lister tous les stores
    Write-Host "=== STRUCTURE OUTLOOK COMPLETE ==="
    foreach ($store in $namespace.Stores) {
        Write-Host ""
        Write-Host "STORE: $($store.DisplayName)"
        Write-Host "Type: $($store.ExchangeStoreType)"
        Write-Host "---"
        
        $rootFolder = $store.GetRootFolder()
        List-FolderStructure $rootFolder
        
        Write-Host ""
    }
    
    Write-Host "=== FIN STRUCTURE ==="
    
} catch {
    Write-Host "ERREUR: $($_.Exception.Message)"
}

# Maintenir le script ouvert
Read-Host "Appuyez sur Entrée pour fermer..."
