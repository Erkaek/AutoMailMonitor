# Script PowerShell pour lister les dossiers Outlook disponibles
try {
    Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Host "🔍 Dossiers Outlook disponibles:" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    
    # Fonction récursive pour lister les dossiers
    function List-Folders($folder, $level = 0) {
        $indent = "  " * $level
        Write-Host "$indent📁 $($folder.Name)" -ForegroundColor Yellow
        
        foreach ($subfolder in $folder.Folders) {
            List-Folders $subfolder ($level + 1)
        }
    }
    
    # Lister les dossiers de la boîte de réception
    $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
    Write-Host "📧 Boîte de réception et sous-dossiers:" -ForegroundColor Cyan
    List-Folders $inbox
    
    Write-Host "`n📁 Autres dossiers principaux:" -ForegroundColor Cyan
    
    # Lister tous les stores (comptes email)
    foreach ($store in $namespace.Stores) {
        Write-Host "`n🏪 Store: $($store.DisplayName)" -ForegroundColor Magenta
        try {
            $rootFolder = $store.GetRootFolder()
            foreach ($folder in $rootFolder.Folders) {
                if ($folder.Name -ne "Inbox") {
                    List-Folders $folder 1
                }
            }
        } catch {
            Write-Host "    ❌ Impossible d'accéder aux dossiers de ce store" -ForegroundColor Red
        }
    }
    
} catch {
    Write-Host "❌ Erreur: $_" -ForegroundColor Red
    Write-Host "💡 Assurez-vous qu'Outlook est ouvert et accessible" -ForegroundColor Yellow
}
