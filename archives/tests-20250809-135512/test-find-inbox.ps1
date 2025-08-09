$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$account = $namespace.Accounts | Where-Object { $_.DisplayName -like "*erkaekanon*" }
$rootFolder = $namespace.Folders($account.DisplayName)

Write-Output "Recherche de la boîte de réception..."
$inbox = $null
foreach ($folder in $rootFolder.Folders) {
    Write-Output "Dossier trouvé: $($folder.Name)"
    if ($folder.Name -like "*ception*" -or $folder.Name -like "*Inbox*") {
        $inbox = $folder
        Write-Output "✓ Boîte de réception identifiée: $($folder.Name)"
        break
    }
}

if ($inbox) {
    Write-Output "Sous-dossiers de $($inbox.Name):"
    foreach ($subfolder in $inbox.Folders) {
        Write-Output "  - $($subfolder.Name) ($($subfolder.Items.Count) éléments)"
    }
} else {
    Write-Output "Boîte de réception non trouvée!"
}
