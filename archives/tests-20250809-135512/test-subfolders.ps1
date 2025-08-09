$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$account = $namespace.Accounts | Where-Object { $_.DisplayName -like "*erkaekanon*" }
$rootFolder = $namespace.Folders($account.DisplayName)
$inbox = $rootFolder.Folders("Boîte de réception")

Write-Output "Sous-dossiers de la boîte de réception:"
foreach ($subfolder in $inbox.Folders) {
    Write-Output "  - $($subfolder.Name) ($($subfolder.Items.Count) éléments)"
    
    # Si c'est testA ou test, afficher le contenu
    if ($subfolder.Name -eq "testA" -or $subfolder.Name -eq "test") {
        Write-Output "    Détails du dossier $($subfolder.Name):"
        $items = $subfolder.Items
        $maxItems = [Math]::Min(3, $items.Count)
        for ($i = 1; $i -le $maxItems; $i++) {
            $mail = $items.Item($i)
            if ($mail.Class -eq 43) {
                $subject = if($mail.Subject) { $mail.Subject } else { "(Sans objet)" }
                Write-Output "      Email: $subject"
            }
        }
    }
}
