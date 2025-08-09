try {
  $outlook = New-Object -ComObject Outlook.Application
  $namespace = $outlook.GetNamespace("MAPI")
  $inbox = $namespace.GetDefaultFolder(6)
  
  Write-Host "DETECTION_TEST: Emails dans Inbox"
  Write-Host "Total items:" $inbox.Items.Count
  Write-Host "Non lus:" $inbox.UnReadItemCount
  
  # Parcourir les premiers emails
  $count = 0
  foreach ($item in $inbox.Items) {
    if ($count -ge 5) { break }
    $status = if ($item.UnRead -eq $false) { "Lu" } else { "Non lu" }
    Write-Host "Email $($count + 1): '$($item.Subject)' - Statut: $status"
    $count++
  }
  
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($inbox) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
  
} catch {
  Write-Host "ERREUR:" $_.Exception.Message
}
