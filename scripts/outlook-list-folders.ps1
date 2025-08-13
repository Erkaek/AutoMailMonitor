param(
  [Parameter(Mandatory=$true)][string]$StoreId,
  [string]$ParentEntryId
)

$ErrorActionPreference = 'Stop'

# Force UTF-8 output
$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc

try {
  try { Add-Type -AssemblyName Microsoft.Office.Interop.Outlook | Out-Null } catch {}
  $olApp = New-Object -ComObject Outlook.Application
  $session = $olApp.Session

  # Constante Inbox
  $olFolderInbox = 6

  # Trouver le Store
  $target = $null
  foreach ($st in $session.Stores) {
    try { if ($st.StoreID -eq $StoreId) { $target = $st; break } } catch {}
  }
  # Fallback: si pas trouvé par EntryID, tenter par DisplayName (égalité stricte puis inclusif)
  if (-not $target -and $StoreId) {
    foreach ($st in $session.Stores) {
      try {
        if ($st.DisplayName -eq $StoreId) { $target = $st; break }
      } catch {}
    }
  }
  if (-not $target -and $StoreId) {
    foreach ($st in $session.Stores) {
      try {
        if ($st.DisplayName -like "*${StoreId}*") { $target = $st; break }
      } catch {}
    }
  }
  # Fallback: si $StoreId ressemble à un SMTP, mapper via Accounts.DeliveryStore
  if (-not $target -and $StoreId -and $StoreId.Contains('@')) {
    try {
      foreach ($acc in $session.Accounts) {
        try {
          if ($acc.SmtpAddress -eq $StoreId) {
            $del = $acc.DeliveryStore
            if ($del) { $target = $del; break }
          }
        } catch {}
      }
    } catch {}
  }
  if (-not $target) { throw "Store non trouvé" }

  # Déterminer le parent: Inbox du store si ParentEntryId vide, sinon GetFolderFromID
  if ([string]::IsNullOrEmpty($ParentEntryId)) {
    $parent = $target.GetDefaultFolder($olFolderInbox)
  } else {
    $parent = $session.GetFolderFromID($ParentEntryId, $StoreId)
  }

  $list = @()
  foreach ($f in $parent.Folders) {
    try {
      $list += [pscustomobject]@{
        Name       = $f.Name
        EntryId    = $f.EntryID
        ChildCount = $f.Folders.Count
      }
    } catch {}
  }

  # Ajouter le nom du dossier parent (utile pour connaître le libellé localisé de la Boîte de réception)
  $parentName = $null
  try { $parentName = $parent.Name } catch {}

  $payload = [pscustomobject]@{
    StoreId       = $StoreId
    ParentEntryId = $ParentEntryId
    ParentName    = $parentName
    Folders       = $list
  }
  $json = $payload | ConvertTo-Json -Depth 6 -Compress
  [Console]::Out.Write($json)
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
