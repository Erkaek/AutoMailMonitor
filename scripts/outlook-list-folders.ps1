param(
  [Parameter(Mandatory=$true)][string]$StoreId,
  [string]$ParentEntryId,
  [switch]$LooseMatch
)

$ErrorActionPreference = 'Stop'

# Force UTF-8 output
$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc

try {
  # Normaliser / trim du StoreId reçu (élimine espaces accidentels)
  try { if ($StoreId) { $StoreId = $StoreId.Trim() } } catch {}
  try { Add-Type -AssemblyName Microsoft.Office.Interop.Outlook | Out-Null } catch {}
  $olApp = New-Object -ComObject Outlook.Application
  $session = $olApp.Session

  # Constante Inbox
  $olFolderInbox = 6

  # ===================== Résolution robuste du Store =====================
  $target = $null

  # 1) Match exact StoreID
  foreach ($st in $session.Stores) { try { if ($st.StoreID -eq $StoreId) { $target = $st; break } } catch {} }

  # 2) Match loose sur début / inclusion si option ou pas trouvé
  if (-not $target) {
    foreach ($st in $session.Stores) {
      try {
        $sid = $st.StoreID
        if ($sid -and ($sid.StartsWith($StoreId) -or $StoreId.StartsWith($sid) -or $sid -like "*${StoreId}*")) { $target = $st; break }
      } catch {}
    }
  }

  # 3) DisplayName exact
  if (-not $target) {
    foreach ($st in $session.Stores) { try { if ($st.DisplayName -eq $StoreId) { $target = $st; break } } catch {} }
  }

  # 4) DisplayName contenant
  if (-not $target) {
    foreach ($st in $session.Stores) { try { if ($st.DisplayName -like "*${StoreId}*") { $target = $st; break } } catch {} }
  }

  # 5) SMTP -> DeliveryStore
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

  # 6) Si toujours pas trouvé: construire payload diagnostique et retourner liste vide au lieu d'erreur dure
  if (-not $target) {
    $storesDiag = @()
    foreach ($st in $session.Stores) {
      try {
        $storesDiag += [pscustomobject]@{ DisplayName=$st.DisplayName; StoreId=($st.StoreID.Substring(0,120)); Len=$st.StoreID.Length }
      } catch {}
    }
    $payload = [pscustomobject]@{
      StoreId       = $StoreId
      ParentEntryId = $ParentEntryId
      ParentName    = $null
      Folders       = @()
      Error         = 'Store non trouvé'
      StoresDiag    = $storesDiag
    }
    $json = $payload | ConvertTo-Json -Depth 6 -Compress
    [Console]::Out.Write($json)
    exit 0
  }

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
