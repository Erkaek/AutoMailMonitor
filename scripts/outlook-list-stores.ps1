param()

$ErrorActionPreference = 'Stop'

# Force UTF-8 output
$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc

try {
  try { Add-Type -AssemblyName Microsoft.Office.Interop.Outlook | Out-Null } catch {}
  $olApp = New-Object -ComObject Outlook.Application
  $session = $olApp.Session

  # Map des comptes -> SMTP
  $accMap = @{}
  try {
    foreach ($acc in $session.Accounts) {
      try { $accMap[$acc.DisplayName] = $acc.SmtpAddress } catch {}
    }
  } catch {}

  $stores = @()
  foreach ($store in $session.Stores) {
    try {
      $display = $store.DisplayName
      $smtp = $null
      if ($accMap.ContainsKey($display)) { $smtp = $accMap[$display] }
      $stores += [pscustomobject]@{
        DisplayName = $display
        StoreId     = $store.StoreID
        SmtpAddress = $smtp
      }
    } catch {}
  }

  $json = $stores | ConvertTo-Json -Depth 5 -Compress
  [Console]::Out.Write($json)
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
