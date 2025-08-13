param(
  [Parameter(Mandatory=$true)][string]$Mailbox,
  [ValidateSet('Inbox','Root')] [string]$Scope = 'Inbox',
  [string]$ParentId,
  [int]$PageSize = 500,
  [string]$DllPath
)

$ErrorActionPreference = 'Stop'
function Get-EwsDll {
  $candidatePaths = @()
  try {
    $base = Split-Path -Parent $MyInvocation.MyCommand.Path
    # 1) Dans le même dossier (dev possible)
    $candidatePaths += (Join-Path $base 'Microsoft.Exchange.WebServices.dll')
  # 2) Dev: projet racine => resources\ews\
  $candidatePaths += (Join-Path $base '..\resources\ews\Microsoft.Exchange.WebServices.dll')
  # 3) Packagé: resources\scripts\... => DLL attendue dans resources\ews\
  $candidatePaths += (Join-Path $base '..\ews\Microsoft.Exchange.WebServices.dll')
  # 4) Fallback additionnel (variantes d'arborescence)
  $candidatePaths += (Join-Path $base '..\..\resources\ews\Microsoft.Exchange.WebServices.dll')
  } catch {}

  foreach ($p in $candidatePaths) {
    if ($p -and (Test-Path $p)) { return $p }
  }
  return $null
}

function Write-Json {
  param($obj)
  $json = $obj | ConvertTo-Json -Depth 6 -Compress
  [Console]::Out.Write($json)
}

try {
  $ewsPath = $null
  if ($DllPath -and (Test-Path $DllPath)) {
    $ewsPath = $DllPath
  } else {
    $ewsPath = Get-EwsDll
  }
  if (-not $ewsPath) { throw "EWS Managed API introuvable. DLL manquante (Microsoft.Exchange.WebServices.dll)" }
  Add-Type -Path $ewsPath

  $version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
  $svc = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new($version)
  $svc.UseDefaultCredentials = $true
  $redirect = { param($url) $true }
  $svc.AutodiscoverUrl($Mailbox, $redirect)

  $view = [Microsoft.Exchange.WebServices.Data.FolderView]::new($PageSize)
  $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
  $props = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
  $props.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
  $props.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::ChildFolderCount)
  $view.PropertySet = $props

  $parentFolderId = $null
  if ($ParentId) {
    $parentFolderId = [Microsoft.Exchange.WebServices.Data.FolderId]::new($ParentId)
  } else {
    if ($Scope -eq 'Inbox') {
      $inbox = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $Mailbox)
      $parentFolderId = $inbox
    } else {
      $root = [Microsoft.Exchange.WebServices.Data.FolderId]::new([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $Mailbox)
      $parentFolderId = $root
    }
  }

  $result = $svc.FindFolders($parentFolderId, $view)
  $list = @()
  foreach ($f in $result.Folders) {
    $list += [pscustomobject]@{
      Id         = $f.Id.UniqueId
      Name       = $f.DisplayName
      ChildCount = $f.ChildFolderCount
    }
  }

  Write-Json @(
    [pscustomobject]@{
      Mailbox  = $Mailbox
      ParentId = $ParentId
      Scope    = $Scope
      Folders  = $list
    }
  )
} catch {
  $err = $_.Exception.Message
  [Console]::Error.WriteLine("EWS-ERR: $err")
  exit 1
}
