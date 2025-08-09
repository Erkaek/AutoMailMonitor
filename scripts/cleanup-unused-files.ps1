# Removes unused legacy server files from the project
$ErrorActionPreference = "SilentlyContinue"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent $root
$server = Join-Path $repo "src/server"

$files = @(
  "server.js",
  "htmlTemplate.js",
  "graphOutlookConnector.js",
  "outlookConnector-backup.js",
  "optimizedOutlookConnector.js",
  "comConnector.js",
  "comConnectorEdge.js",
  "outlookConnector-ultra-robust.js",
  "outlookEventConnector.js"
)

foreach ($f in $files) {
  $p = Join-Path $server $f
  if (Test-Path $p) {
    Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
    Write-Host "Removed: $p"
  } else {
    Write-Host "Not found (skip): $p"
  }
}

Write-Host "Cleanup complete."
