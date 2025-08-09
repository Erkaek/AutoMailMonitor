param(
  [string]$ArchiveDir
)
if (-not $ArchiveDir) { Write-Error "Usage: restore-archived-tests.ps1 -ArchiveDir <path>"; exit 1 }
if (-not (Test-Path $ArchiveDir)) { Write-Error "ArchiveDir introuvable: $ArchiveDir"; exit 1 }

$files = Get-ChildItem -Path $ArchiveDir -File -ErrorAction SilentlyContinue
foreach ($f in $files) {
  $dest = Join-Path (Split-Path -Parent $PSScriptRoot) $f.Name
  Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
}
Write-Host "Restored $($files.Count) files from $ArchiveDir"
