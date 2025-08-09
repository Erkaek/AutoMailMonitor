param(
  [string]$ArchiveDir = "archives/tests-$(Get-Date -Format yyyyMMdd-HHmmss)"
)

New-Item -ItemType Directory -Force -Path $ArchiveDir | Out-Null

$patterns = @('test-*.*', 'fix-*.js', 'fix-*.ps1', 'fix-*.sql', 'fix-*.bat')
$files = Get-ChildItem -Path . -Recurse -File -Include $patterns -ErrorAction SilentlyContinue |
  Where-Object { $_.FullName -notmatch "\\archives\\" }

$count = 0
foreach ($f in $files) {
  $dest = Join-Path $ArchiveDir $f.Name
  Move-Item -LiteralPath $f.FullName -Destination $dest -Force
  $count++
}

Write-Host "Archived $count files to $ArchiveDir"