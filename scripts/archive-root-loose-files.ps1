# Move any loose root-level scripts/configs not part of runtime into archives/loose-<timestamp>
Param(
  [string]$Reason = "declutter"
)

$ErrorActionPreference = "Stop"

$ts = Get-Date -Format "yyyyMMdd-HHmmss"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
$dest = Join-Path -Path $repoRoot -ChildPath "archives/loose-$ts"
New-Item -ItemType Directory -Path $dest -Force | Out-Null

# Allowlist of files to keep at root
$keep = @(
  "package.json","package-lock.json","README.md","LICENSE","CHANGELOG.md","QUICK-WINS-PERFORMANCE.md",
  "RAPPORT-COHERENCE-BDD.md","RAPPORT-FINAL-OPTIMISATION.md","SOLUTION-MODULES-NATIFS.md",
  ".gitignore",".nvmrc"
)

# Only consider files in root
$files = Get-ChildItem -Path $repoRoot -File

$moved = 0
foreach ($f in $files) {
  if ($keep -contains $f.Name) { continue }
  # Archive typical helper/debug formats
  if ($f.Extension -in @(".js",".ps1",".sql",".bat",".psm1",".psd1")) {
    Write-Host "Archiving $($f.Name) -> $dest"
    Move-Item -LiteralPath $f.FullName -Destination $dest -Force
    $moved++
  }
}

"Archived $moved loose files to $dest ($Reason)" | Write-Host
