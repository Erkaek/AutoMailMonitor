# Archives legacy/maintenance scripts from the repo root into a timestamped folder under ./archives
Param(
    [string]$Reason = "cleanup"
)

$ErrorActionPreference = "Stop"

function New-ArchiveFolder {
    $ts = Get-Date -Format "yyyyMMdd-HHmmss"
    $dest = Join-Path -Path (Resolve-Path ".") -ChildPath "archives/maintenance-$ts"
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    return $dest
}

$filesToArchive = @(
    # analysis / checks
    "add-test-folder-config.js",
    "analyze-code-db-coherence.js",
    "analyze-column-usage.js",
    "analyze-database.js",
    "analyze-db-code-coherence.js",
    "analyze-folder-names.ps1",
    "check-all-paths.js",
    "check-missing-columns.js",
    "check-schema-final.js",
    "check-schema.sql",
    # cleanup/maintenance
    "clean-obsolete-references.js",
    "cleanup-database.js",
    "cleanup-unused-columns.js",
    # configure/setup/migrate
    "configure-inbox.sql",
    "configure-real-folders.js",
    "finalize-weekly-setup.js",
    "migrate-database-final.js",
    "migrate-folders-config.js",
    "recreate-weekly-tables.js",
    "setup-db-monitoring.js",
    "setup-inbox-monitoring.js",
    "setup-test-folder.js",
    "update-service-final-schema.js",
    "verify-schema-final.js",
    # debug/explore utilities
    "debug-check-folders.js",
    "debug-folders-config.js",
    "debug-list-outlook-structure.ps1",
    "debug-list-real-folders.ps1",
    "debug-weekly-stats.js",
    "explore-database.js",
    "explore-exact-structure.ps1",
    "inspect-database.js",
    "inspect-db-structure.js",
    # misc convenience
    "start.bat"
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
$destDir = New-ArchiveFolder

$moved = 0
foreach ($name in $filesToArchive) {
    $src = Join-Path $repoRoot $name
    if (Test-Path $src) {
        Write-Host "Archiving $name -> $destDir"
        Move-Item -LiteralPath $src -Destination $destDir -Force
        $moved++
    }
}

"Archived $moved files to $destDir ($Reason)" | Write-Host
