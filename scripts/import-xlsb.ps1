param(
  [Parameter(Mandatory=$true)][string]$Path,
  [Parameter(Mandatory=$true)][int]$Year,
  [string]$Weeks
)
# Uses Excel COM to read fixed cells from sheets S1..S52 and returns JSON { rows, skippedWeeks, year }
# This avoids parsing binary data in Node. Requires Excel installed.

$ErrorActionPreference = 'Stop'

function Get-CellValue($ws, $addr) {
  try { $v = $ws.Range($addr).Value2 } catch { return 0 }
  if ($null -eq $v) { return 0 }
  if ($v -is [string]) { if ([int]::TryParse($v, [ref]$null)) { return [int]$v } else { return 0 } }
  try { return [int]([math]::Truncate([double]$v)) } catch { return 0 }
}

function Get-MondayISO([int]$year, [int]$week) {
  $simple = [datetime]::SpecifyKind([datetime]::new($year,1,1,0,0,0), 'Utc')
  $simple = $simple.AddDays(($week - 1) * 7)
  $dayOfWeek = [int]$simple.DayOfWeek
  $diff = (($dayOfWeek -as [int]) - 1)
  if ($dayOfWeek -le 4) { $diff = $dayOfWeek - 1 } else { $diff = $dayOfWeek - 7 - 1 }
  $isoStart = $simple.AddDays(-$diff)
  return $isoStart.ToString('yyyy-MM-dd')
}

$weeksSet = $null
if ($Weeks -and $Weeks.Trim() -ne '') {
  $weeksSet = New-Object System.Collections.Generic.HashSet[int]
  foreach ($part in $Weeks.Split(',')) {
    $p = $part.Trim()
    if ($p -like '*-*') {
      $a,$b = $p.Split('-')
      if ([int]::TryParse($a, [ref]$null) -and [int]::TryParse($b, [ref]$null)) {
        $start = [math]::Min([int]$a,[int]$b)
        $end = [math]::Max([int]$a,[int]$b)
        for ($k=$start; $k -le $end; $k++) { $null = $weeksSet.Add($k) }
      }
    } else {
      if ([int]::TryParse($p, [ref]$null)) { $null = $weeksSet.Add([int]$p) }
    }
  }
}

$excel = $null
$wb = $null
$rows = New-Object System.Collections.ArrayList
$skipped = 0

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($Path)

  # Read initial stocks from S1
  $initial = @{ Declarations = 0; Reglements = 0; MailSimple = 0 }
  $s1 = $wb.Sheets.Item('S1')
  if ($s1) {
    $initial['Declarations'] = Get-CellValue $s1 'M7'
    $initial['Reglements']   = Get-CellValue $s1 'M10'
    $initial['MailSimple']   = Get-CellValue $s1 'M13'
  }
  $stock = @{ Declarations = $initial['Declarations']; Reglements = $initial['Reglements']; MailSimple = $initial['MailSimple'] }

  for ($k=1; $k -le 52; $k++) {
    if ($weeksSet -and -not $weeksSet.Contains($k)) { continue }
    $sheetName = 'S' + $k
    $sh = $null
    try { $sh = $wb.Sheets.Item($sheetName) } catch { $sh = $null }
    if (-not $sh) { $skipped++ ; continue }

    $cats = @(
      @{ Key='Declarations'; Row=7 },
      @{ Key='Reglements'; Row=10 },
      @{ Key='MailSimple'; Row=13 }
    )
    $weekVals = @{}
    $allZero = $true
    foreach ($cat in $cats) {
      $recu = Get-CellValue $sh ('C' + $cat.Row)
      $traite = Get-CellValue $sh ('D' + $cat.Row)
      $traite_adg = Get-CellValue $sh ('E' + $cat.Row)
      if (($recu -ne 0) -or ($traite -ne 0) -or ($traite_adg -ne 0)) { $allZero = $false }
      $weekVals[$cat.Key] = @{ recu=$recu; traite=$traite; traite_adg=$traite_adg }
    }

    $s1Zero = ($k -ne 1) -or (($initial['Declarations'] -eq 0) -and ($initial['Reglements'] -eq 0) -and ($initial['MailSimple'] -eq 0))
    if ($allZero -and $s1Zero) { $skipped++ ; continue }

    foreach ($cat in $cats) {
      $key = $cat.Key
      $vals = $weekVals[$key]
      $sd = $stock[$key]
      $sf = $sd + $vals['recu'] - ($vals['traite'] + $vals['traite_adg'])
      $weekStart = Get-MondayISO $Year $k
      $row = [ordered]@{
        year = $Year
        week_number = $k
        week_start_date = $weekStart
        category = $key
        recu = $vals['recu']
        traite = $vals['traite']
        traite_adg = $vals['traite_adg']
        stock_debut = $sd
        stock_fin = $sf
      }
      $null = $rows.Add($row)
      $stock[$key] = $sf
    }
  }
}
finally {
  if ($wb) { $wb.Close($false) }
  if ($excel) { $excel.Quit() | Out-Null }
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null 2>$null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null 2>$null
}

$payload = @{ rows = $rows; skippedWeeks = $skipped; year = $Year } | ConvertTo-Json -Depth 4 -Compress
[Console]::Out.WriteLine($payload)
