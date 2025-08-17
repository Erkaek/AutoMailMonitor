param(
	[string]$StoreId = "",
	[string]$StoreName = "",
	[int]$MaxDepth = -1
)

$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

function Get-FolderFlat {
	param(
		[Parameter(Mandatory=$true)] $Folder,
		[Parameter(Mandatory=$true)] [string] $ParentPath,
		[Parameter(Mandatory=$true)] [int] $Depth,
		[Parameter(Mandatory=$true)] [string] $StoreName,
		[Parameter(Mandatory=$true)] [string] $StoreEntryID,
		[Parameter(Mandatory=$true)] [int] $MaxDepth
	)
	$items = @()
	try {
		$name = ''
		try { $name = [string]$Folder.Name } catch {}
		if ([string]::IsNullOrEmpty($name)) { return @() }
		$curPath = if ([string]::IsNullOrEmpty($ParentPath)) { "$StoreName\$name" } else { "$ParentPath\$name" }
		$eid = ''; try { $eid = [string]$Folder.EntryID } catch {}
		$childCount = 0; try { $childCount = [int]$Folder.Folders.Count } catch {}
		$items += @([ordered]@{ StoreDisplayName=$StoreName; StoreEntryID=$StoreEntryID; FolderName=$name; FolderEntryID=$eid; FullPath=$curPath; ChildCount=$childCount })
		if ($MaxDepth -ge 0 -and $Depth -ge $MaxDepth) { return $items }
		if ($childCount -gt 0) {
			foreach ($ch in $Folder.Folders) {
				try { $items += Get-FolderFlat -Folder $ch -ParentPath $curPath -Depth ($Depth+1) -StoreName $StoreName -StoreEntryID $StoreEntryID -MaxDepth $MaxDepth } catch {} finally { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ch) | Out-Null } catch {} }
			}
		}
	} catch {}
	return $items
}

try {
	$ol = New-Object -ComObject Outlook.Application
		$ns = $ol.Session
		try { $null = $ns.Logon() } catch {}
	$stores = @(); foreach ($st in $ns.Stores) { try { $stores += $st } catch {} }
	if ($StoreName -and $StoreName.Trim() -ne '') {
		$filtered = @()
		foreach ($s in $stores) { try { if ($s.DisplayName -eq $StoreName -or $s.DisplayName -like ("*" + $StoreName + "*")) { $filtered += $s } } catch {} }
		$stores = $filtered
	} elseif ($StoreId -and $StoreId.Trim() -ne '') {
		$filtered = @()
		foreach ($s in $stores) { try { if ($s.StoreID -eq $StoreId -or $s.DisplayName -eq $StoreId -or $s.DisplayName -like ("*" + $StoreId + "*")) { $filtered += $s } } catch {} }
		$stores = $filtered
	}
	$all = @()
	foreach ($store in $stores) {
		$root = $null
		try {
			$sName = ''; try { $sName = [string]$store.DisplayName } catch {}
			$sId = ''; try { $sId = [string]$store.StoreID } catch {}
					try { $root = $store.GetRootFolder() } catch {}
					# Some shared stores report an empty root; fix via Namespace.Folders mapping
					try {
						$rc = 0; try { $rc = [int]$root.Folders.Count } catch {}
						if ($rc -eq 0) {
							foreach ($tf in $ns.Folders) { try { if ($tf.Store.StoreID -eq $store.StoreID) { $root = $tf; break } } catch {} }
						}
					} catch {}
					if ($root -ne $null) {
				foreach ($top in $root.Folders) {
					try { $all += Get-FolderFlat -Folder $top -ParentPath $sName -Depth 0 -StoreName $sName -StoreEntryID $sId -MaxDepth $MaxDepth } catch {} finally { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($top) | Out-Null } catch {} }
				}
						# If still nothing found, as a last resort, try going one level deeper where some profiles nest folders under a single child container
						if ($all.Count -eq 0) {
							foreach ($mid in $root.Folders) {
								try {
									foreach ($top2 in $mid.Folders) {
										try { $all += Get-FolderFlat -Folder $top2 -ParentPath $sName -Depth 0 -StoreName $sName -StoreEntryID $sId -MaxDepth $MaxDepth } catch {} finally { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($top2) | Out-Null } catch {} }
									}
								} catch {} finally { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mid) | Out-Null } catch {} }
							}
						}
			}
		} catch {}
		finally { if ($root) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($root) | Out-Null } catch {} } try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($store) | Out-Null } catch {} }
	}
	try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) | Out-Null } catch {}
	try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol) | Out-Null } catch {}
	[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
	@{ success = $true; folders = $all } | ConvertTo-Json -Depth 6 -Compress | Write-Output
} catch {
	$err = $_.Exception.Message
	@{ success = $false; error = $err; folders = @() } | ConvertTo-Json -Depth 3 -Compress | Write-Output
}
