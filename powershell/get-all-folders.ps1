param(
	[string]$StoreId = "",
	[string]$StoreName = "",
	[int]$MaxDepth = -1
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc

function New-FolderNode {
	param(
		[Parameter(Mandatory = $true)] $Folder,
		[Parameter(Mandatory = $true)] [string] $DisplayPath,
		[Parameter(Mandatory = $true)] [string] $StoreId
	)

	$name = ''
	try { $name = [string]$Folder.Name } catch {}
	if ([string]::IsNullOrWhiteSpace($name)) { return $null }

	$eid = ''
	try { $eid = [string]$Folder.EntryID } catch {}
	$childCount = 0
	try { $childCount = [int]$Folder.Folders.Count } catch {}

	return [ordered]@{
		Name        = $name
		EntryId     = $eid
		StoreId     = $StoreId
		DisplayPath = $DisplayPath
		ChildCount  = $childCount
		Children    = @()
	}
}

function Get-StoreRoot {
	param($Store, $Namespace)
	$root = $null
	try { $root = $Store.GetRootFolder() } catch {}
	if ($null -ne $root) {
		try {
			$c = 0; try { $c = [int]$root.Folders.Count } catch {}
			if ($c -gt 0) { return $root }
		} catch {}
	}
	# Certains stores partagés exposent un root vide; fallback via Namespace.Folders
	foreach ($tf in $Namespace.Folders) {
		try { if ($tf.Store.StoreID -eq $Store.StoreID) { return $tf } } catch {}
	}
	return $root
}

try {
	# Ne pas charger Microsoft.Office.Interop.Outlook pour éviter les problèmes de cast (32/64 bits)
	$ol = New-Object -ComObject Outlook.Application
	# Préférer Session pour éviter le cast COM Interop
	$ns = $null
	try { $ns = $ol.Session } catch {}
	if (-not $ns) { try { $ns = $ol.GetNamespace('MAPI') } catch {} }
	if (-not $ns) { throw "Namespace MAPI introuvable" }
	try { $null = $ns.Logon() } catch {}

	$stores = @()
	foreach ($st in $ns.Stores) { try { $stores += $st } catch {} }

	if ($StoreName -and $StoreName.Trim() -ne '') {
		$stores = @($stores | Where-Object { $_.DisplayName -eq $StoreName -or $_.DisplayName -like ("*" + $StoreName + "*") })
	} elseif ($StoreId -and $StoreId.Trim() -ne '') {
		$stores = @($stores | Where-Object { $_.StoreID -eq $StoreId -or $_.DisplayName -eq $StoreId -or $_.DisplayName -like ("*" + $StoreId + "*") })
	}

	$resultStores = @()

	foreach ($store in $stores) {
		$root = $null
		try {
			$storeName = ''
			try { $storeName = [string]$store.DisplayName } catch {}
			$storeIdVal = ''
			try { $storeIdVal = [string]$store.StoreID } catch {}

			$root = Get-StoreRoot -Store $store -Namespace $ns
			if ($null -eq $root) { continue }

			$rootEntryId = ''
			try { $rootEntryId = [string]$root.EntryID } catch {}
			$rootNode = [ordered]@{
				Name        = $storeName
				EntryId     = $rootEntryId
				StoreId     = $storeIdVal
				DisplayPath = $storeName
				ChildCount  = 0
				Children    = @()
			}

			$byId = @{}
			$byId[$rootNode.EntryId] = $rootNode

			$stack = New-Object System.Collections.Stack
			foreach ($top in $root.Folders) {
				try {
					$childPath = "$storeName\$($top.Name)"
					$node = New-FolderNode -Folder $top -DisplayPath $childPath -StoreId $storeIdVal
					if ($null -eq $node) { continue }
					$byId[$node.EntryId] = $node
					$rootNode.Children += $node
					$stack.Push(@{ Folder = $top; Depth = 1; NodeId = $node.EntryId; Path = $childPath })
				} catch {}
			}

			while ($stack.Count -gt 0) {
				$ctx = $stack.Pop()
				$folder = $ctx.Folder
				$depth = $ctx.Depth
				$path = $ctx.Path
				$parentId = $ctx.NodeId

				foreach ($child in $folder.Folders) {
					try {
						$childPath = "$path\$($child.Name)"
						$node = New-FolderNode -Folder $child -DisplayPath $childPath -StoreId $storeIdVal
						if ($null -eq $node) { continue }
						$byId[$node.EntryId] = $node

						if ($byId.ContainsKey($parentId)) {
							$byId[$parentId].Children += $node
						}

						$nextDepth = $depth + 1
						if ($MaxDepth -lt 0 -or $nextDepth -lt $MaxDepth) {
							$stack.Push(@{ Folder = $child; Depth = $nextDepth; NodeId = $node.EntryId; Path = $childPath })
						}
					} catch {}
				}
			}

			foreach ($id in $byId.Keys) {
				$n = $byId[$id]
				if ($n -and $n.Children) { $n.ChildCount = $n.Children.Count }
			}

			$rootNode.ChildCount = $rootNode.Children.Count

			$smtp = $null
			try {
				$accounts = $ol.Session.Accounts
				foreach ($acc in $accounts) {
					try {
						if ($acc.DisplayName -eq $storeName -or $storeName -like "*$($acc.DisplayName)*") {
							if ($acc.SmtpAddress) { $smtp = $acc.SmtpAddress; break }
						}
					} catch {}
				}
			} catch {}

			$resultStores += [ordered]@{
				Name        = $storeName
				StoreId     = $storeIdVal
				SmtpAddress = $smtp
				Root        = $rootNode
			}
		} catch {}
		finally {
			if ($root) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($root) | Out-Null } catch {} }
			try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($store) | Out-Null } catch {}
		}
	}

	try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) | Out-Null } catch {}
	try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol) | Out-Null } catch {}
	[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

	@{ success = $true; stores = $resultStores } | ConvertTo-Json -Depth 64 -Compress | Write-Output
} catch {
	$err = $_.Exception.Message
	@{ success = $false; error = $err; stores = @() } | ConvertTo-Json -Depth 6 -Compress | Write-Output
}
