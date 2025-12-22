param(
    [string]$StoreId = "",
    [string]$StoreName = "",
    [string]$FolderEntryId = "",
    [string]$FolderPath = "",
    [int]$MaxItems = 1000,
    [int]$HoursBack = 8760,
    [switch]$UnreadOnly,
    [switch]$AllItems,
    [switch]$UseLastModificationTime,
    [string]$ModifiedSince = "",
    [string]$ModifiedBefore = ""
)

$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

function Try-ParseDate([string]$s) {
    try {
        if (-not $s) { return $null }
        $t = $s.Trim()
        if ($t -eq '') { return $null }
        return [DateTime]::Parse($t, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal).ToLocalTime()
    } catch {
        return $null
    }
}

function Normalize-Name([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return "" }
    $s2 = $s.Trim().TrimEnd('.')
    $n = $s2.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $n.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant().Trim()
}

function Find-ChildByName($parentFolder, [string]$targetName) {
    $normTarget = Normalize-Name $targetName
    foreach ($f in $parentFolder.Folders) { if ((Normalize-Name $f.Name) -eq $normTarget) { return $f } }
    return $null
}

function Find-FolderFromPath([string]$Path, $Namespace, $FallbackInbox = $false) {
    if (-not $Path) { return $null }
    $parts = $Path -split "\\"
    if ($parts.Length -lt 2) { return $null }

    $accountName = $parts[0]
    $targetStore = $null
    foreach ($store in $Namespace.Stores) {
        try {
            if ((Normalize-Name $store.DisplayName) -eq (Normalize-Name $accountName)) { $targetStore = $store; break }
        } catch {}
    }
    if (-not $targetStore) { $targetStore = $Namespace.DefaultStore }

    try {
        $root = $targetStore.GetRootFolder()
    } catch { $root = $null }
    $cursor = $root

    # Certaines boites partagées ont un root vide; tenter Namespace.Folders mapping
    try {
        $childCount = 0; try { $childCount = [int]$cursor.Folders.Count } catch {}
        if ($childCount -eq 0) {
            foreach ($tf in $Namespace.Folders) { try { if ($tf.Store.StoreID -eq $targetStore.StoreID) { $cursor = $tf; break } } catch {} }
        }
    } catch {}

    # Naviguer dans le chemin (ignorer la première partie: store)
    for ($i = 1; $i -lt $parts.Length; $i++) {
        $name = $parts[$i]
        if (-not $cursor) { break }
        $next = Find-ChildByName -parentFolder $cursor -targetName $name
        if ($next -eq $null) {
            $cursor = $null
            break
        }
        $cursor = $next
    }

    if ($cursor) { return $cursor }

    if ($FallbackInbox) {
        try {
            $inbox = $targetStore.GetDefaultFolder(6)
            if ($inbox -ne $null) {
                $cursor = $inbox
                $startIndex = 1
                if ($parts.Length -gt 1) {
                    $p1 = $parts[1]
                    if ((Normalize-Name $p1) -eq (Normalize-Name $inbox.Name)) { $startIndex = 2 }
                }
                for ($i = $startIndex; $i -lt $parts.Length; $i++) {
                    $name = $parts[$i]
                    $next = Find-ChildByName -parentFolder $cursor -targetName $name
                    if ($next -eq $null) { break }
                    $cursor = $next
                }
                if ($cursor -ne $null) { return $cursor }
            }
        } catch {}
    }

    return $null
}

try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.Session
    try { $null = $ns.Logon() } catch {}

    $targetFolder = $null
    $storeNameOut = ""
    $storeIdOut = ""

    # 1) Résolution directe par EntryID
    if ($FolderEntryId -and $FolderEntryId.Trim() -ne '') {
        try {
            $targetFolder = $ns.GetFolderFromID($FolderEntryId, $StoreId)
        } catch {}
    }

    # 2) Résolution par chemin (store\chemin)
    if (-not $targetFolder -and $FolderPath -and $FolderPath.Trim() -ne '') {
        $targetFolder = Find-FolderFromPath -Path $FolderPath -Namespace $ns -FallbackInbox:$true
    }

    # 3) Fallback: si StoreId/StoreName fournis mais pas de chemin, tenter Inbox
    if (-not $targetFolder -and ($StoreId -or $StoreName)) {
        foreach ($store in $ns.Stores) {
            try {
                $match = $false
                if ($StoreId -and $store.StoreID -eq $StoreId) { $match = $true }
                elseif ($StoreName -and (Normalize-Name $store.DisplayName) -eq (Normalize-Name $StoreName)) { $match = $true }
                if ($match) { $targetFolder = $store.GetDefaultFolder(6); break }
            } catch {}
        }
    }

    if (-not $targetFolder) {
        $err = "Dossier introuvable (EntryID ou chemin)"
        @{ success = $false; error = $err; emails = @(); count = 0; totalInFolder = 0 } | ConvertTo-Json -Depth 4 -Compress | Write-Output
        return
    }

    try { $storeNameOut = [string]$targetFolder.Store.DisplayName } catch {}
    try { $storeIdOut = [string]$targetFolder.Store.StoreID } catch {}

    $items = $targetFolder.Items
    $totalCount = 0; try { $totalCount = [int]$items.Count } catch {}

    # Filtrage / tri
    $filters = @()

    $useModified = $false
    if ($UseLastModificationTime.IsPresent) { $useModified = $true }
    if ($ModifiedSince -and $ModifiedSince.Trim() -ne '') { $useModified = $true }
    if ($ModifiedBefore -and $ModifiedBefore.Trim() -ne '') { $useModified = $true }

    if ($useModified) {
        $sinceDt = Try-ParseDate $ModifiedSince
        $beforeDt = Try-ParseDate $ModifiedBefore
        if ($sinceDt) {
            $sinceStr = $sinceDt.ToString("yyyy-MM-dd HH:mm")
            $filters += "[LastModificationTime] >= '$sinceStr'"
        }
        if ($beforeDt) {
            $beforeStr = $beforeDt.ToString("yyyy-MM-dd HH:mm")
            $filters += "[LastModificationTime] < '$beforeStr'"
        }
    } else {
        if (-not $AllItems.IsPresent -and $HoursBack -gt 0) {
            $from = (Get-Date).AddHours(-1 * [double]$HoursBack)
            $fromStr = $from.ToString("yyyy-MM-dd HH:mm")
            $filters += "[ReceivedTime] >= '$fromStr'"
        }
    }

    if ($UnreadOnly.IsPresent) {
        $filters += "[Unread] = True"
    }

    if ($filters.Count -gt 0) {
        $filter = "(" + ($filters -join ") AND (") + ")"
        try { $items = $items.Restrict($filter) } catch {}
    }

    if ($useModified) {
        try { $items.Sort("[LastModificationTime]", $true) } catch {}
    } else {
        try { $items.Sort("[ReceivedTime]", $true) } catch {}
    }

    $upper = $MaxItems
    try { if ($items.Count -lt $upper) { $upper = [int]$items.Count } } catch {}
    if ($upper -lt 0) { $upper = 0 }

    $list = @()
    $maxLmt = $null
    $minLmt = $null
    for ($i = 1; $i -le $upper; $i++) {
        try {
            $m = $items.Item($i)
            if ($m -and $m.Class -eq 43) {
                $received = $null
                try { $received = $m.ReceivedTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } catch { $received = "" }
                $lmt = $null
                try { $lmt = $m.LastModificationTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } catch { $lmt = "" }
                $imid = $null
                try { $imid = [string]$m.InternetMessageID } catch { $imid = "" }

                if ($lmt -and $lmt.Trim() -ne '') {
                    if (-not $maxLmt) { $maxLmt = $lmt }
                    $minLmt = $lmt
                }
                $list += @([ordered]@{
                    Subject = if($m.Subject){$m.Subject}else{""}
                    SenderName = if($m.SenderName){$m.SenderName}else{""}
                    SenderEmailAddress = if($m.SenderEmailAddress){$m.SenderEmailAddress}else{""}
                    ReceivedTime = $received
                    LastModificationTime = $lmt
                    EntryID = $m.EntryID
                    InternetMessageId = if($imid){$imid}else{""}
                    UnRead = $m.UnRead
                    Importance = $m.Importance
                    Categories = if($m.Categories){$m.Categories}else{""}
                    FlagStatus = $m.FlagStatus
                    Size = $m.Size
                    ConversationTopic = if($m.ConversationTopic){$m.ConversationTopic}else{""}
                    HasAttachments = ($m.Attachments.Count -gt 0)
                    AttachmentCount = $m.Attachments.Count
                    FolderPath = if($targetFolder.FolderPath){$targetFolder.FolderPath}else{$FolderPath}
                    StoreName = $storeNameOut
                    StoreId = $storeIdOut
                })
            }
        } catch {}
    }

    $folderNameOut = try { $targetFolder.Name } catch { $FolderPath }
    $folderPathOut = $FolderPath

    $res = @{
        success = $true
        emails = $list
        count = $list.Count
        totalInFolder = $totalCount
        hasMore = $false
        maxLastModificationTime = $maxLmt
        minLastModificationTime = $minLmt
        folderName = $folderNameOut
        folderPath = $folderPathOut
        storeName = $storeNameOut
        storeId = $storeIdOut
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    try {
        $filteredCount = 0
        try { $filteredCount = [int]$items.Count } catch { $filteredCount = 0 }
        if ($filteredCount -gt $upper) { $res.hasMore = $true }
    } catch {}

    $res | ConvertTo-Json -Depth 6 -Compress | Write-Output

} catch {
    $err = $_.Exception.Message
    @{ success = $false; error = $err; emails = @(); count = 0; totalInFolder = 0 } | ConvertTo-Json -Depth 4 -Compress | Write-Output
}