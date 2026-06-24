param(
    [string]$EwsUrl = "",
    [string]$Mailbox = "",
    [string]$FolderPath = "",
    [string]$FolderEntryId = "",
    [int]$MaxItems = 1000,
    [int]$HoursBack = 8760,
    [switch]$UnreadOnly,
    [switch]$AllItems,
    [switch]$UseLastModificationTime,
    [string]$ModifiedSince = "",
    [string]$ModifiedBefore = ""
)

$ErrorActionPreference = "Stop"
$VerbosePreference = "SilentlyContinue"
$InformationPreference = "SilentlyContinue"
$ProgressPreference = "Continue"

$enc = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $enc
$OutputEncoding = $enc

# ============================================================
# UTILITAIRES
# ============================================================

function Escape-Xml {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-NamespaceManager {
    param([System.Xml.XmlDocument]$Xml)
    $manager = New-Object -TypeName System.Xml.XmlNamespaceManager -ArgumentList $Xml.NameTable
    [void]$manager.AddNamespace("s", "http://schemas.xmlsoap.org/soap/envelope/")
    [void]$manager.AddNamespace("m", "http://schemas.microsoft.com/exchange/services/2006/messages")
    [void]$manager.AddNamespace("t", "http://schemas.microsoft.com/exchange/services/2006/types")
    return $manager
}

function Invoke-EwsSoap {
    param(
        [string]$SoapXml,
        [string]$SoapAction
    )
    $headers = @{ "SOAPAction" = $SoapAction }
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($SoapXml)
    
    try {
        $response = Invoke-WebRequest `
            -Uri $EwsUrl `
            -Method Post `
            -UseDefaultCredentials `
            -UseBasicParsing `
            -Headers $headers `
            -ContentType "text/xml; charset=utf-8" `
            -Body $bodyBytes `
            -TimeoutSec 30
        return $response.Content
    } catch {
        throw "Erreur EWS SOAP: $($_.Exception.Message)"
    }
}

function Get-NodeText {
    param($Node, [string]$XPath, $NamespaceManager)
    $found = $Node.SelectSingleNode($XPath, $NamespaceManager)
    if ($null -eq $found) { return "" }
    return $found.InnerText
}

function Get-NodeBool {
    param($Node, [string]$XPath, $NamespaceManager)
    $text = Get-NodeText -Node $Node -XPath $XPath -NamespaceManager $NamespaceManager
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    return [System.Convert]::ToBoolean($text)
}

function Get-NodeDate {
    param($Node, [string]$XPath, $NamespaceManager)
    $text = Get-NodeText -Node $Node -XPath $XPath -NamespaceManager $NamespaceManager
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return ([datetime]$text).ToLocalTime()
}

function New-DistinguishedFolderXml {
    param([string]$Id, [string]$MailboxSmtp)
    if ([string]::IsNullOrWhiteSpace($MailboxSmtp)) {
        return "<t:DistinguishedFolderId Id=""$Id"" />"
    }
    $safeMailbox = Escape-Xml $MailboxSmtp
    return @"
<t:DistinguishedFolderId Id="$Id">
  <t:Mailbox>
    <t:EmailAddress>$safeMailbox</t:EmailAddress>
  </t:Mailbox>
</t:DistinguishedFolderId>
"@
}

function New-FolderIdXml {
    param([string]$Id, [string]$ChangeKey)
    $safeId = Escape-Xml $Id
    $safeChangeKey = Escape-Xml $ChangeKey
    if ([string]::IsNullOrWhiteSpace($safeChangeKey)) {
        return "<t:FolderId Id=""$safeId"" />"
    }
    return "<t:FolderId Id=""$safeId"" ChangeKey=""$safeChangeKey"" />"
}

function Get-EwsChildFolders {
    param([string]$ParentFolderXml)
    $allFolders = @()
    $offset = 0
    $includesLastFolder = $false
    $pageSize = 100

    while (-not $includesLastFolder) {
        $soap = @"
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </s:Header>
  <s:Body>
    <m:FindFolder Traversal="Shallow">
      <m:FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </m:FolderShape>
      <m:IndexedPageFolderView MaxEntriesReturned="$pageSize" Offset="$offset" BasePoint="Beginning" />
      <m:ParentFolderIds>
        $ParentFolderXml
      </m:ParentFolderIds>
    </m:FindFolder>
  </s:Body>
</s:Envelope>
"@

        $content = Invoke-EwsSoap -SoapXml $soap -SoapAction "http://schemas.microsoft.com/exchange/services/2006/messages/FindFolder"
        [xml]$xml = $content
        $ns = New-NamespaceManager -Xml $xml

        $responseMessage = $xml.SelectSingleNode("//m:FindFolderResponseMessage", $ns)
        if ($null -eq $responseMessage) { throw "Réponse EWS inattendue sur FindFolder." }

        $responseClass = $responseMessage.GetAttribute("ResponseClass")
        $responseCode = Get-NodeText -Node $xml -XPath "//m:ResponseCode" -NamespaceManager $ns

        if ($responseClass -ne "Success") { throw "Erreur EWS FindFolder : $responseCode" }

        $rootFolder = $xml.SelectSingleNode("//m:RootFolder", $ns)
        if ($null -eq $rootFolder) { break }

        $includesLastFolder = [System.Convert]::ToBoolean($rootFolder.GetAttribute("IncludesLastItemInRange"))
        $folderNodes = $xml.SelectNodes("//m:RootFolder/t:Folders/*", $ns)

        foreach ($node in $folderNodes) {
            $displayNameNode = $node.SelectSingleNode("t:DisplayName", $ns)
            $folderIdNode = $node.SelectSingleNode("t:FolderId", $ns)
            $totalCountNode = $node.SelectSingleNode("t:TotalCount", $ns)

            if ($null -ne $folderIdNode) {
                $allFolders += [PSCustomObject]@{
                    DisplayName = if ($displayNameNode) { $displayNameNode.InnerText } else { "" }
                    Id          = $folderIdNode.GetAttribute("Id")
                    ChangeKey   = $folderIdNode.GetAttribute("ChangeKey")
                    TotalCount  = if ($totalCountNode) { [int]$totalCountNode.InnerText } else { 0 }
                }
            }
        }
        $offset += $pageSize
    }
    return $allFolders
}

function Get-EwsFolderByPath {
    param([string[]]$Path, [string]$MailboxSmtp)

    if ($Path.Count -eq 0) { throw "Le chemin du dossier est vide." }

    $currentFolderXml = New-DistinguishedFolderXml -Id "inbox" -MailboxSmtp $MailboxSmtp
    $startIndex = 0

    if ($Path[0] -eq "inbox" -or $Path[0] -eq "Inbox" -or $Path[0] -eq "Boîte de réception") {
        $startIndex = 1
    }

    for ($i = $startIndex; $i -lt $Path.Count; $i++) {
        $wantedName = $Path[$i]
        $children = Get-EwsChildFolders -ParentFolderXml $currentFolderXml
        $match = $children | Where-Object { $_.DisplayName -eq $wantedName } | Select-Object -First 1

        if ($null -eq $match) {
            throw "Dossier introuvable : $wantedName"
        }

        $currentFolderXml = New-FolderIdXml -Id $match.Id -ChangeKey $match.ChangeKey
    }

    return @{
        FolderXml  = $currentFolderXml
        Path       = ($Path -join "\")
        TotalCount = 0
    }
}

# ============================================================
# RECUPERATION DES MAILS EWS
# ============================================================

function Get-EwsMessagesFromFolder {
    param([object]$Folder, [int]$MaxItems)

    $offset = 0
    $includesLastItem = $false
    $pageSize = 200
    $allResults = @()
    $processed = 0

    while (-not $includesLastItem -and $processed -lt $MaxItems) {
        $folderXml = $Folder.FolderXml
        $itemsToRequest = [Math]::Min($pageSize, $MaxItems - $processed)

        $soap = @"
<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </s:Header>
  <s:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="item:Importance" />
          <t:FieldURI FieldURI="item:Categories" />
          <t:FieldURI FieldURI="item:HasAttachments" />
          <t:FieldURI FieldURI="item:LastModifiedTime" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="message:IsRead" />
          <t:FieldURI FieldURI="message:InternetMessageId" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="$itemsToRequest" Offset="$offset" BasePoint="Beginning" />
      <m:ParentFolderIds>
        $folderXml
      </m:ParentFolderIds>
    </m:FindItem>
  </s:Body>
</s:Envelope>
"@

        $content = Invoke-EwsSoap -SoapXml $soap -SoapAction "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem"
        [xml]$xml = $content
        $ns = New-NamespaceManager -Xml $xml

        $responseMessage = $xml.SelectSingleNode("//m:FindItemResponseMessage", $ns)
        if ($null -eq $responseMessage) { throw "Réponse EWS inattendue sur FindItem." }

        $responseClass = $responseMessage.GetAttribute("ResponseClass")
        $responseCode = Get-NodeText -Node $xml -XPath "//m:ResponseCode" -NamespaceManager $ns

        if ($responseClass -ne "Success") { throw "Erreur EWS FindItem : $responseCode" }

        $rootFolder = $xml.SelectSingleNode("//m:RootFolder", $ns)
        if ($null -eq $rootFolder) { break }

        $includesLastItem = [System.Convert]::ToBoolean($rootFolder.GetAttribute("IncludesLastItemInRange"))
        $messages = $xml.SelectNodes("//t:Message", $ns)

        foreach ($msg in $messages) {
            $categoryNodes = $msg.SelectNodes("t:Categories/t:String", $ns)
            $categories = @()
            foreach ($cat in $categoryNodes) {
                $categories += $cat.InnerText
            }

            $receivedTime = Get-NodeDate -Node $msg -XPath "t:DateTimeReceived" -NamespaceManager $ns
            $lastModTime = Get-NodeDate -Node $msg -XPath "t:LastModifiedTime" -NamespaceManager $ns

            $allResults += [PSCustomObject]@{
                Subject               = Get-NodeText -Node $msg -XPath "t:Subject" -NamespaceManager $ns
                SenderName            = Get-NodeText -Node $msg -XPath "t:From/t:Mailbox/t:Name" -NamespaceManager $ns
                SenderEmailAddress    = Get-NodeText -Node $msg -XPath "t:From/t:Mailbox/t:EmailAddress" -NamespaceManager $ns
                ReceivedTime          = if ($receivedTime) { $receivedTime.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { "" }
                LastModificationTime  = if ($lastModTime) { $lastModTime.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { "" }
                EntryID               = Get-NodeText -Node $msg -XPath "t:ItemId" -NamespaceManager $ns
                InternetMessageId     = Get-NodeText -Node $msg -XPath "t:InternetMessageId" -NamespaceManager $ns
                UnRead                = Get-NodeBool -Node $msg -XPath "t:IsRead" -NamespaceManager $ns
                Importance            = Get-NodeText -Node $msg -XPath "t:Importance" -NamespaceManager $ns
                Categories            = ($categories -join ", ")
                HasAttachments        = Get-NodeBool -Node $msg -XPath "t:HasAttachments" -NamespaceManager $ns
                FolderPath            = $Folder.Path
                StoreName             = if ($Mailbox) { $Mailbox } else { "Exchange" }
                StoreId               = ""
            }

            $processed++
            if ($processed -ge $MaxItems) { break }
        }

        $offset += $pageSize
    }

    return $allResults
}

# ============================================================
# EXECUTION PRINCIPALE
# ============================================================

try {
    if (-not $EwsUrl) {
        throw "EwsUrl non fourni"
    }
    if (-not $Mailbox) {
        throw "Mailbox non fournie"
    }

    $maxItems = if ($AllItems) { 99999 } else { $MaxItems }

    # Résoudre le chemin du dossier
    $folderPathArray = if ($FolderPath) {
        @($FolderPath -split "\\" | Where-Object { $_ })
    } else {
        @("Inbox")
    }

    $targetFolder = Get-EwsFolderByPath -Path $folderPathArray -MailboxSmtp $Mailbox

    # Récupérer les emails
    $emails = Get-EwsMessagesFromFolder -Folder $targetFolder -MaxItems $maxItems

    # Calculer les dates min/max
    $maxLmt = $null
    $minLmt = $null
    foreach ($email in $emails) {
        if ($email.LastModificationTime) {
            if (-not $maxLmt) { $maxLmt = $email.LastModificationTime }
            $minLmt = $email.LastModificationTime
        }
    }

    $result = @{
        success                = $true
        emails                 = $emails
        count                  = $emails.Count
        totalInFolder          = $emails.Count
        hasMore                = $false
        maxLastModificationTime = $maxLmt
        minLastModificationTime = $minLmt
        folderName             = ($folderPathArray[-1] ?? "Inbox")
        folderPath             = $FolderPath
        storeName              = $Mailbox
        storeId                = ""
        timestamp              = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    $result | ConvertTo-Json -Depth 6 -Compress

} catch {
    $err = $_.Exception.Message
    @{
        success       = $false
        error         = $err
        emails        = @()
        count         = 0
        totalInFolder = 0
        folderPath    = $FolderPath
    } | ConvertTo-Json -Depth 4 -Compress
}
