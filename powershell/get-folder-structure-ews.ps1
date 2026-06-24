param(
    [string]$EwsUrl = "",
    [string]$Mailbox = "",
    [int]$MaxDepth = 4,
    [string]$RootFolderId = ""
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

# ============================================================
# RECUPERATION RECURSIVE DES DOSSIERS
# ============================================================

function Get-EwsFoldersRecursive {
    param(
        [string]$ParentFolderXml,
        [int]$CurrentDepth,
        [string]$ParentPath,
        [string]$StoreName,
        [string]$StoreId,
        [ref]$AllFolders
    )

    if ($CurrentDepth -ge $MaxDepth) {
        return
    }

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
        if ($null -eq $responseMessage) {
            return
        }

        $responseClass = $responseMessage.GetAttribute("ResponseClass")
        if ($responseClass -ne "Success") {
            return
        }

        $rootFolder = $xml.SelectSingleNode("//m:RootFolder", $ns)
        if ($null -eq $rootFolder) {
            break
        }

        $includesLastFolder = [System.Convert]::ToBoolean($rootFolder.GetAttribute("IncludesLastItemInRange"))
        $folderNodes = $xml.SelectNodes("//m:RootFolder/t:Folders/t:Folder", $ns)

        foreach ($node in $folderNodes) {
            $displayNameNode = $node.SelectSingleNode("t:DisplayName", $ns)
            $folderIdNode = $node.SelectSingleNode("t:FolderId", $ns)
            $totalCountNode = $node.SelectSingleNode("t:TotalCount", $ns)
            $childCountNode = $node.SelectSingleNode("t:ChildFolderCount", $ns)

            if ($null -ne $folderIdNode -and $null -ne $displayNameNode) {
                $displayName = $displayNameNode.InnerText
                $folderId = $folderIdNode.GetAttribute("Id")
                $changeKey = $folderIdNode.GetAttribute("ChangeKey")
                $childCount = if ($childCountNode) { [int]$childCountNode.InnerText } else { 0 }
                
                $fullPath = if ($ParentPath) { "$ParentPath\$displayName" } else { "$StoreName\$displayName" }
                
                $AllFolders.Value += [PSCustomObject]@{
                    StoreDisplayName = $StoreName
                    StoreEntryID     = $StoreId
                    FolderName       = $displayName
                    FolderEntryID    = $folderId
                    FullPath         = $fullPath
                    ChildCount       = $childCount
                    Depth            = $CurrentDepth
                }

                # Récurser si enfants et profondeur pas atteinte
                if ($childCount -gt 0) {
                    $childFolderXml = New-FolderIdXml -Id $folderId -ChangeKey $changeKey
                    Get-EwsFoldersRecursive -ParentFolderXml $childFolderXml `
                                           -CurrentDepth ($CurrentDepth + 1) `
                                           -ParentPath $fullPath `
                                           -StoreName $StoreName `
                                           -StoreId $StoreId `
                                           -AllFolders $AllFolders
                }
            }
        }

        $offset += $pageSize
    }
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

    $allFolders = @()

    # Obtenir la racine des dossiers
    if ($RootFolderId) {
        $rootFolderXml = New-FolderIdXml -Id $RootFolderId -ChangeKey ""
    } else {
        $rootFolderXml = New-DistinguishedFolderXml -Id "root" -MailboxSmtp $Mailbox
    }

    # Récupérer tous les dossiers récursivement
    Get-EwsFoldersRecursive -ParentFolderXml $rootFolderXml `
                           -CurrentDepth 0 `
                           -ParentPath "" `
                           -StoreName $Mailbox `
                           -StoreId "" `
                           -AllFolders ([ref]$allFolders)

    $result = @{
        success = $true
        folders = $allFolders
        count   = $allFolders.Count
        mailbox = $Mailbox
        timestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    $result | ConvertTo-Json -Depth 6 -Compress

} catch {
    $err = $_.Exception.Message
    @{
        success = $false
        error   = $err
        folders = @()
        count   = 0
        mailbox = $Mailbox
    } | ConvertTo-Json -Depth 4 -Compress
}
