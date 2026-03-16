#PowerShell-EWS-Appflow-GetFolderNameById.ps1
 
<#
.SYNOPSIS
  App-only OAuth (client credentials) + EWS Impersonation + Raw SOAP POST
  Get the DisplayName (folder name) for a folder using its EWS FolderId.

.NOTES
  - X-AnchorMailbox MUST be set to the impersonated mailbox when using OAuth + impersonation. (MS Learn)  
  - EWS IDs are case-sensitive; do not modify casing/encoding of FolderId.                 [3](https://microsoft.sharepoint-df.com/teams/ODSPOnboarding/Shared%20Documents/Learning%20Library%20Materials/Understanding%20SharePoint%20Online%20v2.pdf?web=1)
  - GetFolder SOAP structure is per EWS GetFolder operation reference.                    [1](https://morgantechspace.com/2022/03/connect-ews-api-with-modern-authentication-using-powershell.html)
  - Impersonation header is ExchangeImpersonation in SOAP header.                         [2](https://www.youtube.com/watch?v=dOqDNYrE2MU)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =========================
# CONFIG (fill these in)
# =========================
$TenantId     = 'YOUR_TENANT_ID_GUID'         # TODO: Set this 
$ClientId     = 'YOUR_APP_CLIENT_ID_GUID'     # TODO: Set this 
$ClientSecret = 'YOUR_CLIENT_SECRET_VALUE'    # TODO: Set this 

# Target mailbox to impersonate
$ImpersonatedMailboxSmtp = 'user@contoso.com' # TODO: Set this to the SMTP of the mailbox to access

# EWS FolderId (Id attribute from EWS) - DO NOT change casing/encoding.
$FolderEwsId = 'AAMkADIy...AAA='  # TODO: Set this to the folder to access

# Optional change key (can be empty). If you have it, include it; if not, leave blank.
$FolderChangeKey = ''             # TODO: Set if desired - not required

# EWS endpoint
$EwsEndpoint = 'https://outlook.office365.com/EWS/Exchange.asmx'

# Optional: log SOAP request/response (avoid logging secrets/tokens in real environments)
$LogPath = Join-Path $PSScriptRoot 'ews_getfolder_log.txt'

# Add credentials
$include = Join-Path $PSScriptRoot 'PowerShell-EWS-Appflow-FoldersExample_Creds.ps1'  # TODO: Update this file to set application oAuth credentails.
. $include

function Get-EwsAppOnlyAccessToken {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret
    )

    # OAuth app-only token request for EWS uses scope: https://outlook.office365.com/.default  
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://outlook.office365.com/.default'
        grant_type    = 'client_credentials'
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType 'application/x-www-form-urlencoded' -Body $body
    if ([string]::IsNullOrWhiteSpace($resp.access_token)) {
        throw "Failed to obtain access token. Response: $($resp | ConvertTo-Json -Depth 5)"
    }
    return $resp.access_token
}


function Get-EwsFolderDisplayNameById {
    param(
        [Parameter(Mandatory)][string]$EwsEndpoint,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][string]$ImpersonatedMailboxSmtp,
        [Parameter(Mandatory)][string]$FolderEwsId,
        [string]$FolderChangeKey = '',
        [string]$LogPath = ''
    )

    # Build the <t:FolderId .../> element (include ChangeKey only if provided).
    $folderIdXml = if ([string]::IsNullOrWhiteSpace($FolderChangeKey)) {
        "<t:FolderId Id=""$FolderEwsId"" />"
    } else {
        "<t:FolderId Id=""$FolderEwsId"" ChangeKey=""$FolderChangeKey"" />"
    }

    # SOAP GetFolder request; ask for DisplayName via AdditionalProperties.  [1](https://morgantechspace.com/2022/03/connect-ews-api-with-modern-authentication-using-powershell.html)
    # Impersonation via ExchangeImpersonation SOAP header.                   [2](https://www.youtube.com/watch?v=dOqDNYrE2MU)
    $soap = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">

  <soap:Header>
    <t:RequestServerVersion Version="Exchange2016" />
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$ImpersonatedMailboxSmtp</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
  </soap:Header>

  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:DisplayName" />
        </t:AdditionalProperties>
      </m:FolderShape>
      <m:FolderIds>
        $folderIdXml
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>

</soap:Envelope>
"@

    # HTTP headers:
    # - Authorization: Bearer <token>
    # - X-AnchorMailbox: MUST be impersonated mailbox when using OAuth + impersonation. 
    $headers = @{
        Authorization    = "Bearer $AccessToken"
        'X-AnchorMailbox' = $ImpersonatedMailboxSmtp
        Accept           = 'text/xml'
    }

    # Some environments/proxies like having SOAPAction; harmless to include for GetFolder.
    $headers['SOAPAction'] = 'http://schemas.microsoft.com/exchange/services/2006/messages/GetFolder'

    if ($LogPath) {
        "----- SOAP REQUEST -----`r`n$soap`r`n" | Out-File -FilePath $LogPath -Encoding utf8 -Append
    }

    $resp = Invoke-WebRequest -Method Post -Uri $EwsEndpoint -Headers $headers -ContentType 'text/xml; charset=utf-8' -Body $soap

    if ($LogPath) {
        "----- SOAP RESPONSE -----`r`n$($resp.Content)`r`n" | Out-File -FilePath $LogPath -Encoding utf8 -Append
    }

    # Parse XML response
    [xml]$xml = $resp.Content
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('soap','http://schemas.xmlsoap.org/soap/envelope/')
    $ns.AddNamespace('m','http://schemas.microsoft.com/exchange/services/2006/messages')
    $ns.AddNamespace('t','http://schemas.microsoft.com/exchange/services/2006/types')

    # Validate response class
    $msgNode = $xml.SelectSingleNode('//m:GetFolderResponseMessage', $ns)
    if (-not $msgNode) { throw "Unexpected response: missing GetFolderResponseMessage." }

    $rc = $msgNode.GetAttribute('ResponseClass')
    if ($rc -ne 'Success') {
        $code = $xml.SelectSingleNode('//m:ResponseCode', $ns)?.InnerText
        $text = $xml.SelectSingleNode('//m:MessageText', $ns)?.InnerText
        throw "EWS GetFolder failed. ResponseClass=$rc ResponseCode=$code MessageText=$text"
    }

    $displayName = $xml.SelectSingleNode('//t:DisplayName', $ns)?.InnerText
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        throw "DisplayName not found in response."
    }

    return $displayName
}

# =========================
# RUN
# =========================
$token = Get-EwsAppOnlyAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

$folderName = Get-EwsFolderDisplayNameById `
    -EwsEndpoint $EwsEndpoint `
    -AccessToken $token `
    -ImpersonatedMailboxSmtp $ImpersonatedMailboxSmtp `
    -FolderEwsId $FolderEwsId `
    -FolderChangeKey $FolderChangeKey `
    -LogPath $LogPath

[pscustomobject]@{
    Mailbox     = $ImpersonatedMailboxSmtp
    FolderEwsId = $FolderEwsId
    DisplayName = $folderName
}

