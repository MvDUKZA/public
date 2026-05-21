#Requires -Version 5.1
<#
.SYNOPSIS
    Retrieves Qualys Policy Compliance failures with tag-based host filtering.

.DESCRIPTION
    Connects to the Qualys EU platform (qualysapi.qualys.eu), resolves tag IDs
    by name, then retrieves all FAIL posture records for Policy ID 1661380 where
    the host is tagged with "All.Workstations" (any) but NOT tagged with
    "VSI Testing" (any).

    Results are displayed in the console and exported to a timestamped CSV file.

.NOTES
    Platform  : Qualys EU  (qualysapi.qualys.eu)
    Policy ID : 1661380
    Include   : tag ANY of "All.Workstations"
    Exclude   : tag ANY of "VSI Testing"
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure TLS 1.2 (required by Qualys)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region ── Constants ────────────────────────────────────────────────────────────
$BASE_URL       = 'https://qualysapi.qualys.eu'
$POLICY_ID      = 1661380
$INCLUDE_TAG    = 'All.Workstations'
$EXCLUDE_TAG    = 'VSI Testing'
$PAGE_SIZE      = 1000   # max records per page
#endregion

#region ── Banner ───────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '╔══════════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
Write-Host '║        Qualys Policy Compliance  –  Failure Report          ║' -ForegroundColor Cyan
Write-Host '║                                                              ║' -ForegroundColor Cyan
Write-Host "║  Platform   : qualysapi.qualys.eu                           ║" -ForegroundColor Cyan
Write-Host "║  Policy ID  : $POLICY_ID                                  ║" -ForegroundColor Cyan
Write-Host "║  Include    : Any tag '$INCLUDE_TAG'              ║" -ForegroundColor Cyan
Write-Host "║  Exclude    : Any tag '$EXCLUDE_TAG'                  ║" -ForegroundColor Cyan
Write-Host '╚══════════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
Write-Host ''
#endregion

#region ── Credentials ──────────────────────────────────────────────────────────
$cred     = Get-Credential -Message 'Enter your Qualys EU credentials (username / password)'
$username = $cred.UserName
$password = $cred.GetNetworkCredential().Password

$authToken = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))

# Form-encoded headers (for legacy FO API calls)
$foHeaders = @{
    'Authorization'    = "Basic $authToken"
    'X-Requested-With' = 'PowerShell'
    'Content-Type'     = 'application/x-www-form-urlencoded'
}

# XML/JSON headers (for QPS REST API calls)
$qpsHeaders = @{
    'Authorization'    = "Basic $authToken"
    'X-Requested-With' = 'PowerShell'
    'Content-Type'     = 'text/xml'
}
#endregion

#region ── Helper: Invoke-QualysWebRequest (raw XML) ────────────────────────────
function Invoke-QualysWebRequest {
    param(
        [string]   $Uri,
        [string]   $Method  = 'POST',
        [hashtable]$Headers,
        [string]   $Body    = $null,
        [int]      $Retries = 3
    )
    $attempt = 0
    while ($true) {
        try {
            $splat = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ErrorAction = 'Stop'
            }
            if ($Body) { $splat['Body'] = $Body }
            return Invoke-WebRequest @splat
        }
        catch {
            $attempt++
            if ($attempt -ge $Retries) { throw }
            $wait = [math]::Pow(2, $attempt)
            Write-Warning "Request failed (attempt $attempt/$Retries) – retrying in ${wait}s: $_"
            Start-Sleep -Seconds $wait
        }
    }
}

function Invoke-QualysApi {
    param(
        [string]   $Uri,
        [string]   $Method  = 'POST',
        [hashtable]$Headers,
        [string]   $Body    = $null,
        [int]      $Retries = 3
    )
    $raw = Invoke-QualysWebRequest -Uri $Uri -Method $Method -Headers $Headers -Body $Body -Retries $Retries
    [xml]$xml = $raw.Content
    return $xml
}
#endregion

#region ── Step 1: Resolve tag names → IDs via Asset Management API ─────────────
Write-Host '[1/3] Resolving Qualys tag IDs ...' -ForegroundColor Yellow

function Search-QualysTag {
    param([string]$TagName, [string]$Operator)

    $body = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="name" operator="$Operator">$TagName</Criteria>
  </filters>
  <preferences>
    <limitResults>25</limitResults>
  </preferences>
</ServiceRequest>
"@
    $raw = Invoke-QualysWebRequest -Uri "$BASE_URL/qps/rest/2.0/search/am/tag" `
                                   -Method Post -Headers $qpsHeaders -Body $body
    [xml]$xml = $raw.Content
    Write-Verbose "Tag search ($Operator '$TagName'):`n$($raw.Content)"

    $codeNode = $xml.SelectSingleNode('//ServiceResponse/responseCode')
    if (-not $codeNode -or $codeNode.InnerText -ne 'SUCCESS') {
        $detail = $xml.SelectSingleNode('//responseErrorDetails')
        throw "Tag search failed (operator=$Operator, code=$($codeNode.InnerText)): $($detail.InnerText)"
    }

    $countNode = $xml.SelectSingleNode('//ServiceResponse/count')
    $count     = if ($countNode) { [int]$countNode.InnerText } else { 0 }
    return @{ Count = $count; Xml = $xml }
}

function Get-QualysTagId {
    param([string]$TagName)

    # ── 1. Try exact case-sensitive match first (EQUALS per Qualys docs) ──────
    $result = Search-QualysTag -TagName $TagName -Operator 'EQUALS'

    if ($result.Count -gt 0) {
        $idNode = $result.Xml.SelectSingleNode('//ServiceResponse/data/Tag[1]/id')
        if (-not $idNode) {
            Write-Warning "EQUALS match found but <id> missing. Raw:`n$($result.Xml.OuterXml)"
            throw "Unexpected response structure for tag '$TagName'."
        }
        return [long]$idNode.InnerText
    }

    # ── 2. EQUALS returned 0 – EQUALS is case-sensitive in Qualys.
    #       Fall back to CONTAINS so we can suggest the real name. ─────────────
    Write-Warning "Tag '$TagName' not found with exact match (EQUALS is case-sensitive)."
    Write-Warning "Falling back to CONTAINS search to find the closest match..."

    $fuzzy = Search-QualysTag -TagName $TagName -Operator 'CONTAINS'

    if ($fuzzy.Count -eq 0) {
        throw "Tag '$TagName' not found in Qualys with EQUALS or CONTAINS. Verify the tag exists and your account has scope to see it."
    }

    $tagNodes = $fuzzy.Xml.SelectNodes('//ServiceResponse/data/Tag')
    Write-Warning "Found $($fuzzy.Count) tag(s) containing '$TagName':"
    foreach ($t in $tagNodes) {
        $tid   = $t.SelectSingleNode('id').InnerText
        $tname = $t.SelectSingleNode('name').InnerText
        Write-Warning "    ID=$tid  Name='$tname'"
    }

    if ($fuzzy.Count -eq 1) {
        $idNode   = $fuzzy.Xml.SelectSingleNode('//ServiceResponse/data/Tag[1]/id')
        $nameNode = $fuzzy.Xml.SelectSingleNode('//ServiceResponse/data/Tag[1]/name')
        Write-Warning "Auto-selecting the single match: '$($nameNode.InnerText)' (ID=$($idNode.InnerText))"
        return [long]$idNode.InnerText
    }

    throw "Multiple tags match '$TagName'. Update the script constant with the exact name shown above (case-sensitive)."
}

$includeTagId = Get-QualysTagId -TagName $INCLUDE_TAG
$excludeTagId = Get-QualysTagId -TagName $EXCLUDE_TAG

Write-Host "    '$INCLUDE_TAG'  →  ID $includeTagId" -ForegroundColor Green
Write-Host "    '$EXCLUDE_TAG'        →  ID $excludeTagId" -ForegroundColor Green
#endregion

#region ── Step 2: Fetch compliance posture FAIL records (paginated) ─────────────
Write-Host '[2/3] Fetching compliance posture failures (paginated) ...' -ForegroundColor Yellow

$allFailures = [System.Collections.Generic.List[object]]::new()
$page        = 1
$idMin       = 0           # used for keyset pagination

do {
    Write-Host "    Page $page  (id_min=$idMin) ..." -ForegroundColor DarkGray

    $bodyParts = @(
        "action=list",
        "policy_id=$POLICY_ID",
        "status=FAIL",
        "tag_include_selector=any",
        "tag_include_id=$includeTagId",
        "tag_exclude_selector=any",
        "tag_exclude_id=$excludeTagId",
        "truncation_limit=$PAGE_SIZE"
    )
    if ($idMin -gt 0) { $bodyParts += "id_min=$idMin" }
    $body = $bodyParts -join '&'

    $raw  = Invoke-QualysWebRequest -Uri "$BASE_URL/api/2.0/fo/compliance/posture/info/" `
                                    -Method Post -Headers $foHeaders -Body $body
    [xml]$xml = $raw.Content

    $chunkNodes = $xml.SelectNodes('//POSTURE_INFO')
    if ($chunkNodes.Count -gt 0) {
        foreach ($node in $chunkNodes) { $allFailures.Add($node) }
        Write-Host "      Retrieved $($chunkNodes.Count) record(s)  (running total: $($allFailures.Count))" `
                   -ForegroundColor DarkGray
    }

    # Qualys returns a <WARNING><URL>…</URL></WARNING> element when more pages exist
    $warningUrl = $xml.SelectSingleNode('//WARNING/URL')
    if ($warningUrl -and $warningUrl.InnerText -match 'id_min=(\d+)') {
        $idMin = [long]$Matches[1]
        $page++
    }
    else {
        break   # no further pages
    }
} while ($true)

Write-Host "    Total FAIL records: $($allFailures.Count)" -ForegroundColor Green
#endregion

#region ── Step 3: Format, display and export results ────────────────────────────
Write-Host '[3/3] Formatting results ...' -ForegroundColor Yellow

if ($allFailures.Count -eq 0) {
    Write-Host ''
    Write-Host '  No failures found for the specified criteria.' -ForegroundColor Green
    Write-Host ''
    exit 0
}

function Get-NodeText { param($Node, [string]$XPath) $n = $Node.SelectSingleNode($XPath); if ($n) { $n.InnerText } else { '' } }

$results = $allFailures | ForEach-Object {
    $node = $_   # XmlElement
    [PSCustomObject]@{
        PolicyID    = $POLICY_ID
        ControlID   = Get-NodeText $node 'CONTROL_ID'
        ControlText = (Get-NodeText $node 'CONTROL_STATEMENT') -replace '\s+', ' '
        Status      = Get-NodeText $node 'STATUS'
        IP          = Get-NodeText $node 'HOST_ID/IP'
        Hostname    = Get-NodeText $node 'HOST_ID/HOSTNAME'
        OS          = Get-NodeText $node 'HOST_ID/OS'
        Evidence    = (Get-NodeText $node 'EVIDENCE') -replace '\s+', ' '
        LastEval    = Get-NodeText $node 'LAST_EVALUATED'
    }
}

# ── Console: detail table ──
Write-Host ''
Write-Host "── Policy $POLICY_ID  |  Tag: $INCLUDE_TAG  |  Excl: $EXCLUDE_TAG ──" `
           -ForegroundColor Cyan
$results | Format-Table IP, Hostname, ControlID, ControlText, Status, LastEval -AutoSize

# ── Console: failures-per-host summary ──
Write-Host '── Failures per host ──' -ForegroundColor Cyan
$results |
    Group-Object Hostname |
    Sort-Object Count -Descending |
    Format-Table @{L='Hostname'; E={$_.Name}},
                 @{L='Failures'; E={$_.Count}} -AutoSize

# ── Console: failures-per-control summary ──
Write-Host '── Failures per control ──' -ForegroundColor Cyan
$results |
    Group-Object ControlID |
    Sort-Object Count -Descending |
    Format-Table @{L='ControlID';    E={$_.Name}},
                 @{L='AffectedHosts'; E={$_.Count}},
                 @{L='ControlText';  E={($_.Group[0].ControlText | Select-Object -First 1)}} -AutoSize

# ── Export to CSV ──
$csvPath = ".\Qualys_Policy${POLICY_ID}_Failures_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Full results saved to: $csvPath" -ForegroundColor Green
Write-Host ''
#endregion
