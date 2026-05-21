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
            # In PS 5.1, Invoke-WebRequest throws on 4xx/5xx. Extract the response
            # body so we can see the actual Qualys error message rather than just
            # a generic "The remote server returned an error" message.
            $detail = ''
            $ex = $_.Exception
            if ($ex -is [System.Net.WebException] -and $ex.Response) {
                try {
                    $stream = $ex.Response.GetResponseStream()
                    $reader = [System.IO.StreamReader]::new($stream)
                    $detail = " | Qualys response: $($reader.ReadToEnd())"
                } catch {}
            }
            $attempt++
            if ($attempt -ge $Retries) {
                throw "API call to $Uri failed: $($_)$detail"
            }
            $wait = [math]::Pow(2, $attempt)
            Write-Warning "Request failed (attempt $attempt/$Retries) – retrying in ${wait}s: $_"
            Start-Sleep -Seconds $wait
        }
    }
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

#region ── Step 2: Fetch ALL FAIL records for the policy (no host pre-filter) ───
# Pull every failure first – it's one fast paginated call.
# Tag filtering is done in Step 3 against only the hosts that actually have failures,
# which is far smaller than enumerating all tagged hosts up front.
Write-Host '[2/3] Fetching policy failures ...' -ForegroundColor Yellow

function Get-NodeText { param($Node, [string]$XPath) $n = $Node.SelectSingleNode($XPath); if ($n) { $n.InnerText } else { '' } }

$allPosture = [System.Collections.Generic.List[object]]::new()
$idMin      = 0

do {
    $parts = @("action=list", "policy_id=$POLICY_ID", "status=Failed")
    if ($idMin -gt 0) { $parts += "id_min=$idMin" }

    $raw = Invoke-QualysWebRequest -Uri "$BASE_URL/api/2.0/fo/compliance/posture/info/" `
                                   -Method Post -Headers $foHeaders -Body ($parts -join '&')
    [xml]$xml = $raw.Content

    $batch = $xml.SelectNodes('//POSTURE_INFO')
    foreach ($n in $batch) { $allPosture.Add($n) }
    Write-Host "    +$($batch.Count) records (total: $($allPosture.Count))" -ForegroundColor DarkGray

    $warn = $xml.SelectSingleNode('//WARNING/URL')
    if ($warn -and $warn.InnerText -match 'id_min=(\d+)') { $idMin = [long]$Matches[1] } else { break }
} while ($true)

Write-Host "    Raw failures: $($allPosture.Count)" -ForegroundColor Green

if ($allPosture.Count -eq 0) {
    Write-Host "`n  No failures found for policy $POLICY_ID." -ForegroundColor Green
    exit 0
}
#endregion

#region ── Step 3: Tag-filter against only the IPs present in the failure set ───
# Much faster than enumerating all tagged hosts: we only look up the small set
# of IPs that already have failures, then keep those matching the tag criteria.
Write-Host '[3/3] Applying tag filter and exporting ...' -ForegroundColor Yellow

$uniqueIPs = @($allPosture |
    ForEach-Object { Get-NodeText $_ 'HOST_ID/IP' } |
    Where-Object   { $_ -ne '' } |
    Sort-Object -Unique)

Write-Host "    Unique IPs with failures: $($uniqueIPs.Count)" -ForegroundColor DarkGray

# Batch-fetch host assets for those IPs using IN operator, get their tag IDs
$approvedIPs = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)

$BATCH = 500
for ($i = 0; $i -lt $uniqueIPs.Count; $i += $BATCH) {
    $slice  = $uniqueIPs[$i .. ([Math]::Min($i + $BATCH - 1, $uniqueIPs.Count - 1))]
    $ipsCsv = $slice -join ','

    $body = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="address" operator="IN">$ipsCsv</Criteria>
  </filters>
  <preferences>
    <limitResults>$BATCH</limitResults>
  </preferences>
</ServiceRequest>
"@
    $raw = Invoke-QualysWebRequest -Uri "$BASE_URL/qps/rest/2.0/search/am/hostasset" `
                                   -Method Post -Headers $qpsHeaders -Body $body
    [xml]$hxml = $raw.Content

    foreach ($h in $hxml.SelectNodes('//HostAsset')) {
        $addrNode = $h.SelectSingleNode('address')
        if (-not $addrNode) { continue }
        $ip   = $addrNode.InnerText.Trim()
        $tids = @($h.SelectNodes('.//TagSimple/id') | ForEach-Object { [long]$_.InnerText })
        if (($tids -contains $includeTagId) -and ($tids -notcontains $excludeTagId)) {
            [void]$approvedIPs.Add($ip)
        }
    }
}

Write-Host "    IPs passing tag filter: $($approvedIPs.Count)" -ForegroundColor Green

$filtered = @($allPosture | Where-Object {
    $ip = Get-NodeText $_ 'HOST_ID/IP'
    $ip -ne '' -and $approvedIPs.Contains($ip)
})

Write-Host "    Failures after filter: $($filtered.Count)" -ForegroundColor Green

if ($filtered.Count -eq 0) {
    Write-Host "`n  No failures matched the tag criteria." -ForegroundColor Yellow
    exit 0
}

$results = $filtered | ForEach-Object {
    [PSCustomObject]@{
        PolicyID    = $POLICY_ID
        IP          = Get-NodeText $_ 'HOST_ID/IP'
        Hostname    = Get-NodeText $_ 'HOST_ID/HOSTNAME'
        OS          = Get-NodeText $_ 'HOST_ID/OS'
        ControlID   = Get-NodeText $_ 'CONTROL_ID'
        ControlText = (Get-NodeText $_ 'CONTROL_STATEMENT') -replace '\s+', ' '
        Status      = Get-NodeText $_ 'STATUS'
        Evidence    = (Get-NodeText $_ 'EVIDENCE') -replace '\s+', ' '
        LastEval    = Get-NodeText $_ 'LAST_EVALUATED'
    }
}

$csvPath = ".\Qualys_Policy${POLICY_ID}_Failures_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host "CSV exported: $csvPath  ($($results.Count) rows)" -ForegroundColor Green
Write-Host ''
#endregion
