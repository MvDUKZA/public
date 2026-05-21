#Requires -Version 5.1
<#
.SYNOPSIS
    Exports Qualys Policy Compliance failures to CSV using the PCRS API.

.DESCRIPTION
    Uses the Qualys PCRS (Policy Compliance Reporting Streaming) API —
    the same engine that powers the "(SS) Policy Compliance - Workstation Failures"
    dashboard report — to fetch FAILED posture records for Policy ID 1661380,
    filtered to hosts tagged "All.Workstations" but NOT "VSI Testing".

.NOTES
    Platform   : Qualys EU  (qualysapi.qualys.eu)
    Policy ID  : 1661380
    Include tag: All.Workstations
    Exclude tag: VSI Testing
    PCRS docs  : /pcrs/1.0/posture/hostids  +  /pcrs/1.0/posture/postureInfo
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region ── Constants ────────────────────────────────────────────────────────────
$BASE_URL    = 'https://qualysapi.qualys.eu'
$POLICY_ID   = 1661380
$INCLUDE_TAG = 'All.Workstations'
$EXCLUDE_TAG = 'VSI Testing'
#endregion

#region ── Banner ───────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '╔══════════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
Write-Host '║   Qualys Policy Compliance  –  Failure Report (PCRS API)    ║' -ForegroundColor Cyan
Write-Host "║   Policy  : $POLICY_ID                                   ║" -ForegroundColor Cyan
Write-Host "║   Include : $INCLUDE_TAG                          ║" -ForegroundColor Cyan
Write-Host "║   Exclude : $EXCLUDE_TAG                              ║" -ForegroundColor Cyan
Write-Host '╚══════════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
Write-Host ''
#endregion

#region ── Credentials ──────────────────────────────────────────────────────────
$cred      = Get-Credential -Message 'Enter your Qualys EU credentials'
$username  = $cred.UserName
$password  = $cred.GetNetworkCredential().Password
$authToken = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))

$pcrsHeaders = @{
    'Authorization'    = "Basic $authToken"
    'X-Requested-With' = 'PowerShell'
    'Content-Type'     = 'application/json'
    'Accept'           = 'application/json'
}
$qpsHeaders = @{
    'Authorization'    = "Basic $authToken"
    'X-Requested-With' = 'PowerShell'
    'Content-Type'     = 'text/xml'
}
#endregion

#region ── Helper: HTTP with error body extraction ──────────────────────────────
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
            $splat = @{ Uri = $Uri; Method = $Method; Headers = $Headers; ErrorAction = 'Stop' }
            if ($Body) { $splat['Body'] = $Body }
            return Invoke-WebRequest @splat
        }
        catch {
            $detail = ''
            $ex = $_.Exception
            if ($ex -is [System.Net.WebException] -and $ex.Response) {
                try {
                    $stream = $ex.Response.GetResponseStream()
                    $reader = [System.IO.StreamReader]::new($stream)
                    $detail = " | Response: $($reader.ReadToEnd())"
                } catch {}
            }
            $attempt++
            if ($attempt -ge $Retries) { throw "API call to $Uri failed: $($_)$detail" }
            $wait = [math]::Pow(2, $attempt)
            Write-Warning "Retry $attempt/$Retries in ${wait}s: $_"
            Start-Sleep -Seconds $wait
        }
    }
}
#endregion

#region ── Step 1: Resolve tag IDs (needed for AM tag-filter in Step 4) ─────────
Write-Host '[1/4] Resolving tag IDs ...' -ForegroundColor Yellow

function Search-QualysTag {
    param([string]$TagName, [string]$Operator)
    $body = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="name" operator="$Operator">$TagName</Criteria>
  </filters>
  <preferences><limitResults>25</limitResults></preferences>
</ServiceRequest>
"@
    $raw   = Invoke-QualysWebRequest -Uri "$BASE_URL/qps/rest/2.0/search/am/tag" -Method Post -Headers $qpsHeaders -Body $body
    [xml]$xml = $raw.Content
    $code  = $xml.SelectSingleNode('//ServiceResponse/responseCode')
    if (-not $code -or $code.InnerText -ne 'SUCCESS') {
        throw "Tag search failed: $($xml.SelectSingleNode('//responseErrorDetails').InnerText)"
    }
    $count = $xml.SelectSingleNode('//ServiceResponse/count')
    return @{ Count = if ($count) { [int]$count.InnerText } else { 0 }; Xml = $xml }
}

function Get-QualysTagId {
    param([string]$TagName)
    $r = Search-QualysTag -TagName $TagName -Operator 'EQUALS'
    if ($r.Count -gt 0) {
        return [long]$r.Xml.SelectSingleNode('//ServiceResponse/data/Tag[1]/id').InnerText
    }
    Write-Warning "Tag '$TagName' not found by EQUALS (case-sensitive). Trying CONTAINS..."
    $f = Search-QualysTag -TagName $TagName -Operator 'CONTAINS'
    if ($f.Count -eq 0) { throw "Tag '$TagName' not found." }
    foreach ($t in $f.Xml.SelectNodes('//ServiceResponse/data/Tag')) {
        Write-Warning "  Found: ID=$($t.SelectSingleNode('id').InnerText)  Name='$($t.SelectSingleNode('name').InnerText)'"
    }
    if ($f.Count -eq 1) {
        $n = $f.Xml.SelectSingleNode('//ServiceResponse/data/Tag[1]')
        Write-Warning "Auto-selecting '$($n.SelectSingleNode('name').InnerText)' (ID=$($n.SelectSingleNode('id').InnerText))"
        return [long]$n.SelectSingleNode('id').InnerText
    }
    throw "Multiple tags match '$TagName'. Set the exact name in the constants above."
}

$includeTagId = Get-QualysTagId -TagName $INCLUDE_TAG
$excludeTagId = Get-QualysTagId -TagName $EXCLUDE_TAG
Write-Host "    '$INCLUDE_TAG'  → ID $includeTagId" -ForegroundColor Green
Write-Host "    '$EXCLUDE_TAG'  → ID $excludeTagId" -ForegroundColor Green
#endregion

#region ── Step 2: Resolve host IDs for the policy via PCRS ─────────────────────
# GET /pcrs/1.0/posture/hostids returns the QG host IDs scanned by this policy.
Write-Host '[2/4] Getting policy host IDs via PCRS ...' -ForegroundColor Yellow

$policyHostIds = [System.Collections.Generic.List[long]]::new()
$pgNum    = 1
$pgMore   = $true

while ($pgMore) {
    $url = "$BASE_URL/pcrs/1.0/posture/hostids?policyId=$POLICY_ID&pageSize=5000&pageNumber=$pgNum"
    Write-Host "    Page $pgNum ..." -ForegroundColor DarkGray
    $raw  = Invoke-QualysWebRequest -Uri $url -Method Get -Headers $pcrsHeaders

    if ($pgNum -eq 1) {
        Write-Host "    [DIAG] hostids response (first 400 chars): $($raw.Content.Substring(0,[Math]::Min(400,$raw.Content.Length)))" -ForegroundColor DarkGray
    }

    $json = $raw.Content | ConvertFrom-Json
    $ids  = $json.hostIds
    if ($ids -and $ids.Count -gt 0) {
        foreach ($id in $ids) { $policyHostIds.Add([long]$id) }
        Write-Host "    +$($ids.Count) IDs (total: $($policyHostIds.Count))" -ForegroundColor DarkGray
    }
    $pgMore = ($json.hasMoreRecords -eq $true) -and ($ids.Count -gt 0)
    $pgNum++
}

Write-Host "    Total host IDs: $($policyHostIds.Count)" -ForegroundColor Green

if ($policyHostIds.Count -eq 0) {
    Write-Host "`n  No hosts found for policy $POLICY_ID. Check policy ID and account scope." -ForegroundColor Yellow
    exit 0
}
#endregion

#region ── Step 3: Fetch posture info via PCRS ───────────────────────────────────
Write-Host '[3/4] Fetching posture data via PCRS ...' -ForegroundColor Yellow

# Try known PCRS postureInfo endpoint variants in order
$postureInfoCandidates = @(
    "$BASE_URL/pcrs/1.0/posture/postureInfo",
    "$BASE_URL/pcrs/1.0/posture/postureinfo",
    "$BASE_URL/pcrs/2.0/posture/postureInfo",
    "$BASE_URL/pcrs/1.0/postureInfo"
)

function Invoke-PostureInfo {
    param([string]$Url, [string]$BodyJson)
    try {
        return Invoke-QualysWebRequest -Uri $Url -Method Post -Headers $pcrsHeaders -Body $BodyJson -Retries 1
    }
    catch {
        if ($_ -match '404') { return $null }
        throw
    }
}

# Probe for the working URL with a minimal single-host request
$postureInfoUrl = $null
Write-Host "    Probing postureInfo endpoint ..." -ForegroundColor DarkGray
$probeBody = [ordered]@{
    policyId         = [string]$POLICY_ID
    hostIds          = @([string]$policyHostIds[0])
    pageNumber       = 1
    pageSize         = 10
    evidenceRequired = 0
} | ConvertTo-Json -Depth 5

foreach ($candidate in $postureInfoCandidates) {
    Write-Host "    Trying $candidate ..." -ForegroundColor DarkGray
    $probeResult = Invoke-PostureInfo -Url $candidate -BodyJson $probeBody
    if ($null -ne $probeResult) {
        $postureInfoUrl = $candidate
        Write-Host "    Found working endpoint: $postureInfoUrl" -ForegroundColor Green
        Write-Host "    [DIAG] probe response: $($probeResult.Content.Substring(0,[Math]::Min(500,$probeResult.Content.Length)))" -ForegroundColor DarkGray
        break
    }
}

if (-not $postureInfoUrl) {
    throw "Could not find a working postureInfo endpoint. Tried:`n$($postureInfoCandidates -join "`n")"
}

$allRecords  = [System.Collections.Generic.List[object]]::new()
$HCHUNK      = 500   # host IDs per call; keep small to stay within API limits
$totalChunks = [math]::Ceiling($policyHostIds.Count / $HCHUNK)

for ($ci = 0; $ci -lt $policyHostIds.Count; $ci += $HCHUNK) {
    $end   = [Math]::Min($ci + $HCHUNK - 1, $policyHostIds.Count - 1)
    # hostIds must be strings, not integers
    $chunk = @($policyHostIds[$ci..$end] | ForEach-Object { [string]$_ })
    $cNum  = [math]::Floor($ci / $HCHUNK) + 1
    Write-Host "    Chunk $cNum/$totalChunks ($($chunk.Count) hosts) ..." -ForegroundColor DarkGray

    $pgNum2  = 1
    $pgMore2 = $true

    while ($pgMore2) {
        $bodyJson = [ordered]@{
            policyId         = [string]$POLICY_ID
            hostIds          = $chunk
            pageNumber       = $pgNum2
            pageSize         = 100
            evidenceRequired = 1
        } | ConvertTo-Json -Depth 5

        $raw  = Invoke-QualysWebRequest -Uri $postureInfoUrl -Method Post `
                                         -Headers $pcrsHeaders -Body $bodyJson

        $json = $raw.Content | ConvertFrom-Json

        $recs = if     ($null -ne $json.postureInfoList) { $json.postureInfoList }
                elseif ($null -ne $json.data)            { $json.data }
                elseif ($null -ne $json.results)         { $json.results }
                else                                     { @() }

        foreach ($r in $recs) { $allRecords.Add($r) }
        if ($recs.Count -gt 0) {
            Write-Host "      +$($recs.Count) (total: $($allRecords.Count))" -ForegroundColor DarkGray
        }

        $pgMore2 = ($json.hasMoreRecords -eq $true) -and ($recs.Count -gt 0)
        $pgNum2++
    }
}

Write-Host "    Total records (all statuses): $($allRecords.Count)" -ForegroundColor Green

if ($allRecords.Count -eq 0) {
    Write-Warning "PCRS returned 0 records. Check the [DIAG] lines above for the raw response."
    exit 0
}

# Filter FAILED — PCRS uses 'FAILED' (all caps)
$failedRecs = @($allRecords | Where-Object {
    $s = if ($null -ne $_.status) { [string]$_.status }
         elseif ($null -ne $_.postureStatus) { [string]$_.postureStatus }
         else { '' }
    $s.ToUpper() -in @('FAILED','FAIL')
})
Write-Host "    Failed records: $($failedRecs.Count)" -ForegroundColor Green
#endregion

#region ── Step 4: Tag-filter via AM API, then export ───────────────────────────
Write-Host '[4/4] Applying tag filter and exporting ...' -ForegroundColor Yellow

function Get-RecordIp {
    param($rec)
    foreach ($f in @('hostIp','ip','ipAddress','host_ip','ipaddr')) {
        $v = $rec.$f
        if ($v) { return [string]$v }
    }
    return ''
}

$uniqueIPs = @($failedRecs | ForEach-Object { Get-RecordIp $_ } |
    Where-Object { $_ -ne '' } | Sort-Object -Unique)

Write-Host "    Unique IPs in failed records: $($uniqueIPs.Count)" -ForegroundColor DarkGray

$approvedIPs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$BATCH       = 500

for ($i = 0; $i -lt $uniqueIPs.Count; $i += $BATCH) {
    $slice  = $uniqueIPs[$i .. ([Math]::Min($i + $BATCH - 1, $uniqueIPs.Count - 1))]
    $ipsCsv = $slice -join ','

    $body = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="address" operator="IN">$ipsCsv</Criteria>
  </filters>
  <preferences><limitResults>$BATCH</limitResults></preferences>
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

$filtered = @($failedRecs | Where-Object { $approvedIPs.Contains((Get-RecordIp $_)) })
Write-Host "    Records after tag filter: $($filtered.Count)" -ForegroundColor Green

if ($filtered.Count -eq 0) {
    Write-Host "`n  No records matched the tag criteria." -ForegroundColor Yellow
    exit 0
}

$results = $filtered | ForEach-Object {
    $r = $_
    [PSCustomObject]@{
        PolicyID         = $POLICY_ID
        IP               = Get-RecordIp $r
        Hostname         = if ($r.dns)              { $r.dns }
                           elseif ($r.hostname)     { $r.hostname }
                           elseif ($r.hostName)     { $r.hostName }
                           else                     { '' }
        OS               = if ($r.osCpe)            { $r.osCpe }
                           elseif ($r.os)           { $r.os }
                           else                     { '' }
        ControlID        = if ($r.controlId)        { $r.controlId }        else { '' }
        ControlStatement = if ($r.controlStatement) { ([string]$r.controlStatement) -replace '\s+',' ' } else { '' }
        Status           = if ($r.status)           { $r.status }           else { '' }
        EvaluationDate   = if ($r.lastEvaluatedDate){ $r.lastEvaluatedDate }
                           elseif ($r.evaluationDate){ $r.evaluationDate }
                           else                     { '' }
        Evidence         = if ($r.evidence)         { ([string]$r.evidence) -replace '\s+',' ' } else { '' }
    }
}

$csvPath = ".\Qualys_Policy${POLICY_ID}_Failures_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host "CSV exported: $csvPath  ($($results.Count) rows)" -ForegroundColor Green
Write-Host ''
#endregion
