#Requires -Version 5.1
<#
.SYNOPSIS
    Exports Qualys Policy Compliance failures to CSV.

.DESCRIPTION
    Uses the Qualys FO API (/api/2.0/fo/compliance/posture/info/) to fetch posture
    records for Policy ID 1661380 with server-side tag filtering (include
    All.Workstations, exclude VSI Testing). STATUS is filtered to Failed in
    PowerShell. Host IP/DNS and control statements are resolved from the GLOSSARY
    section returned alongside each page of INFO records.

.NOTES
    Platform   : Qualys EU  (qualysapi.qualys.eu)
    Policy ID  : 1661380
    Include tag: All.Workstations  (resolved to tag ID at runtime)
    Exclude tag: VSI Testing       (resolved to tag ID at runtime)

    FO API response structure (details=All):
      POSTURE_INFO_LIST_OUTPUT/RESPONSE/INFO_LIST/INFO  - posture records
        Fields: ID, HOST_ID, CONTROL_ID, STATUS, EVALUATION_DATE, EVIDENCE/BOOLEAN_EXPR
      POSTURE_INFO_LIST_OUTPUT/RESPONSE/GLOSSARY
        HOST_LIST/HOST    - IP, DNS, OS keyed by HOST_ID (<ID>)
        CONTROL_LIST/CONTROL - STATEMENT keyed by CONTROL_ID (<ID>)
      POSTURE_INFO_LIST_OUTPUT/RESPONSE/WARNING/URL     - next-page cursor (id_min=)
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
Write-Host '║   Qualys Policy Compliance  –  Failure Report (FO API)      ║' -ForegroundColor Cyan
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

$foHeaders = @{
    'Authorization'    = "Basic $authToken"
    'X-Requested-With' = 'PowerShell'
    'Content-Type'     = 'application/x-www-form-urlencoded'
}
$qpsHeaders = @{
    'Authorization'    = "Basic $authToken"
    'X-Requested-With' = 'PowerShell'
    'Content-Type'     = 'text/xml'
}
#endregion

#region ── Helper: HTTP with retries and error body extraction ───────────────────
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

#region ── Helper: safe XPath text extractor ────────────────────────────────────
function Get-NodeText {
    param($Node, [string]$XPath)
    $n = $Node.SelectSingleNode($XPath)
    if ($n) { $n.InnerText.Trim() } else { '' }
}
#endregion

#region ── Step 1: Resolve tag IDs ──────────────────────────────────────────────
# Tag names with spaces can't be safely URL-encoded in form bodies;
# using IDs in tag_set_include/exclude avoids that issue entirely.
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
    $raw  = Invoke-QualysWebRequest -Uri "$BASE_URL/qps/rest/2.0/search/am/tag" `
                                    -Method Post -Headers $qpsHeaders -Body $body
    [xml]$xml = $raw.Content
    $code = $xml.SelectSingleNode('//ServiceResponse/responseCode')
    if (-not $code -or $code.InnerText -ne 'SUCCESS') {
        throw "Tag search failed: $(Get-NodeText $xml.DocumentElement '//responseErrorDetails')"
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
    Write-Warning "Tag '$TagName' not found by EQUALS. Trying CONTAINS..."
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

#region ── Step 2: Fetch posture data via FO API ────────────────────────────────
# Fetch ALL statuses (status=Failed is ignored by this platform — filter in PS).
# The GLOSSARY section of each page maps HOST_ID→IP/DNS/OS and CONTROL_ID→STATEMENT.
# Tag filtering is done in Step 3 via the AM API (server-side tag params are silently
# ignored on the Qualys EU platform for this policy).
Write-Host '[2/4] Fetching posture data via FO API ...' -ForegroundColor Yellow

$foPostureUrl  = "$BASE_URL/api/2.0/fo/compliance/posture/info/"
$allInfoNodes  = [System.Collections.Generic.List[object]]::new()
$hostGlossary  = @{}   # HOST_ID  (string) -> @{ IP; DNS; OS }
$ctrlGlossary  = @{}   # CONTROL_ID (string) -> statement text
$idMin         = 0
$pageNum       = 1

while ($true) {
    $bodyParts = @(
        "action=list",
        "policy_id=$POLICY_ID",
        "details=All",
        "truncation_limit=5000"
    )
    if ($idMin -gt 0) { $bodyParts += "id_min=$idMin" }
    $body = $bodyParts -join '&'

    Write-Host "    Page $pageNum (id_min=$idMin) ..." -ForegroundColor DarkGray
    $raw = Invoke-QualysWebRequest -Uri $foPostureUrl -Method Post -Headers $foHeaders -Body $body
    [xml]$xml = $raw.Content

    # Fail fast on API-level errors
    $errNode = $xml.SelectSingleNode('//RESPONSE/ERROR')
    if ($errNode) {
        throw "FO API error $(Get-NodeText $errNode 'NUMBER'): $(Get-NodeText $errNode 'TEXT')"
    }

    # Accumulate GLOSSARY: HOST_ID -> IP / DNS / OS
    foreach ($h in $xml.SelectNodes('//GLOSSARY/HOST_LIST/HOST')) {
        $hid = Get-NodeText $h 'ID'
        if ($hid -and -not $hostGlossary.ContainsKey($hid)) {
            $hostGlossary[$hid] = @{
                IP  = Get-NodeText $h 'IP'
                DNS = Get-NodeText $h 'DNS'
                OS  = Get-NodeText $h 'OS'
            }
        }
    }

    # Accumulate GLOSSARY: CONTROL_ID -> STATEMENT
    foreach ($c in $xml.SelectNodes('//GLOSSARY/CONTROL_LIST/CONTROL')) {
        $cid = Get-NodeText $c 'ID'
        if ($cid -and -not $ctrlGlossary.ContainsKey($cid)) {
            $ctrlGlossary[$cid] = Get-NodeText $c 'STATEMENT'
        }
    }

    # Accumulate INFO nodes
    $infoNodes = $xml.SelectNodes('//INFO_LIST/INFO')
    if ($infoNodes -and $infoNodes.Count -gt 0) {
        foreach ($node in $infoNodes) { $allInfoNodes.Add($node) }
        Write-Host "    +$($infoNodes.Count) records (running total: $($allInfoNodes.Count))" -ForegroundColor DarkGray
    }

    # Pagination via WARNING/URL id_min cursor
    $warnUrl = $xml.SelectSingleNode('//WARNING/URL')
    if ($warnUrl -and ($warnUrl.InnerText -match 'id_min=(\d+)')) {
        $idMin = [long]$Matches[1]
        $pageNum++
    } else {
        break
    }
}

Write-Host "    Total records fetched: $($allInfoNodes.Count)  |  Hosts in glossary: $($hostGlossary.Count)  |  Controls: $($ctrlGlossary.Count)" -ForegroundColor Green

# Filter STATUS = Failed in PowerShell (API-side status=Failed returns 0 on this platform)
$failedNodes = @($allInfoNodes | Where-Object {
    (Get-NodeText $_ 'STATUS') -in @('Failed', 'FAILED', 'Fail', 'FAIL')
})
Write-Host "    Failed records: $($failedNodes.Count)" -ForegroundColor Green

if ($failedNodes.Count -eq 0) {
    Write-Host "`n  No FAILED records found. Verify the policy has been scanned and tags are assigned." -ForegroundColor Yellow
    exit 0
}
#endregion

#region ── Step 3: Tag-filter failed records via AM API ─────────────────────────
# Get unique IPs for failed hosts (from GLOSSARY), then batch-check their tags
# via the AM hostasset API. Only approve hosts that have the include tag AND
# do NOT have the exclude tag.
Write-Host '[3/4] Applying tag filter via AM API ...' -ForegroundColor Yellow

$uniqueIPs = @(
    $failedNodes |
    ForEach-Object { $hid = Get-NodeText $_ 'HOST_ID'; if ($hostGlossary.ContainsKey($hid)) { $hostGlossary[$hid].IP } } |
    Where-Object { $_ -ne '' } |
    Sort-Object -Unique
)
Write-Host "    Unique IPs in failed records: $($uniqueIPs.Count)" -ForegroundColor DarkGray

$approvedIPs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$BATCH       = 500

for ($i = 0; $i -lt $uniqueIPs.Count; $i += $BATCH) {
    $slice  = $uniqueIPs[$i .. ([Math]::Min($i + $BATCH - 1, $uniqueIPs.Count - 1))]
    $ipsCsv = $slice -join ','
    $bNum   = [math]::Floor($i / $BATCH) + 1
    Write-Host "    Tag-checking IP batch $bNum ($($slice.Count) IPs) ..." -ForegroundColor DarkGray

    $amBody = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="address" operator="IN">$ipsCsv</Criteria>
  </filters>
  <preferences><limitResults>$BATCH</limitResults></preferences>
</ServiceRequest>
"@
    $raw  = Invoke-QualysWebRequest -Uri "$BASE_URL/qps/rest/2.0/search/am/hostasset" `
                                    -Method Post -Headers $qpsHeaders -Body $amBody
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

$filteredNodes = @($failedNodes | Where-Object {
    $hid = Get-NodeText $_ 'HOST_ID'
    $ip  = if ($hostGlossary.ContainsKey($hid)) { $hostGlossary[$hid].IP } else { '' }
    $ip -ne '' -and $approvedIPs.Contains($ip)
})
Write-Host "    Failed records after tag filter: $($filteredNodes.Count)" -ForegroundColor Green

if ($filteredNodes.Count -eq 0) {
    Write-Host "`n  No records matched the tag criteria (include '$INCLUDE_TAG', exclude '$EXCLUDE_TAG')." -ForegroundColor Yellow
    exit 0
}
#endregion

#region ── Step 4: Build output objects and export CSV ──────────────────────────
Write-Host '[4/4] Building report and exporting CSV ...' -ForegroundColor Yellow

$results = $filteredNodes | ForEach-Object {
    $hid  = Get-NodeText $_ 'HOST_ID'
    $cid  = Get-NodeText $_ 'CONTROL_ID'
    $h    = if ($hostGlossary.ContainsKey($hid))  { $hostGlossary[$hid]  } else { @{ IP=''; DNS=''; OS='' } }
    $stmt = if ($ctrlGlossary.ContainsKey($cid))  { $ctrlGlossary[$cid]  } else { '' }

    [PSCustomObject]@{
        PolicyID         = $POLICY_ID
        IP               = $h.IP
        Hostname         = $h.DNS
        OS               = $h.OS
        ControlID        = $cid
        ControlStatement = $stmt -replace '\s+', ' '
        Status           = Get-NodeText $_ 'STATUS'
        EvaluationDate   = Get-NodeText $_ 'EVALUATION_DATE'
        Evidence         = (Get-NodeText $_ 'EVIDENCE/BOOLEAN_EXPR') -replace '\s+', ' '
    }
}

$csvPath = ".\Qualys_Policy${POLICY_ID}_Failures_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host "CSV exported: $csvPath  ($($results.Count) rows)" -ForegroundColor Green
Write-Host ''
#endregion
