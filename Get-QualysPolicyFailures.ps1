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

#region ── Step 2: Resolve host IDs by tag via Asset Management API ─────────────
# The PC posture API does not accept tag_include/exclude_selector parameters.
# We resolve host IDs ourselves: include-tag hosts minus exclude-tag hosts.
Write-Host '[2/4] Resolving host IDs by tag ...' -ForegroundColor Yellow

function Get-HostIdsByTagId {
    param([long]$TagId, [string]$Label)

    $ids         = [System.Collections.Generic.List[long]]::new()
    $hasMore     = $true
    $startFromId = 0
    $page        = 1

    while ($hasMore) {
        # startFromId is exclusive – the record with that id is NOT included,
        # so passing the last returned id advances the window correctly.
        $startPref = if ($startFromId -gt 0) { "<startFromId>$startFromId</startFromId>" } else { '' }

        $body = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="tagId" operator="EQUALS">$TagId</Criteria>
  </filters>
  <preferences>
    <limitResults>1000</limitResults>
    $startPref
  </preferences>
</ServiceRequest>
"@
        Write-Host "      '$Label' page $page (startFromId=$startFromId) ..." -ForegroundColor DarkGray

        $raw = Invoke-QualysWebRequest -Uri "$BASE_URL/qps/rest/2.0/search/am/hostasset" `
                                       -Method Post -Headers $qpsHeaders -Body $body
        [xml]$xml = $raw.Content

        $codeNode = $xml.SelectSingleNode('//ServiceResponse/responseCode')
        if (-not $codeNode -or $codeNode.InnerText -ne 'SUCCESS') {
            $detail = $xml.SelectSingleNode('//responseErrorDetails')
            throw "Host search for tag '$Label' (ID=$TagId) failed: $($detail.InnerText)"
        }

        $idNodes = $xml.SelectNodes('//ServiceResponse/data/HostAsset/id')
        foreach ($n in $idNodes) { $ids.Add([long]$n.InnerText) }
        Write-Host "        +$($idNodes.Count) hosts (running total: $($ids.Count))" -ForegroundColor DarkGray

        $hasMoreNode = $xml.SelectSingleNode('//ServiceResponse/hasMoreRecords')
        $hasMore     = ($hasMoreNode -and $hasMoreNode.InnerText -eq 'true')

        if ($hasMore -and $idNodes.Count -gt 0) {
            $startFromId = [long]$idNodes.Item($idNodes.Count - 1).InnerText
            $page++
        } else {
            $hasMore = $false
        }
    }

    return $ids
}

$includeHostIds = Get-HostIdsByTagId -TagId $includeTagId -Label $INCLUDE_TAG
Write-Host "    Hosts with '$INCLUDE_TAG': $($includeHostIds.Count)" -ForegroundColor Green

$excludeHostIds = Get-HostIdsByTagId -TagId $excludeTagId -Label $EXCLUDE_TAG
Write-Host "    Hosts with '$EXCLUDE_TAG':  $($excludeHostIds.Count)" -ForegroundColor Green

$excludeSet    = [System.Collections.Generic.HashSet[long]]::new($excludeHostIds)
$targetHostIds = [long[]]($includeHostIds | Where-Object { -not $excludeSet.Contains($_) })
Write-Host "    Target hosts after exclusion: $($targetHostIds.Count)" -ForegroundColor Green

if ($targetHostIds.Count -eq 0) {
    Write-Host "`n  No hosts match the tag criteria (include minus exclude)." -ForegroundColor Yellow
    exit 0
}
#endregion

#region ── Step 3: Fetch compliance posture FAIL records (chunked by host ID) ────
# The posture API accepts host_id as a comma-separated list.
# We chunk to avoid excessively long request bodies.
Write-Host '[3/4] Fetching compliance posture failures ...' -ForegroundColor Yellow

$allFailures = [System.Collections.Generic.List[object]]::new()
$CHUNK_SIZE  = 300
$totalChunks = [math]::Ceiling($targetHostIds.Count / $CHUNK_SIZE)

for ($offset = 0; $offset -lt $targetHostIds.Count; $offset += $CHUNK_SIZE) {
    $end        = [Math]::Min($offset + $CHUNK_SIZE - 1, $targetHostIds.Count - 1)
    $hostChunk  = ($targetHostIds[$offset..$end]) -join ','
    $chunkNum   = [math]::Floor($offset / $CHUNK_SIZE) + 1
    Write-Host "    Chunk $chunkNum/$totalChunks (hosts $($offset+1)–$($end+1)) ..." -ForegroundColor DarkGray

    $idMin = 0
    do {
        $bodyParts = @(
            "action=list",
            "policy_id=$POLICY_ID",
            "status=FAIL",
            "host_id=$hostChunk",
            "truncation_limit=$PAGE_SIZE"
        )
        if ($idMin -gt 0) { $bodyParts += "id_min=$idMin" }

        $raw = Invoke-QualysWebRequest -Uri "$BASE_URL/api/2.0/fo/compliance/posture/info/" `
                                       -Method Post -Headers $foHeaders -Body ($bodyParts -join '&')
        [xml]$xml = $raw.Content

        $nodes = $xml.SelectNodes('//POSTURE_INFO')
        foreach ($node in $nodes) { $allFailures.Add($node) }
        if ($nodes.Count -gt 0) {
            Write-Host "      +$($nodes.Count) records (total: $($allFailures.Count))" -ForegroundColor DarkGray
        }

        $warningUrl = $xml.SelectSingleNode('//WARNING/URL')
        if ($warningUrl -and $warningUrl.InnerText -match 'id_min=(\d+)') {
            $idMin = [long]$Matches[1]
        } else {
            break
        }
    } while ($true)
}

Write-Host "    Total FAIL records: $($allFailures.Count)" -ForegroundColor Green
#endregion

#region ── Step 4: Format, display and export results ────────────────────────────
Write-Host '[4/4] Formatting results ...' -ForegroundColor Yellow

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
Write-Host "── Policy $POLICY_ID  |  Include: $INCLUDE_TAG  |  Exclude: $EXCLUDE_TAG ──" `
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
