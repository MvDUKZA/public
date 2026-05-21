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

#region ── Helper: Invoke-QualysApi ─────────────────────────────────────────────
function Invoke-QualysApi {
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
            return Invoke-RestMethod @splat
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
#endregion

#region ── Step 1: Resolve tag names → IDs via Asset Management API ─────────────
Write-Host '[1/3] Resolving Qualys tag IDs ...' -ForegroundColor Yellow

function Get-QualysTagId {
    param([string]$TagName)

    $body = @"
<?xml version="1.0" encoding="UTF-8"?>
<ServiceRequest>
  <filters>
    <Criteria field="name" operator="EQUALS">$TagName</Criteria>
  </filters>
</ServiceRequest>
"@
    $r = Invoke-QualysApi -Uri "$BASE_URL/qps/rest/2.0/search/am/tag" `
                          -Method Post -Headers $qpsHeaders -Body $body

    if ($r.ServiceResponse.responseCode -ne 'SUCCESS') {
        throw "Tag search for '$TagName' failed: $($r.ServiceResponse.responseErrorDetails)"
    }

    # May return a single Tag object or an array; grab the first id
    $id = @($r.ServiceResponse.data.Tag)[0].id
    if (-not $id) {
        throw "Tag '$TagName' was not found in Qualys. Check the exact tag name and your subscription."
    }
    return [long]$id
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

    $resp  = Invoke-QualysApi -Uri "$BASE_URL/api/2.0/fo/compliance/posture/info/" `
                               -Method Post -Headers $foHeaders -Body $body
    $chunk = $resp.COMPLIANCE_POSTURE_INFO_OUTPUT.RESPONSE.POSTURE_INFO_LIST.POSTURE_INFO

    if ($chunk) {
        $allFailures.AddRange([object[]]$chunk)
        Write-Host "      Retrieved $($chunk.Count) record(s)  (running total: $($allFailures.Count))" `
                   -ForegroundColor DarkGray
    }

    # Qualys returns a <WARNING><URL>…</URL></WARNING> element when more pages exist
    $warning = $resp.COMPLIANCE_POSTURE_INFO_OUTPUT.RESPONSE.WARNING
    if ($warning -and $warning.URL -match 'id_min=(\d+)') {
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

$results = $allFailures | ForEach-Object {
    [PSCustomObject]@{
        PolicyID    = $POLICY_ID
        ControlID   = $_.CONTROL_ID
        ControlText = ($_.CONTROL_STATEMENT -replace '\s+', ' ').Trim()
        Status      = $_.STATUS
        IP          = $_.HOST_ID.IP
        Hostname    = $_.HOST_ID.HOSTNAME
        OS          = $_.HOST_ID.OS
        Evidence    = ($_.EVIDENCE -replace '\s+', ' ').Trim()
        LastEval    = $_.LAST_EVALUATED
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
