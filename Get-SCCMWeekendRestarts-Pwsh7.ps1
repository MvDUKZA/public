#Requires -Version 7.0
<#
.SYNOPSIS
    (PowerShell 7+) Enumerates SCCM device collections, lets you pick one, then
    reports how many times each machine in that collection was restarted on a
    Saturday or Sunday.

.DESCRIPTION
    PowerShell 7 edition of Get-SCCMWeekendRestarts.ps1. Uses
    ForEach-Object -Parallel (-ThrottleLimit) instead of a manual runspace pool,
    and a picker that degrades gracefully across PS7 hosts.

    It targets ONE specific weekend (a single Saturday + Sunday). With no date
    supplied it defaults to the *previous* weekend relative to today.

    Workflow:
      1. Connects to the SCCM site server WMI namespace (root\SMS\site_<code>).
      2. Lists device collections in a picker (single-select):
           Out-GridView  →  Out-ConsoleGridView  →  numbered console menu.
      3. Resolves the collection's members (SMS_FullCollectionMembership).
      4. For every member it opens a CIM session (WSMan, falling back to DCOM)
         and reads the System event log for restart events that fall on the
         target Saturday or Sunday, bucketing the count by day.

    Restart detection (System log):
      EventID 1074 = a process initiated a shutdown/restart (the classic "who
                     restarted this box" event – includes patch/planned reboots).
      EventID 6005 = the Event Log service started, i.e. the machine booted.

    By default the script counts EventID 1074 entries whose message indicates a
    *restart* (not a bare power-off). Use -CountBoots to instead count every
    system boot (EventID 6005), which captures cold boots and crashes too.

    Two CSVs are written:
      *_Summary.csv  – one row per machine: Saturday / Sunday / weekend totals.
      *_Detail.csv   – one row per individual weekend restart event.

.PARAMETER SiteServer
    SCCM Management Point / Site Server FQDN. Prompted if omitted.

.PARAMETER SiteCode
    SCCM Site Code (e.g. PS1). Prompted if omitted.

.PARAMETER CollectionName
    Device Collection name. Selected via the picker if omitted.

.PARAMETER WeekendOf
    Any date that falls in the weekend you want to check; the script resolves
    that calendar week's Saturday and Sunday. If omitted, defaults to the
    previous weekend relative to today.

.PARAMETER CountBoots
    Count every system boot (EventID 6005) instead of explicit restarts
    (EventID 1074). Captures cold boots / crash recoveries as well.

.PARAMETER OutputPath
    Folder for the CSV output. Defaults to the current directory.

.PARAMETER ThrottleLimit
    Maximum machines queried in parallel. Default: 20.

.EXAMPLE
    .\Get-SCCMWeekendRestarts-Pwsh7.ps1
    # Checks the previous weekend.

.EXAMPLE
    .\Get-SCCMWeekendRestarts-Pwsh7.ps1 -SiteServer SCCM-MP01.corp.local -SiteCode PS1 `
        -CollectionName "All Workstations - Prod" -WeekendOf 2026-06-13 -ThrottleLimit 40

.EXAMPLE
    .\Get-SCCMWeekendRestarts-Pwsh7.ps1 -CountBoots

.NOTES
    Version : 1.0
    Requires: PowerShell 7.0+.
              Read rights on the SCCM site WMI namespace (root\SMS\site_<code>).
              WinRM or DCOM access to the target machines for the event-log read.
              An account with rights to read the remote System event log.
    Note    : Out-GridView / Out-ConsoleGridView are optional. If neither is
              present the script falls back to a numbered console menu, so it
              works on a headless server too.
#>

[CmdletBinding()]
param(
    [string]$SiteServer,
    [string]$SiteCode,
    [string]  $CollectionName,
    [datetime]$WeekendOf,
    [switch]  $CountBoots,
    [string]$OutputPath    = (Get-Location).Path,
    [int]   $ThrottleLimit = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Banner ──────────────────────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║      SCCM Weekend Restart Counter  v1.0  (PS7)           ║" -ForegroundColor Cyan
Write-Host "  ║   Collection → per-machine Saturday / Sunday restarts    ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
#endregion

#region ── Helper: plain text prompt ──────────────────────────────────────────
function Read-Prompt {
    param([string]$Label, [string]$Default = "")
    $display = if ($Default) { "$Label [$Default]" } else { $Label }
    $val     = Read-Host "  $display"
    if ([string]::IsNullOrWhiteSpace($val) -and $Default) { return $Default }
    return $val.Trim()
}
#endregion

#region ── Helper: single-select picker (GUI → console grid → menu) ───────────
function Select-OneItem {
    param(
        [Parameter(Mandatory)] [object[]]$Items,
        [Parameter(Mandatory)] [string]  $Title,
        [string[]]$DisplayProperty
    )

    # 1) Out-GridView (Windows + Microsoft.PowerShell.GraphicalTools)
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        try { return ($Items | Out-GridView -Title $Title -OutputMode Single) } catch { }
    }
    # 2) Out-ConsoleGridView (Microsoft.PowerShell.ConsoleGuiTools) – works headless
    if (Get-Command Out-ConsoleGridView -ErrorAction SilentlyContinue) {
        try { return ($Items | Out-ConsoleGridView -Title $Title -OutputMode Single) } catch { }
    }
    # 3) Plain numbered console menu
    Write-Host ""
    Write-Host "  $Title" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $label = if ($DisplayProperty) {
            ($DisplayProperty | ForEach-Object { $Items[$i].$_ }) -join '  |  '
        } else { "$($Items[$i])" }
        Write-Host ("   [{0,3}] {1}" -f ($i + 1), $label)
    }
    while ($true) {
        $pick = Read-Host "  Enter number (1-$($Items.Count)), or blank to cancel"
        if ([string]::IsNullOrWhiteSpace($pick)) { return $null }
        if ($pick -as [int] -and [int]$pick -ge 1 -and [int]$pick -le $Items.Count) {
            return $Items[[int]$pick - 1]
        }
        Write-Host "  [!] Invalid selection." -ForegroundColor Yellow
    }
}
#endregion

#region ── Step 1 – Site Server & Site Code ───────────────────────────────────
if (-not $SiteServer) { $SiteServer = Read-Prompt "SCCM Site Server / MP FQDN" "SCCM-MP01.corp.local" }
if (-not $SiteCode)   { $SiteCode   = Read-Prompt "Site Code" "PS1" }
$sccmNamespace = "root\SMS\site_$SiteCode"

# Resolve the target weekend (Saturday + Sunday).
if ($PSBoundParameters.ContainsKey('WeekendOf')) {
    # Saturday on-or-before the supplied date = that calendar week's weekend.
    $anchor       = $WeekendOf.Date
    $daysSinceSat = ([int]$anchor.DayOfWeek + 1) % 7        # Sat->0, Sun->1, ... Fri->6
    $targetSat    = $anchor.AddDays(-$daysSinceSat)
}
else {
    # Default: the previous weekend relative to today.
    $today        = (Get-Date).Date
    $daysSinceSat = ([int]$today.DayOfWeek + 1) % 7
    $targetSat    = $today.AddDays(-$daysSinceSat)
    # If today itself is Sat/Sun we're in this weekend – step back to the prior one.
    if ([int]$today.DayOfWeek -eq 6 -or [int]$today.DayOfWeek -eq 0) { $targetSat = $targetSat.AddDays(-7) }
}
$targetSun  = $targetSat.AddDays(1)
$windowEnd  = $targetSat.AddDays(2)        # exclusive upper bound (Monday 00:00)

$eventBasis = if ($CountBoots) { 'system boots (EventID 6005)' }
              else             { 'explicit restarts (EventID 1074)' }
Write-Host ""
Write-Host ("  [*] Weekend : Sat {0:yyyy-MM-dd} + Sun {1:yyyy-MM-dd}" -f $targetSat, $targetSun) -ForegroundColor DarkCyan
Write-Host "  [*] Counting $eventBasis." -ForegroundColor DarkCyan
#endregion

#region ── Step 2 – Collection picker ─────────────────────────────────────────
$CollectionID = $null
if (-not $CollectionName) {
    Write-Host ""
    Write-Host "  [*] Fetching device collections from $SiteServer ..." -ForegroundColor Cyan
    try {
        $allCollections = Get-CimInstance -ComputerName $SiteServer `
                                          -Namespace $sccmNamespace `
                                          -ClassName SMS_Collection `
                                          -Filter "CollectionType = 2" `
                                          -ErrorAction Stop |
                          Select-Object Name,
                                        CollectionID,
                                        @{N='Members';     E={ $_.MemberCount }},
                                        Comment,
                                        @{N='LastRefresh'; E={ $_.LastRefreshTime }} |
                          Sort-Object Name

        if (-not $allCollections) { throw "No device collections returned." }
        Write-Host "  [+] $($allCollections.Count) collections found – select one." -ForegroundColor Green

        $sel = Select-OneItem -Items $allCollections `
                              -Title "Select Device Collection (single-select)" `
                              -DisplayProperty Name, CollectionID, Members

        if (-not $sel) { Write-Error "No collection selected."; exit 1 }
        $CollectionName = $sel.Name
        $CollectionID   = $sel.CollectionID
        Write-Host "  [+] Collection : $CollectionName  ($CollectionID)" -ForegroundColor Green
    }
    catch {
        Write-Warning "  [!] WMI collection query failed: $_"
        $CollectionName = Read-Prompt "Enter Collection Name manually"
        $CollectionID   = $null
    }
}
#endregion

#region ── Step 3 – Resolve collection members ───────────────────────────────
Write-Host ""
Write-Host "  [*] Resolving members of '$CollectionName' ..." -ForegroundColor Cyan
$machines = @()
try {
    if (-not $CollectionID) {
        $cq = "SELECT CollectionID FROM SMS_Collection WHERE Name = '$($CollectionName -replace "'","''")'"
        $co = Get-CimInstance -ComputerName $SiteServer -Namespace $sccmNamespace -Query $cq -ErrorAction Stop
        if (-not $co) { throw "Collection '$CollectionName' not found." }
        $CollectionID = $co.CollectionID
    }
    $mq       = "SELECT Name FROM SMS_FullCollectionMembership WHERE CollectionID = '$CollectionID'"
    $members  = Get-CimInstance -ComputerName $SiteServer -Namespace $sccmNamespace -Query $mq -ErrorAction Stop
    $machines = @($members | Select-Object -ExpandProperty Name | Sort-Object -Unique)
    Write-Host "  [+] $($machines.Count) machines in collection." -ForegroundColor Green
}
catch {
    Write-Warning "  [!] Could not retrieve members via WMI: $_"
    Write-Host "  [?] Enter machine names manually (one per line, blank line to finish):" -ForegroundColor Yellow
    while ($true) {
        $ln = Read-Host "  "
        if ([string]::IsNullOrWhiteSpace($ln)) { break }
        $machines += $ln.Trim()
    }
    $machines = @($machines | Sort-Object -Unique)
    Write-Host "  [+] Using $($machines.Count) manually entered machines." -ForegroundColor Yellow
}

if (-not $machines -or $machines.Count -eq 0) { Write-Error "No machines to process."; exit 1 }
#endregion

#region ── Parallel query (ForEach-Object -Parallel) ─────────────────────────
Write-Host ""
Write-Host "  [*] Querying $($machines.Count) machine(s), up to $ThrottleLimit in parallel ..." -ForegroundColor Cyan

$results = $machines | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $MachineName = $_
    $winStart    = $using:targetSat       # Saturday 00:00 (inclusive)
    $winEnd      = $using:windowEnd        # Monday   00:00 (exclusive)
    $CountBoots  = $using:CountBoots

    $r = [PSCustomObject]@{
        Machine          = $MachineName
        Online           = $false
        SaturdayRestarts = 0
        SundayRestarts   = 0
        WeekendRestarts  = 0
        LastWeekendBoot  = $null
        Notes            = ''
        Detail           = @()        # nested – expanded into the detail CSV later
    }

    # Quick reachability check – avoids long CIM timeouts on dead hosts.
    if (-not (Test-Connection -ComputerName $MachineName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        $r.Notes = 'No ping response'
        return $r
    }
    $r.Online = $true

    # CIM session: try WSMan (default) then fall back to DCOM.
    $cimSess = $null
    try {
        $cimSess = New-CimSession -ComputerName $MachineName -OperationTimeoutSec 30 -ErrorAction Stop
    }
    catch {
        try {
            $opt     = New-CimSessionOption -Protocol Dcom
            $cimSess = New-CimSession -ComputerName $MachineName -SessionOption $opt -OperationTimeoutSec 30 -ErrorAction Stop
        }
        catch { $r.Notes = 'CIM session failed (WSMan + DCOM)'; return $r }
    }

    try {
        # Bound the query server-side to just the target weekend so we don't
        # pull the entire System log over the wire.
        $dmtfStart = [Management.ManagementDateTimeConverter]::ToDmtfDateTime($winStart)
        $dmtfEnd   = [Management.ManagementDateTimeConverter]::ToDmtfDateTime($winEnd)

        $eventCode = if ($CountBoots) { 6005 } else { 1074 }
        $filter    = "Logfile='System' AND EventCode=$eventCode AND TimeWritten>='$dmtfStart' AND TimeWritten<'$dmtfEnd'"

        $evts   = Get-CimInstance -CimSession $cimSess -ClassName Win32_NTLogEvent `
                      -Filter $filter -ErrorAction Stop
        $detail = [System.Collections.Generic.List[object]]::new()

        foreach ($e in @($evts)) {
            $when = $e.TimeGenerated
            if (-not $when) { continue }

            # For 1074, keep only entries that actually describe a *restart*
            # (the same EventID also logs straight power-offs). 6005 boots are
            # always counted.
            if (-not $CountBoots) {
                $msg = "$($e.Message)"
                if ($msg -and $msg -notmatch '(?i)restart') { continue }
            }

            $dow = $when.DayOfWeek
            if ($dow -eq 'Saturday' -or $dow -eq 'Sunday') {
                if ($dow -eq 'Saturday') { $r.SaturdayRestarts++ } else { $r.SundayRestarts++ }
                $detail.Add([PSCustomObject]@{
                    Machine   = $MachineName
                    EventTime = $when.ToString('yyyy-MM-dd HH:mm:ss')
                    DayOfWeek = "$dow"
                    EventID   = $e.EventCode
                    Source    = "$($e.SourceName)"
                    Message   = ("$($e.Message)" -replace '\s+', ' ').Trim()
                })
                if (-not $r.LastWeekendBoot -or $when -gt [datetime]$r.LastWeekendBoot) {
                    $r.LastWeekendBoot = $when.ToString('yyyy-MM-dd HH:mm:ss')
                }
            }
        }

        $r.Detail          = $detail.ToArray()
        $r.WeekendRestarts = $r.SaturdayRestarts + $r.SundayRestarts
        if (-not $r.Notes) { $r.Notes = 'OK' }
    }
    catch { $r.Notes = "Event log query error: $($_.Exception.Message)" }
    finally { Remove-CimSession -CimSession $cimSess -ErrorAction SilentlyContinue }

    return $r
}
#endregion

#region ── Output ────────────────────────────────────────────────────────────
$wkTag       = $targetSat.ToString('yyyy-MM-dd')
$safeColl    = ($CollectionName -replace '[^\w\-]', '_')
$summaryPath = Join-Path $OutputPath "WeekendRestarts_${safeColl}_$wkTag`_Summary.csv"
$detailPath  = Join-Path $OutputPath "WeekendRestarts_${safeColl}_$wkTag`_Detail.csv"

$summary = $results |
    Select-Object Machine, Online, SaturdayRestarts, SundayRestarts, WeekendRestarts, LastWeekendBoot, Notes |
    Sort-Object WeekendRestarts -Descending

$detail  = $results | ForEach-Object { $_.Detail } | Sort-Object Machine, EventTime

$summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
if ($detail) { $detail | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8 }

Write-Host ""
Write-Host "  ── Per-machine weekend restart summary ─────────────────────" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

$totSat    = ($results | Measure-Object SaturdayRestarts -Sum).Sum
$totSun    = ($results | Measure-Object SundayRestarts   -Sum).Sum
$onlineCnt = @($results | Where-Object Online).Count

Write-Host ""
Write-Host ("  ── Totals — '{0}'  (Sat {1:yyyy-MM-dd} + Sun {2:yyyy-MM-dd}) ──" -f $CollectionName, $targetSat, $targetSun) -ForegroundColor Cyan
Write-Host ("    Machines processed     : {0}" -f $results.Count)
Write-Host ("    Machines reachable     : {0}" -f $onlineCnt)
Write-Host ("    Saturday restarts      : {0}" -f $totSat)
Write-Host ("    Sunday restarts        : {0}" -f $totSun)
Write-Host ("    Total weekend restarts : {0}" -f ($totSat + $totSun))
Write-Host ""
Write-Host "  [+] Summary CSV : $summaryPath" -ForegroundColor Green
if ($detail) { Write-Host "  [+] Detail  CSV : $detailPath"  -ForegroundColor Green }
else         { Write-Host "  [i] No weekend restarts found – detail CSV not written." -ForegroundColor Yellow }
Write-Host ""
#endregion
