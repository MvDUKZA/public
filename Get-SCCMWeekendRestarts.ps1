#Requires -Version 5.1
<#
.SYNOPSIS
    Enumerates SCCM device collections, lets you pick one, then reports how many
    times each machine in that collection was restarted on a Saturday or Sunday.

.DESCRIPTION
    Workflow:
      1. Connects to the SCCM site server WMI namespace (root\SMS\site_<code>).
      2. Lists all device collections in an Out-GridView picker (single-select).
      3. Resolves the collection's members (SMS_FullCollectionMembership).
      4. For every member it opens a CIM session (WSMan, falling back to DCOM)
         and reads the System event log for restart events, counting only those
         that fall on a Saturday or Sunday within the look-back window.

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
    Device Collection name. Selected via Out-GridView if omitted.

.PARAMETER DaysBack
    How far back to look for restart events. Default: 90 days.

.PARAMETER CountBoots
    Count every system boot (EventID 6005) instead of explicit restarts
    (EventID 1074). Captures cold boots / crash recoveries as well.

.PARAMETER OutputPath
    Folder for the CSV output. Defaults to the current directory.

.PARAMETER MaxConcurrent
    Parallel runspaces. Default: 20.

.EXAMPLE
    .\Get-SCCMWeekendRestarts.ps1

.EXAMPLE
    .\Get-SCCMWeekendRestarts.ps1 -SiteServer SCCM-MP01.corp.local -SiteCode PS1 `
        -CollectionName "All Workstations - Prod" -DaysBack 180

.EXAMPLE
    .\Get-SCCMWeekendRestarts.ps1 -CountBoots -DaysBack 30

.NOTES
    Version : 1.0
    Requires: Read rights on the SCCM site WMI namespace (root\SMS\site_<code>).
              WinRM or DCOM access to the target machines for the event-log read.
              An account with rights to read the remote System event log.
#>

[CmdletBinding()]
param(
    [string]$SiteServer,
    [string]$SiteCode,
    [string]$CollectionName,
    [int]   $DaysBack      = 90,
    [switch]$CountBoots,
    [string]$OutputPath    = (Get-Location).Path,
    [int]   $MaxConcurrent = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

#region ── Banner ──────────────────────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║         SCCM Weekend Restart Counter  v1.0               ║" -ForegroundColor Cyan
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

#region ── Step 1 – Site Server & Site Code ───────────────────────────────────
if (-not $SiteServer) { $SiteServer = Read-Prompt "SCCM Site Server / MP FQDN" "SCCM-MP01.corp.local" }
if (-not $SiteCode)   { $SiteCode   = Read-Prompt "Site Code" "PS1" }
$sccmNamespace = "root\SMS\site_$SiteCode"

$eventBasis = if ($CountBoots) { 'system boots (EventID 6005)' }
              else             { 'explicit restarts (EventID 1074)' }
Write-Host ""
Write-Host "  [*] Counting $eventBasis over the last $DaysBack day(s)." -ForegroundColor DarkCyan
#endregion

#region ── Step 2 – Collection picker (Out-GridView) ─────────────────────────
$CollectionID = $null
if (-not $CollectionName) {
    Write-Host ""
    Write-Host "  [*] Fetching device collections from $SiteServer ..." -ForegroundColor Cyan
    try {
        $allCollections = Get-WmiObject -ComputerName $SiteServer `
                                        -Namespace $sccmNamespace `
                                        -Class SMS_Collection `
                                        -Filter "CollectionType = 2" `
                                        -ErrorAction Stop |
                          Select-Object Name,
                                        CollectionID,
                                        @{N='Members';     E={ $_.MemberCount }},
                                        Comment,
                                        @{N='LastRefresh'; E={
                                            [Management.ManagementDateTimeConverter]::ToDateTime($_.LastRefreshTime)
                                        }} |
                          Sort-Object Name

        if (-not $allCollections) { throw "No device collections returned." }
        Write-Host "  [+] $($allCollections.Count) collections found – select one and click OK." -ForegroundColor Green

        $sel = $allCollections | Out-GridView `
                   -Title "Select Device Collection  (single-select → OK)" `
                   -OutputMode Single

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
        $co = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace -Query $cq -ErrorAction Stop
        if (-not $co) { throw "Collection '$CollectionName' not found." }
        $CollectionID = $co.CollectionID
    }
    $mq       = "SELECT Name FROM SMS_FullCollectionMembership WHERE CollectionID = '$CollectionID'"
    $members  = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace -Query $mq -ErrorAction Stop
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

#region ── Per-machine worker scriptblock ────────────────────────────────────
#  Runs inside each runspace. Self-contained – no outer-scope references.
$worker = {
    param($MachineName, $DaysBack, $CountBoots)

    $r = [PSCustomObject]@{
        Machine          = $MachineName
        Online           = $false
        SaturdayRestarts = 0
        SundayRestarts   = 0
        WeekendRestarts  = 0
        LastWeekendBoot  = $null
        Notes            = ''
        _Detail          = @()        # carried back, expanded into the detail CSV
    }

    # Quick reachability check (ping) – avoids long CIM timeouts on dead hosts.
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
        # Bound the query server-side with a DMTF datetime so we don't pull the
        # entire System log over the wire.
        $startDate = (Get-Date).AddDays(-1 * [math]::Abs($DaysBack))
        $dmtf      = [Management.ManagementDateTimeConverter]::ToDmtfDateTime($startDate)

        $eventCode = if ($CountBoots) { 6005 } else { 1074 }
        $filter    = "Logfile='System' AND EventCode=$eventCode AND TimeWritten>='$dmtf'"

        $evts = Get-CimInstance -CimSession $cimSess -ClassName Win32_NTLogEvent `
                    -Filter $filter -ErrorAction Stop

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
                $r._Detail += [PSCustomObject]@{
                    Machine   = $MachineName
                    EventTime = $when.ToString('yyyy-MM-dd HH:mm:ss')
                    DayOfWeek = "$dow"
                    EventID   = $e.EventCode
                    Source    = "$($e.SourceName)"
                    Message   = ("$($e.Message)" -replace '\s+', ' ').Trim()
                }
                if (-not $r.LastWeekendBoot -or $when -gt [datetime]$r.LastWeekendBoot) {
                    $r.LastWeekendBoot = $when.ToString('yyyy-MM-dd HH:mm:ss')
                }
            }
        }

        $r.WeekendRestarts = $r.SaturdayRestarts + $r.SundayRestarts
        if (-not $r.Notes) { $r.Notes = 'OK' }
    }
    catch { $r.Notes = "Event log query error: $($_.Exception.Message)" }
    finally { Remove-CimSession -CimSession $cimSess -ErrorAction SilentlyContinue }

    return $r
}
#endregion

#region ── Parallel execution via runspace pool ──────────────────────────────
Write-Host ""
Write-Host "  [*] Querying $($machines.Count) machine(s) with up to $MaxConcurrent in parallel ..." -ForegroundColor Cyan

$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxConcurrent)
$pool.Open()

$jobs = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($m in $machines) {
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($worker).AddArgument($m).AddArgument($DaysBack).AddArgument([bool]$CountBoots)
    $jobs.Add([PSCustomObject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Machine = $m })
}

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$done    = 0
foreach ($j in $jobs) {
    try   { $res = $j.PS.EndInvoke($j.Handle); if ($res) { $results.Add($res) } }
    catch { $results.Add([PSCustomObject]@{ Machine=$j.Machine; Online=$false; SaturdayRestarts=0; SundayRestarts=0; WeekendRestarts=0; LastWeekendBoot=$null; Notes="Runspace error: $($_.Exception.Message)"; _Detail=@() }) }
    finally { $j.PS.Dispose() }
    $done++
    Write-Progress -Activity "Reading event logs" -Status "$done / $($jobs.Count)" -PercentComplete (($done / $jobs.Count) * 100)
}
Write-Progress -Activity "Reading event logs" -Completed
$pool.Close(); $pool.Dispose()
#endregion

#region ── Output ────────────────────────────────────────────────────────────
$stamp       = Get-Date -Format 'yyyyMMdd_HHmmss'
$safeColl    = ($CollectionName -replace '[^\w\-]', '_')
$summaryPath = Join-Path $OutputPath "WeekendRestarts_${safeColl}_$stamp`_Summary.csv"
$detailPath  = Join-Path $OutputPath "WeekendRestarts_${safeColl}_$stamp`_Detail.csv"

$summary = $results |
    Select-Object Machine, Online, SaturdayRestarts, SundayRestarts, WeekendRestarts, LastWeekendBoot, Notes |
    Sort-Object WeekendRestarts -Descending

$detail  = $results | ForEach-Object { $_._Detail } | Sort-Object Machine, EventTime

$summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
if ($detail) { $detail | Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8 }

Write-Host ""
Write-Host "  ── Per-machine weekend restart summary ─────────────────────" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

$totSat     = ($results | Measure-Object SaturdayRestarts -Sum).Sum
$totSun     = ($results | Measure-Object SundayRestarts   -Sum).Sum
$onlineCnt  = @($results | Where-Object Online).Count

Write-Host ""
Write-Host "  ── Totals across collection '$CollectionName' ──────────────" -ForegroundColor Cyan
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
