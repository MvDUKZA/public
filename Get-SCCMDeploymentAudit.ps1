#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Deployment Audit - Queries each machine for update download, install and restart timestamps.

.DESCRIPTION
    Prompts for deployment parameters, then connects to each target machine via WMI/CIM to
    interrogate the Windows Update Agent, SCCM Client (CCMExec) and Event Log for:
      - When the update content was downloaded
      - When the update was installed
      - When the machine last restarted (post-install)
    Results are exported to a timestamped CSV file.

.PARAMETER SiteServer
    SCCM Management Point / Site Server FQDN. Prompted if not supplied.

.PARAMETER SiteCode
    SCCM Site Code (e.g. PS1). Prompted if not supplied.

.PARAMETER CollectionName
    SCCM Device Collection name to audit. Prompted if not supplied.

.PARAMETER KBArticle
    KB number to audit (e.g. KB5034441). Prompted if not supplied.

.PARAMETER OutputPath
    Folder to write the CSV to. Defaults to the current directory.

.PARAMETER MaxConcurrent
    Number of machines to query in parallel (runspaces). Default: 20.

.EXAMPLE
    .\Get-SCCMDeploymentAudit.ps1

.EXAMPLE
    .\Get-SCCMDeploymentAudit.ps1 -SiteServer SCCM-MP01.corp.local -SiteCode PS1 `
        -CollectionName "All Workstations - Prod" -KBArticle KB5034441

.NOTES
    Author  : SCCM Deployment Audit Tool
    Version : 2.0
    Requires: SCCM Admin Console installed (ConfigurationManager module), or
              WMI access to the Site Server, plus WinRM / CIM access to target machines.
              Run as an account with SCCM Read rights and local admin on targets.
#>

[CmdletBinding()]
param(
    [string]$SiteServer,
    [string]$SiteCode,
    [string]$CollectionName,
    [string]$KBArticle,
    [string]$OutputPath = (Get-Location).Path,
    [int]$MaxConcurrent = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

#region ── Banner ──────────────────────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║       SCCM Deployment Audit Tool  v2.0                  ║" -ForegroundColor Cyan
Write-Host "  ║  Download · Install · Restart timestamps per machine    ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
#endregion

#region ── Prompt for parameters if not supplied ───────────────────────────────
function Read-Prompt {
    param([string]$Label, [string]$Default = "")
    $display = if ($Default) { "$Label [$Default]" } else { $Label }
    $val = Read-Host "  $display"
    if ([string]::IsNullOrWhiteSpace($val) -and $Default) { return $Default }
    return $val.Trim()
}

if (-not $SiteServer)    { $SiteServer    = Read-Prompt "SCCM Site Server / MP FQDN" "SCCM-MP01.corp.local" }
if (-not $SiteCode)      { $SiteCode      = Read-Prompt "Site Code" "PS1" }
if (-not $CollectionName){ $CollectionName= Read-Prompt "Device Collection Name" "All Workstations - Prod" }
if (-not $KBArticle)     { $KBArticle     = Read-Prompt "KB Article (e.g. KB5034441)" }

$KBNumber = $KBArticle -replace '[^0-9]', ''   # strip non-numeric for WMI queries

Write-Host ""
Write-Host "  Parameters confirmed:" -ForegroundColor Green
Write-Host "    Site Server  : $SiteServer"
Write-Host "    Site Code    : $SiteCode"
Write-Host "    Collection   : $CollectionName"
Write-Host "    KB Article   : $KBArticle  (numeric: $KBNumber)"
Write-Host "    Output Path  : $OutputPath"
Write-Host ""
#endregion

#region ── Get collection members from SCCM via WMI ───────────────────────────
Write-Host "  [*] Connecting to SCCM Site Server..." -ForegroundColor Cyan

$sccmNamespace = "root\SMS\site_$SiteCode"
$machines = @()

try {
    # Query SMS_FullCollectionMembership for the named collection
    $collectionQuery = "SELECT CollectionID FROM SMS_Collection WHERE Name = '$($CollectionName -replace "'","''")'"
    $collObj = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace `
                             -Query $collectionQuery -ErrorAction Stop

    if (-not $collObj) {
        throw "Collection '$CollectionName' not found on $SiteServer."
    }

    $collectionID = $collObj.CollectionID
    Write-Host "  [+] Collection ID: $collectionID" -ForegroundColor Green

    $memberQuery = "SELECT Name, ResourceID FROM SMS_FullCollectionMembership WHERE CollectionID = '$collectionID'"
    $members = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace `
                             -Query $memberQuery -ErrorAction Stop

    $machines = @($members | Select-Object -ExpandProperty Name | Sort-Object -Unique)
    Write-Host "  [+] Found $($machines.Count) machines in collection." -ForegroundColor Green
}
catch {
    Write-Warning "  [!] Could not retrieve collection members from SCCM WMI: $_"
    Write-Host "  [?] Enter machine names manually (comma or newline separated), then press Enter twice:" -ForegroundColor Yellow
    $manualInput = @()
    while ($true) {
        $line = Read-Host "  "
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $manualInput += $line
    }
    $machines = $manualInput -split '[,\n\r]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | Sort-Object -Unique
    Write-Host "  [+] Using $($machines.Count) manually entered machines." -ForegroundColor Yellow
}

if ($machines.Count -eq 0) {
    Write-Error "No machines to audit. Exiting."
    exit 1
}
Write-Host ""
#endregion

#region ── Per-machine query scriptblock (runs in runspace pool) ───────────────
$queryScriptBlock = {
    param(
        [string]$MachineName,
        [string]$KBNumber,
        [string]$KBArticle
    )

    $result = [PSCustomObject]@{
        Machine          = $MachineName
        Online           = $false
        Downloaded       = 'Unknown'
        DownloadedSource = ''
        Installed        = 'Unknown'
        InstalledBy      = ''
        Restarted        = 'Unknown'
        RestartType      = ''
        ComplianceState  = 'Unknown'
        LastHWInventory  = ''
        ErrorCode        = ''
        Notes            = ''
    }

    # ── Ping test ────────────────────────────────────────────────────────────
    $ping = Test-Connection -ComputerName $MachineName -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $ping) {
        $result.ComplianceState = 'Offline'
        $result.Notes           = 'Machine did not respond to ping'
        return $result
    }
    $result.Online = $true

    $cimSession = $null
    try {
        $cimOpt = New-CimSessionOption -Protocol Wsman
        $cimSession = New-CimSession -ComputerName $MachineName -SessionOption $cimOpt `
                                     -OperationTimeoutSec 30 -ErrorAction Stop
    }
    catch {
        # Fallback to DCOM
        try {
            $cimOpt = New-CimSessionOption -Protocol Dcom
            $cimSession = New-CimSession -ComputerName $MachineName -SessionOption $cimOpt `
                                         -OperationTimeoutSec 30 -ErrorAction Stop
        }
        catch {
            $result.ComplianceState = 'WMI Error'
            $result.Notes           = "CIM connection failed: $($_.Exception.Message)"
            return $result
        }
    }

    try {
        # ── 1. DOWNLOAD TIME ─────────────────────────────────────────────────
        # CCM_SoftwareUpdate in root\ccm\clientsdk holds per-update state
        $swUpdates = Get-CimInstance -CimSession $cimSession `
                                     -Namespace 'root\ccm\clientsdk' `
                                     -ClassName 'CCM_SoftwareUpdate' `
                                     -ErrorAction SilentlyContinue |
                     Where-Object { $_.ArticleID -eq $KBNumber -or $_.BulletinID -like "*$KBNumber*" }

        if ($swUpdates) {
            $upd = $swUpdates | Select-Object -First 1

            # EvaluationState: 8 = PendingInstall (downloaded), 9 = PendingReboot, 13 = Installed
            # ContentDownloadTime is available on some CCM versions
            if ($upd.PSObject.Properties['ContentDownloadTime'] -and $upd.ContentDownloadTime) {
                $result.Downloaded = $upd.ContentDownloadTime.ToString('yyyy-MM-dd HH:mm:ss')
                $result.DownloadedSource = 'CCM_SoftwareUpdate'
            }
            else {
                # Fall back to CacheInfoEx – look for matching content
                $cacheItems = Get-CimInstance -CimSession $cimSession `
                                              -Namespace 'root\ccm\softmgmtagent' `
                                              -ClassName 'CacheInfoEx' `
                                              -ErrorAction SilentlyContinue |
                              Where-Object { $_.ContentID -like "*$KBNumber*" }
                if ($cacheItems) {
                    $dlTime = ($cacheItems | Sort-Object LastReferenced | Select-Object -Last 1).LastReferenced
                    $result.Downloaded       = $dlTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $result.DownloadedSource = 'CacheInfoEx'
                }
                else {
                    $result.Downloaded       = 'Content not in cache'
                    $result.DownloadedSource = ''
                }
            }

            # Map EvaluationState to compliance
            $stateMap = @{
                0  = 'None'
                1  = 'Available'
                2  = 'Submitted'
                3  = 'Detecting'
                4  = 'PreDownload'
                5  = 'Downloading'
                6  = 'WaitInstall'
                7  = 'Installing'
                8  = 'PendingInstall'
                9  = 'PendingReboot'
                10 = 'PendingReboot'
                11 = 'Verifying'
                12 = 'InstallComplete'
                13 = 'Error'
                14 = 'WaitServiceWindow'
            }
            $evalState = [int]$upd.EvaluationState
            $result.ComplianceState = if ($stateMap.ContainsKey($evalState)) { $stateMap[$evalState] } else { "State $evalState" }
            if ($upd.ErrorCode -and $upd.ErrorCode -ne 0) {
                $result.ErrorCode = '0x{0:X8}' -f [uint32]$upd.ErrorCode
            }
            else {
                $result.ErrorCode = '0x00000000'
            }
        }
        else {
            $result.Downloaded      = 'Not found in CCM'
            $result.ComplianceState = 'Not Targeted / NA'
        }

        # ── 2. INSTALL TIME ──────────────────────────────────────────────────
        # Primary: Win32_QuickFixEngineering
        $qfe = Get-CimInstance -CimSession $cimSession -ClassName 'Win32_QuickFixEngineering' `
                               -ErrorAction SilentlyContinue |
               Where-Object { $_.HotFixID -eq "KB$KBNumber" }

        if ($qfe) {
            $installDate = $qfe | Select-Object -First 1
            if ($installDate.InstalledOn) {
                $result.Installed   = $installDate.InstalledOn.ToString('yyyy-MM-dd HH:mm:ss')
                $result.InstalledBy = $installDate.InstalledBy
            }
            else {
                $result.Installed   = 'Installed (date unavailable)'
                $result.InstalledBy = $installDate.InstalledBy
            }
            if ($result.ComplianceState -notin @('PendingReboot','InstallComplete')) {
                $result.ComplianceState = 'Compliant'
            }
        }
        else {
            # Fallback: query Windows Update COM via registry hive
            $regPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install"
            $reg = Invoke-CimMethod -CimSession $cimSession `
                                    -ClassName 'StdRegProv' `
                                    -MethodName 'GetStringValue' `
                                    -Namespace 'root\default' `
                                    -Arguments @{ hDefKey = [uint32]'0x80000002'; sSubKeyName = $regPath; sValueName = 'LastSuccessTime' } `
                                    -ErrorAction SilentlyContinue
            if ($reg -and $reg.sValue) {
                $result.Installed = $reg.sValue
                $result.Notes    += 'Install time from WU registry; '
            }
            else {
                if ($result.ComplianceState -notin @('PendingReboot','InstallComplete','Compliant')) {
                    $result.Installed = 'Not installed'
                }
            }
        }

        # ── 3. RESTART TIME ──────────────────────────────────────────────────
        # Event 6005 (EventLog service start = system boot) or 6006/6008/1074 in System log
        $bootEvents = Get-CimInstance -CimSession $cimSession `
                                      -ClassName 'Win32_NTLogEvent' `
                                      -Filter "Logfile='System' AND (EventCode=6005 OR EventCode=6009 OR EventCode=1074)" `
                                      -ErrorAction SilentlyContinue |
                      Sort-Object TimeGenerated -Descending

        if ($bootEvents) {
            $lastBoot = $bootEvents | Select-Object -First 1
            $result.Restarted    = $lastBoot.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss')
            $result.RestartType  = switch ($lastBoot.EventCode) {
                6005 { 'Clean boot (EventID 6005)' }
                6009 { 'Boot (EventID 6009)' }
                1074 { 'Planned restart (EventID 1074)' }
                default { "EventID $($lastBoot.EventCode)" }
            }
        }
        else {
            # Fallback: LastBootUpTime from Win32_OperatingSystem
            $os = Get-CimInstance -CimSession $cimSession -ClassName 'Win32_OperatingSystem' `
                                  -ErrorAction SilentlyContinue
            if ($os) {
                $result.Restarted   = $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')
                $result.RestartType = 'Win32_OperatingSystem.LastBootUpTime'
            }
            else {
                $result.Restarted = 'Unable to determine'
            }
        }

        # ── 4. SCCM Last HW Inventory (from local CCM) ───────────────────────
        $ccmClient = Get-CimInstance -CimSession $cimSession `
                                     -Namespace 'root\ccm\invagt' `
                                     -ClassName 'InventoryActionStatus' `
                                     -ErrorAction SilentlyContinue |
                     Where-Object { $_.InventoryActionID -eq '{00000000-0000-0000-0000-000000000001}' }
        if ($ccmClient) {
            $result.LastHWInventory = $ccmClient.LastReportDate.ToString('yyyy-MM-dd HH:mm:ss')
        }

        # ── Final compliance state resolution ────────────────────────────────
        if ($result.ComplianceState -eq 'Unknown' -and $result.Installed -ne 'Not installed' -and $result.Installed -ne 'Unknown') {
            if ($result.Restarted -ne 'Unable to determine') {
                $result.ComplianceState = 'Compliant'
            }
        }

    }
    catch {
        $result.ComplianceState = 'Query Error'
        $result.Notes          += "Error during query: $($_.Exception.Message)"
    }
    finally {
        if ($cimSession) { Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue }
    }

    return $result
}
#endregion

#region ── Parallel execution via runspace pool ────────────────────────────────
Write-Host "  [*] Auditing $($machines.Count) machines (up to $MaxConcurrent concurrent)..." -ForegroundColor Cyan
Write-Host ""

$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxConcurrent)
$pool.Open()

$jobs = @()
foreach ($machine in $machines) {
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($queryScriptBlock)
    [void]$ps.AddArgument($machine)
    [void]$ps.AddArgument($KBNumber)
    [void]$ps.AddArgument($KBArticle)
    $jobs += [PSCustomObject]@{
        Machine    = $machine
        PowerShell = $ps
        Handle     = $ps.BeginInvoke()
    }
}

# Collect results with progress
$results         = @()
$completed       = 0
$totalJobs       = $jobs.Count
$progressParams  = @{
    Activity = "SCCM Deployment Audit — $KBArticle"
    Status   = "Querying machines..."
    Id       = 1
}

while ($jobs | Where-Object { -not $_.Handle.IsCompleted }) {
    $done = @($jobs | Where-Object { $_.Handle.IsCompleted })
    foreach ($job in $done) {
        if ($job.PowerShell.HadErrors) {
            Write-Warning "  [!] Errors on $($job.Machine): $($job.PowerShell.Streams.Error[0])"
        }
        $r = $job.PowerShell.EndInvoke($job.Handle)
        if ($r) {
            $results += $r
            $completed++
            $icon  = if ($r.ComplianceState -eq 'Compliant') { '[✓]' } elseif ($r.ComplianceState -like '*Error*' -or $r.ComplianceState -eq 'Offline') { '[✗]' } else { '[~]' }
            $color = if ($r.ComplianceState -eq 'Compliant') { 'Green' } elseif ($r.ComplianceState -like '*Error*' -or $r.ComplianceState -eq 'Offline') { 'Red' } else { 'Yellow' }
            Write-Host ("  {0} {1,-25} {2}" -f $icon, $r.Machine, $r.ComplianceState) -ForegroundColor $color
        }
        $job.PowerShell.Dispose()
        # Remove from jobs list
        $jobs = $jobs | Where-Object { $_.Machine -ne $job.Machine }
    }
    $pctComplete = [int](($completed / $totalJobs) * 100)
    Write-Progress @progressParams -PercentComplete $pctComplete `
                   -CurrentOperation "$completed of $totalJobs complete"
    Start-Sleep -Milliseconds 250
}

# Catch any remaining
foreach ($job in $jobs) {
    $r = $job.PowerShell.EndInvoke($job.Handle)
    if ($r) { $results += $r }
    $job.PowerShell.Dispose()
}

$pool.Close()
$pool.Dispose()
Write-Progress @progressParams -Completed
#endregion

#region ── Summary ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  SUMMARY" -ForegroundColor White
Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor DarkGray

$total       = $results.Count
$compliant   = ($results | Where-Object { $_.ComplianceState -eq 'Compliant' }).Count
$pendRst     = ($results | Where-Object { $_.ComplianceState -like '*Reboot*' -or $_.ComplianceState -like '*Restart*' }).Count
$pendInst    = ($results | Where-Object { $_.ComplianceState -like '*Install*' -or $_.ComplianceState -like '*Available*' }).Count
$offline     = ($results | Where-Object { $_.ComplianceState -eq 'Offline' }).Count
$errors      = ($results | Where-Object { $_.ComplianceState -like '*Error*' }).Count

Write-Host ("  Total machines   : {0}" -f $total)
Write-Host ("  Compliant        : {0}" -f $compliant)  -ForegroundColor Green
Write-Host ("  Pending Restart  : {0}" -f $pendRst)    -ForegroundColor Yellow
Write-Host ("  Pending Install  : {0}" -f $pendInst)   -ForegroundColor Yellow
Write-Host ("  Offline          : {0}" -f $offline)    -ForegroundColor DarkYellow
Write-Host ("  Errors           : {0}" -f $errors)     -ForegroundColor Red
Write-Host ""
#endregion

#region ── Export CSV ─────────────────────────────────────────────────────────
$timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFile     = Join-Path $OutputPath "SCCM_Audit_${KBArticle}_${timestamp}.csv"

$results |
    Select-Object `
        Machine,
        Online,
        ComplianceState,
        Downloaded,
        DownloadedSource,
        Installed,
        InstalledBy,
        Restarted,
        RestartType,
        LastHWInventory,
        ErrorCode,
        Notes |
    Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

if (Test-Path $csvFile) {
    Write-Host "  [+] CSV exported to:" -ForegroundColor Green
    Write-Host "      $csvFile" -ForegroundColor Cyan
    Write-Host ""

    # Optional: open CSV in default application
    $open = Read-Host "  Open CSV now? [Y/N]"
    if ($open -match '^[Yy]') {
        Start-Process $csvFile
    }
}
else {
    Write-Warning "  [!] CSV export failed. Check path: $OutputPath"
}

Write-Host ""
Write-Host "  Audit complete." -ForegroundColor Green
Write-Host ""
#endregion
