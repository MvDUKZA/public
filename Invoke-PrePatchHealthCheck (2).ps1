#Requires -Version 7.0
<#
.SYNOPSIS
    Pre-Patch Health Check Script - MECM Collection Members
.DESCRIPTION
    Reads machines from a specified MECM deployment collection, then runs parallel
    health checks on each machine covering:
      - Online/reachability status
      - C: drive free space (flags < 20GB)
      - Pending reboot detection with HARD vs SOFT classification
      - PendingFileRenameOperations path validation (filters stale/orphaned entries)
      - Last reboot age in days (flags machines not rebooted in > 7 days)
      - Windows Update blockers and service health
      - MECM client service health
    Designed for VDI estates where stale reboot flags are common.
    Results are output to console with colour coding and exported to CSV.
.PARAMETER SiteServer
    MECM Site Server hostname (e.g. "SCCM01")
.PARAMETER SiteCode
    MECM Site Code (e.g. "P01")
.PARAMETER CollectionName
    Name of the MECM collection to target
.PARAMETER DiskThresholdGB
    Free space threshold in GB — machines below this are flagged (default: 20)
.PARAMETER ThrottleLimit
    Number of parallel threads (default: 50)
.PARAMETER OutputPath
    Path for CSV export (default: script directory)
.EXAMPLE
    .\Invoke-PrePatchHealthCheck.ps1 -SiteServer "SCCM01" -SiteCode "P01" -CollectionName "All VDI Machines"
.EXAMPLE
    .\Invoke-PrePatchHealthCheck.ps1 -SiteServer "SCCM01" -SiteCode "P01" -CollectionName "All VDI Machines" -ThrottleLimit 30 -TcpTimeoutMs 500
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteServer,

    [Parameter(Mandatory)]
    [string]$SiteCode,

    [Parameter(Mandatory)]
    [string]$CollectionName,

    [int]$DiskThresholdGB = 20,

    [int]$RebootAgeDaysThreshold = 7,

    # Stage 1 — TCP pre-filter throttle (high, cheap socket checks only)
    [int]$PingThrottleLimit = 200,

    # Stage 2 — full WMI/WinRM health check throttle (lower, expensive)
    [int]$ThrottleLimit = 50,

    # TCP port 5985 timeout in milliseconds — 500ms is plenty on a LAN
    [int]$TcpTimeoutMs = 500,

    [string]$OutputPath = $PSScriptRoot
)

#region ── Helpers ──────────────────────────────────────────────────────────────

function Write-Header {
    $width = 80
    Write-Host ""
    Write-Host ("═" * $width) -ForegroundColor Cyan
    Write-Host "  MECM Pre-Patch Health Check" -ForegroundColor Cyan
    Write-Host "  Collection : $CollectionName" -ForegroundColor White
    Write-Host "  Site       : $SiteCode on $SiteServer" -ForegroundColor White
    Write-Host "  Disk Flag  : < $DiskThresholdGB GB free" -ForegroundColor White
    Write-Host "  Reboot Age : > $RebootAgeDaysThreshold days flagged" -ForegroundColor White
    Write-Host "  TCP Timeout: $TcpTimeoutMs ms  |  Ping Throttle: $PingThrottleLimit  |  Scan Throttle: $ThrottleLimit" -ForegroundColor DarkGray
    Write-Host "  Started    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
    Write-Host ("═" * $width) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Summary {
    param($Results)
    $total        = $Results.Count
    $online       = ($Results | Where-Object { $_.Online -eq $true }).Count
    $offline      = ($Results | Where-Object { $_.Online -eq $false }).Count
    $ready        = ($Results | Where-Object { $_.OverallStatus -eq "READY" }).Count
    $blocked      = ($Results | Where-Object { $_.OverallStatus -eq "BLOCKED" }).Count
    $warning      = ($Results | Where-Object { $_.OverallStatus -eq "WARNING" }).Count
    $hardReboot   = ($Results | Where-Object { $_.RebootSeverity -eq "HARD" }).Count
    $softReboot   = ($Results | Where-Object { $_.RebootSeverity -eq "SOFT" }).Count
    $diskFlagged  = ($Results | Where-Object { $_.DiskFlag -eq $true }).Count
    $rebootAge    = ($Results | Where-Object { $_.RebootAgeFlag -eq $true }).Count

    Write-Host ""
    Write-Host ("─" * 80) -ForegroundColor DarkGray
    Write-Host "  SUMMARY" -ForegroundColor Cyan
    Write-Host ("─" * 80) -ForegroundColor DarkGray
    Write-Host "  Total Machines    : $total"
    Write-Host "  Online            : $online"       -ForegroundColor Green
    Write-Host "  Offline           : $offline"      -ForegroundColor DarkGray
    Write-Host "  Ready             : $ready"        -ForegroundColor Green
    Write-Host "  Blocked           : $blocked"      -ForegroundColor Red
    Write-Host "  Warning           : $warning"      -ForegroundColor Yellow
    Write-Host ("─" * 40) -ForegroundColor DarkGray
    Write-Host "  HARD Reboot       : $hardReboot  (CBS / real PendingFileRename)" -ForegroundColor Red
    Write-Host "  SOFT Reboot       : $softReboot  (WU / CCM / PostReboot — VDI-common)" -ForegroundColor Yellow
    Write-Host "  Low Disk          : $diskFlagged" -ForegroundColor $(if ($diskFlagged -gt 0) {"Red"} else {"Green"})
    Write-Host "  Reboot Age Flag   : $rebootAge   (not rebooted within threshold)" -ForegroundColor $(if ($rebootAge -gt 0) {"Yellow"} else {"Green"})
    Write-Host ("─" * 80) -ForegroundColor DarkGray
    Write-Host ""
}

#endregion

#region ── Load MECM Module & Get Collection Members ───────────────────────────

Write-Header

# NOTE: The ConfigMgr PS module was built for Windows PowerShell 5.1.
# In PowerShell 7, Get-CMCollectionMember can silently return nothing due to
# compatibility shim issues. We query WMI/CIM directly against the SMS Provider
# instead — this is reliable in PS7 and does not require the CM drive at all.

Write-Host "  Querying collection members via SMS Provider..." -ForegroundColor Cyan

try {
    # Step 1 — resolve CollectionID from the friendly name
    $CollectionQuery = Get-CimInstance -ComputerName $SiteServer `
                           -Namespace "root\SMS\site_$SiteCode" `
                           -ClassName "SMS_Collection" `
                           -Filter "Name = '$CollectionName'" `
                           -ErrorAction Stop

    if (-not $CollectionQuery) {
        Write-Error "Collection '$CollectionName' not found on $SiteServer (site $SiteCode). Check the name is exact."
        exit 1
    }

    $CollectionID = $CollectionQuery.CollectionID
    Write-Host "  Resolved CollectionID: $CollectionID" -ForegroundColor DarkGray

    # Step 2 — get all members of that collection
    $CollectionMembers = Get-CimInstance -ComputerName $SiteServer `
                             -Namespace "root\SMS\site_$SiteCode" `
                             -ClassName "SMS_CollectionMember_a" `
                             -Filter "CollectionID = '$CollectionID'" `
                             -ErrorAction Stop

} catch {
    Write-Error "Failed to query SMS Provider on '$SiteServer': $_"
    exit 1
}

if (-not $CollectionMembers -or $CollectionMembers.Count -eq 0) {
    Write-Warning "No members found in collection: $CollectionName (ID: $CollectionID)"
    exit 0
}

$MachineNames = $CollectionMembers | Select-Object -ExpandProperty Name | Sort-Object -Unique
Write-Host "  Found $($MachineNames.Count) machines. Running two-stage health check...`n" -ForegroundColor Green

#endregion

#region ── Stage 1: Fast TCP Pre-Filter ────────────────────────────────────────
# Run a lightweight TCP check on port 5985 (WinRM) against all machines at high
# parallelism. This quickly separates online from offline machines so the expensive
# WMI/WinRM stage only runs against machines that will actually respond.
# Offline machines still get a result object so they appear in the CSV/output.

Write-Host "  Stage 1/2 — TCP pre-filter (port 5985, ${TcpTimeoutMs}ms timeout)..." -ForegroundColor Cyan

$TcpResults = $MachineNames | ForEach-Object -Parallel {

    $MachineName  = $_
    $TcpTimeoutMs = $using:TcpTimeoutMs

    $Reachable = $false
    try {
        $TcpClient = [System.Net.Sockets.TcpClient]::new()
        $Connect   = $TcpClient.BeginConnect($MachineName, 5985, $null, $null)
        $Wait      = $Connect.AsyncWaitHandle.WaitOne($TcpTimeoutMs, $false)
        if ($Wait -and $TcpClient.Connected) { $Reachable = $true }
        $TcpClient.Close()
    } catch { $Reachable = $false }

    [PSCustomObject]@{
        MachineName = $MachineName
        Reachable   = $Reachable
    }

} -ThrottleLimit $PingThrottleLimit

$ReachableMachines = ($TcpResults | Where-Object { $_.Reachable }).MachineName
$OfflineMachines   = ($TcpResults | Where-Object { -not $_.Reachable }).MachineName

Write-Host "  Stage 1 complete — Reachable: $($ReachableMachines.Count)  |  Offline/No WinRM: $($OfflineMachines.Count)" -ForegroundColor $(if ($OfflineMachines.Count -gt 0) {"Yellow"} else {"Green"})
Write-Host ""

# Build offline result objects immediately — no further checks needed
$OfflineResults = $OfflineMachines | ForEach-Object {
    [PSCustomObject]@{
        MachineName       = $_
        Online            = $false
        CDriveFreeGB      = $null
        DiskFlag          = $false
        LastRebootDate    = $null
        DaysSinceReboot   = $null
        RebootAgeFlag     = $false
        PendingReboot     = $false
        RebootSeverity    = "NONE"
        HardRebootSources = @()
        SoftRebootSources = @()
        PFRTotalEntries   = 0
        PFRRealEntries    = 0
        WUService         = "Unknown"
        CCMService        = "Unknown"
        WMIAccessible     = $false
        TrustedInstaller  = "Unknown"
        HardIssues        = @("Machine offline/WinRM unreachable (port 5985)")
        SoftIssues        = @()
        OverallStatus     = "OFFLINE"
        CheckedAt         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

#endregion

#region ── Stage 2: Full Health Check (Reachable Machines Only) ─────────────────

Write-Host "  Stage 2/2 — Full health check on $($ReachableMachines.Count) reachable machines..." -ForegroundColor Cyan
Write-Host ""

# Use a synchronized hashtable for the progress counter
$TotalScan = $ReachableMachines.Count
$SyncHash  = [System.Collections.Hashtable]::Synchronized(@{ Count = 0 })

$OnlineResults = $ReachableMachines | ForEach-Object -Parallel {

    $MachineName            = $_
    $DiskThresholdGB        = $using:DiskThresholdGB
    $RebootAgeDaysThreshold = $using:RebootAgeDaysThreshold
    $SyncHash               = $using:SyncHash
    $TotalScan              = $using:TotalScan

    # ── Small jitter to avoid hitting VMs on same ESXi host simultaneously ────
    Start-Sleep -Milliseconds (Get-Random -Minimum 0 -Maximum 300)

    # ── Result object ─────────────────────────────────────────────────────────
    $Result = [PSCustomObject]@{
        MachineName       = $MachineName
        Online            = $true
        CDriveFreeGB      = $null
        DiskFlag          = $false
        LastRebootDate    = $null
        DaysSinceReboot   = $null
        RebootAgeFlag     = $false
        PendingReboot     = $false
        RebootSeverity    = "NONE"
        HardRebootSources = @()
        SoftRebootSources = @()
        PFRTotalEntries   = 0
        PFRRealEntries    = 0
        WUService         = "Unknown"
        CCMService        = "Unknown"
        WMIAccessible     = $false
        TrustedInstaller  = "Unknown"
        HardIssues        = @()
        SoftIssues        = @()
        OverallStatus     = "READY"
        CheckedAt         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }

    # ── CimSession with timeout — single session for all WMI queries ──────────
    $CimOpt  = New-CimSessionOption -Protocol Dcom
    $CimSess = $null
    try {
        $CimSess = New-CimSession -ComputerName $MachineName `
                       -SessionOption $CimOpt `
                       -OperationTimeoutSec 8 -ErrorAction Stop
    } catch {
        $Result.HardIssues += "CIM session failed: $($_.Exception.Message)"
    }

    if ($CimSess) {

        # ── OS info + last reboot age ─────────────────────────────────────────
        try {
            $OS = Get-CimInstance -CimSession $CimSess `
                      -ClassName Win32_OperatingSystem -ErrorAction Stop
            $Result.WMIAccessible   = $true
            $Result.LastRebootDate  = $OS.LastBootUpTime
            $DaysSince              = [math]::Round((New-TimeSpan -Start $OS.LastBootUpTime -End (Get-Date)).TotalDays, 1)
            $Result.DaysSinceReboot = $DaysSince
            if ($DaysSince -gt $RebootAgeDaysThreshold) {
                $Result.RebootAgeFlag = $true
                $Result.SoftIssues   += "No reboot in $($DaysSince) days (threshold $($RebootAgeDaysThreshold)d)"
            }
        } catch {
            $Result.HardIssues += "WMI OS query failed: $($_.Exception.Message)"
        }

        # ── Disk space ────────────────────────────────────────────────────────
        try {
            $Disk   = Get-CimInstance -CimSession $CimSess `
                          -ClassName Win32_LogicalDisk `
                          -Filter "DeviceID='C:'" -ErrorAction Stop
            $FreeGB = [math]::Round($Disk.FreeSpace / 1GB, 2)
            $Result.CDriveFreeGB = $FreeGB
            if ($FreeGB -lt $DiskThresholdGB) {
                $Result.DiskFlag    = $true
                $Result.HardIssues += "Low disk: $($FreeGB)GB free (threshold $($DiskThresholdGB)GB)"
            }
        } catch {
            $Result.HardIssues += "WMI disk query failed: $($_.Exception.Message)"
        }

        # ── Services ──────────────────────────────────────────────────────────
        try {
            $Services = Get-CimInstance -CimSession $CimSess `
                            -ClassName Win32_Service `
                            -Filter "Name='wuauserv' OR Name='CcmExec' OR Name='TrustedInstaller'" `
                            -ErrorAction Stop
            foreach ($Svc in $Services) {
                switch ($Svc.Name) {
                    "wuauserv" {
                        $Result.WUService = $Svc.State
                        if ($Svc.StartMode -eq "Disabled") {
                            $Result.HardIssues += "Windows Update service is DISABLED"
                        }
                    }
                    "CcmExec" {
                        $Result.CCMService = $Svc.State
                        if ($Svc.State -ne "Running") {
                            $Result.HardIssues += "MECM CcmExec service is $($Svc.State)"
                        }
                    }
                    "TrustedInstaller" { $Result.TrustedInstaller = $Svc.State }
                }
            }
        } catch {
            $Result.SoftIssues += "Service query failed: $($_.Exception.Message)"
        }

        Remove-CimSession -CimSession $CimSess -ErrorAction SilentlyContinue
    }

    # ── Pending Reboot (WinRM) ────────────────────────────────────────────────
    $HardSources = [System.Collections.Generic.List[string]]::new()
    $SoftSources = [System.Collections.Generic.List[string]]::new()

    try {
        $SessionOpt = New-PSSessionOption -OpenTimeout 3000 -OperationTimeout 8000
        $RebootData = Invoke-Command -ComputerName $MachineName `
                          -SessionOption $SessionOpt -ErrorAction Stop -ScriptBlock {

            $r = @{
                CBS                 = $false
                WindowsUpdate       = $false
                PostRebootReporting = $false
                CCMClient           = $false
                PFRTotalEntries     = 0
                PFRRealEntries      = 0
            }

            $r.CBS = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            $r.WindowsUpdate = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            $r.PostRebootReporting = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting"

            $PFR = (Get-ItemProperty `
                        -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                        -Name "PendingFileRenameOperations" `
                        -ErrorAction SilentlyContinue).PendingFileRenameOperations
            if ($PFR) {
                $SourcePaths        = $PFR | Where-Object { $_ -match '\S' }
                $r.PFRTotalEntries  = $SourcePaths.Count
                $RealPaths          = $SourcePaths | Where-Object {
                    try {
                        $Clean = $_ -replace '^\\\?\?\\','' -replace '^\\\?\\',''
                        Test-Path $Clean -ErrorAction SilentlyContinue
                    } catch { $false }
                }
                $r.PFRRealEntries = @($RealPaths).Count
            }

            try {
                $CCM = Invoke-CimMethod -Namespace "root\ccm\clientsdk" `
                           -ClassName "CCM_ClientUtilities" `
                           -MethodName "DetermineIfRebootPending" -ErrorAction Stop
                $r.CCMClient = ($CCM.RebootPending -or $CCM.IsHardRebootPending)
            } catch {}

            return $r
        }

        if ($RebootData.CBS)               { $HardSources.Add("CBS") }
        if ($RebootData.PFRRealEntries -gt 0) {
            $HardSources.Add("PendingFileRename($($RebootData.PFRRealEntries)/$($RebootData.PFRTotalEntries))")
        }
        if ($RebootData.WindowsUpdate)       { $SoftSources.Add("WindowsUpdate") }
        if ($RebootData.PostRebootReporting) { $SoftSources.Add("PostRebootReporting") }
        if ($RebootData.CCMClient)           { $SoftSources.Add("CCMClient") }

        $Result.PFRTotalEntries   = $RebootData.PFRTotalEntries
        $Result.PFRRealEntries    = $RebootData.PFRRealEntries
        $Result.HardRebootSources = $HardSources
        $Result.SoftRebootSources = $SoftSources

        if ($HardSources.Count -gt 0) {
            $Result.PendingReboot  = $true
            $Result.RebootSeverity = "HARD"
            $Result.HardIssues    += "HARD reboot pending: $($HardSources -join ', ')"
        } elseif ($SoftSources.Count -gt 0) {
            $Result.PendingReboot  = $true
            $Result.RebootSeverity = "SOFT"
            $Result.SoftIssues    += "Soft reboot flags (VDI-common): $($SoftSources -join ', ')"
        }

    } catch {
        $Result.SoftIssues += "Reboot check failed: $($_.Exception.Message)"
    }

    # ── Overall Status ────────────────────────────────────────────────────────
    if ($Result.HardIssues.Count -gt 0) {
        $Result.OverallStatus = "BLOCKED"
    } elseif ($Result.SoftIssues.Count -gt 0) {
        $Result.OverallStatus = "WARNING"
    } else {
        $Result.OverallStatus = "READY"
    }

    # ── Live progress counter ─────────────────────────────────────────────────
    $Done = [System.Threading.Interlocked]::Increment([ref]$SyncHash.Count)
    $Pct  = [math]::Round(($Done / $TotalScan) * 100)
    Write-Progress -Activity "Pre-Patch Health Check" `
                   -Status "Checked $Done of $TotalScan ($Pct%)" `
                   -PercentComplete $Pct

    return $Result

} -ThrottleLimit $ThrottleLimit

Write-Progress -Activity "Pre-Patch Health Check" -Completed

# Merge online + offline results
$Results = @($OnlineResults) + @($OfflineResults)

#endregion


#region ── Console Output ──────────────────────────────────────────────────────

$LineWidth = 110
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host ("  {0,-25} {1,-7} {2,-10} {3,-8} {4,-6} {5,-8} {6,-35} {7,-10} {8}" -f `
    "Machine","Online","C: Free","Days Up","Disk","Reboot","Reboot Detail","CCM Svc","Status") -ForegroundColor Cyan
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray

foreach ($R in ($Results | Sort-Object OverallStatus, MachineName)) {

    $StatusColour = switch ($R.OverallStatus) {
        "READY"   { "Green" }
        "BLOCKED" { "Red" }
        "WARNING" { "Yellow" }
        default   { "DarkGray" }
    }

    $OnlineDisplay  = if ($R.Online) { "Yes" } else { "No" }
    $OnlineColour   = if ($R.Online) { "Green" } else { "DarkGray" }

    $DiskDisplay    = if ($null -ne $R.CDriveFreeGB) { "$($R.CDriveFreeGB)GB" } else { "N/A" }
    $DiskColour     = if ($R.DiskFlag) { "Red" } elseif ($null -eq $R.CDriveFreeGB) { "DarkGray" } else { "Green" }

    $DaysDisplay    = if ($null -ne $R.DaysSinceReboot) { "$($R.DaysSinceReboot)d" } else { "N/A" }
    $DaysColour     = if ($R.RebootAgeFlag) { "Yellow" } elseif ($null -eq $R.DaysSinceReboot) { "DarkGray" } else { "Green" }

    $DiskFlag       = if ($R.DiskFlag) { "YES" } else { "No" }
    $DiskFlagColour = if ($R.DiskFlag) { "Red" } else { "DarkGray" }

    # Reboot severity display
    $RebootDisplay  = switch ($R.RebootSeverity) {
        "HARD" { "HARD" }
        "SOFT" { "soft" }
        default { "No" }
    }
    $RebootColour   = switch ($R.RebootSeverity) {
        "HARD" { "Red" }
        "SOFT" { "Yellow" }
        default { "Green" }
    }

    # Reboot detail — combine hard and soft sources concisely
    $AllSources = @()
    if ($R.HardRebootSources.Count -gt 0) { $AllSources += $R.HardRebootSources }
    if ($R.SoftRebootSources.Count -gt 0) { $AllSources += $R.SoftRebootSources }
    $SourceStr = ($AllSources -join ", ")
    if ($SourceStr.Length -gt 35) { $SourceStr = $SourceStr.Substring(0,32) + "..." }
    $SourceStr = $SourceStr.PadRight(35).Substring(0,35)

    $CCMColour = if ($R.CCMService -eq "Running") { "Green" } else { "Red" }

    Write-Host ("  {0,-25}" -f $R.MachineName)    -NoNewline
    Write-Host ("{0,-7}"  -f $OnlineDisplay)       -NoNewline -ForegroundColor $OnlineColour
    Write-Host ("{0,-10}" -f $DiskDisplay)         -NoNewline -ForegroundColor $DiskColour
    Write-Host ("{0,-8}"  -f $DaysDisplay)         -NoNewline -ForegroundColor $DaysColour
    Write-Host ("{0,-6}"  -f $DiskFlag)            -NoNewline -ForegroundColor $DiskFlagColour
    Write-Host ("{0,-8}"  -f $RebootDisplay)       -NoNewline -ForegroundColor $RebootColour
    Write-Host ("{0,-35}" -f $SourceStr)           -NoNewline -ForegroundColor $(if ($R.RebootSeverity -eq "HARD") {"Red"} elseif ($R.RebootSeverity -eq "SOFT") {"Yellow"} else {"DarkGray"})
    Write-Host ("{0,-10}" -f $R.CCMService)        -NoNewline -ForegroundColor $CCMColour
    Write-Host ("{0}"     -f $R.OverallStatus)                -ForegroundColor $StatusColour
}

Write-Summary -Results $Results

# ── Legend ────────────────────────────────────────────────────────────────────
Write-Host "  REBOOT LEGEND:" -ForegroundColor Cyan
Write-Host "  HARD  = CBS or PendingFileRename with real paths — likely to block patching" -ForegroundColor Red
Write-Host "  soft  = WindowsUpdate/CCMClient/PostRebootReporting — common on VDI, warn only" -ForegroundColor Yellow
Write-Host "  Days Up = days since last reboot — flagged if > $RebootAgeDaysThreshold days" -ForegroundColor DarkGray
Write-Host ""

#endregion

#region ── CSV Export ──────────────────────────────────────────────────────────

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvFile   = Join-Path $OutputPath "PrePatchHealthCheck_$Timestamp.csv"

$Results | Select-Object `
    MachineName,
    Online,
    CDriveFreeGB,
    DiskFlag,
    DaysSinceReboot,
    LastRebootDate,
    RebootAgeFlag,
    PendingReboot,
    RebootSeverity,
    @{N="HardRebootSources"; E={ $_.HardRebootSources -join " | " }},
    @{N="SoftRebootSources"; E={ $_.SoftRebootSources -join " | " }},
    PFRTotalEntries,
    PFRRealEntries,
    WUService,
    CCMService,
    TrustedInstaller,
    WMIAccessible,
    OverallStatus,
    @{N="HardIssues"; E={ $_.HardIssues -join " | " }},
    @{N="SoftIssues"; E={ $_.SoftIssues -join " | " }},
    CheckedAt |
Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

Write-Host "  CSV exported to: $CsvFile" -ForegroundColor Cyan
Write-Host ""

#endregion
