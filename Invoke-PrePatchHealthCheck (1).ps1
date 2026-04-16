#Requires -Version 7.0
<#
.SYNOPSIS
    Pre-Patch Health Check Script - MECM Collection Members
.DESCRIPTION
    Reads machines from a specified MECM deployment collection, then runs parallel
    health checks on each machine covering:
      - Online/reachability status
      - C: drive free space (flags < 20GB)
      - Pending reboot detection (multiple sources)
      - Windows Update blockers
      - MECM client service health
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
    .\Invoke-PrePatchHealthCheck.ps1 -SiteServer "SCCM01" -SiteCode "P01" -CollectionName "All Workstations - Patch Tuesday"
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

    [int]$ThrottleLimit = 50,

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
    Write-Host "  Started    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
    Write-Host ("═" * $width) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Summary {
    param($Results)
    $total   = $Results.Count
    $online  = ($Results | Where-Object { $_.Online -eq $true }).Count
    $offline = ($Results | Where-Object { $_.Online -eq $false }).Count
    $ready   = ($Results | Where-Object { $_.OverallStatus -eq "READY" }).Count
    $blocked = ($Results | Where-Object { $_.OverallStatus -eq "BLOCKED" }).Count
    $warning = ($Results | Where-Object { $_.OverallStatus -eq "WARNING" }).Count

    Write-Host ""
    Write-Host ("─" * 80) -ForegroundColor DarkGray
    Write-Host "  SUMMARY" -ForegroundColor Cyan
    Write-Host ("─" * 80) -ForegroundColor DarkGray
    Write-Host "  Total Machines : $total"
    Write-Host "  Online         : $online" -ForegroundColor Green
    Write-Host "  Offline        : $offline" -ForegroundColor DarkGray
    Write-Host "  Ready          : $ready"  -ForegroundColor Green
    Write-Host "  Blocked        : $blocked" -ForegroundColor Red
    Write-Host "  Warning        : $warning" -ForegroundColor Yellow
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
Write-Host "  Found $($MachineNames.Count) machines. Starting parallel health checks...`n" -ForegroundColor Green

#endregion

#region ── Parallel Health Check ───────────────────────────────────────────────

$Results = $MachineNames | ForEach-Object -Parallel {

    $MachineName     = $_
    $DiskThresholdGB = $using:DiskThresholdGB

    # ── Result object ─────────────────────────────────────────────────────────
    $Result = [PSCustomObject]@{
        MachineName       = $MachineName
        Online            = $false
        CDriveFreeGB      = $null
        DiskFlag          = $false
        PendingReboot     = $false
        RebootSources     = @()
        WUService         = "Unknown"
        CCMService        = "Unknown"
        WMIAccessible     = $false
        TrustedInstaller  = "Unknown"
        Issues            = @()
        OverallStatus     = "OFFLINE"
        CheckedAt         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }

    # ── 1. Connectivity check ─────────────────────────────────────────────────
    $Reachable = Test-Connection -ComputerName $MachineName -Count 1 -TimeoutSeconds 3 -Quiet 2>$null
    if (-not $Reachable) {
        $Result.Issues += "Machine offline/unreachable"
        return $Result
    }
    $Result.Online = $true

    # ── 2. Disk Space ─────────────────────────────────────────────────────────
    try {
        $Disk = Get-CimInstance -ClassName Win32_LogicalDisk `
                    -ComputerName $MachineName `
                    -Filter "DeviceID='C:'" `
                    -ErrorAction Stop
        $FreeGB = [math]::Round($Disk.FreeSpace / 1GB, 2)
        $Result.CDriveFreeGB = $FreeGB
        $Result.WMIAccessible = $true

        if ($FreeGB -lt $DiskThresholdGB) {
            $Result.DiskFlag = $true
            $Result.Issues  += "Low disk: $($FreeGB)GB free (threshold $($DiskThresholdGB)GB)"
        }
    } catch {
        $Result.Issues += "WMI disk query failed: $($_.Exception.Message)"
    }

    # ── 3. Pending Reboot Detection ───────────────────────────────────────────
    $RebootSources = [System.Collections.Generic.List[string]]::new()

    try {
        $RegChecks = Invoke-Command -ComputerName $MachineName -ErrorAction Stop -ScriptBlock {

            $sources = @()

            # CBS Reboot Pending
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
                $sources += "CBS"
            }

            # Windows Update Reboot Required
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
                $sources += "WindowsUpdate"
            }

            # Pending File Rename Operations
            $PFR = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                        -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue).PendingFileRenameOperations
            if ($PFR) { $sources += "PendingFileRename" }

            # Post Reboot Reporting
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting") {
                $sources += "PostRebootReporting"
            }

            # SCCM CCM Client reboot pending
            try {
                $CCMReboot = Invoke-CimMethod -Namespace "root\ccm\clientsdk" `
                                -ClassName "CCM_ClientUtilities" `
                                -MethodName "DetermineIfRebootPending" -ErrorAction Stop
                if ($CCMReboot.RebootPending -or $CCMReboot.IsHardRebootPending) {
                    $sources += "CCMClient"
                }
            } catch {}

            return $sources
        }

        foreach ($s in $RegChecks) { $RebootSources.Add($s) }

    } catch {
        $Result.Issues += "Registry/reboot check failed: $($_.Exception.Message)"
    }

    if ($RebootSources.Count -gt 0) {
        $Result.PendingReboot = $true
        $Result.RebootSources = $RebootSources
        $Result.Issues       += "Pending reboot: $($RebootSources -join ', ')"
    }

    # ── 4. Service Health ─────────────────────────────────────────────────────
    try {
        $Services = Get-CimInstance -ClassName Win32_Service `
                        -ComputerName $MachineName `
                        -Filter "Name='wuauserv' OR Name='CcmExec' OR Name='TrustedInstaller'" `
                        -ErrorAction Stop

        foreach ($Svc in $Services) {
            switch ($Svc.Name) {
                "wuauserv" {
                    $Result.WUService = $Svc.State
                    if ($Svc.State -ne "Running" -and $Svc.StartMode -ne "Disabled") {
                        # WU service is on-demand, only flag if StartMode is disabled
                    }
                    if ($Svc.StartMode -eq "Disabled") {
                        $Result.Issues += "Windows Update service is DISABLED"
                    }
                }
                "CcmExec" {
                    $Result.CCMService = $Svc.State
                    if ($Svc.State -ne "Running") {
                        $Result.Issues += "MECM CcmExec service is $($Svc.State)"
                    }
                }
                "TrustedInstaller" {
                    $Result.TrustedInstaller = $Svc.State
                }
            }
        }
    } catch {
        $Result.Issues += "Service query failed: $($_.Exception.Message)"
    }

    # ── 5. WMI Health ─────────────────────────────────────────────────────────
    if (-not $Result.WMIAccessible) {
        try {
            Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $MachineName -ErrorAction Stop | Out-Null
            $Result.WMIAccessible = $true
        } catch {
            $Result.Issues += "WMI not accessible"
        }
    }

    # ── 6. Determine Overall Status ───────────────────────────────────────────
    $BlockingIssues = $Result.Issues | Where-Object {
        $_ -match "Low disk|DISABLED|CcmExec|WMI not accessible|Pending reboot"
    }

    if ($BlockingIssues.Count -gt 0) {
        $Result.OverallStatus = "BLOCKED"
    } elseif ($Result.Issues.Count -gt 0) {
        $Result.OverallStatus = "WARNING"
    } else {
        $Result.OverallStatus = "READY"
    }

    return $Result

} -ThrottleLimit $ThrottleLimit

#endregion

#region ── Console Output ──────────────────────────────────────────────────────

Write-Host ("─" * 80) -ForegroundColor DarkGray
Write-Host ("  {0,-25} {1,-8} {2,-10} {3,-8} {4,-20} {5,-10} {6}" -f `
    "Machine","Online","C: Free GB","Reboot","Reboot Sources","CCM Svc","Status") -ForegroundColor Cyan
Write-Host ("─" * 80) -ForegroundColor DarkGray

foreach ($R in ($Results | Sort-Object OverallStatus, MachineName)) {

    $StatusColour = switch ($R.OverallStatus) {
        "READY"   { "Green" }
        "BLOCKED" { "Red" }
        "WARNING" { "Yellow" }
        default   { "DarkGray" }
    }

    $DiskDisplay   = if ($null -ne $R.CDriveFreeGB) { "$($R.CDriveFreeGB) GB" } else { "N/A" }
    $DiskColour    = if ($R.DiskFlag) { "Red" } elseif ($null -eq $R.CDriveFreeGB) { "DarkGray" } else { "Green" }
    $RebootDisplay = if ($R.PendingReboot) { "YES" } else { "No" }
    $RebootColour  = if ($R.PendingReboot) { "Red" } else { "Green" }
    $OnlineDisplay = if ($R.Online) { "Yes" } else { "No" }
    $OnlineColour  = if ($R.Online) { "Green" } else { "DarkGray" }
    $Sources       = ($R.RebootSources -join ",").PadRight(20).Substring(0,20)

    Write-Host ("  {0,-25}" -f $R.MachineName) -NoNewline
    Write-Host ("{0,-8}" -f $OnlineDisplay)     -NoNewline -ForegroundColor $OnlineColour
    Write-Host ("{0,-10}" -f $DiskDisplay)       -NoNewline -ForegroundColor $DiskColour
    Write-Host ("{0,-8}" -f $RebootDisplay)      -NoNewline -ForegroundColor $RebootColour
    Write-Host ("{0,-20}" -f $Sources)           -NoNewline -ForegroundColor $(if ($R.PendingReboot) {"Yellow"} else {"DarkGray"})
    Write-Host ("{0,-10}" -f $R.CCMService)      -NoNewline -ForegroundColor $(if ($R.CCMService -eq "Running") {"Green"} else {"Red"})
    Write-Host ("{0}" -f $R.OverallStatus)                  -ForegroundColor $StatusColour
}

Write-Summary -Results $Results

#endregion

#region ── CSV Export ──────────────────────────────────────────────────────────

$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvFile    = Join-Path $OutputPath "PrePatchHealthCheck_$Timestamp.csv"

$Results | Select-Object `
    MachineName,
    Online,
    CDriveFreeGB,
    DiskFlag,
    PendingReboot,
    @{N="RebootSources"; E={ $_.RebootSources -join " | " }},
    WUService,
    CCMService,
    TrustedInstaller,
    WMIAccessible,
    OverallStatus,
    @{N="Issues"; E={ $_.Issues -join " | " }},
    CheckedAt |
Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

Write-Host "  CSV exported to: $CsvFile" -ForegroundColor Cyan
Write-Host ""

#endregion
