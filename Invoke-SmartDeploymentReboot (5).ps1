#Requires -Version 7.0
<#
.SYNOPSIS
    Smart Deployment Reboot — multi-deployment picker with deduplication.
.DESCRIPTION
    1. Connects to MECM SMS Provider
    2. Shows ALL software update deployments in Out-GridView — Ctrl+Click to pick multiple
    3. Pulls Failed + Unknown machines from ALL selected deployments
    4. Deduplicates — if a machine failed in multiple deployments it is
       rebooted ONCE but every deployment failure is logged against it
    5. Shows deduplicated list in Out-GridView for final confirmation
    6. Checks each machine for a logged-on user (parallel)
    7. Reboots unattended machines immediately
    8. Creates retry list (user on / offline / failed)
    9. Exports full audit CSV + retry detail CSV
.PARAMETER SiteServer
    MECM Site Server hostname (e.g. "SCCM01")
.PARAMETER SiteCode
    MECM Site Code (e.g. "P01")
.PARAMETER ThrottleLimit
    Parallel threads for user check / reboot (default: 50)
.PARAMETER LogPath
    Where to save logs (default: script directory)
.EXAMPLE
    .\Invoke-SmartDeploymentReboot.ps1 -SiteServer "SCCM01" -SiteCode "P01"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SiteServer,

    [Parameter(Mandatory)]
    [string]$SiteCode,

    [int]$ThrottleLimit = 50,

    [string]$LogPath = $PSScriptRoot
)

#region ── Header ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor Cyan
Write-Host "  MECM Smart Deployment Reboot" -ForegroundColor Cyan
Write-Host "  Site Server : $SiteServer" -ForegroundColor White
Write-Host "  Site Code   : $SiteCode" -ForegroundColor White
Write-Host "  Started     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host ("═" * 70) -ForegroundColor Cyan
Write-Host ""

#endregion

#region ── Load ConfigMgr Module ──────────────────────────────────────────────

# The CM module was built for Windows PowerShell 5.1 but works in PS7 via the
# compatibility shim. We only use it for Get-CMDeploymentStatus which reliably
# returns per-machine per-deployment status — bypassing the compat issues that
# affect other cmdlets like Get-CMCollectionMember.

$ModulePath = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
if (-not (Test-Path $ModulePath)) {
    Write-Error "ConfigMgr PS module not found. Install the MECM console on this machine."
    Write-Error "Expected path: $ModulePath"
    exit 1
}

Import-Module $ModulePath -ErrorAction Stop

if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
}

# Remember original location and switch to CM drive
$OriginalLocation = Get-Location
Set-Location "$($SiteCode):\" -ErrorAction Stop

#endregion

#region ── Step 1: Pick Deployments (multi-select) ────────────────────────────

Write-Host "  Step 1/4 — Loading deployments from MECM..." -ForegroundColor Cyan

try {
    $AllDeployments = Get-CimInstance -ComputerName $SiteServer `
                          -Namespace "root\SMS\site_$SiteCode" `
                          -ClassName "SMS_DeploymentSummary" `
                          -ErrorAction Stop

    Write-Host "  Total deployments (all types) : $($AllDeployments.Count)" -ForegroundColor DarkGray

    $SWUDeployments = $AllDeployments | Where-Object { $_.FeatureType -eq 5 }
    Write-Host "  Software Update deployments   : $($SWUDeployments.Count)" -ForegroundColor DarkGray

    if ($SWUDeployments.Count -eq 0) {
        Write-Host "  No FeatureType=5 found — showing all types" -ForegroundColor Yellow
        $SWUDeployments = $AllDeployments
    }

    $DeploymentList = $SWUDeployments |
                      Select-Object `
                          @{N="DeploymentID";   E={$_.DeploymentID}},
                          @{N="SoftwareName";   E={$_.SoftwareName}},
                          @{N="CollectionName"; E={$_.CollectionName}},
                          @{N="Failed";         E={$_.NumberErrors}},
                          @{N="Unknown";        E={$_.NumberUnknown}},
                          @{N="Compliant";      E={$_.NumberCompliant}},
                          @{N="InProgress";     E={$_.NumberInProgress}},
                          @{N="Total";          E={$_.NumberTotal}},
                          @{N="CreationTime";   E={$_.CreationTime}} |
                      Sort-Object CreationTime -Descending

} catch {
    Write-Error "Failed to query deployments: $_"
    exit 1
}

if (-not $DeploymentList -or $DeploymentList.Count -eq 0) {
    Write-Warning "No deployments found. Check SiteServer and SiteCode."
    exit 0
}

Write-Host "  Opening picker ($($DeploymentList.Count) deployments)..." -ForegroundColor Green
Write-Host "  Tip: Ctrl+Click to select multiple deployments" -ForegroundColor DarkGray
Write-Host ""

$SelectedDeployments = $DeploymentList | Out-GridView `
    -Title "STEP 1 — SELECT DEPLOYMENTS  |  Ctrl+Click to pick multiple  |  Failed+Unknown machines will be combined and deduplicated" `
    -OutputMode Multiple

if (-not $SelectedDeployments -or $SelectedDeployments.Count -eq 0) {
    Write-Warning "No deployments selected. Exiting."
    exit 0
}

Write-Host "  Selected $($SelectedDeployments.Count) deployment(s):" -ForegroundColor Green
foreach ($D in $SelectedDeployments) {
    Write-Host "    → $($D.SoftwareName)  [$($D.CollectionName)]  Failed:$($D.Failed)  Unknown:$($D.Unknown)" -ForegroundColor DarkGray
}
Write-Host ""

#endregion

#region ── Step 2: Query + Deduplicate Failed/Unknown Machines ────────────────

Write-Host "  Step 2/4 — Querying failed/unknown machines across all selected deployments..." -ForegroundColor Cyan

# MachineName -> list of deployment names it failed in
$MachineDeploymentMap = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
# MachineName -> list of "DeploymentName [ErrorCode]" strings
$MachineErrorMap = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($Dep in $SelectedDeployments) {

    Write-Host "    Querying: $($Dep.SoftwareName)..." -ForegroundColor DarkGray

    try {
        # Get-CMDeploymentStatus returns per-machine status for a deployment.
        # StatusType values:
        #   1 = Success
        #   2 = In Progress
        #   3 = Requirements not met / Unknown
        #   4 = Error / Failed
        #   5 = Not applicable
        #
        # We only want 3 (Unknown) and 4 (Failed).

        $AllStatus = Get-CMDeploymentStatus -DeploymentId $Dep.DeploymentID -ErrorAction Stop

        # Filter to Failed or Unknown only — excludes Success, InProgress, NotApplicable
        $FailedOrUnknown = $AllStatus | Where-Object { $_.StatusType -in 3, 4 }

        if (-not $FailedOrUnknown -or $FailedOrUnknown.Count -eq 0) {
            Write-Host "      0 failed/unknown machines in this deployment" -ForegroundColor DarkGray
            continue
        }

        $FailCount = 0
        foreach ($Entry in $FailedOrUnknown) {
            $Name       = $Entry.DeviceName
            $StatusType = $Entry.StatusType

            if (-not $Name) { continue }

            $ErrorStr = switch ($StatusType) {
                3       { "Unknown" }
                4       {
                    if ($Entry.StatusDescription) {
                        "Failed: $($Entry.StatusDescription)"
                    } else {
                        "Failed"
                    }
                }
                default { "Status $StatusType" }
            }

            # Add to maps
            if (-not $MachineDeploymentMap.ContainsKey($Name)) {
                $MachineDeploymentMap[$Name] = [System.Collections.Generic.List[string]]::new()
                $MachineErrorMap[$Name]      = [System.Collections.Generic.List[string]]::new()
            }

            if (-not $MachineDeploymentMap[$Name].Contains($Dep.SoftwareName)) {
                $MachineDeploymentMap[$Name].Add($Dep.SoftwareName)
            }
            $MachineErrorMap[$Name].Add("$($Dep.SoftwareName) [$ErrorStr]")
            $FailCount++
        }

        Write-Host "      $FailCount failed/unknown machines in this deployment" -ForegroundColor DarkGray

    } catch {
        Write-Host "      ERROR querying $($Dep.SoftwareName): $_" -ForegroundColor Red
    }
}

Write-Host ""

if ($MachineDeploymentMap.Count -eq 0) {
    Write-Host "  No failed or unknown machines found across selected deployments." -ForegroundColor Green
    Set-Location $OriginalLocation
    exit 0
}

# Build deduplicated GridView list
$DeduplicatedRaw = foreach ($KV in $MachineDeploymentMap.GetEnumerator()) {
    $Name    = $KV.Key
    $DepList = $KV.Value

    [PSCustomObject]@{
        MachineName         = $Name
        FailedInCount       = $DepList.Count
        DuplicateFlag       = if ($DepList.Count -gt 1) { "YES" } else { "" }
        FailedInDeployments = $DepList -join " | "
        ErrorDetail         = $MachineErrorMap[$Name] -join " | "
    }
}

$DeduplicatedList = $DeduplicatedRaw | Sort-Object -Property @{Expression="FailedInCount";Descending=$true}, @{Expression="MachineName";Descending=$false}

$TotalDupes = ($DeduplicatedList | Where-Object { $_.FailedInCount -gt 1 }).Count

Write-Host "  Unique machines across all deployments : $($DeduplicatedList.Count)" -ForegroundColor Green
Write-Host "  Machines in multiple deployments       : $TotalDupes" -ForegroundColor $(if ($TotalDupes -gt 0) {"Yellow"} else {"DarkGray"})
if ($TotalDupes -gt 0) {
    Write-Host "  Each will be rebooted ONCE regardless of how many deployments they failed in." -ForegroundColor Yellow
}
Write-Host ""

#endregion

#region ── Step 3: Confirm Machine Selection ──────────────────────────────────

Write-Host "  Step 3/4 — Confirm machines to reboot..." -ForegroundColor Cyan
Write-Host "  Machines with DuplicateFlag=YES failed in multiple deployments" -ForegroundColor DarkGray
Write-Host ""

$SelectedMachines = $DeduplicatedList | Out-GridView `
    -Title "STEP 3 — CONFIRM MACHINES TO REBOOT  |  $($DeduplicatedList.Count) unique machines  |  DuplicateFlag=YES = failed in multiple deployments  |  Ctrl+Click to deselect  |  OK to reboot" `
    -OutputMode Multiple

if (-not $SelectedMachines -or $SelectedMachines.Count -eq 0) {
    Write-Warning "No machines selected. Exiting."
    exit 0
}

Write-Host "  $($SelectedMachines.Count) machines confirmed." -ForegroundColor Green
Write-Host ""

#endregion

#region ── Step 4: Check Users + Reboot ───────────────────────────────────────

Write-Host "  Step 4/4 — Checking logged-on users and rebooting unattended machines..." -ForegroundColor Cyan
Write-Host ""

# Switch back to original filesystem location before parallel operations.
# ForEach-Object -Parallel cannot run while current location is a CM drive.
Set-Location $OriginalLocation

$Results = $SelectedMachines | ForEach-Object -Parallel {

    $Machine     = $_
    $MachineName = $Machine.MachineName

    $Result = [PSCustomObject]@{
        MachineName         = $MachineName
        FailedInCount       = $Machine.FailedInCount
        DuplicateFlag       = $Machine.DuplicateFlag
        FailedInDeployments = $Machine.FailedInDeployments
        ErrorDetail         = $Machine.ErrorDetail
        Reachable           = $false
        LoggedOnUser        = $null
        UserPresent         = $false
        Rebooted            = $false
        RebootTime          = $null
        Outcome             = ""
        RetryNeeded         = $false
    }

    # ── Reachability (TCP 135) ────────────────────────────────────────────────
    try {
        $Tcp     = [System.Net.Sockets.TcpClient]::new()
        $Connect = $Tcp.BeginConnect($MachineName, 135, $null, $null)
        $Wait    = $Connect.AsyncWaitHandle.WaitOne(1000, $false)
        if ($Wait -and $Tcp.Connected) { $Result.Reachable = $true }
        $Tcp.Close()
    } catch {}

    if (-not $Result.Reachable) {
        $Result.Outcome     = "Offline / unreachable"
        $Result.RetryNeeded = $true
        return $Result
    }

    # ── Logged-on user check ──────────────────────────────────────────────────
    try {
        $CS   = Get-CimInstance -ClassName Win32_ComputerSystem `
                    -ComputerName $MachineName -ErrorAction Stop
        $User = $CS.UserName

        if ($User -and $User -match '\S') {
            $Result.LoggedOnUser = $User
            $Result.UserPresent  = $true
            $Result.Outcome      = "User logged on: $User — added to retry list"
            $Result.RetryNeeded  = $true
            return $Result
        }
    } catch {
        $Result.Outcome     = "WMI failed: $($_.Exception.Message)"
        $Result.RetryNeeded = $true
        return $Result
    }

    # ── No user present — reboot ──────────────────────────────────────────────
    try {
        $Reboot = Invoke-CimMethod -ClassName Win32_OperatingSystem `
                      -ComputerName $MachineName `
                      -MethodName Win32Shutdown `
                      -Arguments @{ Flags = 6 } `
                      -ErrorAction Stop

        if ($Reboot.ReturnValue -eq 0) {
            $Result.Rebooted   = $true
            $Result.RebootTime = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            $Result.Outcome    = "Rebooted successfully"
        } else {
            $Result.Outcome     = "Reboot call returned: $($Reboot.ReturnValue)"
            $Result.RetryNeeded = $true
        }
    } catch {
        $Result.Outcome     = "Reboot failed: $($_.Exception.Message)"
        $Result.RetryNeeded = $true
    }

    return $Result

} -ThrottleLimit $ThrottleLimit

#endregion

#region ── Console Output ─────────────────────────────────────────────────────

$LineWidth = 100
Write-Host ""
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host ("  {0,-25} {1,-10} {2,-6} {3,-22} {4}" -f "Machine","Result","Deps","User","Detail") -ForegroundColor Cyan
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray

foreach ($R in ($Results | Sort-Object Rebooted -Descending)) {

    $Display = if ($R.Rebooted)           { "REBOOTED" }
               elseif (-not $R.Reachable) { "OFFLINE"  }
               elseif ($R.UserPresent)    { "USER ON"  }
               else                       { "FAILED"   }

    $Colour  = if ($R.Rebooted)           { "Green"    }
               elseif (-not $R.Reachable) { "DarkGray" }
               elseif ($R.UserPresent)    { "Yellow"   }
               else                       { "Red"      }

    $DepsStr = if ($R.FailedInCount -gt 1) { "⚠ $($R.FailedInCount)" } else { "  1" }
    $UserStr = if ($R.LoggedOnUser) { $R.LoggedOnUser } else { "-" }

    Write-Host ("  {0,-25}" -f $R.MachineName)  -NoNewline
    Write-Host ("{0,-10}"   -f $Display)          -NoNewline -ForegroundColor $Colour
    Write-Host ("{0,-6}"    -f $DepsStr)          -NoNewline -ForegroundColor $(if ($R.FailedInCount -gt 1) {"Yellow"} else {"DarkGray"})
    Write-Host ("{0,-22}"   -f $UserStr)          -NoNewline -ForegroundColor $(if ($R.UserPresent) {"Yellow"} else {"DarkGray"})
    Write-Host ("{0}"       -f $R.Outcome)                   -ForegroundColor $Colour
}

$Rebooted = ($Results | Where-Object { $_.Rebooted }).Count
$UserOn   = ($Results | Where-Object { $_.UserPresent }).Count
$Offline  = ($Results | Where-Object { -not $_.Reachable }).Count
$Failed   = ($Results | Where-Object { -not $_.Rebooted -and $_.Reachable -and -not $_.UserPresent }).Count
$Retry    = ($Results | Where-Object { $_.RetryNeeded }).Count
$MultiDep = ($Results | Where-Object { $_.FailedInCount -gt 1 }).Count

Write-Host ""
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host "  Deployments selected       : $($SelectedDeployments.Count)"  -ForegroundColor Cyan
Write-Host "  Rebooted                   : $Rebooted"   -ForegroundColor Green
Write-Host "  Skipped (user logged on)   : $UserOn"     -ForegroundColor Yellow
Write-Host "  Offline                    : $Offline"    -ForegroundColor DarkGray
Write-Host "  Failed to reboot           : $Failed"     -ForegroundColor $(if ($Failed -gt 0) {"Red"} else {"Green"})
Write-Host "  Retry list                 : $Retry"      -ForegroundColor $(if ($Retry -gt 0) {"Yellow"} else {"Green"})
Write-Host "  Multi-deployment machines  : $MultiDep"   -ForegroundColor $(if ($MultiDep -gt 0) {"Yellow"} else {"Green"})
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host ""

#endregion

#region ── Export Logs ────────────────────────────────────────────────────────

$Timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$DepSafeNames = ($SelectedDeployments | Select-Object -First 2 -ExpandProperty SoftwareName |
                 ForEach-Object { $_ -replace '[\\/:*?"<>|]','-' }) -join "_"
if ($SelectedDeployments.Count -gt 2) {
    $DepSafeNames += "_plus$($SelectedDeployments.Count - 2)more"
}

# Full audit log
$AuditLog = Join-Path $LogPath "RebootAudit_${DepSafeNames}_$Timestamp.csv"
$Results | Export-Csv -Path $AuditLog -NoTypeInformation -Encoding UTF8
Write-Host "  Audit log    : $AuditLog" -ForegroundColor Cyan

# Retry list
$RetryMachines = $Results | Where-Object { $_.RetryNeeded }

if ($RetryMachines.Count -gt 0) {

    # Plain text — ready to pipe into Invoke-UnattendedReboot.ps1
    $RetryTxt = Join-Path $LogPath "RetryList_${DepSafeNames}_$Timestamp.txt"
    $RetryMachines.MachineName | Out-File -FilePath $RetryTxt -Encoding UTF8
    Write-Host "  Retry list   : $RetryTxt  ($($RetryMachines.Count) machines)" -ForegroundColor Yellow

    # Retry detail CSV — why each machine needs retrying + which deployments it failed in
    $RetryDetailCsv = Join-Path $LogPath "RetryDetail_${DepSafeNames}_$Timestamp.csv"
    $RetryMachines |
        Select-Object MachineName, FailedInCount, DuplicateFlag, FailedInDeployments, ErrorDetail, LoggedOnUser, Outcome |
        Export-Csv -Path $RetryDetailCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  Retry detail : $RetryDetailCsv" -ForegroundColor Yellow

    Write-Host ""
    Write-Host "  To retry skipped machines:" -ForegroundColor DarkGray
    Write-Host "  .\Invoke-UnattendedReboot.ps1 -ComputerList '$RetryTxt'" -ForegroundColor DarkGray

} else {
    Write-Host "  No retry list needed — all machines processed successfully." -ForegroundColor Green
}

Write-Host ""

#endregion
