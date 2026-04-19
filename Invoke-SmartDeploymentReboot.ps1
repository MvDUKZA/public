#Requires -Version 7.0
<#
.SYNOPSIS
    Smart Deployment Reboot — pick a deployment, reboot failed/unknown unattended machines.
.DESCRIPTION
    1. Connects to MECM SMS Provider
    2. Shows all active deployments in Out-GridView — you pick one
    3. Pulls all Failed + Unknown machines from that deployment
    4. Shows them in Out-GridView — you confirm/deselect as needed
    5. Checks each machine for a logged-on user (parallel)
    6. Reboots unattended machines immediately
    7. Creates a timestamped retry list of machines with users logged on
    8. Exports a full audit log CSV
.PARAMETER SiteServer
    MECM Site Server hostname (e.g. "SCCM01")
.PARAMETER SiteCode
    MECM Site Code (e.g. "P01")
.PARAMETER ThrottleLimit
    Parallel threads for user check (default: 50)
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

#region ── Step 1: Pick a Deployment ──────────────────────────────────────────

Write-Host "  Step 1/4 — Loading deployments from MECM..." -ForegroundColor Cyan

try {
    # Get all software update deployments
    $Deployments = Get-CimInstance -ComputerName $SiteServer `
                       -Namespace "root\SMS\site_$SiteCode" `
                       -ClassName "SMS_DeploymentSummary" `
                       -Filter "FeatureType=5" `
                       -ErrorAction Stop |
                   Select-Object `
                       @{N="DeploymentID";       E={$_.DeploymentID}},
                       @{N="SoftwareName";        E={$_.SoftwareName}},
                       @{N="CollectionName";      E={$_.CollectionName}},
                       @{N="Compliant";           E={$_.NumberCompliant}},
                       @{N="NonCompliant";        E={$_.NumberNonCompliant}},
                       @{N="Failed";              E={$_.NumberErrors}},
                       @{N="Unknown";             E={$_.NumberUnknown}},
                       @{N="InProgress";          E={$_.NumberInProgress}},
                       @{N="TotalMachines";       E={$_.NumberTotal}},
                       @{N="CreationTime";        E={$_.CreationTime}} |
                   Sort-Object CreationTime -Descending

} catch {
    Write-Error "Failed to query deployments from $SiteServer : $_"
    exit 1
}

if (-not $Deployments -or $Deployments.Count -eq 0) {
    Write-Warning "No software update deployments found."
    exit 0
}

Write-Host "  Found $($Deployments.Count) deployments. Opening picker..." -ForegroundColor Green
Write-Host ""

# Show deployment picker — single select
$SelectedDeployment = $Deployments | Out-GridView `
    -Title "SELECT A DEPLOYMENT — Pick one deployment to target for reboots" `
    -OutputMode Single

if (-not $SelectedDeployment) {
    Write-Warning "No deployment selected. Exiting."
    exit 0
}

Write-Host "  Selected: $($SelectedDeployment.SoftwareName)" -ForegroundColor Green
Write-Host "  Collection: $($SelectedDeployment.CollectionName)" -ForegroundColor DarkGray
Write-Host "  Failed: $($SelectedDeployment.Failed)  |  Unknown: $($SelectedDeployment.Unknown)" -ForegroundColor DarkGray
Write-Host ""

#endregion

#region ── Step 2: Get Failed + Unknown Machines ──────────────────────────────

Write-Host "  Step 2/4 — Querying failed and unknown machines..." -ForegroundColor Cyan

try {
    # SMS_DeploymentInfo maps deployment to collection members with status
    # Status: 1=Success, 2=InProgress, 3=Unknown, 4=Error, 5=NotApplicable
    $DeploymentMembers = Get-CimInstance -ComputerName $SiteServer `
                             -Namespace "root\SMS\site_$SiteCode" `
                             -ClassName "SMS_DeploymentInfo" `
                             -Filter "DeploymentID='$($SelectedDeployment.DeploymentID)'" `
                             -ErrorAction Stop

    # Get compliance status for the deployment
    $ComplianceData = Get-CimInstance -ComputerName $SiteServer `
                          -Namespace "root\SMS\site_$SiteCode" `
                          -ClassName "SMS_UpdateComplianceStatus" `
                          -ErrorAction Stop |
                      Where-Object { $_.LastErrorCode -ne $null }

} catch {
    # Fallback — query collection members and cross reference with update status
    Write-Host "  Falling back to collection member query..." -ForegroundColor Yellow
}

# Primary approach — use SMS_ClientOperationStatus via deployment assignment
try {
    $AssignmentData = Get-CimInstance -ComputerName $SiteServer `
                          -Namespace "root\SMS\site_$SiteCode" `
                          -ClassName "SMS_UpdatesAssignment" `
                          -Filter "AssignmentID='$($SelectedDeployment.DeploymentID)'" `
                          -ErrorAction SilentlyContinue

    # Get all machines from the target collection
    $CollectionMembers = Get-CimInstance -ComputerName $SiteServer `
                             -Namespace "root\SMS\site_$SiteCode" `
                             -ClassName "SMS_CollectionMember_a" `
                             -Filter "CollectionID IN (SELECT CollectionID FROM SMS_Collection WHERE Name='$($SelectedDeployment.CollectionName)')" `
                             -ErrorAction Stop

    # Get compliance status — Status 0=Unknown, 1=NotRequired, 2=Required(missing), 3=Installed
    $StatusData = Get-CimInstance -ComputerName $SiteServer `
                      -Namespace "root\SMS\site_$SiteCode" `
                      -ClassName "SMS_UpdateComplianceStatus" `
                      -ErrorAction Stop

    # Build a lookup of machine name -> last error code
    $StatusLookup = @{}
    foreach ($S in $StatusData) {
        if ($S.MachineName -and -not $StatusLookup.ContainsKey($S.MachineName)) {
            $StatusLookup[$S.MachineName] = $S
        }
    }

    # Build machine list with status
    $MachineList = foreach ($Member in $CollectionMembers) {
        $Name   = $Member.Name
        $Status = $StatusLookup[$Name]

        [PSCustomObject]@{
            MachineName   = $Name
            Status        = if ($Status) {
                                switch ($Status.Status) {
                                    0 { "Unknown" }
                                    1 { "Not Required" }
                                    2 { "Required (Missing)" }
                                    3 { "Installed" }
                                    default { "Unknown" }
                                }
                            } else { "Unknown" }
            LastErrorCode = if ($Status.LastErrorCode -and $Status.LastErrorCode -ne 0) {
                                "0x{0:X8}" -f $Status.LastErrorCode
                            } else { "-" }
            LastStatusTime = if ($Status.LastStatusTime) { $Status.LastStatusTime } else { "-" }
        }
    }

} catch {
    Write-Error "Failed to query machine status: $_"
    exit 1
}

# Filter to only Failed and Unknown
$TargetMachines = $MachineList | Where-Object {
    $_.Status -in "Unknown","Required (Missing)" -or $_.LastErrorCode -ne "-"
}

if (-not $TargetMachines -or $TargetMachines.Count -eq 0) {
    Write-Host "  No failed or unknown machines found in this deployment." -ForegroundColor Green
    exit 0
}

Write-Host "  Found $($TargetMachines.Count) failed/unknown machines. Opening selection grid..." -ForegroundColor Yellow
Write-Host ""

#endregion

#region ── Step 3: Confirm Machine Selection ──────────────────────────────────

Write-Host "  Step 3/4 — Select machines to reboot (Ctrl+Click for multi-select)..." -ForegroundColor Cyan
Write-Host ""

$SelectedMachines = $TargetMachines | 
    Sort-Object Status, MachineName |
    Out-GridView `
        -Title "SELECT MACHINES TO REBOOT — Failed/Unknown machines from: $($SelectedDeployment.SoftwareName) | Ctrl+Click to multi-select | OK to confirm" `
        -OutputMode Multiple

if (-not $SelectedMachines -or $SelectedMachines.Count -eq 0) {
    Write-Warning "No machines selected. Exiting."
    exit 0
}

Write-Host "  $($SelectedMachines.Count) machines selected for reboot processing." -ForegroundColor Green
Write-Host ""

#endregion

#region ── Step 4: Check Users + Reboot ───────────────────────────────────────

Write-Host "  Step 4/4 — Checking for logged-on users and rebooting unattended machines..." -ForegroundColor Cyan
Write-Host ""

$MachineNames = $SelectedMachines.MachineName

$Results = $MachineNames | ForEach-Object -Parallel {

    $MachineName = $_

    $Result = [PSCustomObject]@{
        MachineName     = $MachineName
        Reachable       = $false
        LoggedOnUser    = $null
        UserPresent     = $false
        Rebooted        = $false
        RebootTime      = $null
        Outcome         = ""
        RetryNeeded     = $false
    }

    # ── Reachability ──────────────────────────────────────────────────────────
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

    # ── Check for logged-on user ──────────────────────────────────────────────
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

    # ── No user — reboot ──────────────────────────────────────────────────────
    try {
        $Reboot = Invoke-CimMethod -ClassName Win32_OperatingSystem `
                      -ComputerName $MachineName `
                      -MethodName Win32Shutdown `
                      -Arguments @{ Flags = 6 } `
                      -ErrorAction Stop

        if ($Reboot.ReturnValue -eq 0) {
            $Result.Rebooted    = $true
            $Result.RebootTime  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            $Result.Outcome     = "Rebooted successfully"
        } else {
            $Result.Outcome     = "Reboot returned code: $($Reboot.ReturnValue)"
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

$LineWidth = 80
Write-Host ""
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host ("  {0,-25} {1,-10} {2,-22} {3}" -f "Machine","Result","User","Detail") -ForegroundColor Cyan
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray

foreach ($R in ($Results | Sort-Object Rebooted -Descending)) {

    $Display = if ($R.Rebooted)           { "REBOOTED" }
               elseif (-not $R.Reachable) { "OFFLINE" }
               elseif ($R.UserPresent)    { "USER ON" }
               else                       { "FAILED" }

    $Colour  = if ($R.Rebooted)           { "Green" }
               elseif (-not $R.Reachable) { "DarkGray" }
               elseif ($R.UserPresent)    { "Yellow" }
               else                       { "Red" }

    $UserStr = if ($R.LoggedOnUser) { $R.LoggedOnUser } else { "-" }

    Write-Host ("  {0,-25}" -f $R.MachineName) -NoNewline
    Write-Host ("{0,-10}"   -f $Display)        -NoNewline -ForegroundColor $Colour
    Write-Host ("{0,-22}"   -f $UserStr)        -NoNewline -ForegroundColor $(if ($R.UserPresent) {"Yellow"} else {"DarkGray"})
    Write-Host ("{0}"       -f $R.Outcome)                 -ForegroundColor $Colour
}

# Summary
$Rebooted = ($Results | Where-Object { $_.Rebooted }).Count
$UserOn   = ($Results | Where-Object { $_.UserPresent }).Count
$Offline  = ($Results | Where-Object { -not $_.Reachable }).Count
$Failed   = ($Results | Where-Object { -not $_.Rebooted -and $_.Reachable -and -not $_.UserPresent }).Count
$Retry    = ($Results | Where-Object { $_.RetryNeeded }).Count

Write-Host ""
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host "  Rebooted           : $Rebooted"  -ForegroundColor Green
Write-Host "  Skipped (user on)  : $UserOn"    -ForegroundColor Yellow
Write-Host "  Offline            : $Offline"   -ForegroundColor DarkGray
Write-Host "  Failed             : $Failed"    -ForegroundColor $(if ($Failed -gt 0) {"Red"} else {"Green"})
Write-Host "  Retry list count   : $Retry"     -ForegroundColor $(if ($Retry -gt 0) {"Yellow"} else {"Green"})
Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
Write-Host ""

#endregion

#region ── Export Logs ────────────────────────────────────────────────────────

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$SafeName  = $SelectedDeployment.SoftwareName -replace '[\\/:*?"<>|]','-'

# Full audit log
$AuditLog  = Join-Path $LogPath "RebootAudit_${SafeName}_$Timestamp.csv"
$Results | Export-Csv -Path $AuditLog -NoTypeInformation -Encoding UTF8
Write-Host "  Audit log  : $AuditLog" -ForegroundColor Cyan

# Retry list — machines that need another attempt (user on, offline, failed)
$RetryMachines = $Results | Where-Object { $_.RetryNeeded }

if ($RetryMachines.Count -gt 0) {
    # Plain text list — one machine per line, ready to feed back into this script
    $RetryTxt  = Join-Path $LogPath "RetryList_${SafeName}_$Timestamp.txt"
    $RetryMachines.MachineName | Out-File -FilePath $RetryTxt -Encoding UTF8
    Write-Host "  Retry list : $RetryTxt  ($($RetryMachines.Count) machines)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  To retry later, run:" -ForegroundColor DarkGray
    Write-Host "  .\Invoke-SmartDeploymentReboot.ps1 -SiteServer '$SiteServer' -SiteCode '$SiteCode'" -ForegroundColor DarkGray
    Write-Host "  -- OR use the retry list with Invoke-UnattendedReboot.ps1:" -ForegroundColor DarkGray
    Write-Host "  .\Invoke-UnattendedReboot.ps1 -ComputerList '$RetryTxt'" -ForegroundColor DarkGray
} else {
    Write-Host "  No retry list needed — all machines processed successfully." -ForegroundColor Green
}

Write-Host ""

#endregion
