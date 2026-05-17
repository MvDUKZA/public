<#
.SYNOPSIS
    Remediates VDIs with failed or unknown software update deployments.

.DESCRIPTION
    1.  Lists all software update deployments via Get-CMDeployment.
        Select one or more in the grid view.
    2.  Queries failed/unknown assets by running WMI locally on the site
        server via Invoke-Command (avoids remote WMI permission issues).
    3.  Checks each VDI for logged-on users before rebooting.
    4.  Adds machines to 'VDI Maintenance Anytime' collection so the
        anytime maintenance window applies.
    5.  Reboots in batches (default 20, 5 min apart). Machines with a
        logged-on user are skipped from forced reboot and logged.
    6.  After each batch comes back online: triggers Machine Policy,
        Software Updates Scan, Deployment Evaluation via SCCM cmdlets
        and client-side TriggerSchedule.
    7.  Writes a full CSV report of every machine and every action taken.

.PARAMETER SiteCode
    SCCM site code. Default: PRD

.PARAMETER SiteServer
    SCCM site server FQDN. Default: appsmcm101fp.iprod.local

.PARAMETER MaintenanceCollectionName
    Collection granting the Anytime maintenance window.
    Default: 'VDI Maintenance Anytime'

.PARAMETER BatchSize
    Machines per reboot wave. Default: 20

.PARAMETER BatchIntervalMinutes
    Minutes between reboot waves. Default: 5

.PARAMETER IncludeUnknown
    Also remediate machines in Unknown state (not just Failed).
    Prompted interactively if not supplied.

.PARAMETER SkipLoggedOnCheck
    Reboot machines even if a user is logged on. Default: off (skip them).

.PARAMETER OnlineWaitMinutes
    How long to wait for a rebooted machine to come back up. Default: 15

.PARAMETER LogPath
    Output CSV path. Default: C:\Temp\VDIPatchRemediation_<timestamp>.csv

.EXAMPLE
    .\Invoke-VDIPatchRemediation.ps1 -WhatIf -Verbose
    .\Invoke-VDIPatchRemediation.ps1 -IncludeUnknown -BatchSize 10
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string] $SiteCode                  = 'PRD',
    [string] $SiteServer                = 'appsmcm101fp.iprod.local',
    [string] $MaintenanceCollectionName = 'VDI Maintenance Anytime',
    [int]    $BatchSize                 = 20,
    [int]    $BatchIntervalMinutes      = 5,
    [switch] $IncludeUnknown,
    [switch] $SkipLoggedOnCheck,
    [int]    $OnlineWaitMinutes         = 15,
    [string] $LogPath                   = "C:\Temp\VDIPatchRemediation_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# Logging
# ============================================================================

$script:Report = New-Object System.Collections.Generic.List[object]

function Add-Report {
    param(
        [string] $Computer,
        [string] $Action,
        [string] $Result,
        [string] $Detail = ''
    )
    $row = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Computer  = $Computer
        Action    = $Action
        Result    = $Result
        Detail    = $Detail
    }
    $script:Report.Add($row) | Out-Null
    Write-Verbose "[$Computer] $Action => $Result$(if($Detail){" | $Detail"})"
}

function Save-Report {
    $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
    $script:Report | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport saved: $LogPath" -ForegroundColor Green
}

# ============================================================================
# Connect to site
# ============================================================================

function Connect-CMSite {
    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw "SMS_ADMIN_UI_PATH not set. Is the ConfigMgr console installed on this machine?"
    }
    $module = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (-not (Get-Module ConfigurationManager)) {
        Write-Host "Importing ConfigurationManager module ..." -ForegroundColor Cyan
        Import-Module $module -ErrorAction Stop
    }
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
    }
    Set-Location "$SiteCode`:\"
    Write-Host "Connected to site $SiteCode on $SiteServer" -ForegroundColor Green
}

# ============================================================================
# 1. Deployment selection
# ============================================================================

function Select-Deployments {
    Write-Host "Loading software update deployments ..." -ForegroundColor Cyan

    $all = Get-CMDeployment -FeatureType SoftwareUpdate |
           Sort-Object DeploymentTime -Descending

    if (-not $all) { throw "No software update deployments found." }

    # Build a clean display object — keep the source object for later use
    $display = $all | Select-Object `
        @{N='Deployment Name'; E={$_.SoftwareName}},
        @{N='Target Collection'; E={$_.CollectionName}},
        @{N='Targeted'; E={$_.NumberTargeted}},
        @{N='Errors';   E={$_.NumberErrors}},
        @{N='Unknown';  E={$_.NumberUnknown}},
        @{N='Success';  E={$_.NumberSuccess}},
        @{N='Date';     E={$_.DeploymentTime}},
        @{N='DeploymentID'; E={$_.DeploymentID}},    # GUID
        @{N='AssignmentID'; E={$_.AssignmentID}}     # integer

    $picked = $display |
              Out-GridView -Title 'Select deployments to remediate  (Ctrl+Click for multiple)' `
                           -OutputMode Multiple

    if (-not $picked) { return $null }

    # Return the raw CM deployment objects for the selected ones
    $pickedIDs = $picked.DeploymentID
    $all | Where-Object { $_.DeploymentID -in $pickedIDs }
}

# ============================================================================
# 2. Get failed/unknown assets — runs WMI LOCALLY on site server
# ============================================================================

function Get-FailedAssets {
    param(
        [Parameter(Mandatory)] $Deployments,
        [bool] $IncludeUnknown
    )

    # StateType in SMS_UpdateComplianceStatus:
    #   5 = Error / Failed to install
    #   4 = Unknown
    $stateFilter = "StateType = 5"
    if ($IncludeUnknown) { $stateFilter += " OR StateType = 4" }

    $ns = "root\sms\site_$SiteCode"

    $allRows = foreach ($dep in $Deployments) {
        $assignmentID = $dep.AssignmentID   # integer e.g. 16777747
        $depName      = $dep.SoftwareName

        Write-Host "  Querying '$depName' (AssignmentID=$assignmentID) ..." -ForegroundColor Cyan

        # Run WMI locally on the site server via PSRemoting to avoid
        # remote WMI DCOM/permission issues (HRESULT 0x80041001)
        try {
            $rows = Invoke-Command -ComputerName $SiteServer -ScriptBlock {
                param($namespace, $filter, $aid)

                $query = "SELECT MachineID, StateType, LastStatusChangeTime
                          FROM SMS_UpdateComplianceStatus
                          WHERE AssignmentID = $aid AND ($filter)"

                # Get-WmiObject runs locally inside the Invoke-Command session
                $wmiRows = Get-WmiObject -Namespace $namespace -Query $query -ErrorAction Stop

                foreach ($r in $wmiRows) {
                    # Resolve MachineID to computer name
                    $sysQuery = "SELECT Name FROM SMS_R_System WHERE ResourceID = $($r.MachineID)"
                    $sys = Get-WmiObject -Namespace $namespace -Query $sysQuery -ErrorAction SilentlyContinue
                    if ($sys.Name) {
                        [pscustomobject]@{
                            Computer       = $sys.Name
                            StateType      = $r.StateType
                            LastStatusTime = $r.LastStatusChangeTime
                        }
                    }
                }
            } -ArgumentList $ns, $stateFilter, $assignmentID -ErrorAction Stop

            Write-Host ("    Found {0} asset(s)" -f @($rows).Count) -ForegroundColor $(if(@($rows).Count -gt 0){'Yellow'}else{'Green'})

            foreach ($r in $rows) {
                [pscustomobject]@{
                    Deployment     = $depName
                    AssignmentID   = $assignmentID
                    Computer       = $r.Computer
                    Status         = if ($r.StateType -eq 4) { 'Unknown' } else { 'Failed' }
                    LastStatusTime = $r.LastStatusTime
                }
            }
        }
        catch {
            Write-Warning "Could not query assets for '$depName': $($_.Exception.Message)"
        }
    }

    # Deduplicate — if a machine failed multiple deployments, keep the most recent row
    $allRows | Where-Object { $_.Computer } |
               Sort-Object Computer, LastStatusTime -Descending |
               Group-Object Computer |
               ForEach-Object { $_.Group | Select-Object -First 1 }
}

# ============================================================================
# 3. Check for logged-on users
# ============================================================================

function Test-UserLoggedOn {
    param([string] $Computer)

    try {
        $sessions = Get-CimInstance -ComputerName $Computer `
                                    -ClassName Win32_LogonSession `
                                    -Filter "LogonType = 2 OR LogonType = 10" `
                                    -ErrorAction Stop -OperationTimeoutSec 10
        return ($null -ne $sessions)
    }
    catch {
        # If we can't query, assume no user (VDIs are typically unattended)
        Write-Verbose "[$Computer] Could not check logon sessions: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# 4. Add to maintenance collection
# ============================================================================

function Add-ToMaintenanceCollection {
    param([string[]] $Computers)

    $coll = Get-CMDeviceCollection -Name $MaintenanceCollectionName -ErrorAction Stop
    if (-not $coll) { throw "Collection '$MaintenanceCollectionName' not found." }

    $existing = @(Get-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID |
                  Select-Object -ExpandProperty RuleName)

    foreach ($c in $Computers) {
        if ($c -in $existing) {
            Add-Report -Computer $c -Action 'AddToCollection' -Result 'AlreadyMember'
            continue
        }
        try {
            $device = Get-CMDevice -Name $c -Fast -ErrorAction Stop
            if (-not $device) {
                Add-Report -Computer $c -Action 'AddToCollection' -Result 'NotFoundInSCCM'
                continue
            }
            if ($PSCmdlet.ShouldProcess($c, "Add direct membership to '$MaintenanceCollectionName'")) {
                Add-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID `
                                                          -ResourceId $device.ResourceID `
                                                          -ErrorAction Stop
                Add-Report -Computer $c -Action 'AddToCollection' -Result 'Added'
            }
        }
        catch {
            Add-Report -Computer $c -Action 'AddToCollection' -Result 'Error' -Detail $_.Exception.Message
        }
    }

    if ($PSCmdlet.ShouldProcess($MaintenanceCollectionName, 'Refresh collection membership')) {
        Invoke-CMCollectionUpdate -CollectionId $coll.CollectionID -ErrorAction SilentlyContinue
        Write-Host "Collection update triggered. Waiting 30s for membership to propagate ..." -ForegroundColor Cyan
        Start-Sleep -Seconds 30
    }
}

# ============================================================================
# 5. Reboot a single VDI
# ============================================================================

function Invoke-VDIReboot {
    param([string] $Computer)

    # Try Restart-Computer (WinRM) first, fall back to shutdown.exe (RPC)
    try {
        Restart-Computer -ComputerName $Computer -Force -ErrorAction Stop
        Add-Report -Computer $Computer -Action 'Reboot' -Result 'Issued' -Detail 'Restart-Computer'
    }
    catch {
        try {
            $p = Start-Process shutdown.exe `
                 -ArgumentList "/m \\$Computer /r /f /t 5 /c `"SCCM patch remediation`"" `
                 -Wait -PassThru -NoNewWindow -ErrorAction Stop
            if ($p.ExitCode -eq 0) {
                Add-Report -Computer $Computer -Action 'Reboot' -Result 'Issued' -Detail 'shutdown.exe'
            }
            else {
                Add-Report -Computer $Computer -Action 'Reboot' -Result 'Failed' `
                           -Detail "shutdown.exe exit code $($p.ExitCode)"
            }
        }
        catch {
            Add-Report -Computer $Computer -Action 'Reboot' -Result 'Error' -Detail $_.Exception.Message
        }
    }
}

# ============================================================================
# 6. Post-reboot: wait online then trigger SCCM actions
# ============================================================================

function Invoke-PostRebootActions {
    param(
        [string] $Computer,
        [int]    $ResourceID
    )

    # Site-side fast-channel notification (cmdlet)
    try {
        Invoke-CMClientNotification -DeviceId $ResourceID `
                                    -NotificationType RequestMachinePolicyNow `
                                    -ErrorAction Stop
        Add-Report -Computer $Computer -Action 'PolicyNotification' -Result 'Sent'
    }
    catch {
        Add-Report -Computer $Computer -Action 'PolicyNotification' -Result 'Error' `
                   -Detail $_.Exception.Message
    }

    # Client-side schedule triggers (no cmdlet for this — CIM on the client)
    $schedules = [ordered]@{
        'Machine Policy Retrieval'     = '{00000000-0000-0000-0000-000000000021}'
        'Machine Policy Evaluation'    = '{00000000-0000-0000-0000-000000000022}'
        'Software Updates Scan'        = '{00000000-0000-0000-0000-000000000113}'
        'Software Updates Deploy Eval' = '{00000000-0000-0000-0000-000000000108}'
        'State Message Refresh'        = '{00000000-0000-0000-0000-000000000111}'
    }

    foreach ($action in $schedules.Keys) {
        try {
            Invoke-CimMethod -ComputerName $Computer `
                             -Namespace   'root\ccm' `
                             -ClassName   'SMS_Client' `
                             -MethodName  'TriggerSchedule' `
                             -Arguments   @{ sScheduleID = $schedules[$action] } `
                             -ErrorAction Stop | Out-Null
            Add-Report -Computer $Computer -Action $action -Result 'Triggered'
            Start-Sleep -Seconds 2
        }
        catch {
            Add-Report -Computer $Computer -Action $action -Result 'Error' -Detail $_.Exception.Message
        }
    }
}

function Wait-BatchOnline {
    param(
        [array] $Devices,      # objects with .Computer and .ResourceID
        [int]   $TimeoutMinutes
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    # Hashtable: computer -> ResourceID for machines still pending
    $pending  = @{}
    foreach ($d in $Devices) { $pending[$d.Computer] = $d.ResourceID }

    Write-Host "  Waiting up to ${TimeoutMinutes}m for batch to come back online ..." -ForegroundColor DarkGray

    while ($pending.Count -gt 0 -and (Get-Date) -lt $deadline) {
        foreach ($name in @($pending.Keys)) {
            if (Test-Connection -ComputerName $name -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                try {
                    # Confirm OS is up (not just ping)
                    $null = Get-CimInstance -ComputerName $name -ClassName Win32_OperatingSystem `
                                            -ErrorAction Stop -OperationTimeoutSec 10
                    Write-Host "  [$name] Online — triggering SCCM actions" -ForegroundColor Green
                    Invoke-PostRebootActions -Computer $name -ResourceID $pending[$name]
                    $pending.Remove($name) | Out-Null
                }
                catch { <# OS not ready yet, try again next loop #> }
            }
        }
        if ($pending.Count -gt 0) { Start-Sleep -Seconds 20 }
    }

    foreach ($name in $pending.Keys) {
        Write-Warning "[$name] Did not come back online within ${TimeoutMinutes} minutes."
        Add-Report -Computer $name -Action 'WaitOnline' -Result 'TimedOut'
    }
}

# ============================================================================
# Main
# ============================================================================

try {
    Connect-CMSite

    # ---- 1. Select deployments --------------------------------------------
    $selected = Select-Deployments
    if (-not $selected) {
        Write-Warning "No deployments selected. Exiting."
        return
    }
    Write-Host ("Selected {0} deployment(s)." -f @($selected).Count) -ForegroundColor Cyan

    # ---- 2. Include Unknown? ----------------------------------------------
    if (-not $PSBoundParameters.ContainsKey('IncludeUnknown')) {
        $ans = Read-Host "`nAlso include machines in 'Unknown' state? (Y/N)"
        [bool]$IncludeUnknown = $ans -match '^[Yy]'
    }

    # ---- 3. Get failed assets ---------------------------------------------
    Write-Host "`nQuerying failed$(if($IncludeUnknown){' + unknown'}) assets ..." -ForegroundColor Cyan
    $assets = Get-FailedAssets -Deployments $selected -IncludeUnknown $IncludeUnknown

    if (-not $assets) {
        Write-Warning "No failed or unknown assets found. Nothing to do."
        return
    }

    Write-Host ("`nFound {0} unique machine(s) to remediate:" -f @($assets).Count) -ForegroundColor Yellow
    $assets | Group-Object Status | ForEach-Object {
        Write-Host ("  {0,3} x {1}" -f $_.Count, $_.Name) -ForegroundColor White
    }

    # Log initial status for every machine found
    foreach ($a in $assets) {
        Add-Report -Computer $a.Computer -Action 'FoundInDeployment' -Result $a.Status `
                   -Detail "Deployment: $($a.Deployment)"
    }

    # ---- 4. Resolve device records ----------------------------------------
    Write-Host "`nResolving SCCM device records ..." -ForegroundColor Cyan
    $devices = foreach ($a in $assets) {
        $dev = Get-CMDevice -Name $a.Computer -Fast -ErrorAction SilentlyContinue
        if ($dev) {
            [pscustomobject]@{ Computer = $a.Computer; ResourceID = $dev.ResourceID }
        }
        else {
            Write-Warning "$($a.Computer) not found as SCCM device — skipping."
            Add-Report -Computer $a.Computer -Action 'ResolveDevice' -Result 'NotFoundInSCCM'
        }
    }
    $devices = @($devices | Where-Object { $_ })

    if ($devices.Count -eq 0) {
        Write-Warning "No machines resolved to SCCM device records. Exiting."
        Save-Report
        return
    }

    # ---- 5. Confirm -------------------------------------------------------
    Write-Host ""
    if (-not $PSCmdlet.ShouldContinue(
        "Add $($devices.Count) machine(s) to '$MaintenanceCollectionName', check for logged-on users, reboot in batches of $BatchSize every $BatchIntervalMinutes min, then trigger policy and updates?",
        'Confirm VDI Remediation')) {
        Write-Warning "Aborted."
        Save-Report
        return
    }

    # ---- 6. Add to maintenance collection --------------------------------
    Write-Host "`nAdding machines to '$MaintenanceCollectionName' ..." -ForegroundColor Cyan
    Add-ToMaintenanceCollection -Computers $devices.Computer

    # ---- 7. Check for logged-on users ------------------------------------
    $toReboot   = [System.Collections.Generic.List[object]]::new()
    $skipped    = [System.Collections.Generic.List[object]]::new()

    if (-not $SkipLoggedOnCheck) {
        Write-Host "`nChecking for logged-on users ..." -ForegroundColor Cyan
        foreach ($dev in $devices) {
            if (Test-Connection -ComputerName $dev.Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                if (Test-UserLoggedOn -Computer $dev.Computer) {
                    Write-Host "  [$($dev.Computer)] User logged on — SKIPPING reboot" -ForegroundColor Magenta
                    Add-Report -Computer $dev.Computer -Action 'LoggedOnCheck' -Result 'UserLoggedOn-Skipped'
                    $skipped.Add($dev)
                }
                else {
                    Add-Report -Computer $dev.Computer -Action 'LoggedOnCheck' -Result 'NoUserLoggedOn'
                    $toReboot.Add($dev)
                }
            }
            else {
                Write-Host "  [$($dev.Computer)] Offline — adding to reboot list anyway" -ForegroundColor DarkYellow
                Add-Report -Computer $dev.Computer -Action 'LoggedOnCheck' -Result 'Offline-AddedToReboot'
                $toReboot.Add($dev)
            }
        }
    }
    else {
        $toReboot.AddRange($devices)
    }

    Write-Host ("`n  {0} machine(s) queued for reboot" -f $toReboot.Count) -ForegroundColor Cyan
    Write-Host ("  {0} machine(s) skipped (user logged on)" -f $skipped.Count) -ForegroundColor Magenta

    if ($toReboot.Count -eq 0) {
        Write-Warning "No machines to reboot after logged-on check."
        Save-Report
        return
    }

    # ---- 8. Batched reboot + post-reboot triggers ------------------------
    $batchList = for ($i = 0; $i -lt $toReboot.Count; $i += $BatchSize) {
        ,@($toReboot[$i .. [Math]::Min($i + $BatchSize - 1, $toReboot.Count - 1)])
    }

    Write-Host ("`nProcessing {0} batch(es) of up to {1} machine(s)." -f $batchList.Count, $BatchSize) -ForegroundColor Cyan

    for ($b = 0; $b -lt $batchList.Count; $b++) {
        $batch = $batchList[$b]
        Write-Host ("`n=== Batch {0}/{1} — {2} machine(s) ===" -f ($b+1), $batchList.Count, $batch.Count) -ForegroundColor Yellow

        # Reboot each machine in the batch
        foreach ($dev in $batch) {
            if ($PSCmdlet.ShouldProcess($dev.Computer, 'Reboot')) {
                Write-Host "  Rebooting $($dev.Computer) ..." -ForegroundColor White
                Invoke-VDIReboot -Computer $dev.Computer
            }
        }

        # Wait for this batch to come back and trigger SCCM actions
        if ($PSCmdlet.ShouldProcess("Batch $($b+1)", 'Wait for online + trigger SCCM actions')) {
            Wait-BatchOnline -Devices $batch -TimeoutMinutes $OnlineWaitMinutes
        }

        # Pause before next batch
        if ($b -lt $batchList.Count - 1) {
            Write-Host ("`nWaiting {0} minute(s) before next batch ..." -f $BatchIntervalMinutes) -ForegroundColor DarkGray
            Start-Sleep -Seconds ($BatchIntervalMinutes * 60)
        }
    }

    # ---- 9. Final report --------------------------------------------------
    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host "REMEDIATION COMPLETE" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Cyan

    $script:Report | Group-Object Result | Sort-Object Name | ForEach-Object {
        Write-Host ("  {0,3}  {1}" -f $_.Count, $_.Name)
    }

    Save-Report
}
catch {
    Write-Error "Fatal error: $_"
    Save-Report
    throw
}
finally {
    Pop-Location -ErrorAction SilentlyContinue
}
