<#
.SYNOPSIS
    Remediates VDIs with failed or unknown software update deployments.

.DESCRIPTION
    Correct cmdlet chain per Microsoft documentation:
      Get-CMDeployment
        -> Get-CMSoftwareUpdateDeployment -DeploymentId
          -> Get-CMSoftwareUpdateDeploymentStatus -InputObject
            -> Get-CMDeploymentStatusDetails -InputObject (filtered by StatusType)

    Then for each failed machine:
      - Checks for logged-on users before rebooting
      - Adds to 'VDI Maintenance Anytime' collection
      - Reboots in configurable batches
      - Waits for machines to come back online
      - Triggers Machine Policy + Software Updates Scan/Eval
      - Writes a full CSV report

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
    Also remediate Unknown state machines, not just Failed.
    Prompted interactively if not supplied.

.PARAMETER SkipLoggedOnCheck
    Reboot even if a user is logged on.

.PARAMETER OnlineWaitMinutes
    How long to wait per batch for machines to come back. Default: 15

.PARAMETER LogPath
    CSV report output path.

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
# Logging — every action on every machine ends up in this list -> CSV
# ============================================================================

$script:Report = New-Object System.Collections.Generic.List[object]

function Add-Report {
    param(
        [string] $Computer,
        [string] $Action,
        [string] $Result,
        [string] $Detail = ''
    )
    $script:Report.Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Computer  = $Computer
        Action    = $Action
        Result    = $Result
        Detail    = $Detail
    }) | Out-Null
    Write-Verbose "[$Computer] $Action => $Result$(if ($Detail) { " | $Detail" })"
}

function Save-Report {
    $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
    $script:Report | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport saved to: $LogPath" -ForegroundColor Green

    # Print summary to console
    Write-Host "`n--- Summary ---" -ForegroundColor Cyan
    $script:Report |
        Group-Object Action, Result |
        Sort-Object Name |
        ForEach-Object { Write-Host ("  {0,3}  {1}" -f $_.Count, $_.Name) }
}

# ============================================================================
# Site connection
# ============================================================================

function Connect-CMSite {
    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw "SMS_ADMIN_UI_PATH not set. Is the ConfigMgr console installed?"
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

    # Get-CMDeployment with FeatureType SoftwareUpdate returns SMS_DeploymentSummary rows.
    # We need the DeploymentID (GUID) to feed into Get-CMSoftwareUpdateDeployment next.
    $all = Get-CMDeployment -FeatureType SoftwareUpdate | Sort-Object DeploymentTime -Descending

    if (-not $all) { throw "No software update deployments found." }

    $display = $all | Select-Object `
        @{N='Deployment Name';  E={$_.SoftwareName}},
        @{N='Target Collection';E={$_.CollectionName}},
        @{N='Targeted';         E={$_.NumberTargeted}},
        @{N='Errors';           E={$_.NumberErrors}},
        @{N='Unknown';          E={$_.NumberUnknown}},
        @{N='Success';          E={$_.NumberSuccess}},
        @{N='Date';             E={$_.DeploymentTime}},
        @{N='DeploymentID';     E={$_.DeploymentID}}   # GUID, passed to Get-CMSoftwareUpdateDeployment

    $picked = $display |
              Out-GridView -Title 'Select deployments to remediate  (Ctrl+Click for multiple)' `
                           -OutputMode Multiple

    if (-not $picked) { return $null }

    # Return the full raw objects for the selected DeploymentIDs
    $pickedIDs = @($picked.DeploymentID)
    @($all | Where-Object { $_.DeploymentID -in $pickedIDs })
}

# ============================================================================
# 2. Get failed/unknown assets
#    Documented cmdlet chain:
#      Get-CMSoftwareUpdateDeployment -DeploymentId <GUID>
#        | Get-CMSoftwareUpdateDeploymentStatus
#          | Get-CMDeploymentStatusDetails  (filter StatusType 4=Unknown, 5=Error)
# ============================================================================

function Get-FailedAssets {
    param(
        [Parameter(Mandatory)] [array] $Deployments,
        [bool]                         $IncludeUnknown
    )

    # StatusType values in Get-CMDeploymentStatusDetails:
    #   1 = Success
    #   2 = InProgress
    #   4 = Unknown
    #   5 = Error / Failed
    $wantedTypes = @(5)
    if ($IncludeUnknown) { $wantedTypes += 4 }

    $allAssets = foreach ($dep in $Deployments) {

        $depName = $dep.SoftwareName
        $depGuid = $dep.DeploymentID   # GUID e.g. {f96cc997-6315-4a33-8449-c0e18ffe576e}

        Write-Host "  Processing: $depName" -ForegroundColor Cyan

        # Step 1: Get the software update assignment object for this deployment
        $suDeployment = Get-CMSoftwareUpdateDeployment -DeploymentId $depGuid -ErrorAction Stop
        if (-not $suDeployment) {
            Write-Warning "  Get-CMSoftwareUpdateDeployment returned nothing for $depName ($depGuid)"
            continue
        }
        Write-Verbose "  Got SU deployment: $($suDeployment.AssignmentName)"

        # Step 2: Get the per-CI deployment status summary objects
        # Note: returns one row per Configuration Item in the deployment.
        # All rows are valid inputs to Get-CMDeploymentStatusDetails.
        $statusSummaries = Get-CMSoftwareUpdateDeploymentStatus -InputObject $suDeployment -ErrorAction Stop
        if (-not $statusSummaries) {
            Write-Warning "  Get-CMSoftwareUpdateDeploymentStatus returned nothing for $depName"
            continue
        }
        Write-Verbose "  Got $(@($statusSummaries).Count) CI status row(s)"

        # Step 3: Expand each summary to per-asset rows, filter to wanted status types.
        # We iterate all CI summaries; dedupe by machine name at the end.
        foreach ($summary in @($statusSummaries)) {
            $details = Get-CMDeploymentStatusDetails -InputObject $summary -ErrorAction SilentlyContinue
            if (-not $details) { continue }

            $details | Where-Object { $_.StatusType -in $wantedTypes } |
                Select-Object `
                    @{N='Deployment';    E={$depName}},
                    @{N='Computer';      E={$_.DeviceName}},
                    @{N='Status';        E={
                        switch ($_.StatusType) {
                            4 { 'Unknown' }
                            5 { 'Failed'  }
                            default { "Type$($_.StatusType)" }
                        }
                    }},
                    @{N='StatusDescription'; E={$_.StatusDescription}},
                    @{N='LastStatusTime';    E={$_.StatusTime}}
        }
    }

    # Deduplicate across deployments and CIs — keep most recent row per machine
    @($allAssets) |
        Where-Object { $_.Computer } |
        Sort-Object Computer, LastStatusTime -Descending |
        Group-Object Computer |
        ForEach-Object { $_.Group | Select-Object -First 1 }
}

# ============================================================================
# 3. Check for logged-on users (interactive sessions only)
# ============================================================================

function Test-UserLoggedOn {
    param([string] $Computer)
    try {
        $sessions = Get-CimInstance -ComputerName $Computer `
                                    -ClassName Win32_LogonSession `
                                    -Filter 'LogonType = 2 OR LogonType = 10' `
                                    -ErrorAction Stop `
                                    -OperationTimeoutSec 10
        return (@($sessions).Count -gt 0)
    }
    catch {
        Write-Verbose "[$Computer] Logon session query failed: $($_.Exception.Message)"
        return $false   # assume no user if unreachable
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
                Add-Report -Computer $c -Action 'AddToCollection' -Result 'DeviceNotFound'
                continue
            }
            if ($PSCmdlet.ShouldProcess($c, "Add to '$MaintenanceCollectionName'")) {
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
        Write-Host "  Collection update triggered — waiting 30s for membership to propagate ..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 30
    }
}

# ============================================================================
# 5. Reboot
# ============================================================================

function Invoke-VDIReboot {
    param([string] $Computer)

    # Try WinRM-based restart first; fall back to RPC via shutdown.exe
    try {
        Restart-Computer -ComputerName $Computer -Force -ErrorAction Stop
        Add-Report -Computer $Computer -Action 'Reboot' -Result 'Issued' -Detail 'Restart-Computer'
    }
    catch {
        try {
            $p = Start-Process shutdown.exe `
                 -ArgumentList "/m \\$Computer /r /f /t 5 /c `"SCCM patch remediation`"" `
                 -Wait -PassThru -NoNewWindow -ErrorAction Stop
            $result = if ($p.ExitCode -eq 0) { 'Issued' } else { 'Failed' }
            Add-Report -Computer $Computer -Action 'Reboot' -Result $result `
                       -Detail "shutdown.exe exit $($p.ExitCode)"
        }
        catch {
            Add-Report -Computer $Computer -Action 'Reboot' -Result 'Error' -Detail $_.Exception.Message
        }
    }
}

# ============================================================================
# 6. Post-reboot SCCM triggers
# ============================================================================

function Invoke-SCCMClientActions {
    param(
        [string] $Computer,
        [int]    $ResourceID
    )

    # Site-side fast-channel policy notification (cmdlet)
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

    # Client-side schedule triggers — no ConfigMgr cmdlet for these;
    # must use CIM directly against the client's root\ccm namespace
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
                             -Namespace  'root\ccm' `
                             -ClassName  'SMS_Client' `
                             -MethodName 'TriggerSchedule' `
                             -Arguments  @{ sScheduleID = $schedules[$action] } `
                             -ErrorAction Stop | Out-Null
            Add-Report -Computer $Computer -Action $action -Result 'Triggered'
            Start-Sleep -Seconds 2
        }
        catch {
            Add-Report -Computer $Computer -Action $action -Result 'Error' -Detail $_.Exception.Message
        }
    }
}

# ============================================================================
# 7. Wait for a batch to come back online then trigger SCCM actions
# ============================================================================

function Wait-BatchOnlineAndTrigger {
    param(
        [array] $Devices,           # [pscustomobject]@{ Computer; ResourceID }
        [int]   $TimeoutMinutes
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $pending  = @{}
    foreach ($d in $Devices) { $pending[$d.Computer] = $d.ResourceID }

    Write-Host "  Waiting up to ${TimeoutMinutes}m for batch to come back online ..." -ForegroundColor DarkGray

    while ($pending.Count -gt 0 -and (Get-Date) -lt $deadline) {
        foreach ($name in @($pending.Keys)) {
            if (Test-Connection -ComputerName $name -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                try {
                    # Confirm WMI/OS is responsive, not just ICMP
                    $null = Get-CimInstance -ComputerName $name -ClassName Win32_OperatingSystem `
                                            -ErrorAction Stop -OperationTimeoutSec 10
                    Write-Host "  [$name] Online — running SCCM triggers" -ForegroundColor Green
                    Add-Report -Computer $name -Action 'WaitOnline' -Result 'Online'
                    Invoke-SCCMClientActions -Computer $name -ResourceID $pending[$name]
                    $pending.Remove($name) | Out-Null
                }
                catch { <# Pingable but OS/WMI not ready yet — try next loop #> }
            }
        }
        if ($pending.Count -gt 0) { Start-Sleep -Seconds 20 }
    }

    foreach ($name in @($pending.Keys)) {
        Write-Warning "[$name] Did not come back online within ${TimeoutMinutes} minutes."
        Add-Report -Computer $name -Action 'WaitOnline' -Result 'TimedOut'
    }
}

# ============================================================================
# Main
# ============================================================================

try {
    Connect-CMSite

    # ---- 1. Select deployments ----------------------------------------
    $selected = Select-Deployments
    if (-not $selected) { Write-Warning "No deployments selected. Exiting."; return }
    Write-Host ("Selected {0} deployment(s)." -f @($selected).Count) -ForegroundColor Cyan

    # ---- 2. Include Unknown? ------------------------------------------
    if (-not $PSBoundParameters.ContainsKey('IncludeUnknown')) {
        $ans = Read-Host "`nAlso include machines in 'Unknown' state? (Y/N)"
        [bool]$IncludeUnknown = ($ans -match '^[Yy]')
    }

    # ---- 3. Get failed/unknown assets ---------------------------------
    Write-Host "`nQuerying failed$(if($IncludeUnknown){' + unknown'}) assets ..." -ForegroundColor Cyan
    $assets = Get-FailedAssets -Deployments $selected -IncludeUnknown $IncludeUnknown

    if (-not $assets) {
        Write-Warning "No failed or unknown assets found. Nothing to do."
        Save-Report
        return
    }

    Write-Host ("`nFound {0} unique machine(s):" -f @($assets).Count) -ForegroundColor Yellow
    $assets | Group-Object Status | ForEach-Object {
        Write-Host ("  {0,3}  {1}" -f $_.Count, $_.Name) -ForegroundColor White
    }

    # Log initial discovery
    foreach ($a in $assets) {
        Add-Report -Computer $a.Computer -Action 'FoundInDeployment' -Result $a.Status `
                   -Detail "$($a.Deployment) | $($a.StatusDescription)"
    }

    # ---- 4. Resolve SCCM device records ------------------------------
    Write-Host "`nResolving device records ..." -ForegroundColor Cyan
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
        Write-Warning "No machines resolved to SCCM device records."
        Save-Report
        return
    }

    # ---- 5. Confirm ---------------------------------------------------
    Write-Host ""
    if (-not $PSCmdlet.ShouldContinue(
            "Add $($devices.Count) machine(s) to '$MaintenanceCollectionName', check logged-on users, reboot in batches of $BatchSize every $BatchIntervalMinutes min, then trigger SCCM policy and updates?",
            'Confirm VDI Remediation')) {
        Write-Warning "Aborted by user."
        Save-Report
        return
    }

    # ---- 6. Add to maintenance collection ----------------------------
    Write-Host "`nAdding machines to '$MaintenanceCollectionName' ..." -ForegroundColor Cyan
    Add-ToMaintenanceCollection -Computers $devices.Computer

    # ---- 7. Logged-on user check (opt-in only via -SkipLoggedOnCheck:$false) ----
    $toReboot = [System.Collections.Generic.List[object]]::new()
    $skipped  = [System.Collections.Generic.List[object]]::new()

    if ($SkipLoggedOnCheck) {
        # Default: reboot all machines regardless of logged-on state.
        # VDIs in a patch remediation window are rebooted unconditionally.
        $toReboot.AddRange($devices)
        Write-Host "`nSkipping logged-on check — all $($devices.Count) machine(s) queued for reboot." -ForegroundColor Cyan
    }
    else {
        Write-Host "`nChecking for logged-on users ..." -ForegroundColor Cyan
        foreach ($dev in $devices) {
            if (Test-Connection -ComputerName $dev.Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                if (Test-UserLoggedOn -Computer $dev.Computer) {
                    Write-Host "  [$($dev.Computer)] User logged on — SKIPPING reboot" -ForegroundColor Magenta
                    Add-Report -Computer $dev.Computer -Action 'LoggedOnCheck' -Result 'UserLoggedOn-Skipped'
                    $skipped.Add($dev)
                }
                else {
                    Add-Report -Computer $dev.Computer -Action 'LoggedOnCheck' -Result 'Clear'
                    $toReboot.Add($dev)
                }
            }
            else {
                Add-Report -Computer $dev.Computer -Action 'LoggedOnCheck' -Result 'Offline-QueuedForReboot'
                $toReboot.Add($dev)
            }
        }
        Write-Host ""
        Write-Host ("  {0,3}  machine(s) queued for reboot" -f $toReboot.Count) -ForegroundColor Cyan
        Write-Host ("  {0,3}  machine(s) skipped (user logged on)" -f $skipped.Count) -ForegroundColor Magenta
    }

    if ($toReboot.Count -eq 0) {
        Write-Warning "No machines to reboot after logged-on check."
        Save-Report
        return
    }

    # ---- 8. Batched reboot — parallel per batch ----------------------
    #
    # Within each batch:
    #   a) Fire ALL reboots simultaneously (runspace per machine)
    #   b) In parallel, each runspace then waits for its machine to come
    #      back online and fires the SCCM triggers
    # The host thread just waits for the whole batch to finish, then
    # pauses before kicking off the next batch.
    # Results are collected back into $script:Report via a thread-safe bag.

    $batchList = for ($i = 0; $i -lt $toReboot.Count; $i += $BatchSize) {
        ,@($toReboot[$i .. [Math]::Min($i + $BatchSize - 1, $toReboot.Count - 1)])
    }

    Write-Host ("`nProcessing {0} batch(es) of up to {1} in parallel." -f $batchList.Count, $BatchSize) -ForegroundColor Cyan

    # Thread-safe collection for results coming back from runspaces
    $resultBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    # Scriptblock executed in each runspace — one per machine
    $workerScript = {
        param($Computer, $ResourceID, $OnlineWaitMinutes, $WhatIf)

        $results = [System.Collections.Generic.List[object]]::new()
        function Log { param($Action,$Result,$Detail='')
            $results.Add([pscustomobject]@{
                Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Computer  = $Computer
                Action    = $Action
                Result    = $Result
                Detail    = $Detail
            }) | Out-Null
        }

        # -- Reboot --
        if (-not $WhatIf) {
            try {
                Restart-Computer -ComputerName $Computer -Force -ErrorAction Stop
                Log 'Reboot' 'Issued' 'Restart-Computer'
            }
            catch {
                try {
                    $p = Start-Process shutdown.exe `
                         -ArgumentList "/m \\$Computer /r /f /t 5 /c `"SCCM patch remediation`"" `
                         -Wait -PassThru -NoNewWindow -ErrorAction Stop
                    Log 'Reboot' $(if($p.ExitCode -eq 0){'Issued'}else{'Failed'}) "shutdown.exe exit $($p.ExitCode)"
                }
                catch { Log 'Reboot' 'Error' $_.Exception.Message }
            }
        } else {
            Log 'Reboot' 'WhatIf' ''
        }

        # -- Wait online --
        $deadline = (Get-Date).AddMinutes($OnlineWaitMinutes)
        $online   = $false
        while ((Get-Date) -lt $deadline) {
            if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                try {
                    $null = Get-CimInstance -ComputerName $Computer -ClassName Win32_OperatingSystem `
                                            -ErrorAction Stop -OperationTimeoutSec 10
                    $online = $true
                    break
                }
                catch { }
            }
            Start-Sleep -Seconds 20
        }

        if (-not $online) {
            Log 'WaitOnline' 'TimedOut' ''
            return $results
        }
        Log 'WaitOnline' 'Online' ''

        if ($WhatIf) { return $results }

        # -- SCCM client-side triggers --
        $schedules = [ordered]@{
            'Machine Policy Retrieval'     = '{00000000-0000-0000-0000-000000000021}'
            'Machine Policy Evaluation'    = '{00000000-0000-0000-0000-000000000022}'
            'Software Updates Scan'        = '{00000000-0000-0000-0000-000000000113}'
            'Software Updates Deploy Eval' = '{00000000-0000-0000-0000-000000000108}'
            'State Message Refresh'        = '{00000000-0000-0000-0000-000000000111}'
        }
        foreach ($action in $schedules.Keys) {
            try {
                Invoke-CimMethod -ComputerName $Computer -Namespace 'root\ccm' `
                                 -ClassName SMS_Client -MethodName TriggerSchedule `
                                 -Arguments @{ sScheduleID = $schedules[$action] } `
                                 -ErrorAction Stop | Out-Null
                Log $action 'Triggered' ''
                Start-Sleep -Seconds 2
            }
            catch { Log $action 'Error' $_.Exception.Message }
        }

        return $results
    }

    for ($b = 0; $b -lt $batchList.Count; $b++) {
        $batch = $batchList[$b]
        Write-Host ("`n=== Batch {0}/{1} — firing {2} reboot(s) in parallel ===" `
                    -f ($b+1), $batchList.Count, $batch.Count) -ForegroundColor Yellow

        # Create a runspace pool capped at batch size
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $batch.Count)
        $pool.Open()
        $handles = [System.Collections.Generic.List[object]]::new()

        foreach ($dev in $batch) {
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            $null = $ps.AddScript($workerScript).AddParameters(@{
                Computer         = $dev.Computer
                ResourceID       = $dev.ResourceID
                OnlineWaitMinutes = $OnlineWaitMinutes
                WhatIf           = ($WhatIfPreference -eq 'Continue')
            })
            $handles.Add([pscustomobject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Computer = $dev.Computer })
            Write-Host "  [$($dev.Computer)] Reboot + monitor dispatched" -ForegroundColor White
        }

        # Site-side policy notification fires immediately from the host thread
        # (needs the CM drive which isn't available inside runspaces)
        foreach ($dev in $batch) {
            try {
                if ($PSCmdlet.ShouldProcess($dev.Computer, 'Send policy notification')) {
                    Invoke-CMClientNotification -DeviceId $dev.ResourceID `
                                                -NotificationType RequestMachinePolicyNow `
                                                -ErrorAction Stop
                    $resultBag.Add([pscustomobject]@{
                        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                        Computer  = $dev.Computer; Action = 'PolicyNotification'
                        Result    = 'Sent'; Detail = ''
                    })
                }
            }
            catch {
                $resultBag.Add([pscustomobject]@{
                    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Computer  = $dev.Computer; Action = 'PolicyNotification'
                    Result    = 'Error'; Detail = $_.Exception.Message
                })
            }
        }

        Write-Host ("  Waiting for all {0} machine(s) in batch to complete ..." -f $batch.Count) -ForegroundColor DarkGray

        # Collect results as each runspace finishes
        foreach ($h in $handles) {
            try {
                $rows = $h.PS.EndInvoke($h.Handle)
                foreach ($r in $rows) { $resultBag.Add($r) }
                $state = if ($h.PS.HadErrors) { 'RunspaceError' } else { 'Done' }
                Write-Host "  [$($h.Computer)] $state" -ForegroundColor $(if($h.PS.HadErrors){'Red'}else{'Green'})
            }
            catch {
                Write-Warning "[$($h.Computer)] Runspace exception: $($_.Exception.Message)"
            }
            finally {
                $h.PS.Dispose()
            }
        }
        $pool.Close()
        $pool.Dispose()

        if ($b -lt $batchList.Count - 1) {
            Write-Host ("`n  Waiting {0} minute(s) before next batch ..." -f $BatchIntervalMinutes) -ForegroundColor DarkGray
            Start-Sleep -Seconds ($BatchIntervalMinutes * 60)
        }
    }

    # Merge parallel results into main report
    foreach ($r in $resultBag) { $script:Report.Add($r) | Out-Null }

    # ---- 9. Done ------------------------------------------------------
    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host " REMEDIATION COMPLETE" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Save-Report
}
catch {
    Write-Error "Fatal: $_"
    Save-Report
    throw
}
finally {
    Pop-Location -ErrorAction SilentlyContinue
}
