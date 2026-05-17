<#
.SYNOPSIS
    Remediates VDIs that failed (or are unknown) for one or more SCCM
    software update deployments. Cmdlet-driven version.

.DESCRIPTION
    1.  Lists software update deployments via Get-CMDeployment, lets you
        pick one or many in Out-GridView.
    2.  Pulls per-asset status with Get-CMDeploymentStatus + 
        Get-CMDeploymentStatusDetails and filters to Failed and 
        (optionally) Unknown buckets.
    3.  Adds the affected machines as direct members of the maintenance
        collection using Add-CMDeviceCollectionDirectMembershipRule,
        then refreshes with Invoke-CMCollectionUpdate.
    4.  Reboots them in batches (default 20, 5 min apart) using
        Restart-Computer; falls back to shutdown.exe if WinRM is blocked.
    5.  Waits for each batch to come back online, then triggers Machine
        Policy + Software Updates Scan/Eval via Invoke-CMClientNotification
        (site-side cmdlet) and Invoke-CimMethod on the client itself
        (no cmdlet exists for client-side TriggerSchedule).
    6.  Writes a CSV report.

.EXAMPLE
    .\Invoke-VDIPatchRemediation.ps1 -WhatIf -Verbose
    .\Invoke-VDIPatchRemediation.ps1 -IncludeUnknown
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string] $SiteCode                  = 'PRD',
    [string] $SiteServer                = 'appsmcm101fp.iprod.local',
    [string] $MaintenanceCollectionName = 'VDI Maintenance Anytime',
    [int]    $BatchSize                 = 20,
    [int]    $BatchIntervalMinutes      = 5,
    [switch] $IncludeUnknown,
    [int]    $OnlineWaitMinutes         = 15,
    [string] $LogPath                   = "C:\Temp\VDIPatchRemediation_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
)

# ============================================================================
# Helpers
# ============================================================================

$script:Log = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param(
        [string] $Computer,
        [string] $Action,
        [string] $Result,
        [string] $Detail = ''
    )
    $script:Log.Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('s')
        Computer  = $Computer
        Action    = $Action
        Result    = $Result
        Detail    = $Detail
    }) | Out-Null
    Write-Verbose "[$Computer] $Action => $Result $Detail"
}

function Connect-CMSite {
    param([string] $SiteCode, [string] $SiteServer)

    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw "SMS_ADMIN_UI_PATH not set — install the ConfigMgr console on this host."
    }
    $modulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (-not (Get-Module ConfigurationManager)) {
        Import-Module $modulePath -ErrorAction Stop
    }
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
    }
    Set-Location ("{0}:\" -f $SiteCode)
}

# ============================================================================
# 1. Deployment selection
# ============================================================================

function Select-UpdateDeployment {
    # FeatureType 5 = Software Updates
    $raw = Get-CMDeployment -FeatureType SoftwareUpdate

    if (-not $raw) { throw "No software update deployments found." }

    # One-time diagnostic: what properties does this object actually expose?
    Write-Verbose "Available properties on Get-CMDeployment result:"
    $raw[0].PSObject.Properties |
        Where-Object { $_.Name -match 'ID|Unique|Offer|Assign' } |
        ForEach-Object { Write-Verbose ("  {0,-30} = {1}" -f $_.Name, $_.Value) }

    $deployments = $raw |
        Select-Object *,                                    # keep everything; we'll pick the right ID later
                      @{N='Name';           E={$_.SoftwareName}},
                      @{N='Collection';     E={$_.CollectionName}} |
        Sort-Object DeploymentTime -Descending

    # Show the user a clean view, but pass the full objects through
    $picked = $deployments |
        Select-Object Name, Collection, NumberTargeted, NumberErrors, NumberUnknown, DeploymentTime, DeploymentID, AssignmentID, OfferID |
        Out-GridView -Title 'Select one or more update deployments to remediate' -OutputMode Multiple

    if (-not $picked) { return $null }

    # Map the picks back to the full objects (match on DeploymentID which is always populated)
    $deployments | Where-Object { $_.DeploymentID -in $picked.DeploymentID }
}

# ============================================================================
# 2. Per-asset status
# ============================================================================

function Get-FailedAndUnknownAssets {
    <#
        Strategy A: Get-CMUpdateGroupDeployment | Get-CMDeploymentStatus | Get-CMDeploymentStatusDetails
        Strategy B: Invoke-Command on the site server to query SMS_UpdateComplianceStatus
                    locally (avoids WMI remote permission issues — 0x80041001).
    #>
    param(
        [Parameter(Mandatory)] $Deployments,
        [Parameter(Mandatory)] [string] $SiteCode,
        [Parameter(Mandatory)] [string] $SiteServer,
        [switch] $IncludeUnknown
    )

    $wantedTypes = @(5)
    if ($IncludeUnknown) { $wantedTypes += 4 }

    $rows = foreach ($d in $Deployments) {

        # DeploymentID is the GUID, AssignmentID is the integer — confirmed from verbose dump
        $guid      = $d.DeploymentID   # e.g. {2c9b4f52-3919-453d-bd51-e38b6dc867d3}
        $numericId = $d.AssignmentID   # e.g. 16777747

        Write-Verbose "Deployment '$($d.Name)'  GUID=$guid  AssignmentID=$numericId"

        # ---- Strategy A: cmdlet pipeline -----------------------------------
        $found = $null
        try {
            $dep = Get-CMUpdateGroupDeployment -DeploymentId $guid -ErrorAction SilentlyContinue
            if ($dep) {
                $buckets = $dep | Get-CMDeploymentStatus -ErrorAction SilentlyContinue |
                           Where-Object { $_.StatusType -in $wantedTypes }
                Write-Verbose ("  Strategy A: {0} bucket(s)" -f @($buckets).Count)

                if ($buckets) {
                    $found = foreach ($b in $buckets) {
                        $label   = if ($b.StatusType -eq 4) { 'Unknown' } else { 'Failed' }
                        $details = Get-CMDeploymentStatusDetails -InputObject $b -ErrorAction SilentlyContinue
                        Write-Verbose ("    [$label] => $(@($details).Count) asset(s)")
                        $details | Select-Object @{N='Deployment';    E={$d.Name}},
                                                 @{N='Computer';      E={$_.DeviceName}},
                                                 @{N='StatusType';    E={$label}},
                                                 @{N='LastStatusTime';E={$_.StatusTime}}
                    }
                }
            }
        } catch {
            Write-Verbose "  Strategy A failed: $($_.Exception.Message)"
        }

        # ---- Strategy B: Invoke-Command on site server (local WMI) --------
        if (-not $found) {
            Write-Verbose "  Strategy B: Invoke-Command to $SiteServer"
            try {
                $found = Invoke-Command -ComputerName $SiteServer -ErrorAction Stop -ScriptBlock {
                    param($ns, $assignmentId, $wantedTypes, $deploymentName)

                    $stateFilter = ($wantedTypes | ForEach-Object { "StateType = $_" }) -join ' OR '
                    $query = "SELECT * FROM SMS_UpdateComplianceStatus WHERE AssignmentID = $assignmentId AND ($stateFilter)"

                    $rows = Get-CimInstance -Namespace $ns -Query $query -ErrorAction Stop

                    foreach ($r in $rows) {
                        # Resolve ResourceID to machine name locally
                        $res = Get-CimInstance -Namespace $ns `
                               -Query "SELECT Name FROM SMS_R_System WHERE ResourceID = $($r.MachineID)" `
                               -ErrorAction SilentlyContinue
                        if ($res.Name) {
                            [pscustomobject]@{
                                Deployment     = $deploymentName
                                Computer       = $res.Name
                                StatusType     = if ($r.StateType -eq 4) { 'Unknown' } else { 'Failed' }
                                LastStatusTime = $r.LastStatusChangeTime
                            }
                        }
                    }
                } -ArgumentList "root\sms\site_$SiteCode", $numericId, $wantedTypes, $d.Name

                Write-Verbose ("  Strategy B: {0} asset(s)" -f @($found).Count)
            } catch {
                Write-Verbose "  Strategy B failed: $($_.Exception.Message)"
            }
        }

        if (-not $found) {
            Write-Warning "No assets retrieved for deployment '$($d.Name)' via any method."
        }
        $found
    }

    # Dedupe — keep most recent per machine
    $rows | Where-Object Computer |
            Sort-Object Computer, LastStatusTime -Descending |
            Group-Object Computer |
            ForEach-Object { $_.Group | Select-Object -First 1 }
}

# ============================================================================
# 3. Collection membership
# ============================================================================

function Add-ToMaintenanceCollection {
    param(
        [Parameter(Mandatory)] [string[]] $Computers,
        [Parameter(Mandatory)] [string]   $CollectionName
    )

    $coll = Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop
    if (-not $coll) { throw "Collection '$CollectionName' not found." }

    $existing = Get-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID |
                Select-Object -ExpandProperty RuleName
    if (-not $existing) { $existing = @() }

    foreach ($c in $Computers) {
        if ($existing -contains $c) {
            Write-Log -Computer $c -Action 'AddToCollection' -Result 'AlreadyMember'
            continue
        }
        try {
            $device = Get-CMDevice -Name $c -Fast -ErrorAction Stop
            if (-not $device) {
                Write-Log -Computer $c -Action 'AddToCollection' -Result 'NotInSCCM'
                continue
            }
            if ($PSCmdlet.ShouldProcess($c, "Add to $CollectionName")) {
                Add-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID `
                                                          -ResourceId $device.ResourceID `
                                                          -ErrorAction Stop
                Write-Log -Computer $c -Action 'AddToCollection' -Result 'Added'
            }
        } catch {
            Write-Log -Computer $c -Action 'AddToCollection' -Result 'Error' -Detail $_.Exception.Message
        }
    }

    if ($PSCmdlet.ShouldProcess($CollectionName, 'Update collection membership')) {
        Invoke-CMCollectionUpdate -CollectionId $coll.CollectionID -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 30   # let eval propagate before reboots start
    }
}

# ============================================================================
# 4. Reboot
# ============================================================================

function Restart-VDI {
    param([string] $Computer)

    try {
        Restart-Computer -ComputerName $Computer -Force -ErrorAction Stop
        Write-Log -Computer $Computer -Action 'Reboot' -Result 'Issued' -Detail 'Restart-Computer'
    } catch {
        Write-Verbose "[$Computer] Restart-Computer failed, falling back to shutdown.exe: $($_.Exception.Message)"
        try {
            $p = Start-Process -FilePath shutdown.exe `
                               -ArgumentList "/m \\$Computer /r /f /t 5 /c `"SCCM patch remediation`"" `
                               -Wait -PassThru -NoNewWindow -ErrorAction Stop
            if ($p.ExitCode -eq 0) {
                Write-Log -Computer $Computer -Action 'Reboot' -Result 'Issued' -Detail 'shutdown.exe'
            } else {
                Write-Log -Computer $Computer -Action 'Reboot' -Result 'Failed' -Detail "shutdown.exe exit $($p.ExitCode)"
            }
        } catch {
            Write-Log -Computer $Computer -Action 'Reboot' -Result 'Error' -Detail $_.Exception.Message
        }
    }
}

# ============================================================================
# Main
# ============================================================================

try {
    Write-Host "Connecting to site $SiteCode on $SiteServer ..." -ForegroundColor Cyan
    Connect-CMSite -SiteCode $SiteCode -SiteServer $SiteServer

    # ---- 1. choose deployment(s) ------------------------------------------
    Write-Host "Loading software update deployments ..." -ForegroundColor Cyan
    $selected = Select-UpdateDeployment
    if (-not $selected) { Write-Warning 'No deployment selected. Exiting.'; return }

    # ---- 2. Unknown bucket prompt -----------------------------------------
    if (-not $PSBoundParameters.ContainsKey('IncludeUnknown')) {
        $ans = Read-Host "Also include machines in 'Unknown' state? (Y/N)"
        $IncludeUnknown = ($ans -match '^[Yy]')
    }

    # ---- 3. pull failed/unknown assets ------------------------------------
    Write-Host ("Querying failed{0} assets across {1} deployment(s) ..." `
                -f $(if ($IncludeUnknown) {' + unknown'}), $selected.Count) -ForegroundColor Cyan

    $assets = Get-FailedAndUnknownAssets -Deployments $selected `
                                         -SiteCode $SiteCode `
                                         -SiteServer $SiteServer `
                                         -IncludeUnknown:$IncludeUnknown

    if (-not $assets) {
        Write-Warning 'No failed or unknown assets found in the selected deployment(s). Nothing to do.'
        return
    }

    $computers = $assets.Computer | Sort-Object -Unique
    Write-Host ("Found {0} unique machine(s) to remediate." -f $computers.Count) -ForegroundColor Green
    $assets | Group-Object StatusType | Format-Table Count, Name -AutoSize

    # ---- 4. confirm -------------------------------------------------------
    if (-not $PSCmdlet.ShouldContinue(
            ("About to add {0} machine(s) to '{1}', reboot in batches of {2} every {3} min, and trigger policy/updates. Proceed?" `
                -f $computers.Count, $MaintenanceCollectionName, $BatchSize, $BatchIntervalMinutes),
            'Confirm VDI remediation')) {
        Write-Warning 'Aborted by user.'
        return
    }

    # Resolve ResourceIDs once — needed for Invoke-CMClientNotification later
    Write-Host "Resolving device records ..." -ForegroundColor Cyan
    $devices = foreach ($c in $computers) {
        $d = Get-CMDevice -Name $c -Fast -ErrorAction SilentlyContinue
        if ($d) {
            [pscustomobject]@{ Computer = $c; ResourceID = $d.ResourceID }
        } else {
            Write-Log -Computer $c -Action 'ResolveDevice' -Result 'NotInSCCM'
        }
    }
    if (-not $devices) {
        Write-Warning 'None of the failed machines resolved to a device record. Exiting.'
        return
    }

    # ---- 5. add to maintenance collection ---------------------------------
    Add-ToMaintenanceCollection -Computers $devices.Computer -CollectionName $MaintenanceCollectionName

    # ---- 6. batched reboot + post-reboot triggers -------------------------
    $batches = for ($i = 0; $i -lt $devices.Count; $i += $BatchSize) {
        ,@($devices[$i..([Math]::Min($i + $BatchSize - 1, $devices.Count - 1))])
    }
    Write-Host ("Processing {0} batch(es) of up to {1}." -f $batches.Count, $BatchSize) -ForegroundColor Cyan

    $jobs = New-Object System.Collections.Generic.List[object]

    for ($b = 0; $b -lt $batches.Count; $b++) {
        $batch = $batches[$b]
        Write-Host ("`n--- Batch {0}/{1} : {2} machine(s) ---" -f ($b+1), $batches.Count, $batch.Count) -ForegroundColor Yellow

        foreach ($dev in $batch) {
            if ($PSCmdlet.ShouldProcess($dev.Computer, 'Reboot')) {
                Restart-VDI -Computer $dev.Computer
            }
        }

        # Post-reboot wait + triggers as a background job so the 5-min
        # pacing isn't blocked by waiting for this batch to come back.
        $job = Start-Job -Name "PostReboot_Batch$($b+1)" -ScriptBlock {
            param($devList, $timeout)

            $pending = @{}
            foreach ($d in $devList) { $pending[$d.Computer] = $d.ResourceID }
            $deadline = (Get-Date).AddMinutes($timeout)
            $results  = New-Object System.Collections.Generic.List[object]

            $schedules = @(
                @{ Name='Machine Policy Retrieval';    ID='{00000000-0000-0000-0000-000000000021}' },
                @{ Name='Machine Policy Evaluation';   ID='{00000000-0000-0000-0000-000000000022}' },
                @{ Name='Software Updates Scan';       ID='{00000000-0000-0000-0000-000000000113}' },
                @{ Name='Software Updates Deploy Eval';ID='{00000000-0000-0000-0000-000000000108}' },
                @{ Name='State Message Refresh';       ID='{00000000-0000-0000-0000-000000000111}' }
            )

            while ($pending.Count -and (Get-Date) -lt $deadline) {
                foreach ($name in @($pending.Keys)) {
                    if (Test-Connection -ComputerName $name -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                        try {
                            $null = Get-CimInstance -ComputerName $name -ClassName Win32_OperatingSystem `
                                                    -ErrorAction Stop -OperationTimeoutSec 5
                            $results.Add([pscustomobject]@{ Computer=$name; Action='WaitOnline'; Result='Online' })

                            foreach ($s in $schedules) {
                                try {
                                    Invoke-CimMethod -ComputerName $name -Namespace 'root\ccm' `
                                                     -ClassName SMS_Client -MethodName TriggerSchedule `
                                                     -Arguments @{ sScheduleID = $s.ID } -ErrorAction Stop | Out-Null
                                    $results.Add([pscustomobject]@{ Computer=$name; Action=$s.Name; Result='Triggered' })
                                    Start-Sleep -Seconds 2
                                } catch {
                                    $results.Add([pscustomobject]@{ Computer=$name; Action=$s.Name; Result='Error'; Detail=$_.Exception.Message })
                                }
                            }
                            $pending.Remove($name) | Out-Null
                        } catch { }
                    }
                }
                if ($pending.Count) { Start-Sleep -Seconds 20 }
            }
            foreach ($name in $pending.Keys) {
                $results.Add([pscustomobject]@{ Computer=$name; Action='WaitOnline'; Result='TimedOut' })
            }
            $results
        } -ArgumentList (,$batch), $OnlineWaitMinutes

        # Site-side push via cmdlet — best effort, fires immediately
        foreach ($dev in $batch) {
            try {
                Invoke-CMClientNotification -DeviceId $dev.ResourceID `
                                            -NotificationType RequestMachinePolicyNow `
                                            -ErrorAction Stop
                Write-Log -Computer $dev.Computer -Action 'ClientNotify' -Result 'Sent'
            } catch {
                Write-Log -Computer $dev.Computer -Action 'ClientNotify' -Result 'Error' -Detail $_.Exception.Message
            }
        }

        $jobs.Add($job)

        if ($b -lt $batches.Count - 1) {
            Write-Host ("Waiting {0} minute(s) before next batch ..." -f $BatchIntervalMinutes) -ForegroundColor DarkGray
            Start-Sleep -Seconds ($BatchIntervalMinutes * 60)
        }
    }

    Write-Host "`nAll reboots issued. Waiting for outstanding post-reboot jobs ..." -ForegroundColor Cyan
    $jobs | Wait-Job | Out-Null
    foreach ($j in $jobs) {
        $results = Receive-Job -Job $j -ErrorAction SilentlyContinue
        foreach ($r in $results) {
            Write-Log -Computer $r.Computer -Action $r.Action -Result $r.Result -Detail ($r.Detail -as [string])
        }
        Remove-Job -Job $j -Force
    }

    # ---- 7. report --------------------------------------------------------
    $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
    $script:Log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nDone. Log: $LogPath" -ForegroundColor Green

    $script:Log | Group-Object Action, Result |
        Select-Object Count, Name |
        Sort-Object Name |
        Format-Table -AutoSize
}
catch {
    Write-Error $_
    if ($script:Log.Count) {
        $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
        $script:Log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
        Write-Host "Partial log written to $LogPath" -ForegroundColor Yellow
    }
    throw
}
finally {
    Pop-Location -ErrorAction SilentlyContinue
}
