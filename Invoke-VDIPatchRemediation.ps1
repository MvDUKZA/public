<#
.SYNOPSIS
    Remediates VDIs that failed or are in unknown state for one or more SCCM
    software update deployments.

.DESCRIPTION
    1.  Prompts for a Software Update Group / Deployment (multi-select).
    2.  Pulls assets with EnforcementState matching the selected buckets
        (Failed / Unknown / both).
    3.  Adds them as direct members to the "VDI Maintenance Anytime"
        collection and updates the collection.
    4.  Reboots the machines in batches (default 20, 5 min apart) using
        shutdown.exe (reliable on locked-down VDIs); falls back gracefully.
    5.  Waits for each batch to come back online, then triggers:
            - Machine Policy Retrieval + Evaluation
            - Software Updates Scan
            - Software Updates Deployment Evaluation
            - State Message Refresh
    6.  Writes a CSV report to $LogPath.

.PARAMETER SiteCode
    SCCM site code. Defaults to PRD.

.PARAMETER SiteServer
    SCCM site server FQDN. Defaults to appsmcm101fp.iprod.local.

.PARAMETER MaintenanceCollectionName
    Collection that grants Anytime maintenance window. Defaults to
    "VDI Maintenance Anytime".

.PARAMETER BatchSize
    How many machines to reboot per wave. Default 20.

.PARAMETER BatchIntervalMinutes
    Minutes between waves. Default 5.

.PARAMETER IncludeUnknown
    Switch. If set, machines with EnforcementState = Unknown are remediated
    alongside Failed. You will also be prompted interactively.

.PARAMETER OnlineWaitMinutes
    How long to wait for a rebooted machine to come back up before
    skipping its post-reboot triggers. Default 15.

.PARAMETER LogPath
    Output CSV. Defaults to C:\Temp\VDIPatchRemediation_<timestamp>.csv.

.PARAMETER WhatIf
    Standard. Plans the run, makes no changes.

.EXAMPLE
    .\Invoke-VDIPatchRemediation.ps1 -IncludeUnknown -Verbose

.NOTES
    Run on a workstation/jumpbox with the ConfigMgr console installed and
    the active user holding rights on the collection + remote restart on
    the VDIs.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string] $SiteCode                 = 'PRD',
    [string] $SiteServer               = 'appsmcm101fp.iprod.local',
    [string] $MaintenanceCollectionName = 'VDI Maintenance Anytime',
    [int]    $BatchSize                = 20,
    [int]    $BatchIntervalMinutes     = 5,
    [switch] $IncludeUnknown,
    [int]    $OnlineWaitMinutes        = 15,
    [string] $LogPath                  = "C:\Temp\VDIPatchRemediation_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
)

# ---------- helpers ---------------------------------------------------------

$script:Log = New-Object System.Collections.Generic.List[object]

function Write-Log {
    param(
        [string] $Computer,
        [string] $Action,
        [string] $Result,
        [string] $Detail = ''
    )
    $entry = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('s')
        Computer  = $Computer
        Action    = $Action
        Result    = $Result
        Detail    = $Detail
    }
    $script:Log.Add($entry) | Out-Null
    Write-Verbose "[$Computer] $Action => $Result $Detail"
}

function Connect-CMSite {
    param([string] $SiteCode, [string] $SiteServer)

    $modulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (-not (Get-Module ConfigurationManager)) {
        Import-Module $modulePath -ErrorAction Stop
    }
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
    }
    Set-Location ("{0}:\" -f $SiteCode)
}

function Get-FailedAndUnknownAssets {
    <#
        Pulls compliance status straight from WMI on the site server.
        SMS_UpdateComplianceStatus is the per-asset/per-update view but it's
        huge; we instead use the deployment summarizer per-asset class
        SMS_StatMsg / SMS_SUMDeploymentAsset which is keyed by AssignmentID
        (deployment) and gives a per-resource roll-up.
    #>
    param(
        [Parameter(Mandatory)] [int[]] $AssignmentIDs,
        [Parameter(Mandatory)] [string] $SiteCode,
        [Parameter(Mandatory)] [string] $SiteServer,
        [switch] $IncludeUnknown
    )

    # StatusType: 1=Compliant 2=InProgress 3=RequirementsNotMet 4=Unknown 5=Error
    $wantedStatusTypes = @(5)            # Error / Failed
    if ($IncludeUnknown) { $wantedStatusTypes += 4 }

    $filter = ($wantedStatusTypes | ForEach-Object { "StatusType = $_" }) -join ' OR '
    $ns     = "root\sms\site_$SiteCode"

    $all = foreach ($id in $AssignmentIDs) {
        $query = "SELECT * FROM SMS_SUMDeploymentAssetDetails WHERE AssignmentID = $id AND ($filter)"
        Get-CimInstance -ComputerName $SiteServer -Namespace $ns -Query $query -ErrorAction Stop |
            Select-Object @{N='AssignmentID';E={$id}},
                          @{N='Computer';   E={$_.MachineName}},
                          @{N='StatusType'; E={$_.StatusType}},
                          @{N='LastStatus'; E={$_.LastStatusChangeTime}}
    }

    # Dedupe across deployments — one machine may be failing several deployments
    $all | Sort-Object Computer -Unique
}

function Add-ToMaintenanceCollection {
    param(
        [Parameter(Mandatory)] [string[]] $Computers,
        [Parameter(Mandatory)] [string]   $CollectionName
    )

    $coll = Get-CMDeviceCollection -Name $CollectionName -ErrorAction Stop
    if (-not $coll) {
        throw "Collection '$CollectionName' not found."
    }

    # Existing direct-rule members so we don't try to add duplicates
    $existing = (Get-CMCollectionDirectMembershipRule -CollectionId $coll.CollectionID |
                    Select-Object -ExpandProperty RuleName) -as [string[]]
    if (-not $existing) { $existing = @() }

    foreach ($c in $Computers) {
        if ($existing -contains $c) {
            Write-Log -Computer $c -Action 'AddToCollection' -Result 'AlreadyMember'
            continue
        }
        try {
            $res = Get-CMDevice -Name $c -Fast
            if (-not $res) {
                Write-Log -Computer $c -Action 'AddToCollection' -Result 'NotInSCCM'
                continue
            }
            if ($PSCmdlet.ShouldProcess($c, "Add to $CollectionName")) {
                Add-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID `
                                                          -ResourceId $res.ResourceID `
                                                          -ErrorAction Stop
                Write-Log -Computer $c -Action 'AddToCollection' -Result 'Added'
            }
        } catch {
            Write-Log -Computer $c -Action 'AddToCollection' -Result 'Error' -Detail $_.Exception.Message
        }
    }

    if ($PSCmdlet.ShouldProcess($CollectionName, 'Update collection membership')) {
        Invoke-CMCollectionUpdate -CollectionId $coll.CollectionID -ErrorAction SilentlyContinue
        # Give the collection eval a head start before we start rebooting
        Start-Sleep -Seconds 30
    }
}

function Restart-VDI {
    param([string] $Computer)

    # shutdown.exe is more forgiving than Restart-Computer on locked-down VDIs
    $args = "/m \\$Computer /r /f /t 5 /c `"SCCM patch remediation - automated reboot`""
    try {
        $p = Start-Process -FilePath shutdown.exe -ArgumentList $args -Wait -PassThru -NoNewWindow
        if ($p.ExitCode -eq 0) {
            Write-Log -Computer $Computer -Action 'Reboot' -Result 'Issued'
        } else {
            Write-Log -Computer $Computer -Action 'Reboot' -Result 'Failed' -Detail "shutdown.exe exit $($p.ExitCode)"
        }
    } catch {
        Write-Log -Computer $Computer -Action 'Reboot' -Result 'Error' -Detail $_.Exception.Message
    }
}

function Wait-Online {
    param(
        [string[]] $Computers,
        [int]      $TimeoutMinutes
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $pending  = [System.Collections.Generic.HashSet[string]]::new([string[]]$Computers, [System.StringComparer]::OrdinalIgnoreCase)
    $online   = New-Object System.Collections.Generic.List[string]

    # Wait until pingable AND WMI responds (latter means the SCCM client service is likely up)
    while ($pending.Count -and (Get-Date) -lt $deadline) {
        foreach ($c in @($pending)) {
            if (Test-Connection -ComputerName $c -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                try {
                    $null = Get-CimInstance -ComputerName $c -ClassName Win32_OperatingSystem -ErrorAction Stop -OperationTimeoutSec 5
                    $online.Add($c)
                    $pending.Remove($c) | Out-Null
                } catch {
                    # pingable but WMI not ready — try again next loop
                }
            }
        }
        if ($pending.Count) { Start-Sleep -Seconds 20 }
    }

    foreach ($c in $online)  { Write-Log -Computer $c -Action 'WaitOnline' -Result 'Online' }
    foreach ($c in $pending) { Write-Log -Computer $c -Action 'WaitOnline' -Result 'TimedOut' }

    [pscustomobject]@{ Online = $online; Offline = @($pending) }
}

function Invoke-ClientActions {
    param([string] $Computer)

    $actions = [ordered]@{
        'Machine Policy Retrieval'          = '{00000000-0000-0000-0000-000000000021}'
        'Machine Policy Evaluation'         = '{00000000-0000-0000-0000-000000000022}'
        'Software Updates Scan'             = '{00000000-0000-0000-0000-000000000113}'
        'Software Updates Deployment Eval'  = '{00000000-0000-0000-0000-000000000108}'
        'State Message Refresh'             = '{00000000-0000-0000-0000-000000000111}'
    }

    foreach ($name in $actions.Keys) {
        $sched = $actions[$name]
        try {
            Invoke-CimMethod -ComputerName $Computer `
                             -Namespace 'root\ccm' `
                             -ClassName 'SMS_Client' `
                             -MethodName 'TriggerSchedule' `
                             -Arguments @{ sScheduleID = $sched } `
                             -ErrorAction Stop | Out-Null
            Write-Log -Computer $Computer -Action $name -Result 'Triggered'
            Start-Sleep -Seconds 2     # small space between triggers so the client doesn't coalesce them
        } catch {
            Write-Log -Computer $Computer -Action $name -Result 'Error' -Detail $_.Exception.Message
        }
    }
}

# ---------- main ------------------------------------------------------------

try {
    Write-Host "Connecting to site $SiteCode on $SiteServer ..." -ForegroundColor Cyan
    Connect-CMSite -SiteCode $SiteCode -SiteServer $SiteServer

    # ---- 1. choose deployment(s) ------------------------------------------
    Write-Host "Loading software update deployments ..." -ForegroundColor Cyan
    $deployments = Get-CMUpdateGroupDeployment |
        Select-Object AssignmentID, AssignmentName, TargetCollectionID, StartTime,
                      @{N='Collection';E={(Get-CMDeviceCollection -Id $_.TargetCollectionID).Name}} |
        Sort-Object StartTime -Descending

    if (-not $deployments) { throw "No update deployments found." }

    $selected = $deployments | Out-GridView -Title 'Select one or more update deployments to remediate' -OutputMode Multiple
    if (-not $selected) { Write-Warning 'No deployment selected. Exiting.'; return }

    # ---- 2. ask about Unknown bucket --------------------------------------
    if (-not $PSBoundParameters.ContainsKey('IncludeUnknown')) {
        $ans = Read-Host "Also include machines in 'Unknown' state? (Y/N)"
        $IncludeUnknown = ($ans -match '^[Yy]')
    }

    # ---- 3. pull failed/unknown assets ------------------------------------
    Write-Host "Querying failed$(if($IncludeUnknown){' + unknown'}) assets across $($selected.Count) deployment(s) ..." -ForegroundColor Cyan
    $assets = Get-FailedAndUnknownAssets -AssignmentIDs $selected.AssignmentID `
                                         -SiteCode      $SiteCode `
                                         -SiteServer    $SiteServer `
                                         -IncludeUnknown:$IncludeUnknown

    if (-not $assets) { Write-Warning 'No matching assets found. Exiting.'; return }

    $computers = $assets.Computer | Sort-Object -Unique
    Write-Host ("Found {0} unique machines to remediate." -f $computers.Count) -ForegroundColor Green

    # Confirmation gate
    if (-not $PSCmdlet.ShouldContinue(
            ("About to add {0} machines to '{1}', reboot in batches of {2} every {3} min, and trigger policy/updates. Proceed?" `
                -f $computers.Count, $MaintenanceCollectionName, $BatchSize, $BatchIntervalMinutes),
            'Confirm VDI remediation')) {
        Write-Warning 'Aborted by user.'
        return
    }

    # ---- 4. add to maintenance collection ---------------------------------
    Add-ToMaintenanceCollection -Computers $computers -CollectionName $MaintenanceCollectionName

    # ---- 5. batched reboot + post-reboot triggers -------------------------
    $batches = for ($i = 0; $i -lt $computers.Count; $i += $BatchSize) {
        ,@($computers[$i..([Math]::Min($i + $BatchSize - 1, $computers.Count - 1))])
    }
    Write-Host ("Processing {0} batch(es) of up to {1}." -f $batches.Count, $BatchSize) -ForegroundColor Cyan

    for ($b = 0; $b -lt $batches.Count; $b++) {
        $batch = $batches[$b]
        Write-Host ("`n--- Batch {0}/{1} : {2} machine(s) ---" -f ($b+1), $batches.Count, $batch.Count) -ForegroundColor Yellow

        foreach ($c in $batch) {
            if ($PSCmdlet.ShouldProcess($c, 'Reboot')) { Restart-VDI -Computer $c }
        }

        # Wait for THIS batch to come back, then trigger policy/scan on it.
        # We do this in parallel with kicking off the next batch's wait window
        # so the 5-minute pacing is preserved.
        $waitJob = Start-Job -ScriptBlock {
            param($comp, $timeout, $logPath)
            # Inline minimal wait + trigger inside the job so we don't need module import here.
            $pending = [System.Collections.Generic.HashSet[string]]::new([string[]]$comp, [System.StringComparer]::OrdinalIgnoreCase)
            $deadline = (Get-Date).AddMinutes($timeout)
            while ($pending.Count -and (Get-Date) -lt $deadline) {
                foreach ($x in @($pending)) {
                    if (Test-Connection -ComputerName $x -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                        try {
                            $null = Get-CimInstance -ComputerName $x -ClassName Win32_OperatingSystem -ErrorAction Stop -OperationTimeoutSec 5
                            $pending.Remove($x) | Out-Null
                            $schedules = @(
                                '{00000000-0000-0000-0000-000000000021}',
                                '{00000000-0000-0000-0000-000000000022}',
                                '{00000000-0000-0000-0000-000000000113}',
                                '{00000000-0000-0000-0000-000000000108}',
                                '{00000000-0000-0000-0000-000000000111}'
                            )
                            foreach ($s in $schedules) {
                                try {
                                    Invoke-CimMethod -ComputerName $x -Namespace 'root\ccm' `
                                                     -ClassName 'SMS_Client' -MethodName 'TriggerSchedule' `
                                                     -Arguments @{ sScheduleID = $s } -ErrorAction Stop | Out-Null
                                    Start-Sleep -Seconds 2
                                } catch { }
                            }
                            [pscustomobject]@{ Computer=$x; Result='Triggered' }
                        } catch { }
                    }
                }
                if ($pending.Count) { Start-Sleep -Seconds 20 }
            }
            foreach ($x in $pending) { [pscustomobject]@{ Computer=$x; Result='TimedOut' } }
        } -ArgumentList (,$batch), $OnlineWaitMinutes, $LogPath

        $waitJob | Add-Member -NotePropertyName BatchIndex -NotePropertyValue $b -Force

        if ($b -lt $batches.Count - 1) {
            Write-Host ("Waiting {0} minute(s) before next batch ..." -f $BatchIntervalMinutes) -ForegroundColor DarkGray
            Start-Sleep -Seconds ($BatchIntervalMinutes * 60)
        }
    }

    Write-Host "`nAll reboots issued. Waiting for outstanding post-reboot jobs ..." -ForegroundColor Cyan
    Get-Job | Where-Object State -in 'Running','NotStarted' | Wait-Job | Out-Null
    foreach ($j in Get-Job) {
        $results = Receive-Job -Job $j -ErrorAction SilentlyContinue
        foreach ($r in $results) {
            Write-Log -Computer $r.Computer -Action 'PostRebootTriggers' -Result $r.Result
        }
        Remove-Job -Job $j -Force
    }

    # ---- 6. report --------------------------------------------------------
    $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
    $script:Log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nDone. Log: $LogPath" -ForegroundColor Green

    $summary = $script:Log | Group-Object Action, Result | Select-Object Count, Name | Sort-Object Name
    $summary | Format-Table -AutoSize
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
