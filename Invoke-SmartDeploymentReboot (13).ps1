#Requires -Version 7.0
<#
.SYNOPSIS
    Smart Deployment Reboot — multi-deployment picker with deduplication.
.DESCRIPTION
    Queries software update deployments via the SMS Provider, lets you pick
    one or more in a grid view, pulls all Failed and Unknown machines,
    deduplicates them, checks for logged-on users, and reboots the unattended
    ones in parallel. Creates an audit log and retry list.

    STATUS TYPE MAPPING (from MECM SDK):
        1 = Success
        2 = In Progress
        4 = Unknown
        5 = Error / Failed
.PARAMETER SiteServer
    MECM Site Server hostname (e.g. "SCCM01")
.PARAMETER SiteCode
    MECM Site Code (e.g. "P01")
.PARAMETER ThrottleLimit
    Parallel threads for user check / reboot (default: 25)
.PARAMETER LogPath
    Where to save logs (default: script directory)
.EXAMPLE
    .\Invoke-SmartDeploymentReboot.ps1 -SiteServer "SCCM01" -SiteCode "P01"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteServer,
    [Parameter(Mandatory)][string]$SiteCode,
    [int]$ThrottleLimit = 25,
    [string]$LogPath    = $PSScriptRoot,
    [string]$MaintenanceCollectionName = "VDI Maintenance Anytime"
)

$Namespace = "root\SMS\site_$SiteCode"

#region ── Header ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor Cyan
Write-Host "  MECM Smart Deployment Reboot" -ForegroundColor Cyan
Write-Host "  Site Server : $SiteServer"    -ForegroundColor White
Write-Host "  Site Code   : $SiteCode"      -ForegroundColor White
Write-Host "  Started     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host ("═" * 70) -ForegroundColor Cyan
Write-Host ""

#endregion

#region ── Step 1: Load and pick deployments ──────────────────────────────────

Write-Host "  Step 1/4 — Loading software update deployments..." -ForegroundColor Cyan

try {
    # SMS_DeploymentSummary — all deployments at the site
    # FeatureType=5 filters to Software Updates only
    $RawDeployments = Get-CimInstance -ComputerName $SiteServer `
                          -Namespace $Namespace `
                          -ClassName "SMS_DeploymentSummary" `
                          -Filter "FeatureType=5" `
                          -ErrorAction Stop

    Write-Host "  Software Update deployments found: $($RawDeployments.Count)" -ForegroundColor DarkGray

} catch {
    Write-Error "Failed to query SMS_DeploymentSummary: $_"
    exit 1
}

if (-not $RawDeployments -or $RawDeployments.Count -eq 0) {
    Write-Warning "No software update deployments found."
    exit 0
}

# Normalise for grid — include all useful columns the console shows
$DeploymentList = $RawDeployments |
    Select-Object `
        @{N="DeploymentID";     E={$_.DeploymentID}},
        @{N="SoftwareName";     E={$_.SoftwareName}},
        @{N="CollectionName";   E={$_.CollectionName}},
        @{N="Failed";           E={[int]$_.NumberErrors}},
        @{N="Unknown";          E={[int]$_.NumberUnknown}},
        @{N="Compliant";        E={[int]$_.NumberCompliant}},
        @{N="InProgress";       E={[int]$_.NumberInProgress}},
        @{N="Other";            E={[int]$_.NumberOther}},
        @{N="Targeted";         E={[int]$_.NumberTargeted}},
        @{N="Total";            E={[int]$_.NumberTotal}},
        @{N="CreationTime";     E={$_.CreationTime}},
        @{N="ModificationTime"; E={$_.ModificationTime}},
        @{N="DeploymentTime";   E={$_.DeploymentTime}} |
    Sort-Object CreationTime -Descending

Write-Host "  Opening picker — Ctrl+Click to select multiple deployments" -ForegroundColor Green
Write-Host ""

$SelectedDeployments = $DeploymentList | Out-GridView `
    -Title "STEP 1 — SELECT DEPLOYMENTS  |  Ctrl+Click to pick multiple  |  OK to continue" `
    -OutputMode Multiple

if (-not $SelectedDeployments -or $SelectedDeployments.Count -eq 0) {
    Write-Warning "No deployments selected. Exiting."
    exit 0
}

Write-Host "  Selected $($SelectedDeployments.Count) deployment(s):" -ForegroundColor Green
foreach ($D in $SelectedDeployments) {
    Write-Host "    → $($D.SoftwareName) [$($D.CollectionName)]  Failed:$($D.Failed)  Unknown:$($D.Unknown)" -ForegroundColor DarkGray
}
Write-Host ""

#endregion

#region ── Step 2: Resolve numeric AssignmentIDs and query failed/unknown machines

Write-Host "  Step 2/4 — Querying failed and unknown machines..." -ForegroundColor Cyan

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
        # KEY INSIGHT: SMS_DeploymentSummary.DeploymentID is a GUID string like
        # "{c1123130-b58c-4476-864d-c20b994f2a00}", but SMS_SUMDeploymentAssetDetails
        # uses a numeric AssignmentID. We must resolve the GUID → numeric ID first
        # via SMS_UpdateGroupAssignment which has both properties.

        $Assignment = Get-CimInstance -ComputerName $SiteServer `
                          -Namespace $Namespace `
                          -ClassName "SMS_UpdateGroupAssignment" `
                          -Filter "AssignmentUniqueID='$($Dep.DeploymentID)'" `
                          -ErrorAction SilentlyContinue

        if (-not $Assignment) {
            Write-Host "      Could not resolve numeric AssignmentID from GUID $($Dep.DeploymentID)" -ForegroundColor Yellow
            continue
        }

        $NumericAssignmentID = $Assignment.AssignmentID

        # Query per-machine status using numeric AssignmentID
        # StatusType: 1=Success, 2=InProgress, 4=Unknown, 5=Error
        # We want 4 and 5 only
        $StatusData = Get-CimInstance -ComputerName $SiteServer `
                          -Namespace $Namespace `
                          -ClassName "SMS_SUMDeploymentAssetDetails" `
                          -Filter "AssignmentID=$NumericAssignmentID AND (StatusType=4 OR StatusType=5)" `
                          -ErrorAction SilentlyContinue

        if (-not $StatusData) {
            Write-Host "      No failed/unknown machines in this deployment" -ForegroundColor DarkGray
            continue
        }

        $FailCount    = 0
        $UnknownCount = 0

        foreach ($Entry in $StatusData) {
            $Name = $Entry.DeviceName
            if (-not $Name) { continue }

            $IsFailed = $Entry.StatusType -eq 5

            $ErrorStr = if ($IsFailed) {
                            if ($Entry.StatusErrorCode) {
                                "Failed 0x{0:X8}" -f $Entry.StatusErrorCode
                            } else {
                                "Failed"
                            }
                        } else {
                            "Unknown"
                        }

            # Initialise maps for this machine if first time seen
            if (-not $MachineDeploymentMap.ContainsKey($Name)) {
                $MachineDeploymentMap[$Name] = [System.Collections.Generic.List[string]]::new()
                $MachineErrorMap[$Name]      = [System.Collections.Generic.List[string]]::new()
            }

            # Deduplicate deployment name per machine
            if (-not $MachineDeploymentMap[$Name].Contains($Dep.SoftwareName)) {
                $MachineDeploymentMap[$Name].Add($Dep.SoftwareName)
            }
            $MachineErrorMap[$Name].Add("$($Dep.SoftwareName) [$ErrorStr]")

            if ($IsFailed) { $FailCount++ } else { $UnknownCount++ }
        }

        Write-Host "      Failed: $FailCount  |  Unknown: $UnknownCount" -ForegroundColor $(if ($FailCount -gt 0 -or $UnknownCount -gt 0) {"Yellow"} else {"DarkGray"})

    } catch {
        Write-Host "      ERROR: $_" -ForegroundColor Red
    }
}

Write-Host ""

if ($MachineDeploymentMap.Count -eq 0) {
    Write-Host "  No failed or unknown machines found across selected deployments." -ForegroundColor Green
    exit 0
}

# Build deduplicated list
$DeduplicatedRaw = foreach ($KV in $MachineDeploymentMap.GetEnumerator()) {
    [PSCustomObject]@{
        MachineName         = $KV.Key
        FailedInCount       = $KV.Value.Count
        DuplicateFlag       = if ($KV.Value.Count -gt 1) { "YES" } else { "" }
        FailedInDeployments = $KV.Value -join " | "
        ErrorDetail         = $MachineErrorMap[$KV.Key] -join " | "
    }
}

$DeduplicatedList = $DeduplicatedRaw |
    Sort-Object -Property @{Expression="FailedInCount";Descending=$true}, @{Expression="MachineName";Descending=$false}

$TotalDupes = ($DeduplicatedList | Where-Object { $_.FailedInCount -gt 1 }).Count

Write-Host "  Unique machines        : $($DeduplicatedList.Count)" -ForegroundColor Green
Write-Host "  In multiple deployments: $TotalDupes" -ForegroundColor $(if ($TotalDupes -gt 0) {"Yellow"} else {"DarkGray"})
Write-Host ""

#endregion

#region ── Step 3: Confirm machines and choose action ─────────────────────────

Write-Host "  Step 3/4 — Confirm machines..." -ForegroundColor Cyan
Write-Host ""

$SelectedMachines = $DeduplicatedList | Out-GridView `
    -Title "STEP 3 — CONFIRM MACHINES  |  $($DeduplicatedList.Count) unique  |  Ctrl+Click to deselect  |  OK to continue" `
    -OutputMode Multiple

if (-not $SelectedMachines -or $SelectedMachines.Count -eq 0) {
    Write-Warning "No machines selected. Exiting."
    exit 0
}

Write-Host "  $($SelectedMachines.Count) machines confirmed." -ForegroundColor Green
Write-Host ""

# ── Choose what to do with these machines ────────────────────────────────────
# Use Out-GridView for consistent GUI-based input instead of Read-Host which
# can behave inconsistently across PowerShell hosts.

$ActionOptions = @(
    [PSCustomObject]@{
        Action      = "1. Reboot only"
        Description = "Reboot unattended machines, no collection change"
    }
    [PSCustomObject]@{
        Action      = "2. Add to collection"
        Description = "Add to '$MaintenanceCollectionName', no reboot"
    }
    [PSCustomObject]@{
        Action      = "3. Both"
        Description = "Add to collection AND reboot unattended machines"
    }
)

Write-Host "  Opening action picker..." -ForegroundColor Cyan

$ActionResult = $ActionOptions | Out-GridView `
    -Title "CHOOSE ACTION for $($SelectedMachines.Count) machines  |  Select one option and click OK  |  Cancel to exit" `
    -OutputMode Single

if (-not $ActionResult) {
    Write-Host "  No action chosen. Exiting." -ForegroundColor Yellow
    exit 0
}

$DoReboot     = $ActionResult.Action -match '^(1|3)\.'
$DoCollection = $ActionResult.Action -match '^(2|3)\.'

Write-Host "  Selected: $($ActionResult.Action)" -ForegroundColor Green
Write-Host ""

#endregion

#region ── Step 4: Check users + reboot (if enabled) ──────────────────────────

if ($DoReboot) {

    Write-Host "  Step 4 — Checking logged-on users and rebooting unattended machines..." -ForegroundColor Cyan
    Write-Host ""

    $Results = $SelectedMachines | ForEach-Object -Parallel {

    $Machine = $_

    $Result = [PSCustomObject]@{
        MachineName         = $Machine.MachineName
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

    # Reachability — TCP 135 (WMI/RPC endpoint)
    try {
        $Tcp     = [System.Net.Sockets.TcpClient]::new()
        $Connect = $Tcp.BeginConnect($Machine.MachineName, 135, $null, $null)
        $Wait    = $Connect.AsyncWaitHandle.WaitOne(1000, $false)
        if ($Wait -and $Tcp.Connected) { $Result.Reachable = $true }
        $Tcp.Close()
    } catch {}

    if (-not $Result.Reachable) {
        $Result.Outcome     = "Offline / unreachable"
        $Result.RetryNeeded = $true
        return $Result
    }

    # Logged-on user
    try {
        $CS   = Get-CimInstance -ClassName Win32_ComputerSystem `
                    -ComputerName $Machine.MachineName -ErrorAction Stop
        $User = $CS.UserName

        if ($User -and $User -match '\S') {
            $Result.LoggedOnUser = $User
            $Result.UserPresent  = $true
            $Result.Outcome      = "User logged on: $User"
            $Result.RetryNeeded  = $true
            return $Result
        }
    } catch {
        $Result.Outcome     = "WMI failed: $($_.Exception.Message)"
        $Result.RetryNeeded = $true
        return $Result
    }

    # No user — reboot (Flags 6 = forced reboot)
    try {
        $Reboot = Invoke-CimMethod -ClassName Win32_OperatingSystem `
                      -ComputerName $Machine.MachineName `
                      -MethodName Win32Shutdown `
                      -Arguments @{ Flags = 6 } `
                      -ErrorAction Stop

        if ($Reboot.ReturnValue -eq 0) {
            $Result.Rebooted   = $true
            $Result.RebootTime = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            $Result.Outcome    = "Rebooted successfully"
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

} else {
    # Collection-only mode — build stub Results so downstream code still works
    Write-Host "  Step 4 — Skipped (no reboot requested)" -ForegroundColor DarkGray
    Write-Host ""

    $Results = $SelectedMachines | ForEach-Object {
        [PSCustomObject]@{
            MachineName         = $_.MachineName
            FailedInCount       = $_.FailedInCount
            DuplicateFlag       = $_.DuplicateFlag
            FailedInDeployments = $_.FailedInDeployments
            ErrorDetail         = $_.ErrorDetail
            Reachable           = $null
            LoggedOnUser        = $null
            UserPresent         = $false
            Rebooted            = $false
            RebootTime          = $null
            Outcome             = "Reboot skipped (collection-only mode)"
            RetryNeeded         = $false
        }
    }
}

#endregion

#region ── Console output ─────────────────────────────────────────────────────

$LineWidth = 100

if ($DoReboot) {

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

        $Deps    = if ($R.FailedInCount -gt 1) { "! $($R.FailedInCount)" } else { "  1" }
        $User    = if ($R.LoggedOnUser)        { $R.LoggedOnUser        } else { "-"   }

        Write-Host ("  {0,-25}" -f $R.MachineName) -NoNewline
        Write-Host ("{0,-10}"   -f $Display)       -NoNewline -ForegroundColor $Colour
        Write-Host ("{0,-6}"    -f $Deps)          -NoNewline -ForegroundColor $(if ($R.FailedInCount -gt 1) {"Yellow"} else {"DarkGray"})
        Write-Host ("{0,-22}"   -f $User)          -NoNewline -ForegroundColor $(if ($R.UserPresent) {"Yellow"} else {"DarkGray"})
        Write-Host ("{0}"       -f $R.Outcome)                -ForegroundColor $Colour
    }

    $Rebooted = ($Results | Where-Object { $_.Rebooted }).Count
    $UserOn   = ($Results | Where-Object { $_.UserPresent }).Count
    $Offline  = ($Results | Where-Object { -not $_.Reachable }).Count
    $Failed   = ($Results | Where-Object { -not $_.Rebooted -and $_.Reachable -and -not $_.UserPresent }).Count
    $Retry    = ($Results | Where-Object { $_.RetryNeeded }).Count
    $MultiDep = ($Results | Where-Object { $_.FailedInCount -gt 1 }).Count

    Write-Host ""
    Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
    Write-Host "  REBOOT SUMMARY" -ForegroundColor Cyan
    Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
    Write-Host "  Deployments selected       : $($SelectedDeployments.Count)"
    Write-Host "  Rebooted                   : $Rebooted"  -ForegroundColor Green
    Write-Host "  Skipped (user logged on)   : $UserOn"    -ForegroundColor Yellow
    Write-Host "  Offline                    : $Offline"   -ForegroundColor DarkGray
    Write-Host "  Failed                     : $Failed"    -ForegroundColor $(if ($Failed -gt 0) {"Red"} else {"Green"})
    Write-Host "  Retry list                 : $Retry"     -ForegroundColor $(if ($Retry -gt 0) {"Yellow"} else {"Green"})
    Write-Host "  Multi-deployment machines  : $MultiDep"  -ForegroundColor $(if ($MultiDep -gt 0) {"Yellow"} else {"Green"})
    Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
    Write-Host ""
}

#endregion

#region ── Add to maintenance collection (optional) ───────────────────────────
# Prompt the user to add all processed machines to a maintenance collection
# so any that were skipped (user on, offline) can be picked up later by the
# collection's existing maintenance deployments/policies.

Write-Host ""

if ($DoCollection) {

    Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
    Write-Host "  COLLECTION MEMBERSHIP" -ForegroundColor Cyan
    Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
    Write-Host "  Target collection: $MaintenanceCollectionName" -ForegroundColor White
    Write-Host "  Adding $($Results.Count) machines..." -ForegroundColor Cyan
    Write-Host ""

    # Load ConfigMgr module for Add-CMDeviceCollectionDirectMembershipRule
    try {
        $ModulePath = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
        if (-not (Test-Path $ModulePath)) {
            Write-Host "  ERROR: ConfigMgr PS module not found at $ModulePath" -ForegroundColor Red
            Write-Host "  Install the MECM console on this machine to use the collection feature." -ForegroundColor Red
        } else {

            Import-Module $ModulePath -ErrorAction Stop

            if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer | Out-Null
            }

            $OriginalLocation = Get-Location
            Set-Location "$($SiteCode):\" -ErrorAction Stop

            # Verify the collection exists
            $TargetCollection = Get-CMDeviceCollection -Name $MaintenanceCollectionName -ErrorAction SilentlyContinue

            if (-not $TargetCollection) {
                Write-Host "  ERROR: Collection '$MaintenanceCollectionName' not found." -ForegroundColor Red
                Set-Location $OriginalLocation
            } else {

                Write-Host "  Resolved CollectionID: $($TargetCollection.CollectionID)" -ForegroundColor DarkGray
                Write-Host ""

                # Pre-cache existing direct members to avoid duplicate add errors
                $ExistingMembers = @{}
                try {
                    $CurrentMembers = Get-CMCollectionMember -CollectionName $MaintenanceCollectionName -ErrorAction SilentlyContinue
                    if ($CurrentMembers) {
                        foreach ($M in $CurrentMembers) { $ExistingMembers[$M.Name] = $true }
                    }
                } catch {}

                $Added       = 0
                $AlreadyIn   = 0
                $NotFound    = 0
                $FailedToAdd = 0

                foreach ($R in $Results) {

                    $MachineName = $R.MachineName

                    # Skip if already in collection
                    if ($ExistingMembers.ContainsKey($MachineName)) {
                        Write-Host "    $($MachineName.PadRight(25)) already in collection" -ForegroundColor DarkGray
                        $AlreadyIn++
                        continue
                    }

                    # Check the machine exists as a device in MECM
                    $Device = Get-CMDevice -Name $MachineName -Fast -ErrorAction SilentlyContinue |
                              Select-Object -First 1

                    if (-not $Device) {
                        Write-Host "    $($MachineName.PadRight(25)) NOT FOUND in MECM" -ForegroundColor DarkGray
                        $NotFound++
                        continue
                    }

                    # Add direct membership rule
                    try {
                        Add-CMDeviceCollectionDirectMembershipRule `
                            -CollectionName $MaintenanceCollectionName `
                            -ResourceId     $Device.ResourceID `
                            -ErrorAction    Stop

                        Write-Host "    $($MachineName.PadRight(25)) added" -ForegroundColor Green
                        $Added++
                        $ExistingMembers[$MachineName] = $true

                    } catch {
                        Write-Host "    $($MachineName.PadRight(25)) FAILED: $($_.Exception.Message)" -ForegroundColor Red
                        $FailedToAdd++
                    }
                }

                Write-Host ""
                Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
                Write-Host "  Collection update summary:" -ForegroundColor Cyan
                Write-Host "    Added to collection    : $Added"       -ForegroundColor Green
                Write-Host "    Already in collection  : $AlreadyIn"   -ForegroundColor DarkGray
                Write-Host "    Not found in MECM      : $NotFound"    -ForegroundColor Yellow
                Write-Host "    Failed to add          : $FailedToAdd" -ForegroundColor $(if ($FailedToAdd -gt 0) {"Red"} else {"Green"})
                Write-Host ("─" * $LineWidth) -ForegroundColor DarkGray
                Write-Host ""

                if ($Added -gt 0) {
                    Write-Host "  Note: Collection membership updates on next evaluation cycle." -ForegroundColor Yellow
                    Write-Host "  To force it now: right-click the collection in MECM console → 'Update Membership'." -ForegroundColor Yellow
                    Write-Host ""
                }

                # Restore filesystem location
                Set-Location $OriginalLocation
            }
        }

    } catch {
        Write-Host "  ERROR updating collection: $_" -ForegroundColor Red
        if ($OriginalLocation) { Set-Location $OriginalLocation }
    }

}

#endregion

#region ── Export ─────────────────────────────────────────────────────────────

$Timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$DepSafeNames = ($SelectedDeployments | Select-Object -First 2 -ExpandProperty SoftwareName |
                 ForEach-Object { $_ -replace '[\\/:*?"<>|]','-' }) -join "_"
if ($SelectedDeployments.Count -gt 2) {
    $DepSafeNames += "_plus$($SelectedDeployments.Count - 2)more"
}

$LogPrefix = if ($DoReboot -and $DoCollection) { "RebootAndAdd"   }
             elseif ($DoReboot)                { "RebootAudit"    }
             elseif ($DoCollection)            { "CollectionAdd"  }
             else                              { "Audit"          }

$AuditLog = Join-Path $LogPath "${LogPrefix}_${DepSafeNames}_$Timestamp.csv"
$Results | Export-Csv -Path $AuditLog -NoTypeInformation -Encoding UTF8
Write-Host "  Audit log    : $AuditLog" -ForegroundColor Cyan

# Only produce a retry list when reboots were attempted
if ($DoReboot) {
    $RetryMachines = $Results | Where-Object { $_.RetryNeeded }
    if ($RetryMachines.Count -gt 0) {
        $RetryTxt = Join-Path $LogPath "RetryList_${DepSafeNames}_$Timestamp.txt"
        $RetryMachines.MachineName | Out-File -FilePath $RetryTxt -Encoding UTF8
        Write-Host "  Retry list   : $RetryTxt  ($($RetryMachines.Count) machines)" -ForegroundColor Yellow

        $RetryDetailCsv = Join-Path $LogPath "RetryDetail_${DepSafeNames}_$Timestamp.csv"
        $RetryMachines |
            Select-Object MachineName, FailedInCount, DuplicateFlag, FailedInDeployments, ErrorDetail, LoggedOnUser, Outcome |
            Export-Csv -Path $RetryDetailCsv -NoTypeInformation -Encoding UTF8
        Write-Host "  Retry detail : $RetryDetailCsv" -ForegroundColor Yellow
    } else {
        Write-Host "  No retry list needed." -ForegroundColor Green
    }
}

Write-Host ""

#endregion
