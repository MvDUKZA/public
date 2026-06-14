#Requires -Version 5.1
<#
.SYNOPSIS
    End-to-end VDI patch remediation orchestrator.
    Combines MECM deployment status discovery with per-machine WUA/CBS/DISM
    remediation, batched reboots, and post-reboot verification.

.DESCRIPTION
    FLOW PER MACHINE
    ────────────────
    1.  MECM deployment status → identify Failed / Unknown assets
    2.  Add to 'VDI Maintenance Anytime' collection (grants Anytime MW)
    3.  Logged-on user check (unless -SkipLoggedOnCheck)
    4.  PRE-REBOOT REMEDIATION
          Delivers Invoke-VDINodeRemediation scriptblock to each VDI via:
            a) Invoke-CMScript  (MECM Run Script — stays in audit trail)
            b) Invoke-Command   (PSRemoting fallback)
          Scriptblock: reads actual WU/CCM error codes, runs targeted fixes
          (WUA reset, DISM, ACL repair, disk cleanup, DataStore clear)
    5.  REBOOT — parallel within batch, staggered across batches
    6.  Wait for machines to come back (ICMP + WMI confirm)
    7.  POST-REBOOT VERIFICATION
          Re-runs the remediation scriptblock. If error codes are gone → done.
          If still failing → runs remediation again (second pass).
    8.  MECM policy triggers: Machine Policy + SU Scan + SU Deploy Eval
    9.  CSV report of every action on every machine

.PARAMETER SiteCode
    MECM site code. Default: PRD

.PARAMETER SiteServer
    MECM site server FQDN. Default: appsmcm101fp.iprod.local

.PARAMETER MaintenanceCollectionName
    Collection granting the Anytime maintenance window.
    Default: 'VDI Maintenance Anytime'

.PARAMETER CMScriptName
    Name of the approved MECM Run Script that wraps Invoke-VDINodeRemediation.
    If blank or not found, falls back to PSRemoting automatically.
    Default: 'VDI-PatchRemediation'

.PARAMETER BatchSize
    Machines per reboot wave. Default: 20

.PARAMETER BatchIntervalMinutes
    Minutes between waves. Default: 5

.PARAMETER OnlineWaitMinutes
    Max minutes per batch waiting for machines to return. Default: 15

.PARAMETER MinFreeGB
    Minimum C: free space required. Cleanup runs if below this. Default: 20

.PARAMETER IncludeUnknown
    Also remediate Unknown state machines (prompted interactively if omitted).

.PARAMETER SkipLoggedOnCheck
    Reboot regardless of logged-on users (VDI patch windows only).

.PARAMETER LogPath
    CSV output path. Defaults to C:\Temp\VDIPatchOrchestrator_<timestamp>.csv

.EXAMPLE
    # Dry run — see what would happen without touching anything
    .\Invoke-VDIPatchOrchestrator.ps1 -WhatIf -Verbose

.EXAMPLE
    # Production run, include Unknown, 10-machine batches
    .\Invoke-VDIPatchOrchestrator.ps1 -IncludeUnknown -BatchSize 10 -SkipLoggedOnCheck
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string] $SiteCode                  = 'PRD',
    [string] $SiteServer                = 'appsmcm101fp.iprod.local',
    [string] $MaintenanceCollectionName = 'VDI Maintenance Anytime',
    [string] $CMScriptName              = 'VDI-PatchRemediation',
    [int]    $BatchSize                 = 20,
    [int]    $BatchIntervalMinutes      = 5,
    [int]    $OnlineWaitMinutes         = 15,
    [int]    $MinFreeGB                 = 20,
    [switch] $IncludeUnknown,
    [switch] $SkipLoggedOnCheck,
    [string] $LogPath = "C:\Temp\VDIPatchOrchestrator_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
#  SECTION 0 — Report infrastructure
# ============================================================================

$script:Report = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

function Add-Report {
    param(
        [string]$Computer,
        [string]$Phase,
        [string]$Action,
        [string]$Result,
        [string]$Detail = ''
    )
    $script:Report.Add([pscustomobject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Computer  = $Computer
        Phase     = $Phase
        Action    = $Action
        Result    = $Result
        Detail    = $Detail
    })
    Write-Verbose "[$Computer][$Phase] $Action => $Result$(if($Detail){" | $Detail"})"
}

function Save-Report {
    $null = New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force -ErrorAction SilentlyContinue
    $script:Report |
        Sort-Object Timestamp |
        Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport saved: $LogPath" -ForegroundColor Green

    Write-Host "`n─── Summary ──────────────────────────────────────────" -ForegroundColor Cyan
    $script:Report |
        Group-Object Phase, Action, Result |
        Sort-Object Name |
        ForEach-Object { Write-Host ("  {0,4}  {1}" -f $_.Count, $_.Name) }
    Write-Host "──────────────────────────────────────────────────────" -ForegroundColor Cyan
}

# ============================================================================
#  SECTION 1 — MECM site connection
# ============================================================================

function Connect-CMSite {
    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw "SMS_ADMIN_UI_PATH not set. ConfigMgr console must be installed on this machine."
    }
    $module = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    if (-not (Get-Module ConfigurationManager)) {
        Write-Host "Importing ConfigurationManager module ..." -ForegroundColor Cyan
        Import-Module $module -ErrorAction Stop
    }
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
    }
    Push-Location "$SiteCode`:\"
    Write-Host "Connected to site $SiteCode on $SiteServer" -ForegroundColor Green
}

# ============================================================================
#  SECTION 2 — Deployment selection (GridView picker)
# ============================================================================

function Select-Deployments {
    Write-Host "Loading software update deployments ..." -ForegroundColor Cyan
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
        @{N='DeploymentID';     E={$_.DeploymentID}}

    $picked = $display |
              Out-GridView -Title 'Select deployments to remediate  (Ctrl+Click for multiple)' `
                           -OutputMode Multiple
    if (-not $picked) { return $null }

    $pickedIDs = @($picked.DeploymentID)
    @($all | Where-Object { $_.DeploymentID -in $pickedIDs })
}

# ============================================================================
#  SECTION 3 — Get failed / unknown assets
#  Chain: Get-CMDeployment → Get-CMSoftwareUpdateDeployment
#           → Get-CMSoftwareUpdateDeploymentStatus
#             → Get-CMDeploymentStatusDetails (StatusType 4=Unknown, 5=Error)
# ============================================================================

function Get-FailedAssets {
    param(
        [array] $Deployments,
        [bool]  $IncludeUnknown
    )
    $wantedTypes = @(5)
    if ($IncludeUnknown) { $wantedTypes += 4 }

    $allAssets = foreach ($dep in $Deployments) {
        $depName = $dep.SoftwareName
        $depGuid = $dep.DeploymentID
        Write-Host "  Querying: $depName" -ForegroundColor Cyan

        $suDep = Get-CMSoftwareUpdateDeployment -DeploymentId $depGuid -ErrorAction Stop
        if (-not $suDep) { Write-Warning "No SU deployment for $depName ($depGuid)"; continue }

        $summaries = Get-CMSoftwareUpdateDeploymentStatus -InputObject $suDep -ErrorAction Stop
        if (-not $summaries) { Write-Warning "No status rows for $depName"; continue }

        foreach ($summary in @($summaries)) {
            $details = Get-CMDeploymentStatusDetails -InputObject $summary -ErrorAction SilentlyContinue
            if (-not $details) { continue }

            $details | Where-Object { $_.StatusType -in $wantedTypes } |
                Select-Object `
                    @{N='Deployment';       E={$depName}},
                    @{N='Computer';         E={$_.DeviceName}},
                    @{N='Status';           E={switch($_.StatusType){4{'Unknown'}5{'Failed'}default{"Type$($_.StatusType)"}}}},
                    @{N='StatusDescription';E={$_.StatusDescription}},
                    @{N='LastStatusTime';   E={$_.StatusTime}}
        }
    }

    # Deduplicate — keep most recent row per machine
    @($allAssets) |
        Where-Object { $_.Computer } |
        Sort-Object Computer, LastStatusTime -Descending |
        Group-Object Computer |
        ForEach-Object { $_.Group | Select-Object -First 1 }
}

# ============================================================================
#  SECTION 4 — Logged-on user check
# ============================================================================

function Test-UserLoggedOn {
    param([string]$Computer)
    try {
        $s = Get-CimInstance -ComputerName $Computer -ClassName Win32_LogonSession `
                             -Filter 'LogonType=2 OR LogonType=10' `
                             -ErrorAction Stop -OperationTimeoutSec 10
        return (@($s).Count -gt 0)
    } catch {
        return $false  # unreachable = assume no user
    }
}

# ============================================================================
#  SECTION 5 — Maintenance collection membership
# ============================================================================

function Add-ToMaintenanceCollection {
    param([string[]]$Computers)

    $coll = Get-CMDeviceCollection -Name $MaintenanceCollectionName -ErrorAction Stop
    if (-not $coll) { throw "Collection '$MaintenanceCollectionName' not found." }

    $existing = @(
        Get-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID |
        Select-Object -ExpandProperty RuleName
    )

    foreach ($c in $Computers) {
        if ($c -in $existing) {
            Add-Report $c 'Collection' 'AddMember' 'AlreadyMember'; continue
        }
        try {
            $dev = Get-CMDevice -Name $c -Fast -ErrorAction Stop
            if (-not $dev) { Add-Report $c 'Collection' 'AddMember' 'DeviceNotFound'; continue }
            if ($PSCmdlet.ShouldProcess($c, "Add to '$MaintenanceCollectionName'")) {
                Add-CMDeviceCollectionDirectMembershipRule -CollectionId $coll.CollectionID `
                    -ResourceId $dev.ResourceID -Confirm:$false -ErrorAction Stop
                Add-Report $c 'Collection' 'AddMember' 'Added'
            }
        } catch {
            Add-Report $c 'Collection' 'AddMember' 'Error' $_.Exception.Message
        }
    }

    if ($PSCmdlet.ShouldProcess($MaintenanceCollectionName, 'Refresh membership')) {
        Invoke-CMCollectionUpdate -CollectionId $coll.CollectionID -ErrorAction SilentlyContinue
        Write-Host "  Collection update triggered — waiting 30s ..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 30
    }
}

# ============================================================================
#  SECTION 6 — Per-node remediation scriptblock
#              Runs on the VDI itself (via MECM Run Script or PSRemoting)
#              Self-contained — no external dependencies
# ============================================================================

$script:NodeRemediationBlock = {
    param(
        [int]    $MinFreeGB      = 20,
        [bool]   $ScheduleReboot = $false,   # $true = pre-reboot pass (schedule own reboot)
                                              # $false = post-reboot pass (no reboot, just fix+trigger)
        [int]    $RebootDelaySec = 90
    )

    $logPath = 'C:\Windows\Temp\VDIPatchRemediation.log'

    function wl {
        param([string]$m, [string]$l='INFO')
        $e = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$l] $m"
        Add-Content $logPath $e -Encoding UTF8
        $e  # returned to caller / MECM output
    }

    # ── Error code constants (uint32) ──────────────────────────────────────
    $EC = @{
        Shutdown     = [uint32]'0x8007045B'
        PendingReboot= [uint32]'0x87D00651'
        AllUpdates   = [uint32]'0x80240022'
        Unexpected   = [uint32]'0x8000FFFF'
        CBSTrans     = [uint32]'0x800F0820'
        CompStore    = [uint32]'0x80073712'
        MissingDll   = [uint32]'0x8007007E'
        AccessDenied = [uint32]'0x80070005'
        DiskFull     = [uint32]'0x80070070'
        Superseded   = [uint32]'0x8007066A'
        KeyNotFound  = [uint32]'0x80240008'
        DataContract = [uint32]'0x80240439'
    }

    # ── Read actual error codes from this machine ──────────────────────────
    function Get-UpdateErrors {
        $codes = [System.Collections.Generic.HashSet[uint32]]::new()

        # CCM_SoftwareUpdate WMI — MECM's own error record per update article
        try {
            Get-WmiObject -Namespace 'ROOT\ccm\SoftMgmtAgent' -Class 'CCM_SoftwareUpdate' -EA Stop |
                Where-Object { $_.ErrorCode -and $_.ErrorCode -ne 0 } |
                ForEach-Object {
                    $c = [uint32]$_.ErrorCode
                    [void]$codes.Add($c)
                    wl "CCM_SoftwareUpdate KB$($_.ArticleID): 0x$($c.ToString('X8'))"
                }
        } catch { wl "CCM_SoftwareUpdate query failed: $_" 'WARN' }

        # UpdatesDeployment.log — catches codes WMI may have aged out
        try {
            $log = "$env:SystemRoot\CCM\Logs\UpdatesDeployment.log"
            if (Test-Path $log) {
                (Get-Content $log -Tail 500 -EA SilentlyContinue) |
                    Select-String '0x[0-9A-Fa-f]{8}' -AllMatches |
                    ForEach-Object { $_.Matches } |
                    ForEach-Object {
                        try {
                            $c = [uint32]$_.Value
                            if ($c -ne 0 -and $c -ne 0x80070000) { [void]$codes.Add($c) }
                        } catch {}
                    }
            }
        } catch { wl "UpdatesDeployment.log parse failed: $_" 'WARN' }

        # WU registry LastError — final fallback
        try {
            $k = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install'
            if (Test-Path $k) {
                $e = (Get-ItemProperty $k -EA SilentlyContinue).LastError
                if ($e -and $e -ne 0) { [void]$codes.Add([uint32]$e) }
            }
        } catch {}

        return @($codes)
    }

    # ── Disk cleanup ───────────────────────────────────────────────────────
    function Invoke-DiskCleanup {
        wl "Running disk cleanup (target: ${MinFreeGB}GB free)..."
        Remove-Item "$env:SystemRoot\Temp\*" -Recurse -Force -EA SilentlyContinue
        Remove-Item "$env:TEMP\*"            -Recurse -Force -EA SilentlyContinue
        Remove-Item "$env:SystemRoot\Logs\CBS\*.cab" -Force -EA SilentlyContinue

        # SoftwareDistribution\Download — safe to clear, will re-download
        $sdDl = "$env:SystemRoot\SoftwareDistribution\Download"
        if (Test-Path $sdDl) {
            Stop-Service wuauserv -Force -EA SilentlyContinue
            Remove-Item "$sdDl\*" -Recurse -Force -EA SilentlyContinue
            Start-Service wuauserv -EA SilentlyContinue
            wl "Cleared SoftwareDistribution\Download" 'SUCCESS'
        }

        # CCM cache items older than 30 days
        try {
            $mgr   = New-Object -ComObject UIResource.UIResourceMgr -EA Stop
            $cache = $mgr.GetCacheInfo()
            $cut   = (Get-Date).AddDays(-30)
            foreach ($item in $cache.GetCacheElements()) {
                if ($item.LastReferenceTime -lt $cut) {
                    $cache.DeleteCacheElement($item.CacheElementID)
                    wl "  Removed CCM cache: $($item.ContentID)"
                }
            }
        } catch { wl "CCM cache cleanup skipped: $_" 'WARN' }

        # Windows Disk Cleanup (Update Cleanup, Temp Files)
        try {
            $root = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
            @('Update Cleanup','Temporary Files','Windows Upgrade Log Files','Recycle Bin') |
                ForEach-Object {
                    $p = Join-Path $root $_
                    if (Test-Path $p) {
                        Set-ItemProperty $p 'StateFlags0099' 2 -Type DWord -EA SilentlyContinue
                    }
                }
            Start-Process cleanmgr.exe -ArgumentList '/sagerun:99' -Wait -EA SilentlyContinue
        } catch { wl "Disk Cleanup util failed: $_" 'WARN' }

        $free = [math]::Round((Get-PSDrive C -EA SilentlyContinue).Free / 1GB, 2)
        wl "Free space after cleanup: ${free}GB"
        return $free
    }

    # ── WUA full reset ─────────────────────────────────────────────────────
    function Reset-WUA {
        wl "Stopping WUA services..."
        @('wuauserv','bits','cryptsvc','msiserver','ccmexec') |
            ForEach-Object { Stop-Service $_ -Force -EA SilentlyContinue }
        Start-Sleep -Seconds 5

        $ts = Get-Date -Format 'yyyyMMddHHmmss'
        foreach ($p in @("$env:SystemRoot\SoftwareDistribution","$env:SystemRoot\System32\catroot2")) {
            if (Test-Path $p) {
                try { Rename-Item $p "${p}.bak_$ts" -Force; wl "Renamed: $p" 'SUCCESS' }
                catch { wl "Could not rename $p`: $_" 'WARN' }
            }
        }

        @('atl.dll','urlmon.dll','mshtml.dll','shdocvw.dll','browseui.dll','jscript.dll',
          'vbscript.dll','scrrun.dll','msxml3.dll','msxml6.dll','actxprxy.dll','softpub.dll',
          'wintrust.dll','dssenh.dll','rsaenh.dll','cryptdlg.dll','oleaut32.dll','ole32.dll',
          'shell32.dll','initpki.dll','wuapi.dll','wuaueng.dll','wucltui.dll','wups.dll',
          'wups2.dll','wuweb.dll','qmgr.dll','qmgrprxy.dll','wucltux.dll','muweb.dll','wuwebv.dll') |
            ForEach-Object { & regsvr32.exe /s $_ 2>$null }

        & netsh winsock reset | Out-Null
        & netsh winhttp reset proxy | Out-Null

        @('cryptsvc','bits','wuauserv') |
            ForEach-Object { Start-Service $_ -EA SilentlyContinue }
        wl "WUA reset complete." 'SUCCESS'
    }

    # ── DISM repair ────────────────────────────────────────────────────────
    function Invoke-DISMRepair {
        wl "DISM CheckHealth..."
        $chk = & dism.exe /Online /Cleanup-Image /CheckHealth 2>&1
        if ($LASTEXITCODE -eq 0 -and ($chk -join ' ') -notmatch 'repairable|corruption') {
            wl "DISM: clean, no repair needed." 'SUCCESS'; return 'Clean'
        }
        wl "DISM RestoreHealth (may take 10-20 min)..."
        & dism.exe /Online /Cleanup-Image /RestoreHealth /NoRestart 2>&1 | Out-Null
        switch ($LASTEXITCODE) {
            0    { wl "DISM RestoreHealth: success." 'SUCCESS'; return 'Repaired' }
            3010 { wl "DISM RestoreHealth: success, reboot needed." 'SUCCESS'; return 'RepairedRebootNeeded' }
            default {
                wl "DISM RestoreHealth failed (exit $LASTEXITCODE)." 'ERROR'; return 'Failed'
            }
        }
    }

    # ── ACL reset ──────────────────────────────────────────────────────────
    function Reset-SDACL {
        & icacls "$env:SystemRoot\SoftwareDistribution" /reset /T /C /Q | Out-Null
        wl "SoftwareDistribution ACL reset." 'SUCCESS'
    }

    # ── Clear WUA DataStore ────────────────────────────────────────────────
    function Clear-DataStore {
        Stop-Service wuauserv -Force -EA SilentlyContinue
        Start-Sleep -Seconds 3
        $ds = "$env:SystemRoot\SoftwareDistribution\DataStore"
        if (Test-Path $ds) {
            Remove-Item "$ds\*" -Recurse -Force -EA SilentlyContinue
            wl "DataStore cleared." 'SUCCESS'
        }
        Start-Service wuauserv -EA SilentlyContinue
    }

    # ── Schedule reboot via task ───────────────────────────────────────────
    function Schedule-Reboot {
        Unregister-ScheduledTask 'MECM_Patch_Remediation_Reboot' -Confirm:$false -EA SilentlyContinue
        $at = (Get-Date).AddSeconds($RebootDelaySec)
        $a  = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument '-r -t 10 -f'
        $t  = New-ScheduledTaskTrigger -Once -At $at
        Register-ScheduledTask 'MECM_Patch_Remediation_Reboot' -Action $a -Trigger $t `
            -RunLevel Highest -User SYSTEM -Force | Out-Null
        wl "Reboot scheduled at $at" 'SUCCESS'
    }

    # ── Queue MECM triggers via RunOnce (fires after reboot) ───────────────
    function Queue-PostRebootTriggers {
        $cmd = 'powershell.exe -NonInteractive -WindowStyle Hidden -Command ' +
               '"Start-Sleep 60;' +
               'Invoke-WmiMethod -Namespace ROOT\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList \"{00000000-0000-0000-0000-000000000113}\";' +
               'Start-Sleep 10;' +
               'Invoke-WmiMethod -Namespace ROOT\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList \"{00000000-0000-0000-0000-000000000108}\""'
        Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' `
            'MECMPatchEval' $cmd -EA SilentlyContinue
        wl "Post-reboot MECM triggers queued in RunOnce." 'INFO'
    }

    # ── Trigger MECM now (post-reboot pass) ───────────────────────────────
    function Invoke-MECMTriggers {
        try {
            Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' `
                -ArgumentList '{00000000-0000-0000-0000-000000000113}' | Out-Null
            Start-Sleep -Seconds 10
            Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' `
                -ArgumentList '{00000000-0000-0000-0000-000000000108}' | Out-Null
            wl "MECM SU scan + deploy eval triggered." 'SUCCESS'
        } catch { wl "MECM trigger failed: $_" 'WARN' }
    }

    # ══════════════════════════════════════════════════════════════════════
    #  MAIN NODE LOGIC
    # ══════════════════════════════════════════════════════════════════════

    wl "════════ VDI Node Remediation — $env:COMPUTERNAME — Pass: $(if($ScheduleReboot){'PRE-REBOOT'}else{'POST-REBOOT'}) ════════"

    $actions  = [System.Collections.Generic.List[string]]::new()
    $needBoot = $false
    $abort    = $false

    # Step 1 — Read error codes
    $codes = Get-UpdateErrors
    $hexCodes = ($codes | ForEach-Object { '0x' + $_.ToString('X8') }) -join ', '
    wl "Error codes found: $(if($codes.Count){"$hexCodes"}else{'None'})"

    if ($codes.Count -eq 0 -and -not $ScheduleReboot) {
        # Post-reboot pass with no errors = machine is clean
        Invoke-MECMTriggers
        $actions.Add('NoErrors-TriggersOnly')
        # Return result object
        return [PSCustomObject]@{
            Computer    = $env:COMPUTERNAME
            Pass        = 'Post-Reboot'
            ErrorCodes  = ''
            Actions     = 'NoErrors-TriggersOnly'
            NeedsReboot = $false
            Aborted     = $false
            FreeGB      = [math]::Round((Get-PSDrive C -EA SilentlyContinue).Free/1GB, 2)
        }
    }

    # Step 2 — Disk space (must be first; DISM needs headroom)
    $freeGB  = [math]::Round((Get-PSDrive C -EA SilentlyContinue).Free / 1GB, 2)
    $diskLow = ($codes -contains $EC.DiskFull) -or ($freeGB -lt $MinFreeGB)

    if ($diskLow) {
        wl "Disk below ${MinFreeGB}GB (${freeGB}GB). Running cleanup..." 'WARN'
        $freeGB = Invoke-DiskCleanup
        $actions.Add('DiskCleanup')
        if ($freeGB -lt $MinFreeGB) {
            wl "ABORT: ${freeGB}GB free after cleanup — still below ${MinFreeGB}GB minimum." 'ERROR'
            $actions.Add("DiskInsufficient-${freeGB}GB")
            $abort = $true
        }
    }

    # Step 3 — Error-code-driven fixes
    if (-not $abort -and $codes.Count -gt 0) {

        if ($codes -contains $EC.Shutdown) {
            wl "0x8007045B: Shutdown-in-progress. Clean reboot will resolve." 'WARN'
            $needBoot = $true; $actions.Add('0x8007045B-NeedReboot')
        }

        if ($codes -contains $EC.PendingReboot) {
            wl "0x87D00651: Pending reboot blocking updates." 'WARN'
            $needBoot = $true; $actions.Add('0x87D00651-PendingReboot')
        }

        if ($codes -contains $EC.AccessDenied) {
            wl "0x80070005: Access denied — resetting ACL + WUA." 'WARN'
            Reset-SDACL; Reset-WUA
            $needBoot = $true; $actions.Add('0x80070005-ACL+WUAReset')
        }

        if ($codes -contains $EC.AllUpdates) {
            wl "0x80240022: All updates failed — full WUA reset." 'WARN'
            Reset-WUA
            $needBoot = $true; $actions.Add('0x80240022-WUAReset')
        }

        if ($codes -contains $EC.Unexpected) {
            wl "0x8000FFFF: Catastrophic failure — WUA reset + DISM." 'WARN'
            Reset-WUA
            $dr = Invoke-DISMRepair
            $needBoot = $true; $actions.Add("0x8000FFFF-WUA+DISM($dr)")
        }

        $cbsCodes = $codes | Where-Object { $_ -in @($EC.CBSTrans,$EC.CompStore,$EC.MissingDll) }
        if ($cbsCodes) {
            $hex = ($cbsCodes | ForEach-Object { '0x'+$_.ToString('X8') }) -join '+'
            wl "$hex: CBS/component corruption — DISM." 'WARN'
            $dr = Invoke-DISMRepair
            $needBoot = $true; $actions.Add("CBS($hex)-DISM($dr)")
        }

        if (($codes -contains $EC.KeyNotFound) -or ($codes -contains $EC.DataContract)) {
            wl "0x80240008/0x80240439: WUA DataStore issue — clearing cache." 'WARN'
            Clear-DataStore
            $needBoot = $true; $actions.Add('DataStoreCleared')
        }

        if ($codes -contains $EC.Superseded) {
            wl "0x8007066A: Superseded update — will clear on next scan." 'WARN'
            $actions.Add('0x8007066A-Noted')
        }
    }

    # Step 4 — Reboot or trigger
    if (-not $abort) {
        if ($ScheduleReboot -and $needBoot) {
            Queue-PostRebootTriggers
            Schedule-Reboot
            $actions.Add("RebootIn${RebootDelaySec}s")
        } elseif (-not $ScheduleReboot) {
            # Post-reboot pass — trigger MECM now regardless
            Invoke-MECMTriggers
            $actions.Add('MECMTriggered')
        }
    }

    $freeGB = [math]::Round((Get-PSDrive C -EA SilentlyContinue).Free / 1GB, 2)
    wl "Complete. Actions: $($actions -join ' | ') | FreeGB: $freeGB | NeedsReboot: $needBoot | Aborted: $abort"

    [PSCustomObject]@{
        Computer    = $env:COMPUTERNAME
        Pass        = if ($ScheduleReboot) { 'Pre-Reboot' } else { 'Post-Reboot' }
        ErrorCodes  = $hexCodes
        Actions     = $actions -join ' | '
        NeedsReboot = $needBoot
        Aborted     = $abort
        FreeGB      = $freeGB
    }
}

# ============================================================================
#  SECTION 7 — Deliver remediation to a single node
#              Try MECM Run Script first, fall back to PSRemoting
# ============================================================================

function Invoke-NodeRemediation {
    param(
        [string] $Computer,
        [int]    $ResourceID,
        [bool]   $ScheduleReboot,
        [string] $Phase             # 'PreReboot' or 'PostReboot'
    )

    $params = @{ MinFreeGB = $MinFreeGB; ScheduleReboot = $ScheduleReboot; RebootDelaySec = 90 }

    # ── Try MECM Run Script ────────────────────────────────────────────────
    $cmResult = $null
    if ($CMScriptName) {
        try {
            $cmScript = Get-CMScript -ScriptName $CMScriptName -Fast -ErrorAction SilentlyContinue
            if ($cmScript -and $cmScript.ApprovalState -eq 3) {  # 3 = Approved
                if ($PSCmdlet.ShouldProcess($Computer, "Invoke-CMScript '$CMScriptName'")) {
                    $inv = Invoke-CMScript -ScriptGuid $cmScript.ScriptGuid `
                                          -Device (Get-CMDevice -ResourceId $ResourceID -Fast) `
                                          -PassThru -ErrorAction Stop
                    # Wait for result (poll up to 5 min)
                    $deadline = (Get-Date).AddMinutes(5)
                    while ((Get-Date) -lt $deadline) {
                        Start-Sleep -Seconds 15
                        $inv = Get-CMScriptInvocationStatus -OperationId $inv.OperationId -ErrorAction SilentlyContinue
                        if ($inv.Status -in 1,2,3,4) { break }   # Completed states
                    }
                    $cmResult = $inv.ScriptOutput | ConvertFrom-Json -ErrorAction SilentlyContinue
                    Add-Report $Computer $Phase 'Remediation' 'CMScript-OK' ($cmResult.Actions)
                    return $cmResult
                }
            } else {
                Write-Verbose "[$Computer] MECM script '$CMScriptName' not found or not approved — using PSRemoting."
            }
        } catch {
            Write-Verbose "[$Computer] Invoke-CMScript failed ($($_.Exception.Message)) — falling back to PSRemoting."
            Add-Report $Computer $Phase 'Remediation' 'CMScript-Fallback' $_.Exception.Message
        }
    }

    # ── Fall back to PSRemoting ────────────────────────────────────────────
    try {
        if ($PSCmdlet.ShouldProcess($Computer, 'Invoke-Command (PSRemoting) remediation')) {
            $result = Invoke-Command -ComputerName $Computer `
                                     -ScriptBlock $script:NodeRemediationBlock `
                                     -ArgumentList $MinFreeGB, $ScheduleReboot, 90 `
                                     -ErrorAction Stop
            Add-Report $Computer $Phase 'Remediation' 'PSRemoting-OK' ($result.Actions)
            return $result
        }
    } catch {
        Add-Report $Computer $Phase 'Remediation' 'PSRemoting-Error' $_.Exception.Message
        return $null
    }
}

# ============================================================================
#  SECTION 8 — MECM policy triggers (site-side, runs on console host)
# ============================================================================

function Send-MECMPolicyNotification {
    param([int]$ResourceID, [string]$Computer)
    try {
        if ($PSCmdlet.ShouldProcess($Computer, 'CMClientNotification MachinePolicy')) {
            Invoke-CMClientNotification -DeviceId $ResourceID `
                -NotificationType RequestMachinePolicyNow -ErrorAction Stop
            Add-Report $Computer 'PostReboot' 'PolicyNotification' 'Sent'
        }
    } catch {
        Add-Report $Computer 'PostReboot' 'PolicyNotification' 'Error' $_.Exception.Message
    }
}

# ============================================================================
#  SECTION 9 — Per-machine worker (runs in parallel runspace)
#              Handles: pre-remediate → reboot → wait → post-remediate → verify
# ============================================================================

$script:WorkerBlock = {
    param(
        [string]   $Computer,
        [int]      $ResourceID,
        [int]      $OnlineWaitMinutes,
        [bool]     $WhatIf,
        [int]      $MinFreeGB,
        [scriptblock] $RemediationBlock
    )

    $results = [System.Collections.Generic.List[object]]::new()

    function Log {
        param([string]$Phase,[string]$Action,[string]$Result,[string]$Detail='')
        $results.Add([pscustomobject]@{
            Timestamp=$( (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') )
            Computer=$Computer; Phase=$Phase; Action=$Action; Result=$Result; Detail=$Detail
        })
    }

    # ── PRE-REBOOT: run remediation on the node ────────────────────────────
    if (-not $WhatIf) {
        try {
            $pre = Invoke-Command -ComputerName $Computer `
                                  -ScriptBlock $RemediationBlock `
                                  -ArgumentList $MinFreeGB, $true, 90 `
                                  -ErrorAction Stop
            Log 'PreReboot' 'Remediation' $(if($pre.Aborted){'Aborted'}else{'OK'}) `
                "$($pre.ErrorCodes) | $($pre.Actions)"
            if ($pre.Aborted) {
                Log 'PreReboot' 'Workflow' 'AbortedDiskFull' ''
                return $results
            }
        } catch {
            Log 'PreReboot' 'Remediation' 'PSRemoting-Error' $_.Exception.Message
            # Continue to reboot anyway — node may still be partially fixable
        }
    } else {
        Log 'PreReboot' 'Remediation' 'WhatIf' ''
    }

    # ── REBOOT ─────────────────────────────────────────────────────────────
    # The per-node script schedules its own reboot via scheduled task.
    # We additionally issue a remote reboot here as belt-and-braces.
    if (-not $WhatIf) {
        $rebooted = $false
        try {
            Restart-Computer -ComputerName $Computer -Force -ErrorAction Stop
            Log 'Reboot' 'Issue' 'Issued' 'Restart-Computer'
            $rebooted = $true
        } catch {
            try {
                $p = Start-Process shutdown.exe `
                     -ArgumentList "/m \\$Computer /r /f /t 5 /c `"SCCM patch remediation`"" `
                     -Wait -PassThru -NoNewWindow -ErrorAction Stop
                if ($p.ExitCode -in 0,1190) {
                    Log 'Reboot' 'Issue' 'Issued' "shutdown.exe exit $($p.ExitCode)"
                    $rebooted = $true
                } else {
                    Log 'Reboot' 'Issue' 'Failed' "shutdown.exe exit $($p.ExitCode)"
                }
            } catch {
                Log 'Reboot' 'Issue' 'Error' $_.Exception.Message
            }
        }
    } else {
        Log 'Reboot' 'Issue' 'WhatIf' ''
    }

    # ── WAIT ONLINE ────────────────────────────────────────────────────────
    $online   = $false
    $deadline = (Get-Date).AddMinutes($OnlineWaitMinutes)

    while ((Get-Date) -lt $deadline) {
        if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
            try {
                $null = Get-CimInstance -ComputerName $Computer -ClassName Win32_OperatingSystem `
                                        -ErrorAction Stop -OperationTimeoutSec 10
                $online = $true; break
            } catch { }
        }
        Start-Sleep -Seconds 20
    }

    if (-not $online) {
        Log 'PostReboot' 'WaitOnline' 'TimedOut' ">${OnlineWaitMinutes}min"
        return $results
    }
    Log 'PostReboot' 'WaitOnline' 'Online' ''

    if ($WhatIf) { return $results }

    # Brief additional settle time — CCM agent needs a moment after boot
    Start-Sleep -Seconds 30

    # ── POST-REBOOT: verify and re-remediate if still failing ──────────────
    $pass = 0
    do {
        $pass++
        try {
            $post = Invoke-Command -ComputerName $Computer `
                                   -ScriptBlock $RemediationBlock `
                                   -ArgumentList $MinFreeGB, $false, 90 `
                                   -ErrorAction Stop
            Log "PostReboot-Pass$pass" 'Remediation' $(if($post.Aborted){'Aborted'}else{'OK'}) `
                "$($post.ErrorCodes) | $($post.Actions)"

            # If no error codes remain — machine is clean, stop
            if (-not $post.ErrorCodes -or $post.Aborted) { break }

            # Still has errors on pass 1 → run again (pass 2 = final attempt)
        } catch {
            Log "PostReboot-Pass$pass" 'Remediation' 'PSRemoting-Error' $_.Exception.Message
            break
        }
    } while ($pass -lt 2)

    return $results
}

# ============================================================================
#  MAIN
# ============================================================================

try {
    Connect-CMSite

    # ── 1. Select deployments ──────────────────────────────────────────────
    $selected = Select-Deployments
    if (-not $selected) { Write-Warning "No deployments selected. Exiting."; return }
    Write-Host ("Selected {0} deployment(s)." -f @($selected).Count) -ForegroundColor Cyan

    # ── 2. Include Unknown? ────────────────────────────────────────────────
    if (-not $PSBoundParameters.ContainsKey('IncludeUnknown')) {
        $ans = Read-Host "`nAlso include machines in 'Unknown' state? (Y/N)"
        [bool]$IncludeUnknown = ($ans -match '^[Yy]')
    }

    # ── 3. Get failed/unknown assets ───────────────────────────────────────
    Write-Host "`nQuerying failed$(if($IncludeUnknown){' + unknown'}) assets ..." -ForegroundColor Cyan
    $assets = Get-FailedAssets -Deployments $selected -IncludeUnknown $IncludeUnknown

    if (-not $assets) {
        Write-Warning "No failed or unknown assets found."
        Save-Report; return
    }

    Write-Host ("`nFound {0} unique machine(s):" -f @($assets).Count) -ForegroundColor Yellow
    $assets | Group-Object Status | ForEach-Object {
        Write-Host ("  {0,4}  {1}" -f $_.Count, $_.Name)
    }
    foreach ($a in $assets) {
        Add-Report $a.Computer 'Discovery' 'FoundInDeployment' $a.Status `
            "$($a.Deployment) | $($a.StatusDescription)"
    }

    # ── 4. Resolve MECM device records ────────────────────────────────────
    Write-Host "`nResolving device records ..." -ForegroundColor Cyan
    $devices = foreach ($a in $assets) {
        $dev = Get-CMDevice -Name $a.Computer -Fast -ErrorAction SilentlyContinue
        if ($dev) {
            [pscustomobject]@{ Computer = $a.Computer; ResourceID = [int]$dev.ResourceID }
        } else {
            Write-Warning "$($a.Computer) not found in MECM — skipping."
            Add-Report $a.Computer 'Discovery' 'ResolveDevice' 'NotInMECM'
        }
    }
    $devices = @($devices | Where-Object { $_ })
    if ($devices.Count -eq 0) {
        Write-Warning "No devices resolved."; Save-Report; return
    }

    # ── 5. Confirm ────────────────────────────────────────────────────────
    if (-not $PSCmdlet.ShouldContinue(
        "$($devices.Count) machine(s): add to '$MaintenanceCollectionName', " +
        "pre-remediate, reboot in batches of $BatchSize, verify and re-remediate post-reboot?",
        'Confirm VDI Patch Remediation')) {
        Write-Warning "Aborted by user."; Save-Report; return
    }

    # ── 6. Add to maintenance collection ──────────────────────────────────
    Write-Host "`nAdding to '$MaintenanceCollectionName' ..." -ForegroundColor Cyan
    Add-ToMaintenanceCollection -Computers $devices.Computer

    # ── 7. Logged-on user check ────────────────────────────────────────────
    $toProcess = [System.Collections.Generic.List[object]]::new()

    if ($SkipLoggedOnCheck) {
        $toProcess.AddRange($devices)
        Write-Host "`nLogged-on check skipped — $($devices.Count) machine(s) queued." -ForegroundColor Cyan
    } else {
        Write-Host "`nChecking for logged-on users ..." -ForegroundColor Cyan
        foreach ($dev in $devices) {
            if (Test-Connection -ComputerName $dev.Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                if (Test-UserLoggedOn -Computer $dev.Computer) {
                    Write-Host "  [$($dev.Computer)] User logged on — SKIPPED" -ForegroundColor Magenta
                    Add-Report $dev.Computer 'LoggedOnCheck' 'Check' 'UserLoggedOn-Skipped'
                } else {
                    Add-Report $dev.Computer 'LoggedOnCheck' 'Check' 'Clear'
                    $toProcess.Add($dev)
                }
            } else {
                Add-Report $dev.Computer 'LoggedOnCheck' 'Check' 'Offline-Queued'
                $toProcess.Add($dev)
            }
        }
        Write-Host ("  {0,4}  queued | {1,4}  skipped (user logged on)" -f `
            $toProcess.Count, ($devices.Count - $toProcess.Count))
    }

    if ($toProcess.Count -eq 0) {
        Write-Warning "No machines to process."; Save-Report; return
    }

    # ── 8. Batch processing — parallel runspaces per batch ─────────────────
    $batches = for ($i = 0; $i -lt $toProcess.Count; $i += $BatchSize) {
        ,@($toProcess[$i .. [Math]::Min($i + $BatchSize - 1, $toProcess.Count - 1)])
    }

    Write-Host ("`nProcessing {0} batch(es) of up to {1} machines in parallel." `
                -f $batches.Count, $BatchSize) -ForegroundColor Cyan

    $resultBag   = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $remBlock    = $script:NodeRemediationBlock   # capture for closure

    for ($b = 0; $b -lt $batches.Count; $b++) {
        $batch = $batches[$b]
        Write-Host ("`n═══ Batch {0}/{1} — {2} machine(s) ═══" `
                    -f ($b+1), $batches.Count, $batch.Count) -ForegroundColor Yellow

        $pool    = [RunspaceFactory]::CreateRunspacePool(1, $batch.Count)
        $pool.Open()
        $handles = [System.Collections.Generic.List[object]]::new()

        foreach ($dev in $batch) {
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            $null = $ps.AddScript($script:WorkerBlock).AddParameters(@{
                Computer          = $dev.Computer
                ResourceID        = $dev.ResourceID
                OnlineWaitMinutes = $OnlineWaitMinutes
                WhatIf            = ($WhatIfPreference -eq 'Continue')
                MinFreeGB         = $MinFreeGB
                RemediationBlock  = $remBlock
            })
            $handles.Add([pscustomobject]@{
                PS       = $ps
                Handle   = $ps.BeginInvoke()
                Computer = $dev.Computer
                ResID    = $dev.ResourceID
            })
            Write-Host "  [$($dev.Computer)] dispatched" -ForegroundColor White
        }

        # Site-side policy notification fires from host thread (needs CM drive)
        foreach ($dev in $batch) {
            Send-MECMPolicyNotification -ResourceID $dev.ResourceID -Computer $dev.Computer
        }

        # Collect runspace results
        foreach ($h in $handles) {
            try {
                $rows = $h.PS.EndInvoke($h.Handle)
                foreach ($r in $rows) { $resultBag.Add($r) }

                # Surface unhandled stream errors
                foreach ($e in $h.PS.Streams.Error) {
                    $msg = $e.ToString()
                    if ($msg -notmatch '1190|already scheduled') {
                        $resultBag.Add([pscustomobject]@{
                            Timestamp=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                            Computer=$h.Computer; Phase='Worker'; Action='StreamError'
                            Result='Error'; Detail=$msg
                        })
                        Write-Warning "[$($h.Computer)] $msg"
                    }
                }

                # Console status line
                $rebootRow  = $rows | Where-Object { $_.Action -eq 'Issue' -and $_.Phase -eq 'Reboot' } | Select-Object -Last 1
                $onlineRow  = $rows | Where-Object { $_.Action -eq 'WaitOnline' } | Select-Object -Last 1
                $post2Row   = $rows | Where-Object { $_.Phase -like 'PostReboot*' } | Select-Object -Last 1
                $finalState = if ($post2Row) { "$($post2Row.Phase):$($post2Row.Result)" } `
                              elseif ($onlineRow) { "Online:$($onlineRow.Result)" } `
                              else { "Reboot:$($rebootRow.Result)" }
                $colour = if ($finalState -match 'Error|Failed|TimedOut') {'Red'} `
                          elseif ($finalState -match 'OK|Issued|Online') {'Green'} else {'Yellow'}
                Write-Host "  [$($h.Computer)] $finalState" -ForegroundColor $colour
            } catch {
                $resultBag.Add([pscustomobject]@{
                    Timestamp=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    Computer=$h.Computer; Phase='Worker'; Action='EndInvoke'
                    Result='Exception'; Detail=$_.Exception.Message
                })
                Write-Warning "[$($h.Computer)] EndInvoke: $($_.Exception.Message)"
            } finally {
                $h.PS.Dispose()
            }
        }

        $pool.Close(); $pool.Dispose()

        if ($b -lt $batches.Count - 1) {
            Write-Host ("`n  Waiting {0} min before next batch ..." -f $BatchIntervalMinutes) -ForegroundColor DarkGray
            Start-Sleep -Seconds ($BatchIntervalMinutes * 60)
        }
    }

    # Merge parallel results into report
    foreach ($r in $resultBag) { $script:Report.Add($r) }

    # ── 9. Done ────────────────────────────────────────────────────────────
    Write-Host "`n════════════════════════════════════════════" -ForegroundColor Green
    Write-Host " ORCHESTRATION COMPLETE" -ForegroundColor Green
    Write-Host "════════════════════════════════════════════" -ForegroundColor Green
    Save-Report

} catch {
    Write-Error "Fatal: $_"
    Save-Report
    throw
} finally {
    Pop-Location -ErrorAction SilentlyContinue
}
