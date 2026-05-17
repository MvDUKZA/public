# ConfigurationManager module loaded dynamically below
<#
.SYNOPSIS
    Interactive SCCM deployment remediation tool.
    Selects a deployment, filters machines by compliance state via GridView,
    triggers scan cycles, checks reboot pending (CBS/WU/CCM), detects logged-on
    users, warns them before rebooting, and optionally escalates to WUA/CCM repair.

.PARAMETER SiteServer
    SCCM site server FQDN. Defaults to appsmcm101fp.iprod.local.

.PARAMETER SiteCode
    SCCM site code. Defaults to PRD.

.PARAMETER DeploymentName
    Optional. Wildcard-matched against deployment names. If omitted, a GridView picker is shown.

.PARAMETER MaxConcurrent
    Number of machines to process in parallel during Phase 1. Default 50.

.PARAMETER RebootWarningSeconds
    Seconds of warning given to a logged-on user before the reboot fires. Default 900 (15 min).

.PARAMETER LogPath
    Folder for transcript and CSV output. Defaults to C:\Logs\DeploymentRemediation.

.EXAMPLE
    .\Invoke-DeploymentRemediation.ps1
    # Full interactive mode

.EXAMPLE
    .\Invoke-DeploymentRemediation.ps1 -DeploymentName "*June 2026*"
    # Pre-filters by deployment name; GridViews still shown for final pick and machine selection
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SiteServer          = 'appsmcm101fp.iprod.local',
    [string]$SiteCode            = 'PRD',
    [string]$DeploymentName,
    [int]   $MaxConcurrent       = 50,
    [int]   $RebootWarningSeconds = 900,
    [string]$LogPath             = 'C:\Logs\DeploymentRemediation'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Helpers ──────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red    }
        default { Write-Host $line -ForegroundColor Cyan   }
    }
}

# StatusType values per SMS_SUMDeploymentAssetDetails docs
$StateMap = @{
    1 = 'Success'
    2 = 'InProgress'
    3 = 'RebootRequired'
    4 = 'Unknown'
    5 = 'Error'
}

#endregion

#region ── Setup ────────────────────────────────────────────────────────────────

if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$transcript = Join-Path $LogPath "Remediation_$timestamp.log"
$csvPath    = Join-Path $LogPath "Remediation_$timestamp.csv"

Start-Transcript -Path $transcript -Append
Write-Log "=== Invoke-DeploymentRemediation started === (RebootWarning: $([math]::Round($RebootWarningSeconds/60)) min)"

# Load ConfigurationManager module
try {
    $smsKey = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\SMS\Setup' -ErrorAction Stop
    $module = Join-Path (Split-Path $smsKey.UI_Installation_Directory) 'bin\ConfigurationManager.psd1'
    Import-Module $module -ErrorAction Stop
} catch {
    Write-Log "CM module not found via registry — attempting Import-Module ConfigurationManager." WARN
    Import-Module ConfigurationManager -ErrorAction SilentlyContinue
}

$origLocation = Get-Location
try {
    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
        Write-Log "PSDrive $SiteCode created successfully."
    } else {
        Write-Log "PSDrive $SiteCode already exists."
    }
    Set-Location "$SiteCode`:\"
} catch {
    Write-Log "Failed to create PSDrive for site $SiteCode on $SiteServer : $($_.Exception.Message)" ERROR
    Write-Log "Ensure the ConfigurationManager console is installed and the site server is reachable." ERROR
    Stop-Transcript; exit 1
}

#endregion

#region ── Step 1: Deployment Selection ─────────────────────────────────────────

Write-Log "Querying software update deployments..."

# Get-CMDeployment returns SMS_DeploymentSummary objects via the CM PSDrive — no CIM needed here
$allRaw = Get-CMDeployment -FeatureType SoftwareUpdate | Sort-Object DeploymentTime -Descending

if (-not $allRaw) {
    Write-Log "No software update deployments found." ERROR
    Set-Location $origLocation; Stop-Transcript; exit 1
}

# Apply optional name filter
if ($DeploymentName) {
    $allRaw = @($allRaw | Where-Object { $_.SoftwareName -like $DeploymentName })
    if (-not $allRaw) {
        Write-Log "No deployments matched filter '$DeploymentName'." ERROR
        Set-Location $origLocation; Stop-Transcript; exit 1
    }
}

$allDeployments = $allRaw | Select-Object `
    @{N='DeploymentID';        E={$_.DeploymentID}},
    @{N='Name';                E={$_.SoftwareName}},
    @{N='Collection';          E={$_.CollectionName}},
    @{N='Compliant';           E={$_.NumberSuccess}},
    @{N='NonCompliant';        E={$_.NumberNonCompliant}},
    @{N='Error';               E={$_.NumberErrors}},
    @{N='Unknown';             E={$_.NumberUnknown}},
    @{N='Total';               E={$_.NumberTargeted}},
    @{N='DeploymentTime';      E={$_.DeploymentTime}}

Write-Log "Found $($allDeployments.Count) deployment(s). Opening picker..."

$sel = $allDeployments | Sort-Object DeploymentTime -Descending |
    Out-GridView -Title "Select a deployment to remediate  [single-select → OK]" -OutputMode Single

if (-not $sel) {
    Write-Log "No deployment selected. Exiting." WARN
    Set-Location $origLocation; Stop-Transcript; exit 0
}

# Keep reference to the raw CM object for the query chain in Step 2

Write-Log "Deployment selected: '$($sel.Name)'  ID: $($sel.DeploymentID)  Collection: '$($sel.Collection)'"

#endregion

#region ── Step 2: Machine State Query & Picker ──────────────────────────────────

Write-Log "Querying per-machine states for deployment $($sel.DeploymentID)..."

# Proven chain (mirrors Invoke-VDIPatchRemediation.ps1):
#   Get-CMSoftwareUpdateDeployment -DeploymentId <GUID>
#     -> Get-CMSoftwareUpdateDeploymentStatus -InputObject   (one row per CI)
#       -> Get-CMDeploymentStatusDetails -InputObject        (one row per device per CI)
# StatusType: 1=Success  2=InProgress  4=Unknown  5=Error/Failed
# All CIs are iterated and results deduped by machine (most recent row wins).

$machineStates = $null

try {
    Write-Log "Step 1/3 — Get-CMSoftwareUpdateDeployment (DeploymentId: $($sel.DeploymentID))..."
    $suDeployment = Get-CMSoftwareUpdateDeployment -DeploymentId $sel.DeploymentID -ErrorAction Stop
    if (-not $suDeployment) { throw "Get-CMSoftwareUpdateDeployment returned nothing for ID $($sel.DeploymentID)" }

    Write-Log "Step 2/3 — Get-CMSoftwareUpdateDeploymentStatus..."
    $statusSummaries = Get-CMSoftwareUpdateDeploymentStatus -InputObject $suDeployment -ErrorAction Stop
    if (-not $statusSummaries) { throw "Get-CMSoftwareUpdateDeploymentStatus returned nothing" }

    $ciCount = @($statusSummaries).Count
    Write-Log "Got $ciCount CI status row(s). Step 3/3 — expanding per-device details..."

    # Iterate ALL CI summaries — some deployments have many CIs, each may cover different machines
    $rawRows = foreach ($summary in @($statusSummaries)) {
        $details = Get-CMDeploymentStatusDetails -InputObject $summary -ErrorAction SilentlyContinue
        if (-not $details) { continue }
        $details | Select-Object `
            @{N='MachineName';       E={$_.DeviceName}},
            @{N='ResourceID';        E={$_.ResourceID}},
            @{N='StateID';           E={$_.StatusType}},
            @{N='State';             E={
                switch ($_.StatusType) {
                    1 { 'Success'        }
                    2 { 'InProgress'     }
                    4 { 'Unknown'        }
                    5 { 'Error'          }
                    default { "StateType_$($_.StatusType)" }
                }
            }},
            @{N='IsCompliant';       E={$_.IsCompliant}},
            @{N='StatusDescription'; E={$_.StatusDescription}},
            @{N='LastStatusTime';    E={ if ($_.StatusTime) { $_.StatusTime } else { 'Never' } }}
    }

    # Deduplicate — one row per machine, keeping most recent status
    $machineStates = @($rawRows) |
        Where-Object { $_.MachineName } |
        Sort-Object MachineName, LastStatusTime -Descending |
        Group-Object MachineName |
        ForEach-Object { $_.Group | Select-Object -First 1 }

    Write-Log "Returned $($machineStates.Count) unique machine record(s) across $ciCount CI(s)."

} catch {
    Write-Log "Machine state query failed: $($_.Exception.Message)" ERROR
    Write-Log "Ensure you are running from the PRD:\ drive and the CM console is installed." ERROR
    Set-Location $origLocation; Stop-Transcript; exit 1
}

if (-not $machineStates) {
    Write-Log "No machine records returned." WARN
    Set-Location $origLocation; Stop-Transcript; exit 0
}

Write-Log "$($machineStates.Count) machine record(s). Opening machine picker..."

$selectedMachines = $machineStates | Sort-Object State, MachineName |
    Out-GridView -Title "Select machines to remediate  [Ctrl+Click multi-select → OK]" -OutputMode Multiple

if (-not $selectedMachines) {
    Write-Log "No machines selected. Exiting." WARN
    Set-Location $origLocation; Stop-Transcript; exit 0
}

Write-Log "$($selectedMachines.Count) machine(s) selected."

#endregion

#region ── Step 3: Phase 1 — Parallel scan + reboot check + user detection ──────

$results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()

$scanBlock = {
    param($machine, $RebootWarnSecs)

    $r = [PSCustomObject]@{
        MachineName                  = $machine.MachineName
        InitialState                 = $machine.State
        Online                       = $false

        # User
        LoggedOnUser                 = 'None'

        # Reboot sources
        RebootPending                = $false
        RebootSource_CBS             = $false
        RebootSource_WindowsUpdate   = $false
        RebootSource_CCM             = $false
        RebootSource_PendingFileRename = $false

        # Scan cycles
        ScanTriggered                = $false
        DeploymentEvalTriggered      = $false
        MachinePolicyTriggered       = $false
        HardwareInvTriggered         = $false
        PhaseOneError                = $null

        # Reboot (populated in Step 4)
        RebootAction                 = 'None'
        RebootError                  = $null

        # Escalation (populated in Step 5)
        EscalationDone               = $false
        EscalationError              = $null

        FinalNote                    = ''
        Timestamp                    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    # Reachability
    if (-not (Test-Connection -ComputerName $machine.MachineName -Count 1 -Quiet)) {
        $r.FinalNote = 'Offline / unreachable'
        return $r
    }
    $r.Online = $true

    # ── Logged-on user ──────────────────────────────────────────────────────
    try {
        $cs = Get-CimInstance -ComputerName $machine.MachineName -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($cs.UserName) { $r.LoggedOnUser = $cs.UserName }
    } catch {
        $r.LoggedOnUser = "QueryFailed: $($_.Exception.Message)"
    }

    # ── Reboot pending (CBS / WU / CCM / PendingFileRename) ─────────────────
    try {
        $rb = Invoke-Command -ComputerName $machine.MachineName -ErrorAction Stop -ScriptBlock {
            $out = @{ CBS=$false; WU=$false; CCM=$false; PFR=$false }

            # 1. CBS
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
                $out.CBS = $true
            }

            # 2. Windows Update
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
                $out.WU = $true
            }

            # 3. CCM client SDK
            try {
                $state = Invoke-CimMethod -Namespace root/ccm/clientsdk -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending -ErrorAction Stop
                if ($state.RebootPending -or $state.IsHardRebootPending) { $out.CCM = $true }
            } catch {}

            # 4. PendingFileRenameOperations (supplementary)
            $pfr = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
                    -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
            if ($pfr) { $out.PFR = $true }

            return $out
        }

        $r.RebootSource_CBS              = $rb.CBS
        $r.RebootSource_WindowsUpdate    = $rb.WU
        $r.RebootSource_CCM              = $rb.CCM
        $r.RebootSource_PendingFileRename = $rb.PFR
        $r.RebootPending = ($rb.CBS -or $rb.WU -or $rb.CCM)

    } catch {
        $r.FinalNote += " | RebootCheck error: $($_.Exception.Message)"
    }

    # ── Scan cycles ─────────────────────────────────────────────────────────
    try {
        $cimSession = New-CimSession -ComputerName $machine.MachineName -ErrorAction Stop
        $cimArgs = @{ CimSession = $cimSession; Namespace = 'root/ccm'; ClassName = 'SMS_Client'; MethodName = 'TriggerSchedule'; ErrorAction = 'Stop' }
        $null = Invoke-CimMethod @cimArgs -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000113}' }; $r.ScanTriggered           = $true
        $null = Invoke-CimMethod @cimArgs -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000114}' }; $r.DeploymentEvalTriggered = $true
        $null = Invoke-CimMethod @cimArgs -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000021}' }; $r.MachinePolicyTriggered  = $true
        $null = Invoke-CimMethod @cimArgs -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000001}' }; $r.HardwareInvTriggered    = $true
        Remove-CimSession $cimSession -ErrorAction SilentlyContinue
        $r.FinalNote = 'Phase 1 scan cycles OK'
    } catch {
        $r.PhaseOneError = $_.Exception.Message
        $r.FinalNote     = 'Phase 1 FAILED — see PhaseOneError'
    }

    return $r
}

Write-Log "--- Phase 1: scan cycles + reboot check + user detection ---"

$jobs  = [System.Collections.Generic.List[object]]::new()
$queue = [System.Collections.Generic.Queue[object]]::new($selectedMachines)

while ($queue.Count -gt 0 -or $jobs.Count -gt 0) {

    while ($queue.Count -gt 0 -and $jobs.Count -lt $MaxConcurrent) {
        $m = $queue.Dequeue()
        $jobs.Add((Start-Job -ScriptBlock $scanBlock -ArgumentList $m, $RebootWarningSeconds))
    }

    $done = $jobs | Where-Object { $_.State -in 'Completed','Failed' }
    foreach ($j in $done) {
        $rr = Receive-Job $j -ErrorAction SilentlyContinue
        if ($rr) { $results.Add($rr) }
        Remove-Job $j
    }
    foreach ($j in $done) { $jobs.Remove($j) | Out-Null }

    if ($queue.Count -gt 0 -or $jobs.Count -gt 0) { Start-Sleep -Milliseconds 500 }
}

Write-Log "Phase 1 complete — $($results.Count) result(s) collected."

#endregion

#region ── Step 4: Reboot — per-machine operator prompt ─────────────────────────

$rebootNeeded = $results | Where-Object { $_.Online -and $_.RebootPending }

if ($rebootNeeded.Count -gt 0) {

    $warnMins = [math]::Round($RebootWarningSeconds / 60)

    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║              REBOOT PENDING — SUMMARY                       ║" -ForegroundColor Magenta
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta

    foreach ($m in $rebootNeeded) {
        $sources = @()
        if ($m.RebootSource_CBS)           { $sources += 'CBS' }
        if ($m.RebootSource_WindowsUpdate) { $sources += 'Windows Update' }
        if ($m.RebootSource_CCM)           { $sources += 'CCM' }
        $hasUser   = $m.LoggedOnUser -ne 'None'
        $userColor = if ($hasUser) { 'Yellow' } else { 'Green' }
        $userLabel = if ($hasUser) { "⚠  LOGGED ON: $($m.LoggedOnUser)" } else { "No active user session" }

        Write-Host ""
        Write-Host "  Machine  : $($m.MachineName)"          -ForegroundColor White
        Write-Host "  Sources  : $($sources -join ' | ')"    -ForegroundColor Gray
        Write-Host "  User     : $userLabel"                 -ForegroundColor $userColor
    }

    Write-Host ""
    Write-Host "  You will be prompted per machine.  [Y] = reboot  [S] = skip" -ForegroundColor Cyan
    Write-Host "  Logged-on users receive a $warnMins-minute msg.exe pop-up first." -ForegroundColor Cyan
    Write-Host ""

    foreach ($m in $rebootNeeded) {

        $hasUser  = $m.LoggedOnUser -ne 'None'
        $userNote = if ($hasUser) { "  ⚠  User logged on: $($m.LoggedOnUser)" } else { "  No active user" }

        Write-Host "`n  ┌─ $($m.MachineName) ─────────────────────────────" -ForegroundColor White
        Write-Host "  │$userNote" -ForegroundColor $(if ($hasUser) { 'Yellow' } else { 'Gray' })

        $ans = ''
        while ($ans -notin 'Y','S') {
            $ans = (Read-Host "  └─ Reboot this machine? [Y/S]").Trim().ToUpper()
        }

        if ($ans -eq 'S') {
            $m.RebootAction = 'Skipped-OperatorDeclined'
            Write-Log "Reboot SKIPPED for $($m.MachineName) by operator." WARN
            continue
        }

        # Send user warning via msg.exe if someone is logged on
        if ($hasUser) {
            try {
                Invoke-Command -ComputerName $m.MachineName -ErrorAction Stop -ScriptBlock {
                    param($secs, $mins)
                    $msg = "NOTICE: $env:COMPUTERNAME will restart in $mins minutes for mandatory IT patching. " +
                           "Please save all work now. Contact the IT Service Desk if you need assistance."
                    & msg.exe '*' /TIME:$secs $msg 2>$null
                } -ArgumentList $RebootWarningSeconds, $warnMins
                Write-Log "User warning sent to '$($m.LoggedOnUser)' on $($m.MachineName) ($warnMins min)."
                $m.RebootAction = 'Initiated-WithUserWarning'
            } catch {
                Write-Log "msg.exe failed on $($m.MachineName): $($_.Exception.Message) — rebooting anyway." WARN
                $m.RebootAction = 'Initiated-WarningFailed'
            }

            # Schedule reboot after warning window
            try {
                Invoke-Command -ComputerName $m.MachineName -ErrorAction Stop -ScriptBlock {
                    param($secs)
                    shutdown.exe /r /t $secs /c "Mandatory IT patching reboot. Save your work." /f
                } -ArgumentList $RebootWarningSeconds
            } catch {
                $m.RebootAction = 'Failed'
                $m.RebootError  = $_.Exception.Message
                Write-Log "Reboot command FAILED on $($m.MachineName): $($m.RebootError)" ERROR
            }

        } else {
            # No user — immediate reboot with a 60s safety buffer
            try {
                Invoke-Command -ComputerName $m.MachineName -ErrorAction Stop -ScriptBlock {
                    shutdown.exe /r /t 60 /c "Mandatory IT patching reboot." /f
                }
                $m.RebootAction = 'Initiated-NoUser'
                Write-Log "Reboot initiated on $($m.MachineName) (no active user — 60s grace)."
            } catch {
                $m.RebootAction = 'Failed'
                $m.RebootError  = $_.Exception.Message
                Write-Log "Reboot command FAILED on $($m.MachineName): $($m.RebootError)" ERROR
            }
        }
    }

} else {
    Write-Log "No machines flagged as reboot pending — skipping reboot phase."
}

#endregion

#region ── Step 5: Phase 2 Escalation ───────────────────────────────────────────

$escalationCandidates = $results | Where-Object { $_.Online -and $_.PhaseOneError }

if ($escalationCandidates.Count -gt 0) {

    Write-Host ""
    Write-Host "══════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "  $($escalationCandidates.Count) machine(s) failed Phase 1 WMI triggers." -ForegroundColor Yellow
    Write-Host "  [1] WUA cache clear + service restart (recommended)" -ForegroundColor White
    Write-Host "  [2] Full CCM client repair (ccmrepair.exe)"          -ForegroundColor White
    Write-Host "  [S] Skip — save results and exit"                    -ForegroundColor White
    Write-Host "══════════════════════════════════════════════════════" -ForegroundColor Magenta

    $choice = (Read-Host "Choice").Trim().ToUpper()

    if ($choice -in '1','2') {

        $actionLabel = if ($choice -eq '1') { 'WUA cache clear' } else { 'CCM repair' }
        Write-Log "--- Phase 2: $actionLabel on $($escalationCandidates.Count) machine(s) ---"

        $escBlock = {
            param($machine, $ch)
            $e = [PSCustomObject]@{ MachineName = $machine.MachineName; Action = ''; Error = $null }
            try {
                if ($ch -eq '1') {
                    $e.Action = 'WUA cache clear'
                    Invoke-Command -ComputerName $machine.MachineName -ErrorAction Stop -ScriptBlock {
                        Stop-Service wuauserv,bits,ccmexec -Force -ErrorAction SilentlyContinue
                        Remove-Item 'C:\Windows\SoftwareDistribution\*' -Recurse -Force -ErrorAction SilentlyContinue
                        & wuauclt /resetauthorization
                        Start-Service wuauserv,bits,ccmexec -ErrorAction SilentlyContinue
                        Start-Sleep 20
                        $null = Invoke-CimMethod -Namespace root/ccm -ClassName SMS_Client `
                            -MethodName TriggerSchedule -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000113}' }
                    }
                } else {
                    $e.Action = 'CCM repair'
                    Invoke-Command -ComputerName $machine.MachineName -ErrorAction Stop -ScriptBlock {
                        $r = "$env:WinDir\CCM\ccmrepair.exe"
                        if (Test-Path $r) { Start-Process $r -Wait } else { throw "ccmrepair.exe not found" }
                    }
                }
            } catch { $e.Error = $_.Exception.Message }
            return $e
        }

        $escJobs = $escalationCandidates | ForEach-Object {
            Start-Job -ScriptBlock $escBlock -ArgumentList $_, $choice
        }
        $escJobs | Wait-Job | Out-Null

        foreach ($ej in $escJobs) {
            $er = Receive-Job $ej -ErrorAction SilentlyContinue
            if ($er) {
                $match = $results | Where-Object { $_.MachineName -eq $er.MachineName } | Select-Object -First 1
                if ($match) {
                    $match.EscalationDone  = $true
                    $match.EscalationError = $er.Error
                    $match.FinalNote = if ($er.Error) { "Escalation FAILED: $($er.Error)" }
                                       else           { "Escalation ($($er.Action)) OK" }
                }
            }
            Remove-Job $ej
        }
        Write-Log "Phase 2 complete."
    } else {
        Write-Log "Escalation skipped by operator." WARN
    }

} else {
    Write-Log "No Phase 2 escalation needed."
}

#endregion

#region ── Step 6: Output ────────────────────────────────────────────────────────

Set-Location $origLocation

$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV saved: $csvPath"

Write-Host "`n=== REMEDIATION SUMMARY ===" -ForegroundColor Green
$results | Select-Object MachineName, InitialState, Online, LoggedOnUser,
    RebootPending,
    @{N='RebootSources'; E={
        $s = @()
        if ($_.RebootSource_CBS)           { $s += 'CBS' }
        if ($_.RebootSource_WindowsUpdate) { $s += 'WU'  }
        if ($_.RebootSource_CCM)           { $s += 'CCM' }
        if ($s) { $s -join ',' } else { '-' }
    }},
    RebootAction, ScanTriggered, EscalationDone, FinalNote |
    Format-Table -AutoSize

$online        = ($results | Where-Object Online).Count
$offline       = ($results | Where-Object { -not $_.Online }).Count
$withUser      = ($results | Where-Object { $_.Online -and $_.LoggedOnUser -ne 'None' }).Count
$rebootInit    = ($results | Where-Object { $_.RebootAction -like 'Initiated*' }).Count
$rebootSkipped = ($results | Where-Object { $_.RebootAction -like 'Skipped*'   }).Count
$escalated     = ($results | Where-Object EscalationDone).Count
$unresolved    = ($results | Where-Object {
    $_.EscalationError -or ($_.Online -and $_.PhaseOneError -and -not $_.EscalationDone)
}).Count

Write-Log ("=== DONE === Online: $online | Offline: $offline | Users detected: $withUser | " +
           "Reboots initiated: $rebootInit | Reboots skipped: $rebootSkipped | " +
           "Escalated: $escalated | Unresolved errors: $unresolved")
Write-Log "Transcript : $transcript"
Write-Log "CSV        : $csvPath"

Stop-Transcript

#endregion
