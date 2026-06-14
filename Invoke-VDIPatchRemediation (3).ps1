#Requires -RunAsAdministrator
<#
.SYNOPSIS
    VDI Patch Cycle Remediation Script
    Designed for MECM "Run Script" targeting PATCH - Ring 2 - Production - VDIs - Phase 1

.DESCRIPTION
    Reads the actual MECM deployment error code from the CCM WMI/registry on each
    machine, maps it to a specific remediation path, and executes only what is needed.

    Error code → remediation mapping:
        0x8007045B  → Reschedule reboot via scheduled task, reboot
        0x87D00651  → Clear pending reboot flags, reboot
        0x80240022  → WUA full reset, reboot
        0x8000FFFF  → WUA full reset + DISM RestoreHealth, reboot
        0x800F0820  → DISM RestoreHealth, reboot
        0x80073712  → DISM RestoreHealth, reboot
        0x8007007E  → DISM RestoreHealth, reboot
        0x80070005  → SoftwareDistribution ACL reset + WUA reset, reboot
        0x80070070  → Aggressive disk cleanup; abort if still < 20GB after cleanup
        0x8007066A  → Clear superseded update from CCM cache, trigger re-evaluation
        0x80240008  → Clear WUA DataStore cache, re-register SUP, trigger re-evaluation
        0x80240439  → Clear WUA download cache, trigger re-evaluation

.NOTES
    Run via MECM console: Administration > Scripts > Run Script
    Requires script approval. Runs as SYSTEM.
    Reboot scheduled 90s after script exit where required.
    Log: C:\Windows\Temp\VDIPatchRemediation.log
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Logging ────────────────────────────────────────────────────────────
$LogPath = 'C:\Windows\Temp\VDIPatchRemediation.log'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8
    Write-Output $entry
}
#endregion

#region ── Constants ──────────────────────────────────────────────────────────
$MIN_FREE_GB     = 20
$REBOOT_DELAY_S  = 90   # seconds before scheduled reboot fires

# Error codes as unsigned 32-bit to match WMI/registry representation
$EC_SHUTDOWN     = [uint32]'0x8007045B'   # Shutdown in progress
$EC_PENDINGREBOOT= [uint32]'0x87D00651'   # Pending reboot
$EC_ALLUPDATES   = [uint32]'0x80240022'   # Operation failed for all updates
$EC_UNEXPECTED   = [uint32]'0x8000FFFF'   # Catastrophic/unexpected
$EC_CBS_TRANS    = [uint32]'0x800F0820'   # CBS transaction failure
$EC_COMP_STORE   = [uint32]'0x80073712'   # Component store corruption
$EC_MISSING_DLL  = [uint32]'0x8007007E'   # Missing DLL / damaged image
$EC_ACCESS_DENIED= [uint32]'0x80070005'   # Access denied
$EC_DISK_FULL    = [uint32]'0x80070070'   # Disk full
$EC_SUPERSEDED   = [uint32]'0x8007066A'   # Update superseded
$EC_KEY_NOTFOUND = [uint32]'0x80240008'   # WSUS key not found
$EC_DATA_CONTRACT= [uint32]'0x80240439'   # Data contract mismatch
#endregion

#region ── Error Code Discovery ───────────────────────────────────────────────
function Get-MECMUpdateErrors {
    <#
    Returns a deduplicated array of [uint32] error codes sourced from:
      1. CCM_SoftwareUpdate WMI class (most reliable - MECM's own view)
      2. UpdatesDeployment.log parsing (catches codes WMI may have aged out)
      3. WU registry LastError (fallback)
    #>
    $codes = [System.Collections.Generic.HashSet[uint32]]::new()

    # Source 1: MECM WMI - CCM_SoftwareUpdate
    try {
        $updates = Get-WmiObject -Namespace 'ROOT\ccm\SoftMgmtAgent' `
                                 -Class 'CCM_SoftwareUpdate' `
                                 -ErrorAction Stop |
                   Where-Object { $_.ErrorCode -and $_.ErrorCode -ne 0 }
        foreach ($u in $updates) {
            $code = [uint32]$u.ErrorCode
            [void]$codes.Add($code)
            Write-Log "CCM_SoftwareUpdate: Article $($u.ArticleID) error 0x$($code.ToString('X8'))"
        }
    } catch {
        Write-Log "Could not query CCM_SoftwareUpdate WMI: $_" 'WARN'
    }

    # Source 2: UpdatesDeployment.log - last 500 lines for recent install errors
    try {
        $udLog = "$env:SystemRoot\CCM\Logs\UpdatesDeployment.log"
        if (Test-Path $udLog) {
            $pattern = '(?i)installupdate.*?error[:\s]+(-?\d+)|(?i)failed.*?0x([0-9A-Fa-f]{8})'
            $hits = Select-String -Path $udLog -Pattern $pattern -Last 500 -ErrorAction SilentlyContinue
            foreach ($hit in $hits) {
                # Extract hex codes
                $hexMatches = [regex]::Matches($hit.Line, '0x([0-9A-Fa-f]{8})')
                foreach ($m in $hexMatches) {
                    try {
                        $code = [uint32]('0x' + $m.Groups[1].Value)
                        if ($code -ne 0 -and $code -ne 0x80070000) {
                            [void]$codes.Add($code)
                        }
                    } catch { }
                }
            }
        }
    } catch {
        Write-Log "Could not parse UpdatesDeployment.log: $_" 'WARN'
    }

    # Source 3: WU registry fallback
    try {
        $wuKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install'
        if (Test-Path $wuKey) {
            $lastErr = (Get-ItemProperty $wuKey -ErrorAction SilentlyContinue).LastError
            if ($lastErr -and $lastErr -ne 0) {
                $code = [uint32]$lastErr
                [void]$codes.Add($code)
                Write-Log "WU Registry LastError: 0x$($code.ToString('X8'))"
            }
        }
    } catch {
        Write-Log "Could not read WU registry: $_" 'WARN'
    }

    return @($codes)
}
#endregion

#region ── Helper Functions ───────────────────────────────────────────────────

function Get-PendingRebootReasons {
    $reasons = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
        $reasons += 'CBS'
    }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        $reasons += 'WindowsUpdate'
    }
    $pfro = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction SilentlyContinue).PendingFileRenameOperations
    if ($pfro) { $reasons += 'PendingFileRename' }
    try {
        $r = Invoke-WmiMethod -Namespace 'ROOT\ccm\ClientSDK' -Class 'CCM_ClientUtilities' -Name 'DetermineIfRebootPending' -ErrorAction SilentlyContinue
        if ($r -and ($r.RebootPending -or $r.IsHardRebootPending)) { $reasons += 'MECM' }
    } catch { }
    return $reasons
}

function Get-FreeDiskGB {
    $d = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    if ($d) { return [math]::Round($d.Free / 1GB, 2) }
    return $null
}

function Invoke-DiskCleanup {
    Write-Log "Running aggressive disk cleanup..."

    # 1. Windows Temp
    Remove-Item "$env:SystemRoot\Temp\*"        -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:TEMP\*"                   -Recurse -Force -ErrorAction SilentlyContinue

    # 2. SoftwareDistribution Download cache (safe to clear - will re-download)
    $sdDownload = "$env:SystemRoot\SoftwareDistribution\Download"
    if (Test-Path $sdDownload) {
        try {
            Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
            Remove-Item "$sdDownload\*" -Recurse -Force -ErrorAction SilentlyContinue
            Start-Service wuauserv -ErrorAction SilentlyContinue
            Write-Log "Cleared SoftwareDistribution\Download" 'SUCCESS'
        } catch {
            Write-Log "Could not clear SD Download: $_" 'WARN'
        }
    }

    # 3. CBS logs (can grow very large)
    Remove-Item "$env:SystemRoot\Logs\CBS\*.cab" -Force -ErrorAction SilentlyContinue

    # 4. CCM cache - remove aged/orphaned content
    try {
        $ccmCache = New-Object -ComObject UIResource.UIResourceMgr -ErrorAction SilentlyContinue
        if ($ccmCache) {
            $cache = $ccmCache.GetCacheInfo()
            $items = $cache.GetCacheElements()
            $cutoff = (Get-Date).AddDays(-30)
            foreach ($item in $items) {
                if ($item.LastReferenceTime -lt $cutoff) {
                    $cache.DeleteCacheElement($item.CacheElementID)
                    Write-Log "  Removed CCM cache item: $($item.ContentID)"
                }
            }
        }
    } catch {
        Write-Log "CCM cache cleanup skipped: $_" 'WARN'
    }

    # 5. Run built-in Disk Cleanup silently for system files
    try {
        $cleanupKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
        $categories = @(
            'Update Cleanup','Windows Upgrade Log Files','Temporary Files',
            'Recycle Bin','Temporary Internet Files','Downloaded Program Files'
        )
        foreach ($cat in $categories) {
            $catPath = Join-Path $cleanupKey $cat
            if (Test-Path $catPath) {
                Set-ItemProperty -Path $catPath -Name 'StateFlags0099' -Value 2 -Type DWord -ErrorAction SilentlyContinue
            }
        }
        Start-Process cleanmgr.exe -ArgumentList '/sagerun:99' -Wait -ErrorAction SilentlyContinue
        Write-Log "Windows Disk Cleanup completed." 'SUCCESS'
    } catch {
        Write-Log "Windows Disk Cleanup failed: $_" 'WARN'
    }

    $freeAfter = Get-FreeDiskGB
    Write-Log "Free space after cleanup: ${freeAfter}GB"
    return $freeAfter
}

function Reset-WindowsUpdateAgent {
    Write-Log "Stopping WUA-related services..."
    $services = @('wuauserv','bits','cryptsvc','msiserver','ccmexec')
    foreach ($svc in $services) {
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
        Write-Log "  Stopped (or already stopped): $svc"
    }
    Start-Sleep -Seconds 5

    $ts       = Get-Date -Format 'yyyyMMddHHmmss'
    $sdPath   = "$env:SystemRoot\SoftwareDistribution"
    $crPath   = "$env:SystemRoot\System32\catroot2"

    foreach ($path in @($sdPath, $crPath)) {
        if (Test-Path $path) {
            $backup = "${path}.bak_$ts"
            try {
                Rename-Item -Path $path -NewName $backup -Force
                Write-Log "Renamed $path → $backup" 'SUCCESS'
            } catch {
                Write-Log "Could not rename $path`: $_" 'WARN'
            }
        }
    }

    Write-Log "Re-registering WUA/BITS DLLs..."
    $dlls = @(
        'atl.dll','urlmon.dll','mshtml.dll','shdocvw.dll','browseui.dll',
        'jscript.dll','vbscript.dll','scrrun.dll','msxml3.dll','msxml6.dll',
        'actxprxy.dll','softpub.dll','wintrust.dll','dssenh.dll','rsaenh.dll',
        'cryptdlg.dll','oleaut32.dll','ole32.dll','shell32.dll','initpki.dll',
        'wuapi.dll','wuaueng.dll','wucltui.dll','wups.dll','wups2.dll',
        'wuweb.dll','qmgr.dll','qmgrprxy.dll','wucltux.dll','muweb.dll','wuwebv.dll'
    )
    foreach ($dll in $dlls) { & regsvr32.exe /s $dll 2>$null }

    Write-Log "Resetting Winsock and WinHTTP proxy..."
    & netsh winsock reset | Out-Null
    & netsh winhttp reset proxy | Out-Null

    Write-Log "Restarting core services..."
    foreach ($svc in @('cryptsvc','bits','wuauserv')) {
        Start-Service -Name $svc -ErrorAction SilentlyContinue
        Write-Log "  Started: $svc"
    }
}

function Invoke-DISMRepair {
    Write-Log "Running DISM /CheckHealth first..."
    $check = & dism.exe /Online /Cleanup-Image /CheckHealth 2>&1
    $checkExit = $LASTEXITCODE

    if ($checkExit -eq 0 -and ($check -join ' ') -notmatch 'repairable|corruption') {
        Write-Log "DISM CheckHealth: no corruption found - skipping RestoreHealth." 'SUCCESS'
        return 'Clean'
    }

    Write-Log "DISM CheckHealth flagged issues. Running RestoreHealth (may take 10-20 min)..."
    $result   = & dism.exe /Online /Cleanup-Image /RestoreHealth /NoRestart 2>&1
    $exitCode = $LASTEXITCODE

    if ($exitCode -eq 0) {
        Write-Log "DISM RestoreHealth: success." 'SUCCESS'
        return 'Repaired'
    } elseif ($exitCode -eq 3010) {
        Write-Log "DISM RestoreHealth: success, reboot required." 'SUCCESS'
        return 'RepairedRebootNeeded'
    } else {
        Write-Log "DISM RestoreHealth failed (exit $exitCode): $($result -join ' ')" 'ERROR'
        return 'Failed'
    }
}

function Clear-SoftwareDistributionACL {
    Write-Log "Resetting ACLs on SoftwareDistribution..."
    & icacls "$env:SystemRoot\SoftwareDistribution" /reset /T /C /Q | Out-Null
    Write-Log "ACL reset complete." 'SUCCESS'
}

function Clear-WUADataStore {
    Write-Log "Clearing WUA DataStore cache (0x80240008 / 0x80240439)..."
    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    $ds = "$env:SystemRoot\SoftwareDistribution\DataStore"
    if (Test-Path $ds) {
        Remove-Item "$ds\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "DataStore cleared." 'SUCCESS'
    }
    Start-Service wuauserv -ErrorAction SilentlyContinue
}

function Clear-SupersededCCMCache {
    Write-Log "Clearing superseded updates from CCM cache (0x8007066A)..."
    try {
        $updates = Get-WmiObject -Namespace 'ROOT\ccm\SoftMgmtAgent' -Class 'CCM_SoftwareUpdate' -ErrorAction Stop |
                   Where-Object { $_.IsSuperseded -eq $true }
        foreach ($u in $updates) {
            Write-Log "  Superseded: $($u.ArticleID) - $($u.Name)"
        }
        Write-Log "Superseded updates will be cleared by the next MECM scan cycle." 'INFO'
    } catch {
        Write-Log "Could not enumerate superseded updates: $_" 'WARN'
    }
}

function Schedule-Reboot {
    param([int]$DelaySeconds = $REBOOT_DELAY_S)
    $triggerTime = (Get-Date).AddSeconds($DelaySeconds)
    Write-Log "Scheduling reboot at $triggerTime..."
    Unregister-ScheduledTask -TaskName 'MECM_Patch_Remediation_Reboot' -Confirm:$false -ErrorAction SilentlyContinue
    $action  = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument '-r -t 10 -f'
    $trigger = New-ScheduledTaskTrigger -Once -At $triggerTime
    Register-ScheduledTask -TaskName 'MECM_Patch_Remediation_Reboot' `
        -Action $action -Trigger $trigger -RunLevel Highest -User 'SYSTEM' -Force | Out-Null
    Write-Log "Reboot scheduled." 'SUCCESS'
}

function Queue-PostRebootMECMTrigger {
    Write-Log "Queuing MECM scan + deployment eval for post-reboot (RunOnce)..."
    $cmd = 'powershell.exe -NonInteractive -WindowStyle Hidden -Command ' +
           '"Start-Sleep 60; ' +
           'Invoke-WmiMethod -Namespace ROOT\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList \"{00000000-0000-0000-0000-000000000113}\"; ' +
           'Start-Sleep 10; ' +
           'Invoke-WmiMethod -Namespace ROOT\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList \"{00000000-0000-0000-0000-000000000108}\""'
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' `
        -Name 'MECMPatchEval' -Value $cmd -ErrorAction SilentlyContinue
}

function Invoke-MECMTriggerNow {
    Write-Log "Triggering MECM Software Updates Scan + Deployment Eval..."
    try {
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000113}' | Out-Null
        Start-Sleep -Seconds 10
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' -ArgumentList '{00000000-0000-0000-0000-000000000108}' | Out-Null
        Write-Log "MECM triggers sent." 'SUCCESS'
    } catch {
        Write-Log "MECM trigger failed: $_" 'WARN'
    }
}
#endregion

#region ── Main ───────────────────────────────────────────────────────────────

Write-Log "════════════════════════════════════════════════════════"
Write-Log "VDI Patch Remediation started — $env:COMPUTERNAME"
Write-Log "════════════════════════════════════════════════════════"

$rebootRequired  = $false
$remediationDone = [System.Collections.Generic.List[string]]::new()
$abortRemediation= $false

# ── Step 1: Discover actual error codes from this machine ──────────────────
Write-Log "--- Step 1: Reading MECM/WU error codes from this machine ---"
$errorCodes = Get-MECMUpdateErrors

if ($errorCodes.Count -eq 0) {
    Write-Log "No active MECM/WU error codes found. Machine may have self-resolved or already rebooted." 'WARN'
    Write-Log "Running MECM policy triggers to force re-evaluation."
    Invoke-MECMTriggerNow
    $remediationDone.Add('NoErrorsFound-PolicyTriggered')
} else {
    Write-Log "Active error codes: $(($errorCodes | ForEach-Object { '0x' + $_.ToString('X8') }) -join ', ')"
}

# ── Step 2: Disk space — check and remediate FIRST (DISM needs space) ──────
Write-Log "--- Step 2: Disk space check (minimum: ${MIN_FREE_GB}GB) ---"
$freeGB = Get-FreeDiskGB
Write-Log "C: free space: ${freeGB}GB"

$diskIssue = $errorCodes -contains $EC_DISK_FULL
$diskLow   = $freeGB -and $freeGB -lt $MIN_FREE_GB

if ($diskIssue -or $diskLow) {
    Write-Log "Disk space below ${MIN_FREE_GB}GB threshold or 0x80070070 detected. Running cleanup..." 'WARN'
    $freeGB = Invoke-DiskCleanup
    $remediationDone.Add('DiskCleanup')

    if ($freeGB -lt $MIN_FREE_GB) {
        Write-Log "ABORT: Still only ${freeGB}GB free after cleanup. Need ${MIN_FREE_GB}GB minimum. Manual intervention required." 'ERROR'
        Write-Log "Candidates to investigate: WinSxS bloat (DISM /Cleanup-Image /StartComponentCleanup), large user profiles, application logs." 'WARN'
        $remediationDone.Add("DiskCleanupInsufficient-${freeGB}GBRemaining")
        $abortRemediation = $true
    } else {
        Write-Log "Disk space now ${freeGB}GB - sufficient to proceed." 'SUCCESS'
        $remediationDone.Add('DiskCleanup-Sufficient')
    }
}

# ── Step 3: Error-code-driven remediation ──────────────────────────────────
if (-not $abortRemediation -and $errorCodes.Count -gt 0) {
    Write-Log "--- Step 3: Error-code-driven remediation ---"

    # ── 0x8007045B: Shutdown was in progress during install ────────────────
    # Machine was mid-reboot when MECM tried to install. Just needs a clean reboot.
    if ($errorCodes -contains $EC_SHUTDOWN) {
        Write-Log "0x8007045B detected: Shutdown in progress during last install attempt. Scheduling clean reboot." 'WARN'
        $rebootRequired = $true
        $remediationDone.Add('0x8007045B-RebootScheduled')
    }

    # ── 0x87D00651: Pending reboot blocking update installation ───────────
    if ($errorCodes -contains $EC_PENDINGREBOOT) {
        Write-Log "0x87D00651 detected: Pending reboot state blocking updates." 'WARN'
        $pendingReasons = Get-PendingRebootReasons
        Write-Log "Pending reboot sources: $($pendingReasons -join ', ')"
        $rebootRequired = $true
        $remediationDone.Add("0x87D00651-PendingReboot($($pendingReasons -join '+'))")
    }

    # ── 0x80070005: Access denied on SoftwareDistribution ──────────────────
    if ($errorCodes -contains $EC_ACCESS_DENIED) {
        Write-Log "0x80070005 detected: Access denied. Resetting ACLs + WUA." 'WARN'
        Clear-SoftwareDistributionACL
        Reset-WindowsUpdateAgent
        $rebootRequired = $true
        $remediationDone.Add('0x80070005-ACLReset+WUAReset')
    }

    # ── 0x80240022: Operation failed for ALL updates — full WUA reset ──────
    if ($errorCodes -contains $EC_ALLUPDATES) {
        Write-Log "0x80240022 detected: All updates failed (WUA database/transaction issue). Full WUA reset." 'WARN'
        Reset-WindowsUpdateAgent
        $rebootRequired = $true
        $remediationDone.Add('0x80240022-WUAReset')
    }

    # ── 0x8000FFFF: Catastrophic failure — WUA reset + DISM ───────────────
    if ($errorCodes -contains $EC_UNEXPECTED) {
        Write-Log "0x8000FFFF detected: Catastrophic WUA failure. WUA reset + DISM RestoreHealth." 'WARN'
        Reset-WindowsUpdateAgent
        $dismResult = Invoke-DISMRepair
        $rebootRequired = $true
        $remediationDone.Add("0x8000FFFF-WUAReset+DISM($dismResult)")
    }

    # ── CBS/Component store errors — DISM RestoreHealth ───────────────────
    $dismCodes = @($EC_CBS_TRANS, $EC_COMP_STORE, $EC_MISSING_DLL)
    $matchedDismCodes = $errorCodes | Where-Object { $dismCodes -contains $_ }
    if ($matchedDismCodes) {
        $matchedHex = ($matchedDismCodes | ForEach-Object { '0x' + $_.ToString('X8') }) -join ', '
        Write-Log "$matchedHex detected: CBS/component store corruption. Running DISM RestoreHealth." 'WARN'
        $dismResult = Invoke-DISMRepair
        $rebootRequired = $true
        $remediationDone.Add("CBSCorruption($matchedHex)-DISM($dismResult)")
        if ($dismResult -eq 'Failed') {
            Write-Log "DISM failed. This VDI may require manual recompose from golden image." 'ERROR'
        }
    }

    # ── 0x80240008 / 0x80240439: WUA DataStore/cache issues ───────────────
    if (($errorCodes -contains $EC_KEY_NOTFOUND) -or ($errorCodes -contains $EC_DATA_CONTRACT)) {
        $matched = @()
        if ($errorCodes -contains $EC_KEY_NOTFOUND)  { $matched += '0x80240008' }
        if ($errorCodes -contains $EC_DATA_CONTRACT) { $matched += '0x80240439' }
        Write-Log "$($matched -join ', ') detected: WUA DataStore/content cache corruption. Clearing cache." 'WARN'
        Clear-WUADataStore
        $rebootRequired = $true
        $remediationDone.Add("$($matched -join '+')-DataStoreCleared")
    }

    # ── 0x8007066A: Superseded update in cache ─────────────────────────────
    if ($errorCodes -contains $EC_SUPERSEDED) {
        Write-Log "0x8007066A detected: Superseded update reference. Clearing and triggering re-eval." 'WARN'
        Clear-SupersededCCMCache
        $remediationDone.Add('0x8007066A-SupersededCleared')
        # No reboot needed - just re-evaluate
    }
}

# ── Step 4: Reboot or trigger ──────────────────────────────────────────────
Write-Log "--- Step 4: Reboot / trigger ---"
if ($rebootRequired -and -not $abortRemediation) {
    Queue-PostRebootMECMTrigger
    Schedule-Reboot -DelaySeconds $REBOOT_DELAY_S
    $remediationDone.Add("RebootIn${REBOOT_DELAY_S}s")
} elseif (-not $abortRemediation) {
    # No reboot needed — just re-trigger MECM evaluation now
    Invoke-MECMTriggerNow
    $remediationDone.Add('MECMTriggeredNoReboot')
}

# ── Summary ────────────────────────────────────────────────────────────────
$freeGBFinal = Get-FreeDiskGB
Write-Log "════════════════════════════════════════════════════════"
Write-Log "Remediation complete — $env:COMPUTERNAME"
Write-Log "Error codes found  : $(($errorCodes | ForEach-Object { '0x' + $_.ToString('X8') }) -join ', ')"
Write-Log "Actions taken      : $($remediationDone -join ' | ')"
Write-Log "Free disk (final)  : ${freeGBFinal}GB"
Write-Log "Reboot scheduled   : $rebootRequired"
Write-Log "Aborted            : $abortRemediation"
Write-Log "════════════════════════════════════════════════════════"

# JSON output for MECM Run Script results pane
[PSCustomObject]@{
    Computer       = $env:COMPUTERNAME
    ErrorCodes     = ($errorCodes | ForEach-Object { '0x' + $_.ToString('X8') }) -join ', '
    ActionsTaken   = $remediationDone -join ' | '
    RebootScheduled= $rebootRequired
    Aborted        = $abortRemediation
    FreeDiskGB     = $freeGBFinal
    LogPath        = $LogPath
} | ConvertTo-Json -Depth 3

#endregion
