<#
.SYNOPSIS
    VDI patch remediation — right-click Run Script from MECM console.
    Remediates WUA/CBS/disk errors on the local machine, optionally installs
    a specific update, and reboots if anything was fixed.

.DESCRIPTION
    Designed to be run via MECM Administration > Scripts > Run Script against
    machines in a collection (e.g. 0626-Update). Runs as SYSTEM on the target.

    WHAT IT DOES
    ────────────
    0. If KBNumber specified: checks installed state — exits immediately
       with 'AlreadyInstalled' if found. No remediation, no reboot.
    1. Reads actual WU/CCM error codes from the machine
    2. Checks disk space (20 GB minimum) — cleans up if below threshold
    3. Runs the correct fix per error code found:
         0x8007045B  → Clean reboot (shutdown-in-progress during last attempt)
         0x87D00651  → Clear pending reboot state
         0x80070005  → ACL reset + WUA full reset
         0x80240022  → WUA full reset
         0x8000FFFF  → WUA full reset + DISM RestoreHealth
         0x800F0820  → DISM RestoreHealth (CBS transaction failure)
         0x80073712  → DISM RestoreHealth (component store corruption)
         0x8007007E  → DISM RestoreHealth (missing DLL / damaged image)
         0x80240008  → Clear WUA DataStore cache
         0x80240439  → Clear WUA DataStore cache
         0x8007066A  → Noted (clears on next MECM scan)
         0x80070070  → Aggressive disk cleanup
    4. Optionally installs a specific update:
         - UNC path to a single .msu file
         - UNC path to a folder (installs all .msu files found)
         - MECM Package ID (8-char alphanumeric e.g. PRD00042)
    5. Triggers MECM SU scan + deployment eval
    6. Reboots via scheduled task (90s delay) if anything was fixed

    RETURN VALUE (JSON in MECM Run Script Detailed Output pane)
    ────────────
    Computer, KBNumber, KBStatus, ErrorCodes, Actions,
    UpdateInstalled, RebootScheduled, Aborted, FreeGB, LogPath

.PARAMETER KBNumber
    Optional. KB article number to check before doing anything else.
    Accepts with or without the 'KB' prefix: KB5039212 or 5039212.
    Exits immediately with AlreadyInstalled if found.

.PARAMETER UpdatePath
    Optional. UNC path to .msu, folder of .msu files, or MECM Package ID.

.PARAMETER MinFreeGB
    Minimum free space on C: required. Default: 20

.PARAMETER RebootDelaySec
    Seconds between script exit and scheduled reboot. Default: 90

.NOTES
    Set MECM Run Script timeout to 1800 seconds (30 min) to cover DISM.
    Full log: C:\Windows\Temp\VDIPatchRemediation.log
    Script must be approved in MECM console before use.
#>

# MECM Run Script passes parameters as plain strings.
# Do NOT use [CmdletBinding()] — it breaks MECM parameter passing.
param(
    [string] $KBNumber       = '',
    [string] $UpdatePath     = '',
    [int]    $MinFreeGB      = 20,
    [int]    $RebootDelaySec = 90
)

# StrictMode off — MECM runs scripts in constrained environments where
# StrictMode causes spurious terminating errors on null pipeline results.
# We handle nulls explicitly throughout instead.
$ErrorActionPreference = 'SilentlyContinue'

#region ── Logging ─────────────────────────────────────────────────────────────

$LogPath = 'C:\Windows\Temp\VDIPatchRemediation.log'
$null    = New-Item -ItemType Directory -Path 'C:\Windows\Temp' -Force -ErrorAction SilentlyContinue

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
}

#endregion

#region ── Error code constants ────────────────────────────────────────────────
# Stored as strings for reliable -contains comparison across PS versions.
# WMI returns signed int32 error codes; we convert to hex string for matching.

$EC_SHUTDOWN      = '8007045B'
$EC_PENDINGREBOOT = '87D00651'
$EC_ALLUPDATES    = '80240022'
$EC_UNEXPECTED    = '8000FFFF'
$EC_CBS_TRANS     = '800F0820'
$EC_COMP_STORE    = '80073712'
$EC_MISSING_DLL   = '8007007E'
$EC_ACCESS_DENIED = '80070005'
$EC_DISK_FULL     = '80070070'
$EC_SUPERSEDED    = '8007066A'
$EC_KEY_NOTFOUND  = '80240008'
$EC_DATA_CONTRACT = '80240439'

#endregion

#region ── Helper functions ────────────────────────────────────────────────────

function Get-FreeDiskGB {
    try {
        $d = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
        if ($d) { return [math]::Round($d.FreeSpace / 1GB, 2) }
    } catch {}
    return 99   # assume fine if we can't query — don't block on unknown
}

function ConvertTo-HexErrorCode {
    # Converts a WMI/CCM error code (may be signed int32 or uint32) to
    # an 8-character uppercase hex string for reliable comparison.
    param($Code)
    try {
        $i = [int64]$Code
        # Negative values are signed representations of HRESULT codes
        if ($i -lt 0) { $i = $i + [int64]'0x100000000' }
        return $i.ToString('X8')
    } catch { return $null }
}

function Get-UpdateErrorCodes {
    # Returns a [string[]] of 8-char hex codes (no 0x prefix) from three sources.
    # Returns empty array (never $null) so callers can safely use -contains.
    $codes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    # Source 1: CCM_SoftwareUpdate WMI — MECM's own per-article error record
    try {
        $updates = Get-WmiObject -Namespace 'ROOT\ccm\SoftMgmtAgent' `
                                 -Class 'CCM_SoftwareUpdate' `
                                 -ErrorAction Stop
        if ($updates) {
            foreach ($u in @($updates)) {
                if ($u.ErrorCode -and $u.ErrorCode -ne 0) {
                    $hex = ConvertTo-HexErrorCode $u.ErrorCode
                    if ($hex) {
                        [void]$codes.Add($hex)
                        Write-Log "CCM_SoftwareUpdate KB$($u.ArticleID): 0x$hex"
                    }
                }
            }
        }
    } catch { Write-Log "CCM_SoftwareUpdate query failed: $_" 'WARN' }

    # Source 2: UpdatesDeployment.log — last 500 lines
    try {
        $udLog = "$env:SystemRoot\CCM\Logs\UpdatesDeployment.log"
        if (Test-Path $udLog) {
            $lines = Get-Content $udLog -Tail 500 -ErrorAction Stop
            foreach ($line in $lines) {
                $matches = [regex]::Matches($line, '0x([0-9A-Fa-f]{8})')
                foreach ($m in $matches) {
                    $hex = $m.Groups[1].Value.ToUpper()
                    if ($hex -ne '00000000' -and $hex -ne '80070000') {
                        [void]$codes.Add($hex)
                    }
                }
            }
        }
    } catch { Write-Log "UpdatesDeployment.log parse failed: $_" 'WARN' }

    # Source 3: WU registry LastError — fallback
    try {
        $k = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install'
        if (Test-Path $k) {
            $val = (Get-ItemProperty -Path $k -ErrorAction Stop).LastError
            if ($val -and $val -ne 0) {
                $hex = ConvertTo-HexErrorCode $val
                if ($hex) { [void]$codes.Add($hex) }
            }
        }
    } catch {}

    Write-Log "Error codes found: $(if($codes.Count -gt 0){ ($codes | ForEach-Object {"0x$_"}) -join ', ' }else{'None'})"
    return @([string[]]$codes)
}

function Test-KBInstalled {
    param([string]$KB)
    $kbID  = if ($KB -match '^[Kk][Bb]') { $KB.ToUpper() } else { "KB$KB" }
    $kbNum = $kbID -replace '^KB',''

    Write-Log "Checking if $kbID is installed..."

    # Method 1: Get-HotFix
    try {
        $hf = Get-HotFix -Id $kbID -ErrorAction SilentlyContinue
        if ($hf) {
            Write-Log "$kbID found via Get-HotFix." 'SUCCESS'
            return [PSCustomObject]@{ Installed = $true; Method = 'HotFix' }
        }
    } catch {}

    # Method 2: Win32_QuickFixEngineering
    try {
        $wmi = Get-WmiObject -Class Win32_QuickFixEngineering `
                             -Filter "HotFixID='$kbID'" -ErrorAction Stop
        if ($wmi) {
            Write-Log "$kbID found via WMI QFE." 'SUCCESS'
            return [PSCustomObject]@{ Installed = $true; Method = 'WMI-QFE' }
        }
    } catch {}

    # Method 3: CBS registry (most reliable for Win11 CUs)
    try {
        $cbsPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages'
        $found = Get-ChildItem -Path $cbsPath -ErrorAction Stop |
                 Where-Object { $_.PSChildName -like "*$kbNum*" } |
                 Select-Object -First 1
        if ($found) {
            $state = (Get-ItemProperty -Path $found.PSPath -ErrorAction SilentlyContinue).CurrentState
            if ($state -eq 112) {
                Write-Log "$kbID found in CBS registry (state=112)." 'SUCCESS'
                return [PSCustomObject]@{ Installed = $true; Method = 'CBS' }
            }
        }
    } catch {}

    # Method 4: DISM (slowest — authoritative fallback)
    try {
        $dismOut = & "$env:SystemRoot\System32\dism.exe" /Online /Get-Packages /Format:Table 2>&1
        if ($dismOut -match $kbNum) {
            Write-Log "$kbID found via DISM." 'SUCCESS'
            return [PSCustomObject]@{ Installed = $true; Method = 'DISM' }
        }
    } catch {}

    Write-Log "$kbID NOT found on this machine." 'WARN'
    return [PSCustomObject]@{ Installed = $false; Method = 'None' }
}

function Invoke-DiskCleanup {
    Write-Log "Running disk cleanup..."

    # Temp folders
    Remove-Item "$env:SystemRoot\Temp\*"        -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:TEMP\*"                   -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:SystemRoot\Logs\CBS\*.cab" -Force  -ErrorAction SilentlyContinue

    # SoftwareDistribution\Download — safe to clear, re-downloads automatically
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        $sdDl = "$env:SystemRoot\SoftwareDistribution\Download"
        if (Test-Path $sdDl) {
            Remove-Item "$sdDl\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Cleared SoftwareDistribution\Download" 'SUCCESS'
        }
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    } catch { Write-Log "SD\Download clear failed: $_" 'WARN' }

    # CCM cache items >30 days old
    try {
        $mgr = New-Object -ComObject UIResource.UIResourceMgr -ErrorAction Stop
        $cache = $mgr.GetCacheInfo()
        $cut = (Get-Date).AddDays(-30)
        foreach ($item in @($cache.GetCacheElements())) {
            if ($item.LastReferenceTime -lt $cut) {
                $cache.DeleteCacheElement($item.CacheElementID)
                Write-Log "  Removed CCM cache: $($item.ContentID)"
            }
        }
    } catch { Write-Log "CCM cache cleanup skipped: $_" 'WARN' }

    # Windows Disk Cleanup utility
    try {
        $root = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
        @('Update Cleanup','Temporary Files','Windows Upgrade Log Files','Recycle Bin') | ForEach-Object {
            $p = Join-Path $root $_
            if (Test-Path $p) {
                Set-ItemProperty -Path $p -Name 'StateFlags0099' -Value 2 -Type DWord -ErrorAction SilentlyContinue
            }
        }
        Start-Process -FilePath cleanmgr.exe -ArgumentList '/sagerun:99' -Wait -ErrorAction SilentlyContinue
    } catch { Write-Log "Disk Cleanup utility failed: $_" 'WARN' }

    $free = Get-FreeDiskGB
    Write-Log "Free space after cleanup: ${free} GB"
    return $free
}

function Reset-WUA {
    Write-Log "Stopping WUA services..."
    @('wuauserv','bits','cryptsvc','msiserver','ccmexec') | ForEach-Object {
        Stop-Service -Name $_ -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 5

    $ts = Get-Date -Format 'yyyyMMddHHmmss'
    foreach ($p in @("$env:SystemRoot\SoftwareDistribution","$env:SystemRoot\System32\catroot2")) {
        if (Test-Path $p) {
            try {
                Rename-Item -Path $p -NewName "${p}.bak_$ts" -Force -ErrorAction Stop
                Write-Log "Renamed: $(Split-Path $p -Leaf)" 'SUCCESS'
            } catch { Write-Log "Could not rename $p`: $_" 'WARN' }
        }
    }

    @('atl.dll','urlmon.dll','mshtml.dll','shdocvw.dll','browseui.dll','jscript.dll',
      'vbscript.dll','scrrun.dll','msxml3.dll','msxml6.dll','actxprxy.dll','softpub.dll',
      'wintrust.dll','dssenh.dll','rsaenh.dll','cryptdlg.dll','oleaut32.dll','ole32.dll',
      'shell32.dll','initpki.dll','wuapi.dll','wuaueng.dll','wucltui.dll','wups.dll',
      'wups2.dll','wuweb.dll','qmgr.dll','qmgrprxy.dll','wucltux.dll','muweb.dll','wuwebv.dll'
    ) | ForEach-Object {
        & "$env:SystemRoot\System32\regsvr32.exe" /s $_ 2>$null
    }

    & "$env:SystemRoot\System32\netsh.exe" winsock reset | Out-Null
    & "$env:SystemRoot\System32\netsh.exe" winhttp reset proxy | Out-Null

    @('cryptsvc','bits','wuauserv') | ForEach-Object {
        Start-Service -Name $_ -ErrorAction SilentlyContinue
    }
    Write-Log "WUA reset complete." 'SUCCESS'
}

function Invoke-DISMRepair {
    Write-Log "Running DISM CheckHealth..."
    $chkOut  = & "$env:SystemRoot\System32\dism.exe" /Online /Cleanup-Image /CheckHealth 2>&1
    $chkExit = $LASTEXITCODE

    if ($chkExit -eq 0 -and ($chkOut -join ' ') -notmatch 'repairable|corruption') {
        Write-Log "DISM CheckHealth: clean — skipping RestoreHealth." 'SUCCESS'
        return 'Clean'
    }

    Write-Log "DISM RestoreHealth starting (may take 10-25 min)..."
    & "$env:SystemRoot\System32\dism.exe" /Online /Cleanup-Image /RestoreHealth /NoRestart 2>&1 | Out-Null
    switch ($LASTEXITCODE) {
        0    { Write-Log "DISM RestoreHealth: success." 'SUCCESS'; return 'Repaired' }
        3010 { Write-Log "DISM RestoreHealth: success, reboot needed." 'SUCCESS'; return 'RepairedRebootNeeded' }
        default {
            Write-Log "DISM RestoreHealth: failed (exit $LASTEXITCODE)." 'ERROR'
            return 'Failed'
        }
    }
}

function Reset-SDACL {
    Write-Log "Resetting SoftwareDistribution ACLs..."
    & "$env:SystemRoot\System32\icacls.exe" "$env:SystemRoot\SoftwareDistribution" /reset /T /C /Q 2>&1 | Out-Null
    Write-Log "ACL reset complete." 'SUCCESS'
}

function Clear-WUADataStore {
    Write-Log "Clearing WUA DataStore cache..."
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    $ds = "$env:SystemRoot\SoftwareDistribution\DataStore"
    if (Test-Path $ds) {
        Remove-Item "$ds\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "DataStore cleared." 'SUCCESS'
    }
    Start-Service -Name wuauserv -ErrorAction SilentlyContinue
}

function Invoke-MECMTriggers {
    Write-Log "Triggering MECM SU scan + deployment eval..."
    # Must use -Arguments (hashtable) not -ArgumentList for SMS_Client.TriggerSchedule
    try {
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' `
            -Name 'TriggerSchedule' `
            -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000113}' } `
            -ErrorAction Stop | Out-Null
        Write-Log "SU Scan triggered." 'SUCCESS'
        Start-Sleep -Seconds 10
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' `
            -Name 'TriggerSchedule' `
            -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000108}' } `
            -ErrorAction Stop | Out-Null
        Write-Log "SU Deploy Eval triggered." 'SUCCESS'
    } catch { Write-Log "MECM trigger failed: $_" 'WARN' }
}

function Schedule-Reboot {
    param([int]$DelaySec = 90)
    Write-Log "Scheduling reboot in ${DelaySec}s..."
    # Remove existing task silently first
    $null = schtasks.exe /Delete /TN 'MECM_Patch_Remediation_Reboot' /F 2>&1
    $at = (Get-Date).AddSeconds($DelaySec).ToString('HH:mm:ss')
    $null = schtasks.exe /Create /TN 'MECM_Patch_Remediation_Reboot' `
        /TR 'shutdown.exe /r /t 10 /f' `
        /SC ONCE /ST $at /RU SYSTEM /RL HIGHEST /F 2>&1
    Write-Log "Reboot scheduled at $at." 'SUCCESS'
}

function Install-Update {
    param([string]$Path)
    Write-Log "Update installation requested: '$Path'"

    # MECM Package ID: 3 letters + 5 digits e.g. PRD00042
    if ($Path -match '^[A-Za-z]{3}\d{5}$') {
        return Install-FromPackage -PackageID $Path.ToUpper()
    }
    # Single .msu file
    if ($Path -match '\.msu$') {
        if (-not (Test-Path $Path)) {
            Write-Log "MSU not found: $Path" 'ERROR'; return 'FileNotFound'
        }
        return Install-MSU -FilePath $Path
    }
    # Folder of .msu files
    if (Test-Path $Path -PathType Container) {
        $files = @(Get-ChildItem -Path $Path -Filter '*.msu' -ErrorAction SilentlyContinue)
        if (-not $files) { Write-Log "No .msu files in: $Path" 'WARN'; return 'NoMSUFound' }
        $results = $files | ForEach-Object { Install-MSU -FilePath $_.FullName }
        return $results -join ' | '
    }
    Write-Log "Invalid UpdatePath: '$Path'" 'ERROR'
    return 'InvalidPath'
}

function Install-MSU {
    param([string]$FilePath)
    $kb = if ($FilePath -match '(KB\d+)') { $Matches[1] } else { [System.IO.Path]::GetFileNameWithoutExtension($FilePath) }
    Write-Log "Installing $kb from $FilePath..."

    # Check already installed
    $check = Test-KBInstalled -KB $kb
    if ($check.Installed) {
        Write-Log "$kb already installed." 'SUCCESS'
        return "$kb-AlreadyInstalled"
    }

    try {
        $logArg = "/log:`"C:\Windows\Temp\wusa_$kb.log`""
        $p = Start-Process -FilePath "$env:SystemRoot\System32\wusa.exe" `
             -ArgumentList "`"$FilePath`" /quiet /norestart $logArg" `
             -Wait -PassThru -NoNewWindow -ErrorAction Stop
        switch ($p.ExitCode) {
            0       { Write-Log "$kb installed." 'SUCCESS'; return "$kb-Installed" }
            3010    { Write-Log "$kb installed, reboot needed." 'SUCCESS'; return "$kb-InstalledRebootNeeded" }
            2359302 { Write-Log "$kb already installed (wusa 2359302)." 'SUCCESS'; return "$kb-AlreadyInstalled" }
            default { Write-Log "$kb failed: wusa exit $($p.ExitCode)." 'ERROR'; return "$kb-Failed($($p.ExitCode))" }
        }
    } catch {
        Write-Log "$kb exception: $_" 'ERROR'
        return "$kb-Exception"
    }
}

function Install-FromPackage {
    param([string]$PackageID)
    Write-Log "Checking CCM cache for Package $PackageID..."
    try {
        $mgr    = New-Object -ComObject UIResource.UIResourceMgr -ErrorAction Stop
        $cache  = $mgr.GetCacheInfo()
        $cached = @($cache.GetCacheElements()) | Where-Object { $_.ContentID -eq $PackageID } | Select-Object -First 1

        if (-not $cached) {
            Write-Log "Package not in cache — triggering policy..."
            Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' `
                -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000021}' } | Out-Null
            Start-Sleep -Seconds 30
            $cached = @($cache.GetCacheElements()) | Where-Object { $_.ContentID -eq $PackageID } | Select-Object -First 1
        }

        if ($cached) {
            $msuFiles = @(Get-ChildItem -Path $cached.Location -Filter '*.msu' -Recurse -ErrorAction SilentlyContinue)
            if ($msuFiles) {
                $results = $msuFiles | ForEach-Object { Install-MSU -FilePath $_.FullName }
                return "Package($PackageID): $($results -join ' | ')"
            }
            Write-Log "No .msu found in package content at $($cached.Location)." 'WARN'
            return "Package($PackageID)-NoMSUInContent"
        }
        Write-Log "Package $PackageID not available after policy trigger." 'WARN'
        return "Package($PackageID)-ContentUnavailable"
    } catch {
        Write-Log "Package install failed: $_" 'ERROR'
        return "Package($PackageID)-Exception"
    }
}

#endregion

#region ── MAIN ────────────────────────────────────────────────────────────────

# Initialise ALL state variables before anything else runs
$actions       = [System.Collections.Generic.List[string]]::new()
$reboot        = $false
$abort         = $false
$updateResult  = 'NotRequested'
$hexCodes      = 'None'
$kbFinalStatus = 'NotChecked'
$freeGBFinal   = $null

Write-Log "════════════════════════════════════════════════════════"
Write-Log "VDI Patch Remediation started on $env:COMPUTERNAME"
Write-Log "KBNumber    : $(if($KBNumber){$KBNumber}else{'(none)'})"
Write-Log "UpdatePath  : $(if($UpdatePath){$UpdatePath}else{'(none)'})"
Write-Log "MinFreeGB   : $MinFreeGB"
Write-Log "RebootDelay : ${RebootDelaySec}s"
Write-Log "════════════════════════════════════════════════════════"

# ── Step 0: KB pre-check ──────────────────────────────────────────────────────
if ($KBNumber) {
    Write-Log "--- Step 0: KB pre-check ---"
    $kbID     = if ($KBNumber -match '^[Kk][Bb]') { $KBNumber.ToUpper() } else { "KB$KBNumber" }
    $kbStatus = Test-KBInstalled -KB $kbID

    if ($kbStatus.Installed) {
        Write-Log "$kbID already installed — exiting." 'SUCCESS'
        $kbFinalStatus = "AlreadyInstalled-$($kbStatus.Method)"
        $abort         = $true
        $actions.Add('KBAlreadyInstalled-Skipped')
    } else {
        Write-Log "$kbID not installed — proceeding."
        $actions.Add("KB-NotInstalled($kbID)")
    }
}

# ── Step 1: Read error codes ──────────────────────────────────────────────────
$errorCodes = @()   # always an array, never $null
if (-not $abort) {
    Write-Log "--- Step 1: Reading WU/CCM error codes ---"
    $errorCodes = Get-UpdateErrorCodes   # returns [string[]] of hex codes, never $null
    $hexCodes   = if ($errorCodes.Count -gt 0) { ($errorCodes | ForEach-Object { "0x$_" }) -join ', ' } else { 'None' }
    Write-Log "Error codes: $hexCodes"
}

# ── Step 2: Disk space ────────────────────────────────────────────────────────
if (-not $abort) {
    Write-Log "--- Step 2: Disk space check ---"
    $freeGB = Get-FreeDiskGB
    Write-Log "C: free: ${freeGB} GB (minimum: ${MinFreeGB} GB)"

    $diskLow = ($freeGB -is [double] -or $freeGB -is [int]) -and ($freeGB -lt $MinFreeGB)

    if (($errorCodes -contains $EC_DISK_FULL) -or $diskLow) {
        Write-Log "Disk below threshold — running cleanup..." 'WARN'
        $freeGB = Invoke-DiskCleanup
        $actions.Add('DiskCleanup')

        if ($freeGB -lt $MinFreeGB) {
            Write-Log "ABORT: ${freeGB} GB after cleanup — still below ${MinFreeGB} GB." 'ERROR'
            $actions.Add("DiskInsufficient-${freeGB}GB")
            $abort = $true
        } else {
            Write-Log "Disk now ${freeGB} GB — OK." 'SUCCESS'
            $reboot = $true
        }
    }
}

# ── Step 3: Error-code-driven remediation ─────────────────────────────────────
if (-not $abort) {
    Write-Log "--- Step 3: Remediation ---"

    if ($errorCodes -contains $EC_SHUTDOWN) {
        Write-Log "0x8007045B: Shutdown-in-progress — clean reboot will fix." 'WARN'
        $reboot = $true
        $actions.Add('0x8007045B-NeedReboot')
    }

    if ($errorCodes -contains $EC_PENDINGREBOOT) {
        Write-Log "0x87D00651: Pending reboot blocking updates." 'WARN'
        $reboot = $true
        $actions.Add('0x87D00651-PendingReboot')
    }

    if ($errorCodes -contains $EC_ACCESS_DENIED) {
        Write-Log "0x80070005: Access denied — ACL reset + WUA reset." 'WARN'
        Reset-SDACL
        Reset-WUA
        $reboot = $true
        $actions.Add('0x80070005-ACL+WUAReset')
    }

    if ($errorCodes -contains $EC_ALLUPDATES) {
        Write-Log "0x80240022: All updates failed — WUA reset." 'WARN'
        Reset-WUA
        $reboot = $true
        $actions.Add('0x80240022-WUAReset')
    }

    if ($errorCodes -contains $EC_UNEXPECTED) {
        Write-Log "0x8000FFFF: Catastrophic failure — WUA reset + DISM." 'WARN'
        Reset-WUA
        $dr = Invoke-DISMRepair
        $reboot = $true
        $actions.Add("0x8000FFFF-WUA+DISM($dr)")
    }

    $cbsCodes = $errorCodes | Where-Object { $_ -in @($EC_CBS_TRANS, $EC_COMP_STORE, $EC_MISSING_DLL) }
    if ($cbsCodes) {
        $hex = ($cbsCodes | ForEach-Object { "0x$_" }) -join '+'
        Write-Log "$hex: CBS/component corruption — DISM RestoreHealth." 'WARN'
        $dr = Invoke-DISMRepair
        $reboot = $true
        $actions.Add("CBS($hex)-DISM($dr)")
        if ($dr -eq 'Failed') {
            Write-Log "DISM failed — VDI may need recomposing." 'ERROR'
        }
    }

    if (($errorCodes -contains $EC_KEY_NOTFOUND) -or ($errorCodes -contains $EC_DATA_CONTRACT)) {
        $which = @()
        if ($errorCodes -contains $EC_KEY_NOTFOUND)  { $which += '0x80240008' }
        if ($errorCodes -contains $EC_DATA_CONTRACT) { $which += '0x80240439' }
        Write-Log "$($which -join '+') — clearing WUA DataStore." 'WARN'
        Clear-WUADataStore
        $reboot = $true
        $actions.Add("$($which -join '+')-DataStoreCleared")
    }

    if ($errorCodes -contains $EC_SUPERSEDED) {
        Write-Log "0x8007066A: Superseded update — clears on next scan." 'WARN'
        $actions.Add('0x8007066A-Noted')
    }

    if ($errorCodes.Count -eq 0) {
        Write-Log "No error codes — triggering policy only."
        $actions.Add('NoErrorCodes-PolicyTriggerOnly')
    }
}

# ── Step 4: Optional update install ───────────────────────────────────────────
if (-not $abort -and $UpdatePath) {
    Write-Log "--- Step 4: Update install ---"
    $updateResult = Install-Update -Path $UpdatePath
    Write-Log "Update result: $updateResult"
    if ($updateResult -notmatch 'AlreadyInstalled|NotRequested|Failed|Exception|NotFound|Invalid|Unavailable|NoMSU') {
        $reboot = $true
    }
    $actions.Add("UpdateInstall:$updateResult")
} elseif (-not $abort) {
    Write-Log "--- Step 4: No UpdatePath — skipping ---"
}

# ── Step 5: MECM triggers ─────────────────────────────────────────────────────
if (-not $abort) {
    Write-Log "--- Step 5: MECM triggers ---"
    Invoke-MECMTriggers
    $actions.Add('MECMTriggered')
}

# ── Step 6: Reboot ────────────────────────────────────────────────────────────
Write-Log "--- Step 6: Reboot ---"
if ($reboot -and -not $abort) {
    # Use schtasks instead of Register-ScheduledTask for broader compatibility
    Schedule-Reboot -DelaySec $RebootDelaySec
    $actions.Add("RebootIn${RebootDelaySec}s")
} elseif ($abort -and $kbFinalStatus -notmatch 'AlreadyInstalled') {
    Write-Log "Reboot skipped — aborted (disk space)." 'WARN'
} elseif ($abort) {
    Write-Log "Reboot skipped — KB already installed." 'INFO'
} else {
    Write-Log "No reboot needed." 'INFO'
}

# ── Post-remediation KB check ─────────────────────────────────────────────────
if ($KBNumber -and $kbFinalStatus -eq 'NotChecked') {
    $kbID2      = if ($KBNumber -match '^[Kk][Bb]') { $KBNumber.ToUpper() } else { "KB$KBNumber" }
    $kbFinal    = Test-KBInstalled -KB $kbID2
    $kbFinalStatus = if ($kbFinal.Installed) { "Installed-$($kbFinal.Method)" } else { 'StillMissing' }
    Write-Log "Post-remediation KB status: $kbFinalStatus"
}

# ── Summary ───────────────────────────────────────────────────────────────────
$freeGBFinal = Get-FreeDiskGB
Write-Log "════════════════════════════════════════════════════════"
Write-Log "Complete — $env:COMPUTERNAME"
Write-Log "ErrorCodes  : $hexCodes"
Write-Log "Actions     : $($actions -join ' | ')"
Write-Log "KBStatus    : $kbFinalStatus"
Write-Log "UpdateResult: $updateResult"
Write-Log "Reboot      : $reboot"
Write-Log "Aborted     : $abort"
Write-Log "FreeGB      : $freeGBFinal"
Write-Log "════════════════════════════════════════════════════════"

# Single Write-Output — the ONLY thing MECM captures in Detailed Output
$kbOut = if ($KBNumber) { if ($KBNumber -match '^[Kk][Bb]') { $KBNumber.ToUpper() } else { "KB$KBNumber" } } else { 'N/A' }
Write-Output ([PSCustomObject]@{
    Computer        = $env:COMPUTERNAME
    KBNumber        = $kbOut
    KBStatus        = $kbFinalStatus
    ErrorCodes      = $hexCodes
    Actions         = $actions -join ' | '
    UpdateInstalled = $updateResult
    RebootScheduled = $reboot
    Aborted         = $abort
    FreeGB          = $freeGBFinal
    LogPath         = $LogPath
} | ConvertTo-Json -Depth 2 -Compress)

#endregion
