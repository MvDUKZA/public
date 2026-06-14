<#
.SYNOPSIS
    VDI patch remediation — right-click Run Script from MECM console.

.PARAMETER KBNumber
    Optional. KB article to check first e.g. KB5039212

.PARAMETER UpdatePath
    Optional. UNC path to .msu, folder of .msu files, or MECM Package ID.

.PARAMETER MinFreeGB
    Minimum free space on C: required. Default: 20

.PARAMETER RebootDelaySec
    Seconds before scheduled reboot fires. Default: 90
#>

# NOTE: No param() block here.
# MECM Run Script injects parameters as variables directly into scope.
# Declaring param() causes a parse conflict in MECM's execution wrapper.
# Parameters are declared in the MECM console script definition instead.
# Variables available: $KBNumber, $UpdatePath, $MinFreeGB, $RebootDelaySec

# Safe defaults for any parameter not supplied
if (-not $KBNumber)       { $KBNumber       = '' }
if (-not $UpdatePath)     { $UpdatePath      = '' }
if (-not $MinFreeGB)      { $MinFreeGB       = 20 }
if (-not $RebootDelaySec) { $RebootDelaySec  = 90 }

$ErrorActionPreference = 'SilentlyContinue'

#region ── Logging ─────────────────────────────────────────────────────────────

$LogPath = 'C:\Windows\Temp\VDIPatchRemediation.log'
$null = New-Item -ItemType Directory -Path 'C:\Windows\Temp' -Force -ErrorAction SilentlyContinue
$null = New-Item -ItemType File -Path $LogPath -Force -ErrorAction SilentlyContinue

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue
}

#endregion

#region ── Error code constants (hex strings, no 0x prefix) ───────────────────

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

#region ── Functions ───────────────────────────────────────────────────────────

function Get-FreeDiskGB {
    try {
        $d = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
        if ($d) { return [math]::Round($d.FreeSpace / 1GB, 2) }
    } catch {}
    return 99
}

function ConvertTo-HexCode {
    param($Code)
    try {
        $i = [int64]$Code
        if ($i -lt 0) { $i = $i + 4294967296 }
        return $i.ToString('X8').ToUpper()
    } catch { return $null }
}

function Get-UpdateErrorCodes {
    $codes = New-Object 'System.Collections.Generic.HashSet[string]'

    # Source 1: CCM_SoftwareUpdate WMI
    try {
        $updates = Get-WmiObject -Namespace 'ROOT\ccm\SoftMgmtAgent' `
                                 -Class 'CCM_SoftwareUpdate' -ErrorAction Stop
        if ($updates) {
            foreach ($u in @($updates)) {
                if ($null -ne $u.ErrorCode -and $u.ErrorCode -ne 0) {
                    $hex = ConvertTo-HexCode $u.ErrorCode
                    if ($hex) {
                        [void]$codes.Add($hex)
                        Write-Log "CCM KB$($u.ArticleID) error: 0x$hex"
                    }
                }
            }
        }
    } catch { Write-Log "CCM_SoftwareUpdate query failed: $_" 'WARN' }

    # Source 2: UpdatesDeployment.log last 500 lines
    try {
        $udLog = "$env:SystemRoot\CCM\Logs\UpdatesDeployment.log"
        if (Test-Path $udLog) {
            $lines = Get-Content $udLog -Tail 500 -ErrorAction Stop
            foreach ($line in @($lines)) {
                $rxMatches = [regex]::Matches($line, '0x([0-9A-Fa-f]{8})')
                foreach ($m in $rxMatches) {
                    $hex = $m.Groups[1].Value.ToUpper()
                    if ($hex -ne '00000000' -and $hex -ne '80070000') {
                        [void]$codes.Add($hex)
                    }
                }
            }
        }
    } catch { Write-Log "UpdatesDeployment.log parse failed: $_" 'WARN' }

    # Source 3: WU registry
    try {
        $k = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install'
        if (Test-Path $k) {
            $val = (Get-ItemProperty -Path $k -ErrorAction Stop).LastError
            if ($null -ne $val -and $val -ne 0) {
                $hex = ConvertTo-HexCode $val
                if ($hex) { [void]$codes.Add($hex) }
            }
        }
    } catch {}

    $arr = @([string[]]$codes)
    Write-Log "Error codes: $(if($arr.Count -gt 0){ ($arr | ForEach-Object {"0x$_"}) -join ', ' }else{'None'})"
    return $arr
}

function Test-KBInstalled {
    param([string]$KB)
    $kbID  = if ($KB -match '^[Kk][Bb]') { $KB.ToUpper() } else { "KB$KB" }
    $kbNum = $kbID -replace '^KB', ''
    Write-Log "Checking $kbID..."

    # Method 1: Get-HotFix
    try {
        $hf = Get-HotFix -Id $kbID -ErrorAction SilentlyContinue
        if ($hf) {
            Write-Log "$kbID found via HotFix." 'SUCCESS'
            return [PSCustomObject]@{ Installed = $true; Method = 'HotFix' }
        }
    } catch {}

    # Method 2: WMI QFE
    try {
        $wmi = Get-WmiObject -Class Win32_QuickFixEngineering `
               -Filter "HotFixID='$kbID'" -ErrorAction Stop
        if ($wmi) {
            Write-Log "$kbID found via WMI QFE." 'SUCCESS'
            return [PSCustomObject]@{ Installed = $true; Method = 'WMI-QFE' }
        }
    } catch {}

    # Method 3: CBS registry
    try {
        $cbsPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages'
        $found = Get-ChildItem -Path $cbsPath -ErrorAction Stop |
                 Where-Object { $_.PSChildName -like "*$kbNum*" } |
                 Select-Object -First 1
        if ($found) {
            $state = (Get-ItemProperty -Path $found.PSPath -ErrorAction SilentlyContinue).CurrentState
            if ($state -eq 112) {
                Write-Log "$kbID found via CBS (state=112)." 'SUCCESS'
                return [PSCustomObject]@{ Installed = $true; Method = 'CBS' }
            }
        }
    } catch {}

    # Method 4: DISM
    try {
        $dismOut = & "$env:SystemRoot\System32\dism.exe" /Online /Get-Packages /Format:Table 2>&1
        if (($dismOut -join ' ') -match $kbNum) {
            Write-Log "$kbID found via DISM." 'SUCCESS'
            return [PSCustomObject]@{ Installed = $true; Method = 'DISM' }
        }
    } catch {}

    Write-Log "$kbID NOT found." 'WARN'
    return [PSCustomObject]@{ Installed = $false; Method = 'None' }
}

function Invoke-DiskCleanup {
    Write-Log "Running disk cleanup..."
    Remove-Item "$env:SystemRoot\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:SystemRoot\Logs\CBS\*.cab" -Force -ErrorAction SilentlyContinue

    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        $sdDl = "$env:SystemRoot\SoftwareDistribution\Download"
        if (Test-Path $sdDl) {
            Remove-Item "$sdDl\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Cleared SD\Download" 'SUCCESS'
        }
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    } catch {}

    try {
        $mgr = New-Object -ComObject UIResource.UIResourceMgr -ErrorAction Stop
        $cache = $mgr.GetCacheInfo()
        $cut = (Get-Date).AddDays(-30)
        foreach ($item in @($cache.GetCacheElements())) {
            if ($item.LastReferenceTime -lt $cut) {
                $cache.DeleteCacheElement($item.CacheElementID)
            }
        }
    } catch {}

    $free = Get-FreeDiskGB
    Write-Log "Free after cleanup: ${free}GB"
    return $free
}

function Reset-WUA {
    Write-Log "WUA reset starting..."
    @('wuauserv','bits','cryptsvc','msiserver','ccmexec') | ForEach-Object {
        Stop-Service -Name $_ -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 5

    $ts = Get-Date -Format 'yyyyMMddHHmmss'
    foreach ($p in @("$env:SystemRoot\SoftwareDistribution","$env:SystemRoot\System32\catroot2")) {
        if (Test-Path $p) {
            Rename-Item -Path $p -NewName "${p}.bak_$ts" -Force -ErrorAction SilentlyContinue
            Write-Log "Renamed: $(Split-Path $p -Leaf)"
        }
    }

    @('atl.dll','urlmon.dll','mshtml.dll','shdocvw.dll','browseui.dll','jscript.dll',
      'vbscript.dll','scrrun.dll','msxml3.dll','msxml6.dll','actxprxy.dll','softpub.dll',
      'wintrust.dll','dssenh.dll','rsaenh.dll','cryptdlg.dll','oleaut32.dll','ole32.dll',
      'shell32.dll','initpki.dll','wuapi.dll','wuaueng.dll','wucltui.dll','wups.dll',
      'wups2.dll','wuweb.dll','qmgr.dll','qmgrprxy.dll','wucltux.dll','muweb.dll','wuwebv.dll'
    ) | ForEach-Object { & "$env:SystemRoot\System32\regsvr32.exe" /s $_ 2>$null }

    & "$env:SystemRoot\System32\netsh.exe" winsock reset 2>&1 | Out-Null
    & "$env:SystemRoot\System32\netsh.exe" winhttp reset proxy 2>&1 | Out-Null

    @('cryptsvc','bits','wuauserv') | ForEach-Object {
        Start-Service -Name $_ -ErrorAction SilentlyContinue
    }
    Write-Log "WUA reset complete." 'SUCCESS'
}

function Invoke-DISMRepair {
    Write-Log "DISM CheckHealth..."
    $chk  = & "$env:SystemRoot\System32\dism.exe" /Online /Cleanup-Image /CheckHealth 2>&1
    $exit = $LASTEXITCODE
    if ($exit -eq 0 -and ($chk -join ' ') -notmatch 'repairable|corruption') {
        Write-Log "DISM: clean." 'SUCCESS'
        return 'Clean'
    }
    Write-Log "DISM RestoreHealth (10-25 min)..."
    & "$env:SystemRoot\System32\dism.exe" /Online /Cleanup-Image /RestoreHealth /NoRestart 2>&1 | Out-Null
    switch ($LASTEXITCODE) {
        0    { Write-Log "DISM: repaired." 'SUCCESS'; return 'Repaired' }
        3010 { Write-Log "DISM: repaired, reboot needed." 'SUCCESS'; return 'RepairedRebootNeeded' }
        default { Write-Log "DISM: failed (exit $LASTEXITCODE)." 'ERROR'; return 'Failed' }
    }
}

function Reset-SDACL {
    Write-Log "Resetting SoftwareDistribution ACLs..."
    & "$env:SystemRoot\System32\icacls.exe" "$env:SystemRoot\SoftwareDistribution" /reset /T /C /Q 2>&1 | Out-Null
    Write-Log "ACL reset done." 'SUCCESS'
}

function Clear-WUADataStore {
    Write-Log "Clearing WUA DataStore..."
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
    Write-Log "Triggering MECM SU scan + deploy eval..."
    try {
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' `
            -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000113}' } `
            -ErrorAction Stop | Out-Null
        Write-Log "SU Scan triggered." 'SUCCESS'
        Start-Sleep -Seconds 10
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' `
            -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000108}' } `
            -ErrorAction Stop | Out-Null
        Write-Log "SU Deploy Eval triggered." 'SUCCESS'
    } catch { Write-Log "MECM trigger failed: $_" 'WARN' }
}

function Schedule-Reboot {
    param([int]$DelaySec = 90)
    Write-Log "Scheduling reboot in ${DelaySec}s..."
    $null = & "$env:SystemRoot\System32\schtasks.exe" /Delete /TN 'MECM_Patch_Remediation_Reboot' /F 2>&1
    $at   = (Get-Date).AddSeconds($DelaySec).ToString('HH:mm:ss')
    $null = & "$env:SystemRoot\System32\schtasks.exe" /Create /TN 'MECM_Patch_Remediation_Reboot' `
            /TR 'shutdown.exe /r /t 10 /f' /SC ONCE /ST $at /RU SYSTEM /RL HIGHEST /F 2>&1
    Write-Log "Reboot scheduled at $at." 'SUCCESS'
}

function Install-MSU {
    param([string]$FilePath)
    $kb    = if ($FilePath -match '(KB\d+)') { $Matches[1] } else { [System.IO.Path]::GetFileNameWithoutExtension($FilePath) }
    $check = Test-KBInstalled -KB $kb
    if ($check.Installed) { Write-Log "$kb already installed." 'SUCCESS'; return "$kb-AlreadyInstalled" }

    Write-Log "Installing $kb..."
    try {
        $p = Start-Process -FilePath "$env:SystemRoot\System32\wusa.exe" `
             -ArgumentList "`"$FilePath`" /quiet /norestart /log:`"C:\Windows\Temp\wusa_$kb.log`"" `
             -Wait -PassThru -NoNewWindow -ErrorAction Stop
        switch ($p.ExitCode) {
            0       { Write-Log "$kb installed." 'SUCCESS'; return "$kb-Installed" }
            3010    { Write-Log "$kb installed, reboot needed." 'SUCCESS'; return "$kb-InstalledRebootNeeded" }
            2359302 { Write-Log "$kb already installed." 'SUCCESS'; return "$kb-AlreadyInstalled" }
            default { Write-Log "$kb failed: exit $($p.ExitCode)." 'ERROR'; return "$kb-Failed($($p.ExitCode))" }
        }
    } catch { Write-Log "$kb exception: $_" 'ERROR'; return "$kb-Exception" }
}

function Install-Update {
    param([string]$Path)
    Write-Log "Install-Update: '$Path'"
    if ($Path -match '^[A-Za-z]{3}\d{5}$') {
        # MECM Package ID
        try {
            $mgr    = New-Object -ComObject UIResource.UIResourceMgr -ErrorAction Stop
            $cache  = $mgr.GetCacheInfo()
            $cached = @($cache.GetCacheElements()) | Where-Object { $_.ContentID -eq $Path.ToUpper() } | Select-Object -First 1
            if (-not $cached) {
                Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' `
                    -Arguments @{ sScheduleID = '{00000000-0000-0000-0000-000000000021}' } | Out-Null
                Start-Sleep -Seconds 30
                $cached = @($cache.GetCacheElements()) | Where-Object { $_.ContentID -eq $Path.ToUpper() } | Select-Object -First 1
            }
            if ($cached) {
                $files = @(Get-ChildItem -Path $cached.Location -Filter '*.msu' -Recurse -ErrorAction SilentlyContinue)
                if ($files) { return ($files | ForEach-Object { Install-MSU $_.FullName }) -join ' | ' }
                return "Package($Path)-NoMSUInContent"
            }
            return "Package($Path)-ContentUnavailable"
        } catch { return "Package($Path)-Exception:$_" }
    }
    if ($Path -match '\.msu$') {
        if (-not (Test-Path $Path)) { return 'FileNotFound' }
        return Install-MSU -FilePath $Path
    }
    if (Test-Path $Path -PathType Container) {
        $files = @(Get-ChildItem -Path $Path -Filter '*.msu' -ErrorAction SilentlyContinue)
        if (-not $files) { return 'NoMSUFound' }
        return ($files | ForEach-Object { Install-MSU $_.FullName }) -join ' | '
    }
    return 'InvalidPath'
}

#endregion

#region ── MAIN ────────────────────────────────────────────────────────────────

# All state variables initialised here — before any step
$actions       = New-Object 'System.Collections.Generic.List[string]'
$reboot        = $false
$abort         = $false
$updateResult  = 'NotRequested'
$hexCodes      = 'None'
$kbFinalStatus = 'NotChecked'
$freeGBFinal   = 99

Write-Log "════════════════════════════════════════════════════════"
Write-Log "VDI Patch Remediation started on $env:COMPUTERNAME"
Write-Log "KBNumber    : $(if($KBNumber){$KBNumber}else{'(none)'})"
Write-Log "UpdatePath  : $(if($UpdatePath){$UpdatePath}else{'(none)'})"
Write-Log "MinFreeGB   : $MinFreeGB"
Write-Log "RebootDelay : ${RebootDelaySec}s"
Write-Log "════════════════════════════════════════════════════════"

# ── Step 0: KB pre-check ──────────────────────────────────────────────────────
if ($KBNumber -and $KBNumber -ne '') {
    Write-Log "--- Step 0: KB pre-check ---"
    $kbID     = if ($KBNumber -match '^[Kk][Bb]') { $KBNumber.ToUpper() } else { "KB$KBNumber" }
    $kbStatus = Test-KBInstalled -KB $kbID
    if ($kbStatus.Installed) {
        Write-Log "$kbID already installed — skipping remediation." 'SUCCESS'
        $kbFinalStatus = "AlreadyInstalled-$($kbStatus.Method)"
        $abort         = $true
        $actions.Add('KBAlreadyInstalled-Skipped')
    } else {
        Write-Log "$kbID not installed — proceeding."
        $actions.Add("KB-NotInstalled($kbID)")
    }
}

# ── Step 1: Read error codes ──────────────────────────────────────────────────
$errorCodes = @()
if (-not $abort) {
    Write-Log "--- Step 1: Reading error codes ---"
    $errorCodes = Get-UpdateErrorCodes
    if ($errorCodes.Count -gt 0) {
        $hexCodes = ($errorCodes | ForEach-Object { "0x$_" }) -join ', '
    }
    Write-Log "Hex codes: $hexCodes"
}

# ── Step 2: Disk space ────────────────────────────────────────────────────────
if (-not $abort) {
    Write-Log "--- Step 2: Disk space ---"
    $freeGB  = Get-FreeDiskGB
    $diskLow = ($freeGB -lt $MinFreeGB)
    Write-Log "C: free ${freeGB}GB  (min ${MinFreeGB}GB)  low=$diskLow"

    if (($errorCodes -contains $EC_DISK_FULL) -or $diskLow) {
        Write-Log "Running disk cleanup..." 'WARN'
        $freeGB = Invoke-DiskCleanup
        $actions.Add('DiskCleanup')
        if ($freeGB -lt $MinFreeGB) {
            Write-Log "ABORT: ${freeGB}GB still below ${MinFreeGB}GB." 'ERROR'
            $actions.Add("DiskInsufficient-${freeGB}GB")
            $abort = $true
        } else {
            Write-Log "Disk OK: ${freeGB}GB." 'SUCCESS'
            $reboot = $true
        }
    }
}

# ── Step 3: Remediation ───────────────────────────────────────────────────────
if (-not $abort) {
    Write-Log "--- Step 3: Remediation ---"

    if ($errorCodes -contains $EC_SHUTDOWN) {
        Write-Log "0x8007045B: Shutdown-in-progress — reboot will fix." 'WARN'
        $reboot = $true; $actions.Add('0x8007045B-NeedReboot')
    }
    if ($errorCodes -contains $EC_PENDINGREBOOT) {
        Write-Log "0x87D00651: Pending reboot." 'WARN'
        $reboot = $true; $actions.Add('0x87D00651-PendingReboot')
    }
    if ($errorCodes -contains $EC_ACCESS_DENIED) {
        Write-Log "0x80070005: Access denied — ACL + WUA reset." 'WARN'
        Reset-SDACL; Reset-WUA
        $reboot = $true; $actions.Add('0x80070005-ACL+WUAReset')
    }
    if ($errorCodes -contains $EC_ALLUPDATES) {
        Write-Log "0x80240022: All updates failed — WUA reset." 'WARN'
        Reset-WUA
        $reboot = $true; $actions.Add('0x80240022-WUAReset')
    }
    if ($errorCodes -contains $EC_UNEXPECTED) {
        Write-Log "0x8000FFFF: Catastrophic — WUA reset + DISM." 'WARN'
        Reset-WUA; $dr = Invoke-DISMRepair
        $reboot = $true; $actions.Add("0x8000FFFF-WUA+DISM($dr)")
    }

    $cbsCodes = @($errorCodes | Where-Object { $_ -in @($EC_CBS_TRANS,$EC_COMP_STORE,$EC_MISSING_DLL) })
    if ($cbsCodes.Count -gt 0) {
        $hex = ($cbsCodes | ForEach-Object { "0x$_" }) -join '+'
        Write-Log "$hex: CBS corruption — DISM." 'WARN'
        $dr = Invoke-DISMRepair
        $reboot = $true; $actions.Add("CBS($hex)-DISM($dr)")
    }

    if (($errorCodes -contains $EC_KEY_NOTFOUND) -or ($errorCodes -contains $EC_DATA_CONTRACT)) {
        $which = @()
        if ($errorCodes -contains $EC_KEY_NOTFOUND)  { $which += '0x80240008' }
        if ($errorCodes -contains $EC_DATA_CONTRACT) { $which += '0x80240439' }
        Write-Log "$($which -join '+') — clearing DataStore." 'WARN'
        Clear-WUADataStore
        $reboot = $true; $actions.Add("$($which -join '+')-DataStoreCleared")
    }
    if ($errorCodes -contains $EC_SUPERSEDED) {
        Write-Log "0x8007066A: Superseded — clears on next scan." 'WARN'
        $actions.Add('0x8007066A-Noted')
    }
    if ($errorCodes.Count -eq 0) {
        $actions.Add('NoErrorCodes-PolicyTriggerOnly')
    }
}

# ── Step 4: Update install ────────────────────────────────────────────────────
if ((-not $abort) -and ($UpdatePath -and $UpdatePath -ne '')) {
    Write-Log "--- Step 4: Update install ---"
    $updateResult = Install-Update -Path $UpdatePath
    Write-Log "Update result: $updateResult"
    if ($updateResult -notmatch 'AlreadyInstalled|NotRequested|Failed|Exception|NotFound|Invalid|NoMSU|Unavailable') {
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
Write-Log "--- Step 6: Reboot (reboot=$reboot abort=$abort) ---"
if ($reboot -and (-not $abort)) {
    Schedule-Reboot -DelaySec $RebootDelaySec
    $actions.Add("RebootIn${RebootDelaySec}s")
} elseif ($abort -and ($kbFinalStatus -match 'AlreadyInstalled')) {
    Write-Log "No reboot — KB already installed." 'INFO'
} elseif ($abort) {
    Write-Log "No reboot — aborted (disk space)." 'WARN'
} else {
    Write-Log "No reboot needed." 'INFO'
}

# ── Post-remediation KB check ─────────────────────────────────────────────────
if (($KBNumber -and $KBNumber -ne '') -and ($kbFinalStatus -eq 'NotChecked')) {
    $kbID2     = if ($KBNumber -match '^[Kk][Bb]') { $KBNumber.ToUpper() } else { "KB$KBNumber" }
    $kbFinal   = Test-KBInstalled -KB $kbID2
    $kbFinalStatus = if ($kbFinal.Installed) { "Installed-$($kbFinal.Method)" } else { 'StillMissing' }
    Write-Log "Post-remediation KB: $kbFinalStatus"
}

# ── Summary and JSON output ───────────────────────────────────────────────────
$freeGBFinal = Get-FreeDiskGB
$kbOut = if ($KBNumber -and $KBNumber -ne '') {
    if ($KBNumber -match '^[Kk][Bb]') { $KBNumber.ToUpper() } else { "KB$KBNumber" }
} else { 'N/A' }

Write-Log "════════════════════════════════════════════════════════"
Write-Log "COMPLETE $env:COMPUTERNAME | KB=$kbOut | Status=$kbFinalStatus"
Write-Log "Codes=$hexCodes | Actions=$($actions -join ' | ')"
Write-Log "Reboot=$reboot | Abort=$abort | FreeGB=$freeGBFinal"
Write-Log "════════════════════════════════════════════════════════"

# This is the ONLY Write-Output in the script.
# MECM captures this as the Script Output in the results pane.
$jsonOut = '{"Computer":"' + $env:COMPUTERNAME + '",' +
           '"KBNumber":"' + $kbOut + '",' +
           '"KBStatus":"' + $kbFinalStatus + '",' +
           '"ErrorCodes":"' + $hexCodes + '",' +
           '"Actions":"' + ($actions -join ' | ') + '",' +
           '"UpdateInstalled":"' + $updateResult + '",' +
           '"RebootScheduled":' + ($reboot.ToString().ToLower()) + ',' +
           '"Aborted":' + ($abort.ToString().ToLower()) + ',' +
           '"FreeGB":' + $freeGBFinal + ',' +
           '"LogPath":"' + $LogPath + '"}'

Write-Output $jsonOut

#endregion
