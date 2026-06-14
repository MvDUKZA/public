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
    0. If KBNumber specified: checks all four methods (HotFix, WMI, CBS registry,
       DISM) — exits immediately with 'AlreadyInstalled' if found. No remediation,
       no reboot, no wasted time.
    1. Reads actual WU/CCM error codes from the machine itself
    2. Checks disk space (20 GB minimum) — cleans up if below threshold
    3. Runs the correct fix per error code:
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
         0x80070070  → Aggressive disk cleanup; aborts if still < 20 GB
    4. Optionally installs a specific update:
         - UNC path to a single .msu file
         - UNC path to a folder (installs all .msu files found)
         - MECM Package ID (triggers CCM content download + install)
    5. Triggers MECM SU scan + deployment eval
    6. Reboots via scheduled task (90s delay) if anything was fixed
       or if the update install requires a reboot

    RETURN VALUE
    ────────────
    JSON object visible in MECM Run Script results pane:
      Computer, KBNumber, KBStatus, ErrorCodes, Actions, UpdateInstalled,
      RebootScheduled, Aborted, FreeGB, LogPath

.PARAMETER KBNumber
    Optional. KB article number to check before doing anything else.
    Accepts with or without the 'KB' prefix: KB5039212 or 5039212
    If specified and the KB is already installed, the script exits immediately
    with status 'AlreadyInstalled' — no remediation, no reboot, nothing.
    If specified and NOT installed, remediation proceeds normally and the KB
    status is re-checked at the end and reported in the JSON output.

.PARAMETER UpdatePath
    Optional. One of:
      \\server\share\KB1234567.msu          → installs that single file
      \\server\share\updates\               → installs all .msu in folder
      PKG00001                              → MECM Package ID (8-char alphanumeric)

.PARAMETER MinFreeGB
    Minimum free space on C: required. Default: 20

.PARAMETER RebootDelaySec
    Seconds between script exit and scheduled reboot firing. Default: 90

.NOTES
    Set MECM Run Script timeout to 1800 seconds (30 minutes) to cover DISM.
    Log written to C:\Windows\Temp\VDIPatchRemediation.log on each target.
    Script must be approved in MECM console before use.
#>

# ── Parameters ────────────────────────────────────────────────────────────────
# MECM Run Script passes parameters as plain strings — no [CmdletBinding()] here.
param(
    [string] $KBNumber       = '',   # e.g. KB5039212 or 5039212 — exits immediately if already installed
    [string] $UpdatePath     = '',
    [int]    $MinFreeGB      = 20,
    [int]    $RebootDelaySec = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Logging ───────────────────────────────────────────────────────────────────
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

# ── Error code constants ──────────────────────────────────────────────────────
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

# ══════════════════════════════════════════════════════════════════════════════
#  FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Get-FreeDiskGB {
    $d = Get-PSDrive -Name C -ErrorAction SilentlyContinue
    if ($d) { return [math]::Round($d.Free / 1GB, 2) }
    return $null
}

# ── KB installation check ────────────────────────────────────────────────────
function Test-KBInstalled {
    param([string]$KB)

    # Normalise — strip 'KB' prefix for numeric comparison, keep full ID for display
    $kbID  = if ($KB -match '^KB') { $KB } else { "KB$KB" }
    $kbNum = $kbID -replace '^KB',''

    Write-Log "Checking if $kbID is installed..."

    # Method 1: Get-HotFix (fast, catches most CUs)
    try {
        $hf = Get-HotFix -Id $kbID -ErrorAction SilentlyContinue
        if ($hf) {
            Write-Log "$kbID found via Get-HotFix (installed: $($hf.InstalledOn))." 'SUCCESS'
            return @{ Installed = $true; Method = 'HotFix'; InstalledOn = $hf.InstalledOn }
        }
    } catch {}

    # Method 2: WMI Win32_QuickFixEngineering (catches some that Get-HotFix misses)
    try {
        $wmi = Get-WmiObject -Class Win32_QuickFixEngineering `
                             -Filter "HotFixID='$kbID'" -EA Stop
        if ($wmi) {
            Write-Log "$kbID found via WMI Win32_QuickFixEngineering." 'SUCCESS'
            return @{ Installed = $true; Method = 'WMI'; InstalledOn = $wmi.InstalledOn }
        }
    } catch {}

    # Method 3: CBS registry — most reliable for Windows 11 CUs
    # Cumulative updates register under HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages
    try {
        $cbsPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages'
        $found   = Get-ChildItem $cbsPath -EA Stop |
                   Where-Object { $_.PSChildName -match $kbNum } |
                   Select-Object -First 1
        if ($found) {
            $state = (Get-ItemProperty $found.PSPath -EA SilentlyContinue).CurrentState
            # CurrentState 112 = Installed
            if ($state -eq 112) {
                Write-Log "$kbID found in CBS registry (CurrentState=112)." 'SUCCESS'
                return @{ Installed = $true; Method = 'CBS'; InstalledOn = 'Unknown' }
            }
        }
    } catch {}

    # Method 4: DISM Get-WindowsPackage (authoritative but slower)
    try {
        $dismOut = & dism.exe /Online /Get-Packages /Format:Table 2>&1 |
                   Select-String $kbNum
        if ($dismOut) {
            Write-Log "$kbID found via DISM Get-Packages." 'SUCCESS'
            return @{ Installed = $true; Method = 'DISM'; InstalledOn = 'Unknown' }
        }
    } catch {}

    Write-Log "$kbID NOT found on this machine." 'WARN'
    return @{ Installed = $false; Method = 'None'; InstalledOn = $null }
}

# ── Read actual error codes from this machine ─────────────────────────────────
function Get-UpdateErrors {
    $codes = [System.Collections.Generic.HashSet[uint32]]::new()

    # Source 1: CCM_SoftwareUpdate WMI — MECM's own per-article error record
    try {
        Get-WmiObject -Namespace 'ROOT\ccm\SoftMgmtAgent' -Class 'CCM_SoftwareUpdate' -EA Stop |
            Where-Object { $_.ErrorCode -and $_.ErrorCode -ne 0 } |
            ForEach-Object {
                $c = [uint32]$_.ErrorCode
                [void]$codes.Add($c)
                Write-Log "CCM_SoftwareUpdate KB$($_.ArticleID): 0x$($c.ToString('X8'))"
            }
    } catch { Write-Log "CCM_SoftwareUpdate query failed: $_" 'WARN' }

    # Source 2: UpdatesDeployment.log — catches codes WMI may have aged out
    try {
        $udLog = "$env:SystemRoot\CCM\Logs\UpdatesDeployment.log"
        if (Test-Path $udLog) {
            (Get-Content $udLog -Tail 500 -EA SilentlyContinue) |
                Select-String '0x[0-9A-Fa-f]{8}' -AllMatches |
                ForEach-Object { $_.Matches } |
                ForEach-Object {
                    try {
                        $c = [uint32]$_.Value
                        if ($c -ne 0 -and $c -ne 0x80070000) { [void]$codes.Add($c) }
                    } catch {}
                }
        }
    } catch { Write-Log "UpdatesDeployment.log parse failed: $_" 'WARN' }

    # Source 3: WU registry LastError — fallback
    try {
        $k = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install'
        if (Test-Path $k) {
            $e = (Get-ItemProperty $k -EA SilentlyContinue).LastError
            if ($e -and $e -ne 0) {
                $c = [uint32]$e
                [void]$codes.Add($c)
                Write-Log "WU Registry LastError: 0x$($c.ToString('X8'))"
            }
        }
    } catch { Write-Log "WU registry read failed: $_" 'WARN' }

    return @($codes)
}

# ── Disk cleanup ──────────────────────────────────────────────────────────────
function Invoke-DiskCleanup {
    Write-Log "Running disk cleanup (target ≥ ${MinFreeGB} GB)..."

    # Windows temp folders
    Remove-Item "$env:SystemRoot\Temp\*" -Recurse -Force -EA SilentlyContinue
    Remove-Item "$env:TEMP\*"            -Recurse -Force -EA SilentlyContinue

    # CBS log cabs (can grow very large)
    Remove-Item "$env:SystemRoot\Logs\CBS\*.cab" -Force -EA SilentlyContinue

    # SoftwareDistribution\Download — safe to clear, will re-download
    $sdDl = "$env:SystemRoot\SoftwareDistribution\Download"
    if (Test-Path $sdDl) {
        Stop-Service wuauserv -Force -EA SilentlyContinue
        Remove-Item "$sdDl\*" -Recurse -Force -EA SilentlyContinue
        Start-Service wuauserv -EA SilentlyContinue
        Write-Log "Cleared SoftwareDistribution\Download" 'SUCCESS'
    }

    # CCM cache — remove items not referenced in last 30 days
    try {
        $mgr   = New-Object -ComObject UIResource.UIResourceMgr -EA Stop
        $cache = $mgr.GetCacheInfo()
        $cut   = (Get-Date).AddDays(-30)
        foreach ($item in $cache.GetCacheElements()) {
            if ($item.LastReferenceTime -lt $cut) {
                $cache.DeleteCacheElement($item.CacheElementID)
                Write-Log "  Removed CCM cache item: $($item.ContentID)"
            }
        }
    } catch { Write-Log "CCM cache cleanup skipped: $_" 'WARN' }

    # Windows Disk Cleanup utility — Update Cleanup + Temp Files
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
        Write-Log "Windows Disk Cleanup complete." 'SUCCESS'
    } catch { Write-Log "Disk Cleanup utility failed: $_" 'WARN' }

    $free = Get-FreeDiskGB
    Write-Log "Free space after cleanup: ${free} GB"
    return $free
}

# ── WUA full reset ────────────────────────────────────────────────────────────
function Reset-WUA {
    Write-Log "Stopping WUA services..."
    @('wuauserv','bits','cryptsvc','msiserver','ccmexec') |
        ForEach-Object { Stop-Service $_ -Force -EA SilentlyContinue }
    Start-Sleep -Seconds 5

    $ts = Get-Date -Format 'yyyyMMddHHmmss'
    foreach ($p in @(
        "$env:SystemRoot\SoftwareDistribution",
        "$env:SystemRoot\System32\catroot2"
    )) {
        if (Test-Path $p) {
            try { Rename-Item $p "${p}.bak_$ts" -Force; Write-Log "Renamed: $p" 'SUCCESS' }
            catch { Write-Log "Could not rename $p`: $_" 'WARN' }
        }
    }

    Write-Log "Re-registering WUA/BITS DLLs..."
    @(
        'atl.dll','urlmon.dll','mshtml.dll','shdocvw.dll','browseui.dll','jscript.dll',
        'vbscript.dll','scrrun.dll','msxml3.dll','msxml6.dll','actxprxy.dll','softpub.dll',
        'wintrust.dll','dssenh.dll','rsaenh.dll','cryptdlg.dll','oleaut32.dll','ole32.dll',
        'shell32.dll','initpki.dll','wuapi.dll','wuaueng.dll','wucltui.dll','wups.dll',
        'wups2.dll','wuweb.dll','qmgr.dll','qmgrprxy.dll','wucltux.dll','muweb.dll','wuwebv.dll'
    ) | ForEach-Object { & regsvr32.exe /s $_ 2>$null }

    & netsh winsock reset  | Out-Null
    & netsh winhttp reset proxy | Out-Null

    @('cryptsvc','bits','wuauserv') |
        ForEach-Object { Start-Service $_ -EA SilentlyContinue }
    Write-Log "WUA reset complete." 'SUCCESS'
}

# ── DISM repair ───────────────────────────────────────────────────────────────
function Invoke-DISMRepair {
    Write-Log "DISM CheckHealth..."
    $chk = & dism.exe /Online /Cleanup-Image /CheckHealth 2>&1
    if ($LASTEXITCODE -eq 0 -and ($chk -join ' ') -notmatch 'repairable|corruption') {
        Write-Log "DISM: no corruption found." 'SUCCESS'
        return 'Clean'
    }
    Write-Log "DISM RestoreHealth (may take up to 25 min)..."
    & dism.exe /Online /Cleanup-Image /RestoreHealth /NoRestart 2>&1 | Out-Null
    switch ($LASTEXITCODE) {
        0    { Write-Log "DISM RestoreHealth: success." 'SUCCESS'; return 'Repaired' }
        3010 { Write-Log "DISM RestoreHealth: success, reboot needed." 'SUCCESS'; return 'RepairedRebootNeeded' }
        default {
            Write-Log "DISM RestoreHealth failed (exit $LASTEXITCODE)." 'ERROR'
            return 'Failed'
        }
    }
}

# ── ACL reset on SoftwareDistribution ─────────────────────────────────────────
function Reset-SDACL {
    Write-Log "Resetting SoftwareDistribution ACLs..."
    & icacls "$env:SystemRoot\SoftwareDistribution" /reset /T /C /Q | Out-Null
    Write-Log "ACL reset complete." 'SUCCESS'
}

# ── Clear WUA DataStore ────────────────────────────────────────────────────────
function Clear-DataStore {
    Write-Log "Clearing WUA DataStore cache..."
    Stop-Service wuauserv -Force -EA SilentlyContinue
    Start-Sleep -Seconds 3
    $ds = "$env:SystemRoot\SoftwareDistribution\DataStore"
    if (Test-Path $ds) {
        Remove-Item "$ds\*" -Recurse -Force -EA SilentlyContinue
        Write-Log "DataStore cleared." 'SUCCESS'
    }
    Start-Service wuauserv -EA SilentlyContinue
}

# ── MECM triggers ─────────────────────────────────────────────────────────────
function Invoke-MECMTriggers {
    Write-Log "Triggering MECM SU scan + deployment eval..."
    try {
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' `
            -Name 'TriggerSchedule' `
            -ArgumentList '{00000000-0000-0000-0000-000000000113}' | Out-Null
        Start-Sleep -Seconds 10
        Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' `
            -Name 'TriggerSchedule' `
            -ArgumentList '{00000000-0000-0000-0000-000000000108}' | Out-Null
        Write-Log "MECM triggers sent." 'SUCCESS'
    } catch { Write-Log "MECM trigger failed: $_" 'WARN' }
}

# ── Schedule reboot ───────────────────────────────────────────────────────────
function Schedule-Reboot {
    Write-Log "Scheduling reboot in ${RebootDelaySec}s..."
    Unregister-ScheduledTask 'MECM_Patch_Remediation_Reboot' -Confirm:$false -EA SilentlyContinue
    $at = (Get-Date).AddSeconds($RebootDelaySec)
    $a  = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument '-r -t 10 -f'
    $t  = New-ScheduledTaskTrigger -Once -At $at
    Register-ScheduledTask 'MECM_Patch_Remediation_Reboot' `
        -Action $a -Trigger $t -RunLevel Highest -User SYSTEM -Force | Out-Null
    Write-Log "Reboot scheduled at $at." 'SUCCESS'
}

# ── Update installation ───────────────────────────────────────────────────────
function Install-Update {
    param([string]$Path)

    Write-Log "Update installation requested: '$Path'"

    # ── Detect input type ─────────────────────────────────────────────────

    # MECM Package ID: 8-char alphanumeric e.g. PRD00042
    if ($Path -match '^[A-Z0-9]{3}\d{5}$') {
        Write-Log "Detected MECM Package ID: $Path"
        return Install-UpdateFromPackage -PackageID $Path
    }

    # Single .msu file
    if ($Path -match '\.msu$') {
        if (-not (Test-Path $Path)) {
            Write-Log "MSU file not found: $Path" 'ERROR'
            return 'FileNotFound'
        }
        return Install-MSU -FilePath $Path
    }

    # Folder — install all .msu files found
    if (Test-Path $Path -PathType Container) {
        $files = Get-ChildItem -Path $Path -Filter '*.msu' -ErrorAction SilentlyContinue
        if (-not $files) {
            Write-Log "No .msu files found in folder: $Path" 'WARN'
            return 'NoMSUFound'
        }
        Write-Log "Found $($files.Count) .msu file(s) in folder."
        $results = foreach ($f in $files) { Install-MSU -FilePath $f.FullName }
        return $results -join ' | '
    }

    Write-Log "UpdatePath '$Path' is not a valid MSU file, folder, or Package ID." 'ERROR'
    return 'InvalidPath'
}

function Install-MSU {
    param([string]$FilePath)

    $kb = if ($FilePath -match '(KB\d+)') { $Matches[1] } else { Split-Path $FilePath -Leaf }
    Write-Log "Installing $kb from $FilePath ..."

    # Check if already installed
    $installed = Get-HotFix -Id $kb -ErrorAction SilentlyContinue
    if ($installed) {
        Write-Log "$kb already installed (installed: $($installed.InstalledOn))." 'SUCCESS'
        return "$kb-AlreadyInstalled"
    }

    try {
        # wusa.exe /quiet /norestart — MECM will handle the reboot
        $p = Start-Process wusa.exe `
             -ArgumentList "`"$FilePath`" /quiet /norestart /log:`"C:\Windows\Temp\wusa_$kb.log`"" `
             -Wait -PassThru -NoNewWindow -ErrorAction Stop

        switch ($p.ExitCode) {
            0     { Write-Log "$kb installed successfully." 'SUCCESS'; return "$kb-Installed" }
            3010  { Write-Log "$kb installed — reboot required." 'SUCCESS'; return "$kb-InstalledRebootNeeded" }
            2359302 {
                Write-Log "$kb already installed (wusa exit 2359302)." 'SUCCESS'
                return "$kb-AlreadyInstalled"
            }
            default {
                Write-Log "$kb install failed — wusa exit $($p.ExitCode)." 'ERROR'
                return "$kb-Failed($($p.ExitCode))"
            }
        }
    } catch {
        Write-Log "$kb install exception: $_" 'ERROR'
        return "$kb-Exception"
    }
}

function Install-UpdateFromPackage {
    param([string]$PackageID)

    Write-Log "Requesting MECM content download for Package $PackageID ..."
    try {
        # Check if package is already in CCM cache
        $mgr   = New-Object -ComObject UIResource.UIResourceMgr -EA Stop
        $cache = $mgr.GetCacheInfo()
        $cached = $cache.GetCacheElements() |
                  Where-Object { $_.ContentID -eq $PackageID } |
                  Select-Object -First 1

        if (-not $cached) {
            Write-Log "Package $PackageID not in CCM cache — requesting download..."
            # Trigger Machine Policy to pull the package assignment
            Invoke-WmiMethod -Namespace 'ROOT\ccm' -Class 'SMS_Client' `
                -Name 'TriggerSchedule' `
                -ArgumentList '{00000000-0000-0000-0000-000000000021}' | Out-Null
            Start-Sleep -Seconds 30

            # Re-check cache
            $cached = $cache.GetCacheElements() |
                      Where-Object { $_.ContentID -eq $PackageID } |
                      Select-Object -First 1
        }

        if ($cached) {
            Write-Log "Package content found at: $($cached.Location)"
            # Look for .msu files in the cached content
            $msuFiles = Get-ChildItem -Path $cached.Location -Filter '*.msu' -Recurse -EA SilentlyContinue
            if ($msuFiles) {
                $results = foreach ($f in $msuFiles) { Install-MSU -FilePath $f.FullName }
                return "Package($PackageID): $($results -join ' | ')"
            }
            # Look for setup.exe or install.cmd as fallback
            $installer = Get-ChildItem -Path $cached.Location `
                         -Include 'setup.exe','install.cmd','install.bat' `
                         -Recurse -EA SilentlyContinue | Select-Object -First 1
            if ($installer) {
                Write-Log "Running installer: $($installer.FullName)"
                $p = Start-Process $installer.FullName -ArgumentList '/quiet /norestart' `
                     -Wait -PassThru -NoNewWindow
                return "Package($PackageID)-Installer:exit$($p.ExitCode)"
            }
            Write-Log "No .msu or installer found in package content." 'WARN'
            return "Package($PackageID)-NoInstaller"
        } else {
            Write-Log "Package $PackageID content not available after policy trigger." 'WARN'
            return "Package($PackageID)-ContentUnavailable"
        }
    } catch {
        Write-Log "Package install failed: $_" 'ERROR'
        return "Package($PackageID)-Exception"
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

Write-Log "════════════════════════════════════════════════════════"
Write-Log "VDI Patch Remediation started — $env:COMPUTERNAME"
Write-Log "KBNumber    : $(if($KBNumber){"'$KBNumber'"}else{'(none)'})"
Write-Log "UpdatePath  : $(if($UpdatePath){"'$UpdatePath'"}else{'(none)'})"
Write-Log "MinFreeGB   : $MinFreeGB"
Write-Log "RebootDelay : ${RebootDelaySec}s"
Write-Log "════════════════════════════════════════════════════════"

$actions      = [System.Collections.Generic.List[string]]::new()
$reboot       = $false
$abort        = $false
$updateResult = 'NotRequested'

# ── Step 0: KB pre-check — exit immediately if already installed ──────────────
if ($KBNumber) {
    Write-Log "--- Step 0: KB pre-check ---"
    $kbID     = if ($KBNumber -match '^KB') { $KBNumber } else { "KB$KBNumber" }
    $kbStatus = Test-KBInstalled -KB $kbID

    if ($kbStatus.Installed) {
        Write-Log "$kbID already installed on $env:COMPUTERNAME — nothing to do." 'SUCCESS'

        # Return early — no remediation, no reboot, clean exit
        [PSCustomObject]@{
            Computer        = $env:COMPUTERNAME
            KBNumber        = $kbID
            KBStatus        = "AlreadyInstalled ($($kbStatus.Method))"
            ErrorCodes      = 'NotChecked'
            Actions         = 'KBAlreadyInstalled-Skipped'
            UpdateInstalled = 'AlreadyInstalled'
            RebootScheduled = $false
            Aborted         = $false
            FreeGB          = (Get-FreeDiskGB)
            LogPath         = $LogPath
        } | ConvertTo-Json -Depth 3
        return
    }

    Write-Log "$kbID not installed — proceeding with remediation."
    $actions.Add("KB-NotInstalled($kbID)")
}

# ── Step 1: Read error codes ──────────────────────────────────────────────────
Write-Log "--- Step 1: Reading WU/CCM error codes ---"
$codes    = Get-UpdateErrors
$hexCodes = ($codes | ForEach-Object { '0x' + $_.ToString('X8') }) -join ', '

if ($codes.Count -eq 0) {
    Write-Log "No active error codes found on this machine." 'WARN'
} else {
    Write-Log "Error codes: $hexCodes"
}

# ── Step 2: Disk space ────────────────────────────────────────────────────────
Write-Log "--- Step 2: Disk space check ---"
$freeGB  = Get-FreeDiskGB
Write-Log "C: free: ${freeGB} GB (minimum: ${MinFreeGB} GB)"

if (($codes -contains $EC.DiskFull) -or ($freeGB -lt $MinFreeGB)) {
    Write-Log "Below threshold — running cleanup..." 'WARN'
    $freeGB = Invoke-DiskCleanup
    $actions.Add('DiskCleanup')
    if ($freeGB -lt $MinFreeGB) {
        Write-Log "ABORT: ${freeGB} GB free after cleanup — still below ${MinFreeGB} GB minimum." 'ERROR'
        Write-Log "Manual intervention required: check WinSxS, user profiles, app logs." 'WARN'
        $actions.Add("DiskInsufficient-${freeGB}GB")
        $abort = $true
    } else {
        Write-Log "Disk now ${freeGB} GB — sufficient." 'SUCCESS'
        $reboot = $true   # disk cleanup warrants a reboot
    }
}

# ── Step 3: Error-code-driven remediation ────────────────────────────────────
if (-not $abort) {
    Write-Log "--- Step 3: Remediation ---"

    if ($codes -contains $EC.Shutdown) {
        Write-Log "0x8007045B: Shutdown-in-progress during last install. Reboot will resolve." 'WARN'
        $reboot = $true
        $actions.Add('0x8007045B-NeedReboot')
    }

    if ($codes -contains $EC.PendingReboot) {
        Write-Log "0x87D00651: Pending reboot blocking updates." 'WARN'
        $reboot = $true
        $actions.Add('0x87D00651-PendingReboot')
    }

    if ($codes -contains $EC.AccessDenied) {
        Write-Log "0x80070005: Access denied — ACL reset + WUA reset." 'WARN'
        Reset-SDACL
        Reset-WUA
        $reboot = $true
        $actions.Add('0x80070005-ACL+WUAReset')
    }

    if ($codes -contains $EC.AllUpdates) {
        Write-Log "0x80240022: All updates failed — WUA reset." 'WARN'
        Reset-WUA
        $reboot = $true
        $actions.Add('0x80240022-WUAReset')
    }

    if ($codes -contains $EC.Unexpected) {
        Write-Log "0x8000FFFF: Catastrophic failure — WUA reset + DISM." 'WARN'
        Reset-WUA
        $dr = Invoke-DISMRepair
        $reboot = $true
        $actions.Add("0x8000FFFF-WUA+DISM($dr)")
    }

    $cbsCodes = $codes | Where-Object { $_ -in @($EC.CBSTrans, $EC.CompStore, $EC.MissingDll) }
    if ($cbsCodes) {
        $hex = ($cbsCodes | ForEach-Object { '0x' + $_.ToString('X8') }) -join '+'
        Write-Log "$hex: CBS/component corruption — DISM RestoreHealth." 'WARN'
        $dr = Invoke-DISMRepair
        $reboot = $true
        $actions.Add("CBS($hex)-DISM($dr)")
        if ($dr -eq 'Failed') {
            Write-Log "DISM failed — this VDI may need recomposing from golden image." 'ERROR'
        }
    }

    if (($codes -contains $EC.KeyNotFound) -or ($codes -contains $EC.DataContract)) {
        $matched = @()
        if ($codes -contains $EC.KeyNotFound)  { $matched += '0x80240008' }
        if ($codes -contains $EC.DataContract) { $matched += '0x80240439' }
        Write-Log "$($matched -join '+') — clearing WUA DataStore." 'WARN'
        Clear-DataStore
        $reboot = $true
        $actions.Add("$($matched -join '+')-DataStoreCleared")
    }

    if ($codes -contains $EC.Superseded) {
        Write-Log "0x8007066A: Superseded update — will clear on next MECM scan." 'WARN'
        $actions.Add('0x8007066A-Noted')
    }

    if ($codes.Count -eq 0) {
        $actions.Add('NoErrorCodes-PolicyTriggerOnly')
    }
}

# ── Step 4: Optional update install ──────────────────────────────────────────
if (-not $abort -and $UpdatePath) {
    Write-Log "--- Step 4: Update installation ---"
    $updateResult = Install-Update -Path $UpdatePath
    Write-Log "Update result: $updateResult"
    # Any install attempt (success or reboot-needed) warrants a reboot
    if ($updateResult -notmatch 'AlreadyInstalled|NotRequested|Failed|Exception|NotFound|Invalid|Unavailable') {
        $reboot = $true
    }
    $actions.Add("UpdateInstall:$updateResult")
} else {
    Write-Log "--- Step 4: No update path specified — skipping ---"
}

# ── Step 5: MECM triggers ─────────────────────────────────────────────────────
Write-Log "--- Step 5: MECM triggers ---"
if (-not $abort) {
    Invoke-MECMTriggers
    $actions.Add('MECMTriggered')
}

# ── Step 6: Reboot ────────────────────────────────────────────────────────────
Write-Log "--- Step 6: Reboot ---"
if ($reboot -and -not $abort) {
    Schedule-Reboot
    $actions.Add("RebootIn${RebootDelaySec}s")
} elseif ($abort) {
    Write-Log "Reboot skipped — aborted due to disk space." 'WARN'
} else {
    Write-Log "Nothing fixed — no reboot needed." 'INFO'
}

# ── Summary ───────────────────────────────────────────────────────────────────
$freeGBFinal = Get-FreeDiskGB
Write-Log "════════════════════════════════════════════════════════"
Write-Log "Complete — $env:COMPUTERNAME"
Write-Log "Error codes  : $(if($hexCodes){$hexCodes}else{'None'})"
Write-Log "Actions      : $($actions -join ' | ')"
Write-Log "Update result: $updateResult"
Write-Log "Reboot       : $reboot"
Write-Log "Aborted      : $abort"
Write-Log "Free disk    : ${freeGBFinal} GB"
Write-Log "════════════════════════════════════════════════════════"

# Post-remediation KB check — confirm whether the KB is now present
$kbFinalStatus = 'NotChecked'
if ($KBNumber) {
    $kbID         = if ($KBNumber -match '^KB') { $KBNumber } else { "KB$KBNumber" }
    $kbFinal      = Test-KBInstalled -KB $kbID
    $kbFinalStatus= if ($kbFinal.Installed) { "Installed ($($kbFinal.Method))" } else { 'StillMissing' }
    Write-Log "Post-remediation KB status: $kbFinalStatus"
}

# JSON returned to MECM Run Script results pane
[PSCustomObject]@{
    Computer        = $env:COMPUTERNAME
    KBNumber        = if ($KBNumber) { if ($KBNumber -match '^KB') { $KBNumber } else { "KB$KBNumber" } } else { 'N/A' }
    KBStatus        = $kbFinalStatus
    ErrorCodes      = if ($hexCodes) { $hexCodes } else { 'None' }
    Actions         = $actions -join ' | '
    UpdateInstalled = $updateResult
    RebootScheduled = $reboot
    Aborted         = $abort
    FreeGB          = $freeGBFinal
    LogPath         = $LogPath
} | ConvertTo-Json -Depth 3
