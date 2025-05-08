<#
.SYNOPSIS
    Build a WinPE ISO, inject VMware drivers and custom scripts.

.DESCRIPTION
    - Requires Windows ADK + WinPE Add-on installed.
    - Uses copype and MakeWinPEMedia from ADK.
    - Logs every step, aborts on error.
#>

#region Configuration
# Architecture (only amd64 supported for Win11/Server)
$WinPEArch          = 'amd64'

# Paths (hardcoded)
$WinPEWorkRoot      = 'C:\WinPE_Custom'           # working folder
$VMWareDriverRoot   = 'C:\Drivers\VMWare'         # folder with *.inf (and subfolders)
$CustomScriptsRoot  = 'C:\WinPE\CustomScripts'    # contains capture_restore.cmd + scripts\
$OutputISO          = 'C:\WinPE_Custom.iso'       # output ISO path

# ADK tools location
$ADKWinPEPath       = "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment"
$CopypeCmd          = Join-Path $ADKWinPEPath 'copype.cmd'
$MakeWinPEMediaCmd  = Join-Path $ADKWinPEPath 'MakeWinPEMedia.cmd'

# Log file
$LogFile            = Join-Path $WinPEWorkRoot 'BuildWinPE.log'
#endregion

#region Logging Function
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    if ($Level -eq 'ERROR') {
        Write-Error $Message
    } else {
        Write-Host $line
    }
}
#endregion

try {
    Write-Log "===== Starting WinPE build ====="

    # Validate ADK tools
    if (-not (Test-Path $CopypeCmd)) {
        throw "copype.cmd not found at $CopypeCmd"
    }
    if (-not (Test-Path $MakeWinPEMediaCmd)) {
        throw "MakeWinPEMedia.cmd not found at $MakeWinPEMediaCmd"
    }
    Write-Log "Found ADK tools."

    # Validate driver and scripts folder
    foreach ($path in @($VMWareDriverRoot, $CustomScriptsRoot)) {
        if (-not (Test-Path $path)) {
            throw "Required path missing: $path"
        }
    }
    Write-Log "Driver and scripts folders exist."

    # Clean up prior workspace
    if (Test-Path $WinPEWorkRoot) {
        Write-Log "Removing existing workspace $WinPEWorkRoot"
        Remove-Item -Recurse -Force $WinPEWorkRoot
    }
    New-Item -ItemType Directory -Path $WinPEWorkRoot | Out-Null

    # Step 1: copype
    Write-Log "Running copype for $WinPEArch"
    & "$CopypeCmd" $WinPEArch $WinPEWorkRoot 2>&1 | ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "copype failed" }

    # Paths derived
    $MediaDir = Join-Path $WinPEWorkRoot 'media'
    $MountDir = Join-Path $WinPEWorkRoot 'mount'

    # Step 2: Mount boot.wim
    $BootWim = Join-Path $MediaDir 'sources\boot.wim'
    Write-Log "Creating mount folder $MountDir"
    New-Item -ItemType Directory -Path $MountDir | Out-Null

    Write-Log "Mounting $BootWim to $MountDir"
    dism /Mount-Wim /WimFile:$BootWim /Index:1 /MountDir:$MountDir /ReadOnly:No 2>&1 |
        ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "Failed to mount WIM" }

    # Step 3: Inject VMware drivers
    Write-Log "Injecting VMware drivers from $VMWareDriverRoot"
    dism /Image:$MountDir /Add-Driver /Driver:$VMWareDriverRoot /Recurse 2>&1 |
        ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "Failed to add VMware drivers" }

    # Step 4: Copy custom scripts
    Write-Log "Copying custom scripts from $CustomScriptsRoot"
    Copy-Item -Path (Join-Path $CustomScriptsRoot '*') -Destination $MountDir -Recurse -Force |
        ForEach-Object { Write-Log "Copied $_" }

    # Ensure startnet.cmd exists and only runs wpeinit
    $StartNet    = Join-Path $MountDir 'Windows\System32\startnet.cmd'
    $StartNetDir = Split-Path $StartNet -Parent

    if (-not (Test-Path $StartNetDir)) {
        Write-Log "Creating directory for startnet.cmd: $StartNetDir"
        New-Item -ItemType Directory -Path $StartNetDir -Force | Out-Null
    }

    if (-not (Test-Path $StartNet)) {
        Write-Log "startnet.cmd not found. Creating new file."
    } else {
        Write-Log "startnet.cmd already exists. Overwriting content."
    }
    Set-Content -Path $StartNet -Value 'wpeinit'
    Write-Log "startnet.cmd set to only run wpeinit."

    # Step 5: Unmount & commit
    Write-Log "Unmounting and committing changes"
    dism /Unmount-Wim /MountDir:$MountDir /Commit 2>&1 |
        ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "Failed to unmount WIM" }

    # Step 6: Build ISO
    Write-Log "Creating ISO $OutputISO"
    & "$MakeWinPEMediaCmd" "/ISO" $WinPEWorkRoot $OutputISO 2>&1 |
        ForEach-Object { Write-Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "MakeWinPEMedia failed" }

    Write-Log "WinPE ISO successfully created at $OutputISO"
    Write-Host "Done. ISO located at: $OutputISO"
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    exit 1
}
finally {
    Write-Log "===== Build script finished ====="
}
