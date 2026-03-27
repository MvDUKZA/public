#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Reads the NVIDIA vGPU Manager version stamped by vCenter into GuestInfo,
    compares it to the currently installed NVIDIA guest driver, and installs
    the matching driver if required.

.DESCRIPTION
    Intended to be deployed as an SCCM Application or Script, run as SYSTEM.
    
    Logic:
      1. Read guestinfo.nvidia.vgpumanager.version via vmtoolsd.exe
      2. Read currently installed NVIDIA driver version from the registry
      3. Compare major branch versions (first two octets, e.g. 17.4 → "17")
      4. If mismatch (or no driver installed), run the silent installer from
         a UNC share / local cache path you specify in $DriverSourceRoot
      5. Exit with appropriate code for SCCM (0 = success, 1 = reboot needed,
         non-zero = failure)

.NOTES
    Set $DriverSourceRoot to a UNC path or local path where driver installers
    are organised as subfolders named by vGPU branch version, e.g.:
        \\sccm-share\nvidia-drivers\
            17\setup.exe
            16\setup.exe

    The GuestInfo key must have been stamped by Set-NvidiaGuestInfo.ps1 first.

    SCCM Detection Method: Use a script detection rule calling this script with
    -DetectOnly, which exits 0 if driver is current, 1 if not.
#>

[CmdletBinding()]
param(
    # Path to folder containing branch-named subfolders with setup.exe
    [string]$DriverSourceRoot = "\\your-sccm-share\nvidia-vgpu-drivers",

    # Switch: only check/detect — do not install. For SCCM detection rules.
    [switch]$DetectOnly,

    # GuestInfo key to read (must match Set-NvidiaGuestInfo.ps1)
    [string]$GuestInfoKey = "guestinfo.nvidia.vgpumanager.version",

    # Silent installer arguments
    [string]$InstallerArgs = "-s -noreboot"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Logging ───────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # SCCM script output goes to stdout; no colour codes needed
    Write-Output "[$ts] [$Level] $Message"
}

# ── Exit codes ────────────────────────────────────────────────────────────────
$EXIT_OK           = 0
$EXIT_REBOOT       = 1
$EXIT_NOT_CURRENT  = 1   # Used in DetectOnly mode to signal "not compliant"
$EXIT_ERROR        = 2

# ── Step 1: Read GuestInfo from VMware Tools ──────────────────────────────────
Write-Log "Reading GuestInfo key: $GuestInfoKey"

$vmToolsExe = Join-Path $env:ProgramFiles "VMware\VMware Tools\vmtoolsd.exe"
if (-not (Test-Path $vmToolsExe)) {
    Write-Log "VMware Tools not found at expected path. Is this a VMware guest?" "ERROR"
    exit $EXIT_ERROR
}

try {
    $hostVibVersion = & $vmToolsExe --cmd "info-get $GuestInfoKey" 2>&1
    $hostVibVersion = $hostVibVersion.Trim()

    if ([string]::IsNullOrWhiteSpace($hostVibVersion) -or $hostVibVersion -match "No value found") {
        Write-Log "GuestInfo key not set. Run Set-NvidiaGuestInfo.ps1 on vCenter first." "ERROR"
        exit $EXIT_ERROR
    }

    Write-Log "Host vGPU Manager version (from GuestInfo): $hostVibVersion"
} catch {
    Write-Log "Failed to read GuestInfo: $_" "ERROR"
    exit $EXIT_ERROR
}

# ── Step 2: Parse the branch from the VIB version ────────────────────────────
# NVIDIA vGPU VIB versions look like: 550.90.05-1OEM.800.1.0.20613240
# The guest driver branch maps to the first segment major version.
# We extract the leading major number (e.g. 550 → branch "550", or use
# the vGPU release branch like 17.x where VIB version starts with that).
# Adjust the regex below to match your environment's versioning scheme.

# This regex extracts the first numeric group before a dot or dash
if ($hostVibVersion -match '^(\d+)[\.\-]') {
    $hostBranch = $Matches[1]
    Write-Log "Host vGPU branch: $hostBranch"
} else {
    Write-Log "Cannot parse branch from VIB version: '$hostVibVersion'" "ERROR"
    exit $EXIT_ERROR
}

# ── Step 3: Read installed NVIDIA guest driver version ────────────────────────
Write-Log "Checking installed NVIDIA guest driver..."

$installedBranch = $null

try {
    # Check Display Adapters via registry (works without nvidia-smi)
    $nvidiaRegPaths = @(
        "HKLM:\SOFTWARE\NVIDIA Corporation\Global\NvCplApi\Policies",
        "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000"
    )

    # Prefer Win32_VideoController for simplicity
    $gpu = Get-CimInstance Win32_VideoController |
           Where-Object { $_.Name -match "NVIDIA" } |
           Select-Object -First 1

    if ($gpu) {
        # DriverVersion is e.g. "31.0.15.5009" on Windows — last segment is
        # the NVIDIA version (550.09 → 15.5009 on Windows driver model)
        # Extract the meaningful NVIDIA version from the Windows driver string
        $driverVer = $gpu.DriverVersion  # e.g. "31.0.15.5009"
        Write-Log "Installed Windows driver string: $driverVer"

        # The last two segments of the Windows driver version encode the NVIDIA
        # driver version. E.g. 15.5009 → 550.09, 15.4741 → 474.1
        # Extract first digit group to get the branch
        if ($driverVer -match '(\d+)\.(\d+)\.(\d+)\.(\d+)') {
            $thirdOctet  = [int]$Matches[3]   # e.g. 15
            $fourthOctet = [int]$Matches[4]   # e.g. 5009
            # Reconstruct NVIDIA-style version: concatenate 3rd+4th with leading digit stripped
            $nvidiaStyleVer = "$thirdOctet" + ($fourthOctet.ToString().PadLeft(4,'0'))
            # Branch is the first 3 digits
            $installedBranch = $nvidiaStyleVer.Substring(0, [Math]::Min(3, $nvidiaStyleVer.Length))
            Write-Log "Derived NVIDIA driver branch: $installedBranch (from $nvidiaStyleVer)"
        }
    } else {
        Write-Log "No NVIDIA GPU found via WMI. Driver may not be installed." "WARN"
    }
} catch {
    Write-Log "Error reading installed driver: $_" "WARN"
}

# ── Step 4: Compare branches ──────────────────────────────────────────────────
if ($installedBranch -and $installedBranch -eq $hostBranch) {
    Write-Log "Driver branch matches host vGPU Manager branch ($hostBranch). No action needed." "OK" # This is also used for SCCM detection — not detected = compliant
    if ($DetectOnly) {
        # SCCM: output something so the detection script finds the application
        Write-Output "NVIDIA_DRIVER_CURRENT:$installedBranch"
        exit $EXIT_OK
    }
    exit $EXIT_OK
}

Write-Log "Branch mismatch. Installed: '$installedBranch', Required: '$hostBranch'." "WARN"

if ($DetectOnly) {
    # In detection mode — exit non-zero signals SCCM the app is not installed/current
    Write-Log "DetectOnly mode — exiting to trigger SCCM installation." "WARN"
    exit $EXIT_NOT_CURRENT
}

# ── Step 5: Locate and run the installer ─────────────────────────────────────
$installerPath = Join-Path $DriverSourceRoot "$hostBranch\setup.exe"
Write-Log "Looking for installer at: $installerPath"

if (-not (Test-Path $installerPath)) {
    Write-Log "Installer not found: $installerPath" "ERROR"
    Write-Log "Ensure the driver package for branch $hostBranch is staged at $DriverSourceRoot\$hostBranch\setup.exe" "ERROR"
    exit $EXIT_ERROR
}

Write-Log "Starting silent install: $installerPath $InstallerArgs"
try {
    $proc = Start-Process -FilePath $installerPath `
                          -ArgumentList $InstallerArgs `
                          -Wait `
                          -PassThru `
                          -NoNewWindow

    Write-Log "Installer exited with code: $($proc.ExitCode)"

    switch ($proc.ExitCode) {
        0       { Write-Log "Installation succeeded. Reboot required." "OK"; exit $EXIT_REBOOT }
        1       { Write-Log "Installation succeeded (reboot flagged by installer)." "OK"; exit $EXIT_REBOOT }
        default {
            Write-Log "Installer returned unexpected exit code: $($proc.ExitCode)" "WARN"
            exit $proc.ExitCode
        }
    }
} catch {
    Write-Log "Failed to launch installer: $_" "ERROR"
    exit $EXIT_ERROR
}
