<# 
.SYNOPSIS
  Sets HKLM:\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability\ScanOnDemand (DWORD) to a desired value.
  - Runs as SYSTEM (self-elevates via a temporary Scheduled Task if needed)
  - Forces 64-bit registry write
  - Logs and verifies

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File .\Set-QualysScanOnDemand.ps1
  # Sets value to 1 (default)

.PARAMETER Value
  DWORD value to set. Default: 1

.PARAMETER LogPath
  Log file path. Default: C:\ProgramData\Qualys\Scripts\Set-QualysScanOnDemand.log
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [ValidateRange(0,2147483647)]
  [int]$Value = 1,

  [Parameter(Mandatory=$false)]
  [string]$LogPath = "C:\ProgramData\Qualys\Scripts\Set-QualysScanOnDemand.log",

  # Internal switch used when the scheduled task relaunches as SYSTEM
  [switch]$__RunAsSystemAlready
)

# --- Config ---
$KeyRelativePath = "SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability"
$ValueName       = "ScanOnDemand"
$TaskName        = "\_Temp_SetQualysScanOnDemand"
$Sentinel        = "$env:ProgramData\Qualys\Scripts\SetQualysScanOnDemand.done"
$EnsureDirs      = @([IO.Path]::GetDirectoryName($LogPath), [IO.Path]::GetDirectoryName($Sentinel))

# --- Logging helper ---
function Write-Log {
  param([string]$Message, [string]$Level = "INFO")
  $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[$timestamp] [$Level] $Message"
  Write-Host $line
  try {
    $null = New-Item -ItemType Directory -Path (Split-Path -Parent $LogPath) -Force -ErrorAction SilentlyContinue
    Add-Content -LiteralPath $LogPath -Value $line -ErrorAction SilentlyContinue
  } catch { }
}

# Ensure folders
foreach ($d in $EnsureDirs) { if ($d -and -not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

# --- Detect SYSTEM ---
function Test-IsSystem { ([Security.Principal.WindowsIdentity]::GetCurrent().Name -eq "NT AUTHORITY\SYSTEM") }

# --- Self-elevate via Scheduled Task (SYSTEM) if needed ---
if (-not (Test-IsSystem) -and -not $__RunAsSystemAlready) {
  try {
    Write-Log "Not running as SYSTEM. Creating temporary scheduled task to relaunch as SYSTEM..."
    # Clean any previous task
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }

    $action = New-ScheduledTaskAction -Execute "powershell.exe" `
      -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Value $Value -LogPath `"$LogPath`" -__RunAsSystemAlready"

    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddSeconds(5))
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
    Register-ScheduledTask -TaskName $TaskName -InputObject $task | Out-Null

    Start-ScheduledTask -TaskName $TaskName
    Write-Log "SYSTEM task started. Waiting up to 120 seconds for completion signal..."

    $waited = 0
    while ($waited -lt 120) {
      if (Test-Path $Sentinel) { Write-Log "SYSTEM run signaled completion."; break }
      Start-Sleep -Seconds 2
      $waited += 2
    }

    # Clean up task
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }

    if (-not (Test-Path $Sentinel)) {
      Write-Log "Timed out waiting for SYSTEM run to complete." "ERROR"
      exit 2
    } else {
      Remove-Item $Sentinel -Force -ErrorAction SilentlyContinue
      Write-Log "All done via SYSTEM task."
      exit 0
    }
  } catch {
    Write-Log "Failed to elevate via Scheduled Task: $($_.Exception.Message)" "ERROR"
    exit 3
  }
}

# --- Main work (runs here directly if already SYSTEM, or inside the scheduled task) ---
try {
  Write-Log "Running as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
  Write-Log "Forcing 64-bit registry view. Target: HKLM\$KeyRelativePath\$ValueName = $Value (DWORD)"

  # Use .NET Registry APIs to force Registry64 and create path if missing
  $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
  $subKey = $base.CreateSubKey($KeyRelativePath, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree)
  if (-not $subKey) { throw "Unable to open or create HKLM\$KeyRelativePath" }

  $subKey.SetValue($ValueName, [int]$Value, [Microsoft.Win32.RegistryValueKind]::DWord)
  $subKey.Flush()
  $subKey.Dispose()
  $base.Dispose()

  # Verify
  $base2 = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
  $check = $base2.OpenSubKey($KeyRelativePath)
  $actual = $check.GetValue($ValueName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
  $check.Dispose()
  $base2.Dispose()

  if ($null -eq $actual) { throw "Verification failed: value not found." }
  if ([int]$actual -ne [int]$Value) { throw "Verification mismatch: expected $Value, got $actual." }

  Write-Log "Success. Verified HKLM\$KeyRelativePath\$ValueName = $actual (DWORD)."
  
  # Drop sentinel if we were the SYSTEM task instance
  if ($__RunAsSystemAlready) { New-Item -ItemType File -Path $Sentinel -Force | Out-Null }

  exit 0
}
catch {
  Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
  if ($__RunAsSystemAlready) {
    # Signal failure but still drop a sentinel so the launcher stops waiting
    New-Item -ItemType File -Path $Sentinel -Force | Out-Null
  }
  exit 1
}


# Define the registry path and value
$RegPath = "HKLM:\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability"
$Name    = "ScanOnDemand"
$Value   = 1

# Create the key if it doesn’t exist
If (-not (Test-Path $RegPath)) {
    New-Item -Path $RegPath -Force | Out-Null
}

# Set the DWORD value
New-ItemProperty -Path $RegPath -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null


# Force x64 registry write regardless of process bitness
$Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
$View = [Microsoft.Win32.RegistryView]::Registry64

$Base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
$SubPath = "SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability"

# Create/open the key
$Key = $Base.CreateSubKey($SubPath, [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree)

# Set DWORD value
$Key.SetValue("ScanOnDemand", 1, [Microsoft.Win32.RegistryValueKind]::DWord)

# (Optional) verify
$actual = $Key.GetValue("ScanOnDemand")
if ($actual -ne 1) { throw "Expected 1 but found $actual at HKLM\$SubPath\ScanOnDemand (x64 view)" }

# Cleanup
$Key.Close()
$Base.Close()


