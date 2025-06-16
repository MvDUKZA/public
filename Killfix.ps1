# === CONFIGURATION ===
$logPath         = Join-Path $env:LOCALAPPDATA 'KillStartHosts.log'
$process1        = 'SearchHost'
$process2        = 'StartMenuExperienceHost'
$checkInterval   = 1    # seconds between process checks
$killDelay       = 3    # seconds to wait between kills

# === ENSURE LOG FOLDER EXISTS ===
$dir = Split-Path $logPath
if (-not (Test-Path $dir)) {
    try { New-Item -Path $dir -ItemType Directory -Force | Out-Null } catch { Exit 1 }
}

# === LOGGING FUNCTION ===
function Write-Log {
    param([string]$msg)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts`t$msg" | Out-File -FilePath $logPath -Append -Encoding utf8
}

# === WAIT & KILL FUNCTION ===
function Wait-And-Kill {
    param([string]$name)
    Write-Log "Waiting for process '$name'..."
    while (-not (Get-Process -Name $name -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds $checkInterval
    }
    Write-Log "Detected '$name'; attempting to stop."
    try {
        Stop-Process -Name $name -Force -ErrorAction Stop
        Write-Log "Successfully stopped '$name'."
    } catch {
        Write-Log "ERROR stopping '$name': $($_.Exception.Message)"
    }
}

# === MAIN ===
Write-Log "=== Script start ==="
Wait-And-Kill -name $process1
Start-Sleep -Seconds $killDelay
Wait-And-Kill -name $process2
Write-Log "=== Script complete ==="
