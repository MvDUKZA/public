# === CONFIGURATION VARIABLES ===
$processesToWatch = @('explorer', 'SearchHost', 'StartMenuExperienceHost')
$processesToKill  = @('SearchHost', 'StartMenuExperienceHost')
$timeoutSeconds   = 60
$intervalSeconds  = 2
$logPath          = Join-Path $env:LOCALAPPDATA 'KillStartHosts.log'

# === ENSURE LOG FOLDER EXISTS ===
$logDir = Split-Path $logPath
if (-not (Test-Path $logDir)) {
    Try {
        New-Item -ItemType Directory -Path $logDir -ErrorAction Stop | Out-Null
    } Catch {
        # If we can’t even create the folder, bail out
        Exit 1
    }
}

# === LOGGING FUNCTION ===
Function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts`t$Message" | Out-File -FilePath $logPath -Append -Encoding utf8
}

# === MAIN WAIT-AND-KILL LOGIC ===
$sw = [Diagnostics.Stopwatch]::StartNew()
Write-Log "Script started. Waiting for all target processes: $($processesToWatch -join ', ')"

while ($sw.Elapsed.TotalSeconds -lt $timeoutSeconds) {
    # check each required process
    $running = $processesToWatch | ForEach-Object {
        if (Get-Process -Name $_ -ErrorAction SilentlyContinue) { $true } else { $false }
    }

    if ($running -notcontains $false) {
        Write-Log "All target processes detected. Terminating hosts now."
        foreach ($name in $processesToKill) {
            Try {
                Stop-Process -Name $name -Force -ErrorAction Stop
                Write-Log "Stopped $name.exe"
            } Catch {
                Write-Log "Failed to stop $name.exe: $($_.Exception.Message)"
            }
        }
        break
    }

    Start-Sleep -Seconds $intervalSeconds
}

if ($sw.Elapsed.TotalSeconds -ge $timeoutSeconds) {
    Write-Log "Timeout ($timeoutSeconds s) reached before all processes appeared."
}
$sw.Stop()
Write-Log "Script completed."
