# === CONFIGURATION ===
$logPath = "C:\Logs\Kill-StartMenuSearch.log"
$timeoutSeconds = 60
$intervalSeconds = 3

# === LOGGING FUNCTION ===
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp`t$Message" | Out-File -Append -FilePath $logPath -Encoding utf8
}

# === WAIT LOOP ===
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
Write-Log "Startup script launched. Waiting for all target processes."

while ($stopwatch.Elapsed.TotalSeconds -lt $timeoutSeconds) {
    $searchHost = Get-Process -Name "SearchHost" -ErrorAction SilentlyContinue
    $startMenu = Get-Process -Name "StartMenuExperienceHost" -ErrorAction SilentlyContinue
    $explorer   = Get-Process -Name "explorer" -ErrorAction SilentlyContinue

    if ($searchHost -and $startMenu -and $explorer) {
        Write-Log "All target processes detected. Proceeding to terminate."
        Try {
            Stop-Process -Name "SearchHost" -Force -ErrorAction Stop
            Write-Log "SearchHost.exe terminated."
        } Catch {
            Write-Log "Error stopping SearchHost: $($_.Exception.Message)"
        }

        Try {
            Stop-Process -Name "StartMenuExperienceHost" -Force -ErrorAction Stop
            Write-Log "StartMenuExperienceHost.exe terminated."
        } Catch {
            Write-Log "Error stopping StartMenuExperienceHost: $($_.Exception.Message)"
        }

        break
    }

    Start-Sleep -Seconds $intervalSeconds
}

if ($stopwatch.Elapsed.TotalSeconds -ge $timeoutSeconds) {
    Write-Log "Timeout reached — processes did not all start in time."
}

$stopwatch.Stop()
