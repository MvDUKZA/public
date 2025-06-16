# ======== DEBUGGING VERSION ========
# Put this on a share and run it manually first:
#   .\KillStartHosts.ps1

$Log = "$env:USERPROFILE\Desktop\KillStartHosts_Debug.log"
"`n===== $(Get-Date) =====" | Out-File $Log -Append

function Wait-And-Kill {
    param(
        [string]$ProcName
    )

    "`$(Get-Date -Format 'HH:mm:ss') Waiting for $ProcName…" | Out-File $Log -Append
    while (-not (Get-Process -Name $ProcName -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds 1
    }

    "`$(Get-Date -Format 'HH:mm:ss') $ProcName found, killing…" | Out-File $Log -Append
    try {
        Stop-Process -Name $ProcName -Force -ErrorAction Stop
        "`$(Get-Date -Format 'HH:mm:ss') $ProcName stopped." | Out-File $Log -Append
    } catch {
        "`$(Get-Date -Format 'HH:mm:ss') ERROR stopping $ProcName: $($_.Exception.Message)" |
            Out-File $Log -Append
    }
}

# Wait for each host in turn
Wait-And-Kill -ProcName 'SearchHost'
Wait-And-Kill -ProcName 'StartMenuExperienceHost'

"`$(Get-Date -Format 'HH:mm:ss') Done." | Out-File $Log -Append
