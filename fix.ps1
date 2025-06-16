Start-Transcript "C:\Logs\SearchRepair.log"

# Uninstall problematic patch
wusa /uninstall /kb:5060842 /quiet /norestart

# Run system repairs
sfc /scannow
DISM /Online /Cleanup-Image /RestoreHealth

# Restart shell services
Stop-Process -Name explorer -Force
Stop-Process -Name SearchHost -Force
Stop-Process -Name StartMenuExperienceHost -Force
Start-Process explorer.exe

# Reset Search Index
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Search" -Name "SetupCompletedSuccessfully" -Value 0
Restart-Service -Name "WSearch"

# Re-register apps
Get-AppxPackage -AllUsers | ForEach-Object {
    Try {
        Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction Stop
    } Catch {
        Write-Output "Failed to register $($_.Name)"
    }
}

Stop-Transcript
