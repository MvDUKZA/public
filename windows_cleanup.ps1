# Run this script as Administrator

Write-Host "==== Starting Post-Upgrade Cleanup ===="

# 1. Delete C:\Windows.old
if (Test-Path "C:\Windows.old") {
    Write-Host "`nDeleting Windows.old..."
    takeown /F "C:\Windows.old" /R /D Y
    icacls "C:\Windows.old" /grant Administrators:F /T /C
    Remove-Item "C:\Windows.old" -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path "C:\Windows.old")) {
        Write-Host "Windows.old deleted."
    } else {
        Write-Host "Failed to delete Windows.old."
    }
} else {
    Write-Host "No Windows.old folder found."
}

# 2. Delete C:\$WINDOWS.~BT
if (Test-Path "C:\$WINDOWS.~BT") {
    Write-Host "`nDeleting $WINDOWS.~BT..."
    takeown /F "C:\$WINDOWS.~BT" /R /D Y
    icacls "C:\$WINDOWS.~BT" /grant Administrators:F /T /C
    Remove-Item "C:\$WINDOWS.~BT" -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path "C:\$WINDOWS.~BT")) {
        Write-Host "$WINDOWS.~BT deleted."
    } else {
        Write-Host "Failed to delete $WINDOWS.~BT."
    }
} else {
    Write-Host "No $WINDOWS.~BT folder found."
}

# 3. Clean C:\Windows\Temp
Write-Host "`nCleaning C:\Windows\Temp..."
Get-ChildItem -Path "C:\Windows\Temp" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
Write-Host "Windows Temp cleaned."

# 4. Clean user temp folder
Write-Host "`nCleaning User Temp Folder..."
$UserTemp = [System.IO.Path]::GetTempPath()
Get-ChildItem -Path $UserTemp -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
Write-Host "User Temp folder cleaned."

# 5. Clear Windows Update cache
Write-Host "`nClearing Windows Update Cache..."
net stop wuauserv
net stop bits
Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
net start wuauserv
net start bits
Write-Host "Windows Update cache cleared."

# 6. Component Store Cleanup
Write-Host "`nRunning DISM Cleanup..."
DISM /Online /Cleanup-Image /StartComponentCleanup /Quiet /NoRestart
Write-Host "Component store cleanup complete."

# 7. Done
Write-Host "`nPost-upgrade cleanup complete."
