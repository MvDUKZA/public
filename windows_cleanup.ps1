Start-Process cleanmgr -ArgumentList "/sagerun:1" -Wait

# Create the registry entries for /sagerun:1
$cleanMgrSettings = @"
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations]
"StateFlags0001"=dword:00000002
"@

$regFilePath = "$env:TEMP\CleanMgr.reg"
$cleanMgrSettings | Out-File -Encoding ASCII -FilePath $regFilePath

# Apply registry settings
reg import $regFilePath

# Run Disk Cleanup with those settings
cleanmgr /sagerun:1
