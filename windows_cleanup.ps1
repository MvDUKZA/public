# Create the .reg file to enable Disk Cleanup for Previous Windows Installation(s)
$regContent = @"
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations]
"StateFlags0001"=dword:00000002
"@

$regPath = "$env:TEMP\enable_windows_old_cleanup.reg"

# Save the .reg file with proper encoding (UTF-16 LE)
$regContent | Out-File -FilePath $regPath -Encoding Unicode

# Import the registry file
reg import "$regPath"

# Run Disk Cleanup for preset #1
cleanmgr /sagerun:1
