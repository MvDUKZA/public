# Write the required registry key directly to enable cleanup of Previous Windows Installations
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations" `
                 -Name "StateFlags0001" -Value 2 -Type DWord

# Now run Disk Cleanup with sagerun preset 1
cleanmgr /sagerun:1
