# Stop CCM
Stop-Service CcmExec -Force
Start-Sleep -Seconds 10

# Delete and re-register the CCM WMI provider
$ccmPath = "C:\Windows\CCM"

# Remove the stale WMI namespace
Get-WmiObject -Namespace root -Query "SELECT * FROM __Namespace WHERE Name='ccm'" | Remove-WmiObject

# Re-register CCM WMI providers
Get-ChildItem "$ccmPath\*.dll" | ForEach-Object {
    regsvr32.exe /s $_.FullName
}

# Force CCM to rebuild its WMI registration
& "$ccmPath\CCMRepair.exe"

Start-Sleep -Seconds 30
Start-Service CcmExec
