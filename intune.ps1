# Variables
$deviceName = "DEVICE-NAME-HERE"
$userUPN = "user@domain.com"

# Ensure module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.DeviceManagement)) {
    Install-Module -Name Microsoft.Graph.DeviceManagement -Force -ErrorAction Stop
}
Import-Module Microsoft.Graph.DeviceManagement

# Connect to Graph with required scope
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All"

# Find the device by name
$device = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$deviceName'" -ErrorAction Stop

if (-not $device) {
    Write-Error "Device '$deviceName' not found in Intune."
    exit 1
}

$deviceId = $device.Id

# Assign the user to the device
$uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$deviceId/assignUserToDevice"

$body = @{
    userPrincipalName = $userUPN
} | ConvertTo-Json

$response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ErrorAction Stop

Write-Output "Successfully assigned '$userUPN' as primary user to device '$deviceName'."