# ——— CONFIGURE THESE VARIABLES ———
$OutputCount    = 50                           # How many deleted devices to retrieve at once
$RestoreObjectId = ''                          # If you want to restore a device, put its ObjectId here

# ——— LOGGING FUNCTION ———
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-Host "$ts [$Level] $Message"
}

# ——— 1. Install & Import Microsoft.Graph ———
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Log "Microsoft.Graph module not found; installing…" 'INFO'
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
        Write-Log "Microsoft.Graph installed." 'INFO'
    } catch {
        Write-Log "Failed to install Microsoft.Graph: $_" 'ERROR'
        exit 1
    }
}
Import-Module Microsoft.Graph -ErrorAction Stop

# ——— 2. Connect to MS Graph ———
try {
    Write-Log "Connecting to Microsoft Graph…" 'INFO'
    Connect-MgGraph -Scopes "Directory.Read.All","Directory.AccessAsUser.All" -ErrorAction Stop
    Write-Log "Connected as $(Get-MgUserMe).UserPrincipalName" 'INFO'
} catch {
    Write-Log "Graph connection failed: $_" 'ERROR'
    exit 1
}

# ——— 3. List Deleted Devices ———
try {
    Write-Log "Retrieving deleted device objects…" 'INFO'
    $deletedDevices = Get-MgDirectoryDeletedItem -DirectoryObjectType "microsoft.graph.device" -All:$true `
                       | Select-Object DisplayName, DeviceId, Id, DeletionTimestamp
    if ($deletedDevices.Count -gt 0) {
        $deletedDevices | Format-Table -AutoSize
        Write-Log "Found $($deletedDevices.Count) deleted devices." 'INFO'
    } else {
        Write-Log "No deleted devices found." 'WARN'
    }
} catch {
    Write-Log "Error retrieving deleted devices: $_" 'ERROR'
}

# ——— 4. (Optional) Restore a Deleted Device ———
if ($RestoreObjectId) {
    try {
        Write-Log "Restoring device with ObjectId $RestoreObjectId…" 'INFO'
        Restore-MgDirectoryDeletedItem -DirectoryObjectId $RestoreObjectId -ErrorAction Stop
        Write-Log "Device restored successfully." 'INFO'
    } catch {
        Write-Log "Failed to restore device: $_" 'ERROR'
    }
}
