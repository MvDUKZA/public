# ----------------------------------------
# Configuration
# ----------------------------------------
$CsvFilePath        = "C:\Scripts\user_device_mapping.csv"
$LogFilePath        = "C:\Scripts\avd_intune_assignment.log"
$RequiredGraphScopes= @(
    "DeviceManagementManagedDevices.ReadWrite.All",
    "User.Read.All"
)

# ----------------------------------------
# Logging (unchanged from yours)
# ----------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts [$Level] $Message" | Tee-Object -FilePath $LogFilePath -Append
}

# ----------------------------------------
# Ensure a module is imported (install if needed)
# ----------------------------------------
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "Installing module $Name" 'INFO'
        try { Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop }
        catch { Write-Log "Failed to install $Name: $_" 'ERROR'; throw }
    }
    Import-Module -Name $Name -ErrorAction Stop
    Write-Log "Module loaded: $Name" 'INFO'
}

# ----------------------------------------
# Test that a user has an Intune license
# ----------------------------------------
function Test-UserHasIntuneLicense {
    param([string]$UPN)
    # Intune Service Plan GUID
    $intunePlanId = '8e9ff0ff-aa7a-4b20-83ad-eba658c935bf'
    try {
        $details = Get-MgUserLicenseDetail -UserId $UPN -ErrorAction Stop
        return $details.ServicePlans |
               Where-Object { $_.ServicePlanId -eq $intunePlanId -and $_.ProvisioningStatus -eq 'Success' }
    }
    catch {
        Write-Log "License lookup failed for $UPN: $_" 'ERROR'
        return $false
    }
}

# ----------------------------------------
# Validate user exists, enabled, is member & licensed
# ----------------------------------------
function Test-UserEligibility {
    param([string]$UPN)
    try {
        $u = Get-MgUser -UserId $UPN -Property AccountEnabled,UserType -ErrorAction Stop
        if (-not $u.AccountEnabled) {
            Write-Log "User disabled: $UPN" 'WARN'; return $false
        }
        if ($u.UserType -ne 'Member') {
            Write-Log "User is not Member: $UPN" 'WARN'; return $false
        }
        if (-not (Test-UserHasIntuneLicense $UPN)) {
            Write-Log "User lacks Intune license: $UPN" 'WARN'; return $false
        }
        return $true
    }
    catch {
        Write-Log "User lookup error for $UPN: $_" 'ERROR'
        return $false
    }
}

# ----------------------------------------
# Grab the freshest Windows-managed device record
# ----------------------------------------
function Get-IntuneDevice {
    param([string]$Name)
    try {
        $filter = "deviceName eq '$Name' and operatingSystem eq 'Windows'"
        $matches = Get-MgDeviceManagementManagedDevice -Filter $filter -ErrorAction Stop
        $dev = $matches |
               Where-Object { $_.ManagementState -in @('mdm','coManaged') } |
               Sort-Object LastSyncDateTime -Descending |
               Select-Object -First 1
        if (-not $dev) {
            Write-Log "No Intune device found matching $Name" 'WARN'
            return $null
        }
        return $dev
    }
    catch {
        Write-Log "Error querying Intune device $Name: $_" 'ERROR'
        return $null
    }
}

# ----------------------------------------
# Your original Set-PrimaryUser, fixed for PSCore escaping
# ----------------------------------------
function Set-PrimaryUser {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][string]$UserId
    )
    try {
        $uri  = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$DeviceId')/users/`$ref"
        $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/users/$UserId" } |
                ConvertTo-Json
        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "Intune primary user set: DeviceId=$DeviceId User=$UserId" 'INFO'
        return $true
    }
    catch {
        Write-Log "Failed Set-PrimaryUser for $DeviceId -> $UserId: $_" 'ERROR'
        return $false
    }
}

# ----------------------------------------
# Assign AVD session host to a user
# ----------------------------------------
function Set-AVDPrimaryUser {
    param(
        [string]$ResourceGroupName,
        [string]$HostPoolName,
        [string]$SessionHostName,
        [string]$UserUPN
    )
    try {
        $host = Get-AzWvdSessionHost `
                    -ResourceGroupName $ResourceGroupName `
                    -HostPoolName $HostPoolName `
                    -ErrorAction Stop |
                Where-Object { $_.Name -ieq $SessionHostName }
        if (-not $host) {
            throw "Host not found: $SessionHostName"
        }
        if ($host.AssignedUser -and $host.AssignedUser -ne $UserUPN) {
            Write-Log "AVD host already assigned to $($host.AssignedUser); skipping" 'WARN'
            return
        }
        Update-AzWvdSessionHost `
            -ResourceGroupName $ResourceGroupName `
            -HostPoolName $HostPoolName `
            -Name $SessionHostName `
            -AssignedUser $UserUPN -ErrorAction Stop
        Write-Log "AVD host assigned: $SessionHostName -> $UserUPN" 'INFO'
    }
    catch {
        Write-Log "AVD assignment error for $SessionHostName: $_" 'ERROR'
    }
}

# ----------------------------------------
# Process one CSV record
# ----------------------------------------
function Process-Record {
    param($rec)

    $upn       = $rec.UserUPN.Trim()
    $vmName    = $rec.VDIName.Trim()
    $pool      = $rec.HostPoolName.Trim()
    $rg        = $rec.ResourceGroupName.Trim()

    Write-Log "=== Processing $upn -> $vmName ===" 'INFO'

    if (-not (Test-UserEligibility $upn)) { return }

    Set-AVDPrimaryUser -ResourceGroupName $rg `
                       -HostPoolName $pool `
                       -SessionHostName $vmName `
                       -UserUPN $upn

    $intuneDev = Get-IntuneDevice -Name $vmName
    if ($intuneDev) {
        # skip if already assigned to someone else
        if ($intuneDev.UserPrincipalName -and $intuneDev.UserPrincipalName -ne $upn) {
            Write-Log "Intune device $vmName already owned by $($intuneDev.UserPrincipalName); skipping" 'WARN'
        }
        else {
            Set-PrimaryUser -DeviceId $intuneDev.Id -UserId $upn
        }
    }
}

# ----------------------------------------
# Main
# ----------------------------------------
try {
    # modules
    'Az.Accounts','Az.DesktopVirtualization','Microsoft.Graph' ,
    'Microsoft.Graph.DeviceManagement','Microsoft.Graph.Users' |
    ForEach-Object { Ensure-Module $_ }

    # auth
    Connect-AzAccount -ErrorAction Stop
    Connect-MgGraph -Scopes $RequiredGraphScopes -ErrorAction Stop
    foreach ($s in $RequiredGraphScopes) {
        if ((Get-MgContext).Scopes -notcontains $s) {
            throw "Missing Graph scope: $s"
        }
    }

    # import CSV
    Write-Log "Importing CSV $CsvFilePath" 'INFO'
    $records = Import-Csv -Path $CsvFilePath -ErrorAction Stop
    $cols    = ($records | Get-Member -MemberType NoteProperty).Name
    $reqCols = 'UserUPN','VDIName','HostPoolName','ResourceGroupName'
    if (Compare-Object $reqCols $cols) {
        throw "CSV schema mismatch: need $($reqCols -join ', ')"
    }

    # iterate
    foreach ($r in $records) { Process-Record $r }

    Write-Log "All done." 'INFO'
}
catch {
    Write-Log "Fatal error: $_" 'ERROR'
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
