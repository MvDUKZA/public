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
# Logging
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
# Note: Ensure-Module is a custom function, not built-in. Consider renaming to EnsureRequiredModule for clarity if desired.
function Ensure-Module {
    param(
        [string]$Name,
        [string]$MinimumVersion = $null
    )
    $existingModule = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($existingModule -and ($MinimumVersion -eq $null -or $existingModule.Version -ge [Version]$MinimumVersion)) {
        Write-Log "Module $Name (Version $($existingModule.Version)) already available" 'INFO'
    }
    else {
        Write-Log "Installing module $Name (MinimumVersion: $MinimumVersion)" 'INFO'
        try {
            $installParams = @{
                Name = $Name
                Scope = 'CurrentUser'
                ErrorAction = 'Stop'
            }
            if ($MinimumVersion) {
                $installParams.MinimumVersion = $MinimumVersion
            }
            Install-Module @installParams
        }
        catch {
            Write-Log "Failed to install $Name: $_" 'ERROR'
            throw
        }
    }
    try {
        Import-Module -Name $Name -ErrorAction Stop
        Write-Log "Module $Name loaded successfully" 'INFO'
    }
    catch {
        Write-Log "Failed to import $Name: $_" 'ERROR'
        throw
    }
}

# ----------------------------------------
# Test that a user has an Intune license
# ----------------------------------------
function Test-UserHasIntuneLicense {
    param([string]$UPN)
    # Intune Service Plan GUID (Microsoft Intune service plan, part of EMS, M365 E3/E5, or standalone Intune licenses)
    # Source: Standard Microsoft 365 licensing data as of 2025
    $intunePlanId = '8e9ff0ff-aa7a-4b20-83ad-eba658c935bf'
    
    try {
        $details = Get-MgUserLicenseDetail -UserId $UPN -ErrorAction Stop
        $hasIntune = $details.ServicePlans |
                     Where-Object { $_.ServicePlanId -eq $intunePlanId -and $_.ProvisioningStatus -eq 'Success' }
        
        if ($hasIntune) {
            return $true
        }
        
        # Fallback: Dynamically check for Intune service plan if hardcoded GUID doesn't match
        # Note: Requires Directory.Read.All scope, which is not included in $RequiredGraphScopes
        Write-Log "Hardcoded Intune GUID ($intunePlanId) not found for $UPN; attempting dynamic lookup" 'WARN'
        try {
            $skus = Get-MgSubscribedSku -ErrorAction Stop
            $intuneServicePlan = $skus.ServicePlans |
                                 Where-Object { $_.ServicePlanName -like '*INTUNE*' -and $_.ProvisioningStatus -eq 'Success' } |
                                 Select-Object -First 1
        }
        catch {
            Write-Log "Dynamic lookup failed (missing Directory.Read.All scope?): $_" 'ERROR'
            return $false
        }
        
        if ($intuneServicePlan) {
            Write-Log "Dynamic lookup found Intune service plan ID: $($intuneServicePlan.ServicePlanId)" 'INFO'
            $hasIntune = $details.ServicePlans |
                         Where-Object { $_.ServicePlanId -eq $intuneServicePlan.ServicePlanId -and $_.ProvisioningStatus -eq 'Success' }
            return [bool]$hasIntune
        }
        
        Write-Log "No Intune license found for $UPN" 'WARN'
        return $false
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
            Write-Log "User $UPN is disabled" 'WARN'
            return $false
        }
        if ($u.UserType -ne 'Member') {
            Write-Log "User $UPN is not a Member account" 'WARN'
            return $false
        }
        if (-not (Test-UserHasIntuneLicense $UPN)) {
            Write-Log "User $UPN lacks an Intune license" 'WARN'
            return $false
        }
        return $true
    }
    catch {
        Write-Log "Error looking up user $UPN: $_" 'ERROR'
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
            Write-Log "No Intune device found matching name $Name" 'WARN'
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
# Set primary user for Intune device
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
        Write-Log "Set Intune primary user for DeviceId $DeviceId to UserId $UserId" 'INFO'
        return $true
    }
    catch {
        Write-Log "Failed to set primary user for DeviceId $DeviceId to UserId $UserId: $_" 'ERROR'
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
        $sessionHost = Get-AzWvdSessionHost `
                       -ResourceGroupName $ResourceGroupName `
                       -HostPoolName $HostPoolName `
                       -ErrorAction Stop |
                       Where-Object { $_.Name -ieq $SessionHostName }
        if (-not $sessionHost) {
            throw "Session host $SessionHostName not found"
        }
        if ($sessionHost.AssignedUser -and $sessionHost.AssignedUser -ne $UserUPN) {
            Write-Log "AVD session host $SessionHostName already assigned to user $($sessionHost.AssignedUser); skipping" 'WARN'
            return
        }
        Update-AzWvdSessionHost `
            -ResourceGroupName $ResourceGroupName `
            -HostPoolName $HostPoolName `
            -Name $SessionHostName `
            -AssignedUser $UserUPN -ErrorAction Stop
        Write-Log "Assigned AVD session host $SessionHostName to user $UserUPN" 'INFO'
    }
    catch {
        Write-Log "Error assigning AVD session host $SessionHostName: $_" 'ERROR'
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

    Write-Log "Processing user $upn to device $vmName" 'INFO'

    if (-not (Test-UserEligibility $upn)) { return }

    Set-AVDPrimaryUser -ResourceGroupName $rg `
                       -HostPoolName $pool `
                       -SessionHostName $vmName `
                       -UserUPN $upn

    $intuneDev = Get-IntuneDevice -Name $vmName
    if ($intuneDev) {
        if ($intuneDev.UserPrincipalName -and $intuneDev.UserPrincipalName -ne $upn) {
            Write-Log "Intune device $vmName already assigned to user $($intuneDev.UserPrincipalName); skipping" 'WARN'
        }
        else {
            try {
                $user = Get-MgUser -UserId $upn -ErrorAction Stop
                if (-not $user.Id) {
                    Write-Log "Failed to retrieve user ID for $upn" 'ERROR'
                    return
                }
                Set-PrimaryUser -DeviceId $intuneDev.Id -UserId $user.Id
            }
            catch {
                Write-Log "Failed to resolve user ID for $upn: $_" 'ERROR'
            }
        }
    }
}

# ----------------------------------------
# Main
# ----------------------------------------
try {
    # modules with minimum versions
    $modules = @(
        @{ Name = 'Az.Accounts'; MinimumVersion = '2.0.0' },
        @{ Name = 'Az.DesktopVirtualization'; MinimumVersion = '0.1.0' },
        @{ Name = 'Microsoft.Graph'; MinimumVersion = '1.0.0' },
        @{ Name = 'Microsoft.Graph.DeviceManagement'; MinimumVersion = '1.0.0' },
        @{ Name = 'Microsoft.Graph.Users'; MinimumVersion = '1.0.0' }
    )
    foreach ($mod in $modules) {
        Ensure-Module -Name $mod.Name -MinimumVersion $mod.MinimumVersion
    }

    # auth
    Connect-AzAccount -ErrorAction Stop
    Connect-MgGraph -Scopes $RequiredGraphScopes -ErrorAction Stop
    foreach ($s in $RequiredGraphScopes) {
        if ((Get-MgContext).Scopes -notcontains $s) {
            throw "Missing required Graph scope: $s"
        }
    }

    # import CSV
    Write-Log "Importing CSV file $CsvFilePath" 'INFO'
    $records = Import-Csv -Path $CsvFilePath -ErrorAction Stop
    $cols    = ($records | Get-Member -MemberType NoteProperty).Name
    $reqCols = 'UserUPN','VDIName','HostPoolName','ResourceGroupName'
    if (Compare-Object $reqCols $cols) {
        throw "CSV schema mismatch: required columns are $($reqCols -join ', ')"
    }

    # iterate
    foreach ($r in $records) { Process-Record $r }

    Write-Log "Processing completed successfully" 'INFO'
}
catch {
    Write-Log "Fatal error occurred: $_" 'ERROR'
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
