$ErrorActionPreference = 'Stop'

<#
.SYNOPSIS
Assigns Horizon Desktop machines to users based on a CSV file.

.DESCRIPTION
This script assigns users to Horizon desktop machines in dedicated desktop pools using a CSV file input.
The CSV file must contain columns: UserName, Domain, DesktopPool, MachineName.
It requires VMware PowerCLI and is compatible with Omnissa Horizon 8.12.1.

.NOTES
- Requires VMware PowerCLI installed: `Install-Module -Name VMware.PowerCLI -Force -AllowClobber -Scope AllUsers`
- The account running the script needs read-only administrator permissions in Horizon.
- CSV file path and Horizon Connection Server FQDN must be specified.

.LINK
https://code.vmware.com/web/tool/12.0.0/vmware-powercli
https://kb.vmware.com/s/article/2143853

.COMPONENT
VMware PowerCLI
#>

# Parameters
param (
    [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file containing user and machine assignments.")]
    [string]$CsvPath,
    [Parameter(Mandatory = $true, HelpMessage = "FQDN of the Horizon Connection Server.")]
    [string]$HVConnectionServerFQDN,
    [Parameter(Mandatory = $true, HelpMessage = "PSCredential object for Horizon authentication.")]
    [PSCredential]$Credential
)

# Function to output messages or errors
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$Warning,
        [Parameter(Mandatory = $false)]
        [switch]$Stop,
        [Parameter(Mandatory = $false)]
        $Exception
    )

    if ($Exception) {
        Write-Warning "$Message`n$($Exception.CategoryInfo.Category)"
        Write-Error "$Message`n$($Exception.Exception.Message)`n$($Exception.CategoryInfo)" -ErrorAction Stop
    }
    elseif ($Stop) {
        Write-Warning "Error: $Message"
        Throw $Message
    }
    elseif ($Warning) {
        Write-Warning $Message
    }
    else {
        Write-Output $Message
    }
}

# Function to load VMware PowerCLI modules
function Load-VMwareModules {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Components
    )

    foreach ($component in $Components) {
        try {
            Import-Module -Name VMware.$component -ErrorAction Stop
        }
        catch {
            try {
                Add-PSSnapin -Name VMware -ErrorAction Stop
            }
            catch {
                Write-Log -Message "Required VMware modules not found. Install VMware PowerCLI." -Stop
            }
        }
    }
}

# Function to connect to Horizon Connection Server
function Connect-HorizonConnectionServer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$HVConnectionServerFQDN,
        [Parameter(Mandatory = $true)]
        [PSCredential]$Credential
    )

    try {
        Connect-HVServer -Server $HVConnectionServerFQDN -Credential $Credential -ErrorAction Stop
    }
    catch {
        Write-Log -Message "Failed to connect to Horizon Connection Server: $HVConnectionServerFQDN" -Exception $_
    }
}

# Function to disconnect from Horizon Connection Server
function Disconnect-HorizonConnectionServer {
    param (
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        Disconnect-HVServer -Server $HVConnectionServer -Confirm:$false -ErrorAction Stop
    }
    catch {
        Write-Log -Message "Failed to disconnect from Horizon Connection Server." -Exception $_
    }
}

# Function to get Horizon Desktop Pool
function Get-HVDesktopPool {
    param (
        [Parameter(Mandatory = $true)]
        [string]$HVPoolName,
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        $queryService = New-Object VMware.Hv.QueryServiceService
        $defn = New-Object VMware.Hv.QueryDefinition
        $defn.queryEntityType = 'DesktopSummaryView'
        $defn.Filter = New-Object VMware.Hv.QueryFilterEquals -Property @{'memberName'='desktopSummaryData.displayName'; 'value' = $HVPoolName}
        [array]$queryResults = ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)

        if (!$queryResults) {
            Write-Log -Message "Desktop pool '$HVPoolName' not found." -Stop
        }
        return $queryResults
    }
    catch {
        Write-Log -Message "Error retrieving Horizon Desktop Pool '$HVPoolName'." -Exception $_
    }
}

# Function to get Horizon Desktop Machine
function Get-HVDesktopMachine {
    param (
        [Parameter(Mandatory = $true)]
        [VMware.Hv.DesktopId]$HVPoolID,
        [Parameter(Mandatory = $true)]
        [string]$HVMachineName,
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        $queryService = New-Object VMware.Hv.QueryServiceService
        $defn = New-Object VMware.Hv.QueryDefinition
        $defn.queryEntityType = 'MachineDetailsView'
        $poolFilter = New-Object VMware.Hv.QueryFilterEquals -Property @{'memberName'='desktopData.id'; 'value' = $HVPoolID}
        $machineFilter = New-Object VMware.Hv.QueryFilterEquals -Property @{'memberName'='data.name'; 'value' = $HVMachineName}
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = @($poolFilter, $machineFilter)
        $defn.Filter = $filterAnd
        [array]$queryResults = ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)

        if (!$queryResults) {
            Write-Log -Message "Machine '$HVMachineName' not found in pool." -Stop
        }
        return $queryResults
    }
    catch {
        Write-Log -Message "Error retrieving machine '$HVMachineName'." -Exception $_
    }
}

# Function to get Horizon Desktop Pool specification
function Get-HVPoolSpec {
    param (
        [Parameter(Mandatory = $true)]
        [VMware.Hv.DesktopId]$HVPoolID,
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        $HVConnectionServer.ExtensionData.Desktop.Desktop_Get($HVPoolID)
    }
    catch {
        Write-Log -Message "Error retrieving desktop pool details." -Exception $_
    }
}

# Function to assign a desktop to a user
function New-HVDesktopAssignment {
    param (
        [Parameter(Mandatory = $true)]
        [VMware.Hv.MachineId]$HVMachineID,
        [Parameter(Mandatory = $true)]
        [VMware.Hv.UserOrGroupId]$HVUserID,
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        $machineService = New-Object VMware.Hv.MachineService
        $machineInfoHelper = $machineService.read($HVConnectionServer.extensionData, $HVMachineID)
        $machineInfoHelper.getBaseHelper().setUser($HVUserID)
        $machineService.update($HVConnectionServer.extensionData, $machineInfoHelper)
        Write-Log -Message "Desktop assigned successfully."
    }
    catch {
        Write-Log -Message "Error assigning desktop." -Exception $_
    }
}

# Function to get Horizon User
function Get-HVUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$HVUserLoginName,
        [Parameter(Mandatory = $true)]
        [string]$HVDomain,
        [Parameter(Mandatory = $true)]
        [VMware.VimAutomation.HorizonView.Impl.V1.ViewObjectImpl]$HVConnectionServer
    )

    try {
        $queryService = New-Object VMware.Hv.QueryServiceService
        $defn = New-Object VMware.Hv.QueryDefinition
        $defn.queryEntityType = 'ADUserOrGroupSummaryView'
        $userFilter = New-Object VMware.Hv.QueryFilterEquals -Property @{'memberName'='base.loginName'; 'value' = $HVUserLoginName}
        $domainFilter = New-Object VMware.Hv.QueryFilterEquals -Property @{'memberName'='base.domain'; 'value' = $HVDomain}
        $filterAnd = New-Object VMware.Hv.QueryFilterAnd
        $filterAnd.Filters = @($userFilter, $domainFilter)
        $defn.Filter = $filterAnd
        [array]$queryResults = ($queryService.queryService_create($HVConnectionServer.extensionData, $defn)).results
        $queryService.QueryService_DeleteAll($HVConnectionServer.extensionData)

        if (!$queryResults) {
            Write-Log -Message "User '$HVUserLoginName' not found in domain '$HVDomain'." -Stop
        }
        return $queryResults
    }
    catch {
        Write-Log -Message "Error retrieving user '$HVUserLoginName'." -Exception $_
    }
}

# Main script logic
try {
    # Validate CSV file
    if (-not (Test-Path $CsvPath)) {
        Write-Log -Message "CSV file '$CsvPath' not found." -Stop
    }

    # Import CSV
    $assignments = Import-Csv -Path $CsvPath
    if (-not $assignments[0].PSObject.Properties.Match('UserName') -or
        -not $assignments[0].PSObject.Properties.Match('Domain') -or
        -not $assignments[0].PSObject.Properties.Match('DesktopPool') -or
        -not $assignments[0].PSObject.Properties.Match('MachineName')) {
        Write-Log -Message "CSV file must contain columns: UserName, Domain, DesktopPool, MachineName." -Stop
    }

    # Load VMware PowerCLI modules
    Load-VMwareModules -Components @('VimAutomation.HorizonView')

    # Connect to Horizon Connection Server
    $objHVConnectionServer = Connect-HorizonConnectionServer -HVConnectionServerFQDN $HVConnectionServerFQDN -Credential $Credential

    # Process each assignment
    foreach ($assignment in $assignments) {
        $HVUserName = $assignment.UserName
        $HVDomain = $assignment.Domain
        $HVDesktopPoolName = $assignment.DesktopPool
        $HVMachineName = $assignment.MachineName

        Write-Log -Message "Processing assignment for user '$HVUserName' to machine '$HVMachineName' in pool '$HVDesktopPoolName'."

        # Retrieve Desktop Pool
        $HVPool = Get-HVDesktopPool -HVPoolName $HVDesktopPoolName -HVConnectionServer $objHVConnectionServer
        $HVPoolID = $HVPool.id

        # Retrieve Pool Specification
        $HVPoolSpec = Get-HVPoolSpec -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID

        # Check if pool is dedicated
        if ($HVPoolSpec.type -eq "AUTOMATED" -and $HVPoolSpec.AutomatedDesktopData.userAssignment.userAssignment -ne "DEDICATED") {
            Write-Log -Message "Pool '$HVDesktopPoolName' is not a dedicated desktop pool." -Warning
            continue
        }
        elseif ($HVPoolSpec.type -eq "MANUAL" -and $HVPoolSpec.ManualDesktopData.userAssignment.userAssignment -ne "DEDICATED") {
            Write-Log -Message "Pool '$HVDesktopPoolName' is not a dedicated desktop pool." -Warning
            continue
        }

        # Retrieve Machine
        $HVMachine = Get-HVDesktopMachine -HVConnectionServer $objHVConnectionServer -HVPoolID $HVPoolID -HVMachineName $HVMachineName
        $HVMachineID = $HVMachine.id

        # Retrieve User
        $HVUser = Get-HVUser -HVConnectionServer $objHVConnectionServer -HVUserLoginName $HVUserName -HVDomain $HVDomain
        $HVUserID = $HVUser.id

        # Assign Desktop
        New-HVDesktopAssignment -HVConnectionServer $objHVConnectionServer -HVMachineID $HVMachineID -HVUserID $HVUserID
    }
}
finally {
    # Disconnect from Horizon Connection Server
    if ($objHVConnectionServer) {
        Disconnect-HorizonConnectionServer -HVConnectionServer $objHVConnectionServer
    }
}
