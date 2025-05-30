<#
.SYNOPSIS
    Assign Horizon users to virtual machines from a CSV using VMware Horizon 8.12.1 PowerCLI.
.DESCRIPTION
    Reads a CSV (HorizonServer, UserUPN, MachineName), connects to each Horizon server, and assigns users to machines.
.PARAMETER AssignmentListPath
    Path to the CSV file containing assignments.
.PARAMETER LogFile
    Path to the log file (default: timestamped file under C:\temp\scripts\logs).
.NOTES
    Author: Marinus van Deventer
    Version: 1.0.1
    Requires: VMware.VimAutomation.HorizonView, VMware.Hv.Helper
    WorkingDirectory: C:\temp\scripts
    LogDirectory: C:\temp\scripts\logs
    Date: 2025-05-30
#>

#region Parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$AssignmentListPath = "C:\temp\scripts\Assignments.csv",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogFile
)

if (-not $LogFile) {
    $LogDirectory = "C:\temp\scripts\logs"
    if (-not (Test-Path $LogDirectory)) { New-Item -Path $LogDirectory -ItemType Directory | Out-Null }
    $LogFile = Join-Path $LogDirectory "HorizonAssignment_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
}
#endregion

#region Logging Function
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]$Message
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$($script:ScriptVersion)] $timestamp - $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Host $entry
}
#endregion

#region Script Version
$script:ScriptVersion = '1.0.1'
#endregion

#region Module Check and Import
$requiredModules = 'VMware.VimAutomation.HorizonView','VMware.Hv.Helper'
foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Log "ERROR: Required module '$mod' not available. Install via PowerCLI or copy helper module."
        throw "Missing module $mod"
    }
    Import-Module $mod -ErrorAction Stop
    Write-Log "Imported module: $mod"
}
#endregion

#region CSV Import
if (-not (Test-Path $AssignmentListPath)) {
    Write-Log "ERROR: CSV file not found at $AssignmentListPath"
    throw "CSV path invalid"
}
$assignments = Import-Csv -Path $AssignmentListPath
if (-not $assignments) {
    Write-Log "ERROR: No assignments found in CSV"
    throw "CSV contains no data"
}
Write-Log "Loaded $($assignments.Count) assignments"
#endregion

#region Credentials
$cred = Get-Credential -Message 'Enter Horizon Admin credentials'
Write-Log 'Credentials obtained'
#endregion

#region Helper Function: Get-UserId
function Get-UserId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][VMware.Hv.Services]$Services,
        [Parameter(Mandatory)][string]$UserName
    )
    $def = [VMware.Hv.QueryDefinition]::new()
    $def.queryEntityType = 'ADUserOrGroupSummaryView'
    $filterUser = [VMware.Hv.QueryFilterEquals]::new()
    $filterUser.PropertyName = 'base.name'
    $filterUser.Value = $UserName
    $filterGroup = [VMware.Hv.QueryFilterEquals]::new()
    $filterGroup.PropertyName = 'base.group'
    $filterGroup.Value = $false
    $andFilter = [VMware.Hv.QueryFilterAnd]::new()
    $andFilter.Filters = @($filterUser, $filterGroup)
    $def.Filter = $andFilter
    $result = $Services.QueryService.Query($def)
    if (-not $result.Results) { throw "User '$UserName' not found" }
    return $result.Results[0]
}
#endregion

#region Process Assignments
foreach ($entry in $assignments) {
    $server    = $entry.HorizonServer
    $userUPN   = $entry.UserUPN
    $vmName    = $entry.MachineName

    if (-not ($server -and $userUPN -and $vmName)) {
        Write-Log "WARNING: Incomplete entry, skipping: $($entry | Out-String)"
        continue
    }
    Write-Log "Starting assignment: $userUPN -> $vmName on $server"

    try {
        $hvConn = Connect-HVServer -Server $server -Credential $cred -ErrorAction Stop
        Write-Log "Connected to $server"
    } catch {
        Write-Log "ERROR: Connection to $server failed: $_"
        continue
    }

    try {
        # Find VM via client filter
        $vm = Get-HVMachine -Server $hvConn | Where-Object { $_.Base.Name -eq $vmName }
        if (-not $vm) { throw "VM '$vmName' not found" }

        # Find User with helper function
        $userObj = Get-UserId -Services $hvConn.ExtensionData -UserName $userUPN

        # Perform assignment
        $assignSpec = [VMware.Hv.MachineAssignmentSpec]::new()
        $assignSpec.Id   = $vm.Id
        $assignSpec.User = $userUPN
        $hvConn.ExtensionData.Machine.AssignUser($assignSpec)
        Write-Log "SUCCESS: Assigned $userUPN to $vmName"
    } catch {
        Write-Log "ERROR: Assignment error: $_"
    } finally {
        Disconnect-HVServer -Confirm:$false
        Write-Log "Disconnected from $server"
    }
}
#endregion

Write-Log "Script completed successfully at $(Get-Date)"
