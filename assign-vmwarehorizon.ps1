<#
.SYNOPSIS
    Assign Horizon users to virtual machines from a CSV.
.DESCRIPTION
    Reads a CSV of HorizonServer, UserUPN, MachineName; connects and assigns each user.
.PARAMETER AssignmentListPath
    Path to the CSV file.
.PARAMETER LogFile
    Path to the log file.
.PARAMETER WorkingDirectory
    Root script directory.
.NOTES
    Author: Marinus van Deventer
    Version: 1.0.0
    Date: 2025-05-30
#>

#region Parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkingDirectory = 'C:\temp\scripts',

    [Parameter(Mandatory = $false, HelpMessage = 'CSV: HorizonServer,UserUPN,MachineName')]
    [ValidateNotNullOrEmpty()]
    [string]$AssignmentListPath = "$WorkingDirectory\Assignments.csv",

    [Parameter(Mandatory = $false, HelpMessage = 'Log file path; default in logs folder')]
    [string]$LogFile = "$WorkingDirectory\logs\HorizonAssignment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)
#endregion

#region Initialization
# Ensure directories exist
if (-not (Test-Path $WorkingDirectory)) { New-Item -Path $WorkingDirectory -ItemType Directory | Out-Null }
$logDirectory = "$WorkingDirectory\logs"
if (-not (Test-Path $logDirectory)) { New-Item -Path $logDirectory -ItemType Directory | Out-Null }

# Script version
$scriptVersion = '1.0.0'
#endregion

#region Logging Function
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Message
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "$timestamp [$scriptVersion] - $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Host $entry -ForegroundColor Cyan
}
#endregion

#region Module Validation and Import
$requiredModules = @('VMware.VimAutomation.HorizonView','VMware.Hv.Helper')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "ERROR: Module '$module' not found. Install via PowerCLI or download helper module."
        throw "Missing module: $module"
    }
    try {
        Import-Module $module -ErrorAction Stop
        Write-Log "Imported module: $module"
    } catch {
        Write-Log "ERROR: Import failed for module '$module': $_"
        throw
    }
}
#endregion

#region Import Assignments
if (-not (Test-Path $AssignmentListPath)) {
    Write-Log "ERROR: CSV not found at $AssignmentListPath"
    throw 'CSV missing'
}
try {
    $assignments = Import-Csv -Path $AssignmentListPath
    if ($assignments.Count -eq 0) {
        Write-Log 'ERROR: No entries in assignment CSV.'
        throw 'Empty CSV'
    }
    Write-Log "Loaded $($assignments.Count) assignments."
} catch {
    Write-Log "ERROR: Import-Csv failed: $_"
    throw
}
#endregion

#region Credential Prompt
$credential = Get-Credential -Message 'Enter Horizon Admin credentials'
#endregion

#region Helper: Get-UserId
function Get-UserId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [VMware.Hv.Services]$Services,

        [Parameter(Mandatory)]
        [string]$UserName
    )
    $defn = [VMware.Hv.QueryDefinition]::new()
    $defn.queryEntityType = 'ADUserOrGroupSummaryView'

    $nameFilter = [VMware.Hv.QueryFilterEquals]::new()
    $nameFilter.PropertyName = 'base.name'
    $nameFilter.Value        = $UserName

    $groupFilter = [VMware.Hv.QueryFilterEquals]::new()
    $groupFilter.PropertyName = 'base.group'
    $groupFilter.Value        = $false

    $andFilter = [VMware.Hv.QueryFilterAnd]::new()
    $andFilter.Filters = @($nameFilter, $groupFilter)
    $defn.Filter       = $andFilter

    $result = $Services.QueryService.Query($defn)
    if (-not $result.Results) {
        throw "User '$UserName' not found via API."
    }
    return $result.Results[0]
}
#endregion

#region Process Assignments
foreach ($entry in $assignments) {
    $server      = $entry.HorizonServer
    $userUpn     = $entry.UserUPN
    $machineName = $entry.MachineName

    if (-not ($server -and $userUpn -and $machineName)) {
        Write-Log "WARNING: Skipping incomplete entry: $($entry | Out-String)"
        continue
    }

    Write-Log "Processing assignment: Server=$server, User=$userUpn, VM=$machineName"

    try {
        $hvConnection = Connect-HVServer -Server $server -Credential $credential -ErrorAction Stop
        Write-Log "Connected to $server"
    } catch {
        Write-Log "ERROR: Could not connect to $server: $_"
        continue
    }

    try {
        # Machine lookup
        $vmObj = Get-HVMachine -Server $hvConnection | Where-Object { $_.Base.Name -eq $machineName }
        if (-not $vmObj) {
            Write-Log "ERROR: VM '$machineName' not found on $server"
            continue
        }

        # User lookup
        $services = $hvConnection.ExtensionData
        $userInfo = Get-UserId -Services $services -UserName $userUpn

        # Assign
        $assignmentSpec = [VMware.Hv.MachineAssignmentSpec]::new()
        $assignmentSpec.Id   = $vmObj.Id
        $assignmentSpec.User = $userUpn
        $services.Machine.AssignUser($assignmentSpec)

        Write-Log "SUCCESS: Assigned $userUpn to $machineName"
    } catch {
        Write-Log "ERROR: Assignment error: $_"
    } finally {
        try {
            Disconnect-HVServer -Server $server -Confirm:$false
            Write-Log "Disconnected from $server"
        } catch {
            Write-Log "WARNING: Disconnect failed: $_"
        }
    }
}
#endregion

Write-Log "Script completed successfully at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
