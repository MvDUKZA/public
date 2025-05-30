<#
.SYNOPSIS
    Assigns Horizon users to virtual machines from a CSV input using VMware modules.
.DESCRIPTION
    Reads a CSV with HorizonServer, UserUPN, and MachineName, connects to each Horizon environment,
    and assigns each user to the specified machine.
.PARAMETER AssignmentListPath
    Path to the CSV input file containing assignments.
.PARAMETER LogFile
    Path to the log file.
.EXAMPLE
    .\Assign-HorizonUsers.ps1 -AssignmentListPath "C:\temp\scripts\Assignments.csv"
.NOTES
    Author: Marinus van Deventer
    Version: 1.1
    Requires: VMware.VimAutomation.HorizonView, VMware.Hv.Helper
    Date: 2025-05-30
#>

#region Parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Path to the assignment CSV file.")]
    [ValidateNotNullOrEmpty()]
    [string]$AssignmentListPath = "C:\temp\scripts\Assignments.csv",

    [Parameter(Mandatory = $false, HelpMessage = "Path to the log file.")]
    [string]$LogFile
)

# Set dynamic default log file path if not specified
if (-not $LogFile) {
    $LogFile = "C:\temp\scripts\logs\HorizonAssignment_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
}
#endregion

#region Logging Function
function Write-Log {
    param (
        [Parameter(Mandatory)]
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Host $entry
}
#endregion

#region Module Validation and Import
$requiredModules = @("VMware.VimAutomation.HorizonView", "VMware.Hv.Helper")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "ERROR: Required module '$module' not found. Install via PowerCLI or VMware PowerShell Gallery."
        throw "Module '$module' is missing. Script cannot continue."
    }
    try {
        Import-Module $module -ErrorAction Stop
        Write-Log "Imported module: $module"
    } catch {
        Write-Log "ERROR: Failed to import module '$module'. $_"
        throw
    }
}
#endregion

#region Validate CSV Path
if (-not (Test-Path $AssignmentListPath)) {
    Write-Log "ERROR: CSV file not found at $AssignmentListPath"
    throw "CSV file missing. Provide a valid AssignmentListPath."
}
#endregion

#region Import CSV Assignments
try {
    $assignments = Import-Csv -Path $AssignmentListPath
    if (-not $assignments) {
        Write-Log "ERROR: No assignments found in CSV at $AssignmentListPath"
        throw "CSV contains no data."
    }
    Write-Log "Imported $($assignments.Count) assignments from $AssignmentListPath"
} catch {
    Write-Log "ERROR: Failed to import CSV. $_"
    throw
}
#endregion

#region Secure Credential Prompt
$cred = Get-Credential -Message "Enter Horizon Admin Credentials"
#endregion

#region Process Assignments
foreach ($assignment in $assignments) {
    $server = $assignment.HorizonServer
    $userUPN = $assignment.UserUPN
    $machineName = $assignment.MachineName

    # Validate assignment entry
    if (-not $server -or -not $userUPN -or -not $machineName) {
        Write-Log "WARNING: Incomplete assignment entry found. Skipping: $($assignment | Out-String)"
        continue
    }

    Write-Log "Processing assignment: Server=$server, User=$userUPN, Machine=$machineName"

    try {
        $hvServer = Connect-HVServer -Server $server -Credential $cred -ErrorAction Stop
        Write-Log "Connected to $server"
    } catch {
        Write-Log "ERROR: Failed to connect to $server. $_"
        continue
    }

    try {
        # Retrieve machine and user using summaries
        $machine = Get-HVMachineSummary | Where-Object { $_.base.name -eq $machineName }
        if (-not $machine) {
            Write-Log "ERROR: Machine '$machineName' not found on $server."
            continue
        }

        $user = Get-HVUserSummary | Where-Object { $_.base.name -eq $userUPN }
        if (-not $user) {
            Write-Log "ERROR: User '$userUPN' not found on $server."
            continue
        }

        # Perform assignment
        $services = $hvServer.ExtensionData
        $machineService = $services.Machine
        $assignmentSpec = New-Object VMware.Hv.MachineAssignmentSpec
        $assignmentSpec.Id = $machine.Id
        $assignmentSpec.User = $userUPN
        $machineService.AssignUser($assignmentSpec)

        Write-Log "SUCCESS: Assigned $machineName to $userUPN on $server"
    } catch {
        Write-Log "ERROR: Failed assignment for $userUPN on $server. $_"
    } finally {
        try {
            Disconnect-HVServer -Server $server -Confirm:$false
            Write-Log "Disconnected from $server"
        } catch {
            Write-Log "WARNING: Could not disconnect from $server. $_"
        }
    }
}
#endregion

Write-Log "Script completed at $(Get-Date)"
