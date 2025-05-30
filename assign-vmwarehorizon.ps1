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

#region Module Check and Installation
$requiredModules = @(
    @{ Name = "VMware.VimAutomation.HorizonView"; Source = "VMware.PowerCLI" },
    @{ Name = "VMware.Hv.Helper"; Source = "VMware.Hv.Helper" }
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module.Name)) {
        Write-Log "Module '$($module.Name)' not found. Attempting installation..."
        try {
            # Install VMware.PowerCLI for core modules
            if ($module.Source -eq "VMware.PowerCLI") {
                Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Log "Installed VMware.PowerCLI module, which includes $($module.Name)"
            }
            elseif ($module.Source -eq "VMware.Hv.Helper") {
                # VMware.Hv.Helper might not be on PowerShell Gallery; try PSGallery or prompt user
                try {
                    Install-Module -Name VMware.Hv.Helper -Scope CurrentUser -Force -ErrorAction Stop
                    Write-Log "Installed VMware.Hv.Helper module."
                } catch {
                    Write-Log "WARNING: VMware.Hv.Helper is not available on PSGallery. Install it manually from GitHub."
                    Write-Log "URL: https://github.com/vmware/PowerCLI-Example-Scripts/tree/master/Modules/VMware.Hv.Helper"
                    throw
                }
            }
        } catch {
            Write-Log "ERROR: Failed to install module '$($module.Name)'. $_"
            throw
        }
    }

    try {
        Import-Module $module.Name -ErrorAction Stop
        Write-Log "Imported module: $($module.Name)"
    } catch {
        Write-Log "ERROR: Failed to import module '$($module.Name)'. $_"
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
        $machine = Get-HVMachine -Name $machineName -ErrorAction Stop
        if (-not $machine) {
            Write-Log "ERROR: Machine '$machineName' not found on $server."
            continue
        }

        $user = Get-HVUser -UserName $userUPN -ErrorAction Stop
        if (-not $user) {
            Write-Log "ERROR: User '$userUPN' not found on $server."
            continue
        }

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
