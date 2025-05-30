<#
.SYNOPSIS
    Assigns Horizon users to virtual machines from a CSV input.
.DESCRIPTION
    Reads a CSV with HorizonServer, UserUPN, and MachineName, connects to each Horizon environment,
    and assigns each user to the specified machine.
.PARAMETER AssignmentListPath
    Path to the CSV input file containing assignments.
.PARAMETER LogFile
    Path to the log file.
.NOTES
    Author: Marinus van Deventer
    Version: 1.4
    Requires: Omnissa.VimAutomation.HorizonView, Omnissa.Horizon.Helper
    Date: 2025-05-29
#>

#region Parameters
param (
    [Parameter(Mandatory = $false)]
    [string]$AssignmentListPath = "C:\temp\scripts\Assignments.csv",

    [Parameter(Mandatory = $false)]
    [string]$LogFile = "C:\temp\scripts\logs\OmnissaAssignment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)
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

#region Ensure Log Directory Exists
$logDir = Split-Path -Path $LogFile -Parent
if (-not (Test-Path $logDir)) {
    try {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Write-Host "Created log directory: $logDir"
    } catch {
        Write-Host "ERROR: Failed to create log directory: $logDir. $_"
        throw
    }
}
#endregion

#region Module Validation and Import
$requiredModules = @("Omnissa.VimAutomation.HorizonView", "Omnissa.Horizon.Helper")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "ERROR: Required module '$module' not found. Please install from https://developer.omnissa.com/horizon-powercli/download/"
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

        $spec = New-Object Omnissa.Horizon.Helper.HVMachineAssignmentSpec
        $spec.User = $userUPN
        $spec.Machine = $machineName

        Set-HVMachineAssignment -AssignmentSpec $spec -ErrorAction Stop
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
