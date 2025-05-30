<#
.SYNOPSIS
    Assigns Horizon users to machines from a CSV for Horizon 8.12.1 environments.
.DESCRIPTION
    Reads a CSV (HorizonServer, UserUPN, MachineName), connects to each Horizon server,
    and assigns each user to the correct VM, using only supported APIs.
.PARAMETER AssignmentListPath
    Path to the CSV file.
.PARAMETER LogFile
    Path to the log file (default: C:\temp\scripts\logs\HorizonAssignment_yyyymmdd_hhmmss.log).
.NOTES
    Author: Marinus van Deventer
    Version: 1.0.0
    Date: 2025-05-30
#>

#region Parameters
[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$AssignmentListPath = "C:\temp\scripts\Assignments.csv",

    [Parameter()]
    [string]$LogFile
)

# Ensure log directory exists and set default log path if missing
if (-not $LogFile) {
    $LogDirectory = "C:\temp\scripts\logs"
    if (-not (Test-Path $LogDirectory)) { New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null }
    $LogFile = Join-Path $LogDirectory ("HorizonAssignment_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
#endregion

#region Logging Function
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ver = "[1.0.0]"
    $entry = "$ts $ver [$Level] - $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Host $entry
}
#endregion

#region Module Checks and Import
$modules = @("VMware.VimAutomation.HorizonView", "VMware.Hv.Helper")
foreach ($m in $modules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Log "ERROR: Module $m missing." "ERROR"
        throw "Module $m missing."
    }
    Import-Module $m -ErrorAction Stop
    Write-Log "Imported module: $m"
}
#endregion

#region CSV Import
if (-not (Test-Path $AssignmentListPath)) {
    Write-Log "ERROR: CSV not found at $AssignmentListPath" "ERROR"
    throw "CSV file missing"
}
$assignments = Import-Csv $AssignmentListPath
if (-not $assignments) {
    Write-Log "ERROR: No assignments in CSV" "ERROR"
    throw "CSV empty"
}
Write-Log "Loaded $($assignments.Count) assignments"
#endregion

#region Credentials
$cred = Get-Credential -Message "Enter Horizon Admin credentials"
#endregion

#region Get-UserId helper
function Get-UserId {
    param(
        [VMware.Hv.Services]$Services,
        [string]$UserName
    )
    $defn = New-Object VMware.Hv.QueryDefinition
    $defn.queryEntityType = 'ADUserOrGroupSummaryView'
    $groupFilter = New-Object VMware.Hv.QueryFilterEquals
    $groupFilter.PropertyName = 'base.group'
    $groupFilter.Value = $false
    $userFilter = New-Object VMware.Hv.QueryFilterEquals
    $userFilter.PropertyName = 'base.name'
    $userFilter.Value = $UserName
    $andFilter = New-Object VMware.Hv.QueryFilterAnd
    $andFilter.Filters = @($userFilter, $groupFilter)
    $defn.Filter = $andFilter
    $res = $Services.QueryService.Query($defn)
    if (-not $res.Results) {
        throw "User '$UserName' not found"
    }
    return $res.Results[0]
}
#endregion

#region Process Assignments
foreach ($a in $assignments) {
    $server      = $a.HorizonServer
    $userUPN     = $a.UserUPN
    $machineName = $a.MachineName

    if (-not ($server -and $userUPN -and $machineName)) {
        Write-Log "Skipping incomplete entry: $($a | Out-String)" "WARN"
        continue
    }

    Write-Log "Processing: Server=$server, User=$userUPN, Machine=$machineName"

    try {
        $hv = Connect-HVServer -Server $server -Credential $cred -ErrorAction Stop
        $services = $hv.ExtensionData
        Write-Log "Connected to $server"
    } catch {
        Write-Log "Connection failed: $_" "ERROR"
        continue
    }

    try {
        # Find the machine by name
        $vm = Get-HVMachine | Where-Object { $_.Base.Name -eq $machineName }
        if (-not $vm) {
            Write-Log "Machine '$machineName' not found" "ERROR"
            continue
        }

        # Find the user by UPN (using the helper)
        $user = Get-UserId -Services $services -UserName $userUPN

        # Assign user to machine
        $spec = New-Object VMware.Hv.MachineAssignmentSpec
        $spec.Id   = $vm.Id
        $spec.User = $userUPN

        $services.Machine.AssignUser($spec)
        Write-Log "Assigned $userUPN to $machineName"
    } catch {
        Write-Log "Assignment error: $_" "ERROR"
    } finally {
        try {
            Disconnect-HVServer -Confirm:$false
            Write-Log "Disconnected from $server"
        } catch {
            Write-Log "Disconnect failed: $_" "WARN"
        }
    }
}
#endregion

Write-Log "Script completed successfully at $(Get-Date)"
