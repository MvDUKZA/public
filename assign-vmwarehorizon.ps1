<#
.SYNOPSIS
    Assigns Horizon users to virtual machines from a CSV input using VMware modules.
.DESCRIPTION
    Reads a CSV with HorizonServer, UserUPN, and MachineName, connects to each Horizon environment,
    and assigns each user to the specified machine via the Horizon View API.
.PARAMETER AssignmentListPath
    Path to the CSV input file containing assignments.
.PARAMETER LogFile
    Path to the log file.
.NOTES
    Author: Marinus van Deventer
    Version: 1.4
    Requires: VMware.VimAutomation.HorizonView, VMware.Hv.Helper
    Date: 2025-05-30
#>

#region Parameters
[CmdletBinding()]
param (
    [Parameter(HelpMessage = "Path to the assignment CSV file.")]
    [ValidateNotNullOrEmpty()]
    [string]$AssignmentListPath = "C:\temp\scripts\Assignments.csv",

    [Parameter(HelpMessage = "Path to the log file.")]
    [string]$LogFile
)

if (-not $LogFile) {
    $LogFile = "C:\temp\scripts\logs\HorizonAssignment_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
}
#endregion

#region Logging Function
function Write-Log {
    param ([string]$Message)
    $t = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$t - $Message"
    Add-Content -Path $LogFile -Value $entry
    Write-Host  $entry
}
#endregion

#region Module Check & Import
foreach ($m in "VMware.VimAutomation.HorizonView","VMware.Hv.Helper") {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Log "ERROR: Module '$m' missing. Install VMware PowerCLI or the helper manually."
        throw "Missing module $m"
    }
    Import-Module -Name $m -ErrorAction Stop
    Write-Log "Imported module: $m"
}
#endregion

#region CSV Import
if (-not (Test-Path $AssignmentListPath)) {
    Write-Log "ERROR: CSV not found at $AssignmentListPath"
    throw "CSV file missing"
}
$assignments = Import-Csv -Path $AssignmentListPath
if (-not $assignments) {
    Write-Log "ERROR: No records in CSV"
    throw "CSV contains no data"
}
Write-Log "Loaded $($assignments.Count) assignments"
#endregion

#region Get Credentials
$cred = Get-Credential -Message "Enter Horizon Admin credentials"
#endregion

#region Process Assignments
foreach ($a in $assignments) {
    $svr       = $a.HorizonServer
    $userUPN   = $a.UserUPN
    $vmName    = $a.MachineName

    if (-not ($svr -and $userUPN -and $vmName)) {
        Write-Log "WARN: Incomplete row, skipping: $($a|Out-String)"
        continue
    }
    Write-Log ">>> $svr: assign $userUPN → $vmName"

    # Connect
    try {
        $hv = Connect-HVServer -Server $svr -Credential $cred -ErrorAction Stop
        Write-Log "Connected to $svr"
    } catch {
        Write-Log "ERROR: Connection to $svr failed: $_"
        continue
    }

    try {
        $svc = $hv.ExtensionData
        $qs  = $svc.QueryService

        # — Find VM —
        $mq = [VMware.Hv.QueryDefinition]::new()
        $mq.queryEntityType = 'MachineSummaryView'
        $filter = [VMware.Hv.QueryFilterEquals]::new()
        $filter.PropertyName = 'base.name'
        $filter.Value        = $vmName
        $mq.filter = ,$filter

        $mres = $qs.Query($mq)
        $vm   = $mres.Results | Select-Object -First 1
        if (-not $vm) {
            Write-Log "ERROR: VM '$vmName' not found"
            continue
        }

        # — Find User —
        $uq = [VMware.Hv.QueryDefinition]::new()
        $uq.queryEntityType = 'UserSummaryView'
        $uf = [VMware.Hv.QueryFilterEquals]::new()
        $uf.PropertyName = 'base.userName'
        $uf.Value        = $userUPN
        $uq.filter = ,$uf

        $ures = $qs.Query($uq)
        $usr  = $ures.Results | Select-Object -First 1
        if (-not $usr) {
            Write-Log "ERROR: User '$userUPN' not found"
            continue
        }

        # — Assign —
        $spec = [VMware.Hv.MachineAssignmentSpec]::new()
        $spec.Id   = $vm.Id
        $spec.User = $userUPN

        $svc.Machine.AssignUser($spec)
        Write-Log "SUCCESS: $userUPN → $vmName"
    } catch {
        Write-Log "ERROR: Assignment for $userUPN → $vmName failed: $_"
    } finally {
        try {
            Disconnect-HVServer -Server $svr -Confirm:$false
            Write-Log "Disconnected from $svr"
        } catch {
            Write-Log "WARN: Disconnect failed: $_"
        }
    }
}
#endregion

Write-Log "All done at $(Get-Date)"
