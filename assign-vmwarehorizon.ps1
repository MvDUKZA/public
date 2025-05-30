$ErrorActionPreference = 'Stop'
<#
.SYNOPSIS
    Assign a dedicated-pool Horizon desktop to a user (Horizon 8.12.x).

.DESCRIPTION
    Called from ControlUp with five arguments:
      0 – Connection-server FQDN
      1 – Desktop-pool display-name
      2 – VM / Machine name
      3 – User login-name (sAMAccountName, **not** UPN)
      4 – AD domain (NETBIOS or DNS)

    The script:
      • Loads PowerCLI (HorizonView) modules
      • Connects to the Connection-server with stored creds
      • Looks up pool → machine → user → IDs via QueryService
      • Confirms pool is DEDICATED
      • Uses MachineService to assign the user
      • Logs progress and errors to console + file

.NOTES
    Author :  Wouter Kursten (original) / Marinus van Deventer (re-fit)
    Version : 1.1  (30-May-2025)
    Requires: PowerShell 5.1+, VMware.PowerCLI ≤ 12.x
#>

#region ----- args → strongly typed variables -----
[string]$hvConnectionServer = $args[0]
[string]$hvDesktopPoolName  = $args[1]
[string]$hvMachineName      = $args[2]
[string]$hvUserName         = $args[3]
[string]$hvDomain           = $args[4]
#endregion

#region ----- helpers -----
$logRoot = 'C:\temp\scripts\logs'
if (-not (Test-Path $logRoot)) { New-Item $logRoot -ItemType Directory -Force | Out-Null }
$logFile = Join-Path $logRoot ("HvAssign_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$ts [1.1][$Level] $Message" | Tee-Object -FilePath $logFile -Append
}

function Test-ArgsCount {
    param(
        [int]$Expected = 5,
        [string]$Reason = 'Argument mismatch'
    )
    if ($args.Count -lt $Expected) {
        Write-Log "$Reason – expected $Expected, got $($args.Count)" 'ERROR'
        exit 1
    }
}
#endregion

Test-ArgsCount

#region ----- load PowerCLI -----
try {
    Import-Module VMware.VimAutomation.HorizonView -ErrorAction Stop
    Write-Log 'Imported VMware.VimAutomation.HorizonView'
} catch {
    try { Add-PSSnapin VMware -ErrorAction Stop }
    catch { Write-Log 'PowerCLI modules/snap-ins not found.' 'ERROR'; throw }
}
#endregion

#region ----- credential & connection -----
$creds = Import-Clixml "$env:ProgramData\ControlUp\ScriptSupport\$($env:USERNAME)_HorizonView_Cred.xml"
$hv   = Connect-HVServer -Server $hvConnectionServer -Credential $creds -ErrorAction Stop
$svc  = $hv.ExtensionData
$qsvc = New-Object VMware.Hv.QueryServiceService
Write-Log "Connected to $hvConnectionServer"
#endregion

#region ----- lookup pool -----
$qPool = [VMware.Hv.QueryDefinition]@{ queryEntityType = 'DesktopSummaryView' }
$qPool.Filter = New-Object VMware.Hv.QueryFilterEquals -Property @{
    MemberName = 'desktopSummaryData.displayName'; Value = $hvDesktopPoolName
}
$pool = ($qsvc.queryService_create($svc, $qPool)).Results[0] ; $qsvc.QueryService_DeleteAll($svc)
if (-not $pool) { Write-Log "Pool '$hvDesktopPoolName' not found" 'ERROR'; Disconnect-HVServer -Confirm:$false; exit }
$poolId   = $pool.id
$poolSpec = $svc.Desktop.Desktop_Get($poolId)
if (($poolSpec.Type -eq 'AUTOMATED' -and $poolSpec.AutomatedDesktopData.UserAssignment.UserAssignment -ne 'DEDICATED') -or
    ($poolSpec.Type -eq 'MANUAL'    -and $poolSpec.ManualDesktopData.UserAssignment.UserAssignment    -ne 'DEDICATED')) {
    Write-Log "Pool '$hvDesktopPoolName' is not DEDICATED – assignment aborted" 'ERROR'
    Disconnect-HVServer -Confirm:$false; exit
}
#endregion

#region ----- lookup machine -----
$qMach = [VMware.Hv.QueryDefinition]@{ queryEntityType = 'MachineDetailsView' }
$qMach.Filter = New-Object VMware.Hv.QueryFilterAnd
$qMach.Filter.Filters = @(
    New-Object VMware.Hv.QueryFilterEquals -Property @{ MemberName='desktopData.id'; Value=$poolId },
    New-Object VMware.Hv.QueryFilterEquals -Property @{ MemberName='data.name';     Value=$hvMachineName }
)
$machine = ($qsvc.queryService_create($svc,$qMach)).Results[0] ; $qsvc.QueryService_DeleteAll($svc)
if (-not $machine) { Write-Log "Machine '$hvMachineName' not found" 'ERROR'; Disconnect-HVServer -Confirm:$false; exit }
$machineId = $machine.id
#endregion

#region ----- lookup user -----
$qUser = [VMware.Hv.QueryDefinition]@{ queryEntityType = 'ADUserOrGroupSummaryView' }
$qUser.Filter = New-Object VMware.Hv.QueryFilterAnd
$qUser.Filter.Filters = @(
    New-Object VMware.Hv.QueryFilterEquals -Property @{ MemberName='base.loginName'; Value=$hvUserName },
    New-Object VMware.Hv.QueryFilterEquals -Property @{ MemberName='base.domain';    Value=$hvDomain },
    New-Object VMware.Hv.QueryFilterEquals -Property @{ MemberName='base.group';     Value=$false  }
)
$user = ($qsvc.queryService_create($svc,$qUser)).Results[0] ; $qsvc.QueryService_DeleteAll($svc)
if (-not $user) { Write-Log "User '$hvUserName@$hvDomain' not found" 'ERROR'; Disconnect-HVServer -Confirm:$false; exit }
$userId = $user.id
#endregion

#region ----- perform assignment -----
try {
    $mSvc  = New-Object VMware.Hv.MachineService
    $mInfo = $mSvc.Read($svc, $machineId)
    $mInfo.GetBaseHelper().SetUser($userId)
    $mSvc.Update($svc, $mInfo)
    Write-Log "SUCCESS – $hvUserName assigned to $hvMachineName"
} catch {
    Write-Log "Assignment failed: $_" 'ERROR'
}
#endregion

Disconnect-HVServer -Confirm:$false
Write-Log "Disconnected; script finished"
