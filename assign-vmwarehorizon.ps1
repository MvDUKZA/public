#Requires -Version 5.1
#Requires -Modules VMware.VimAutomation.HorizonView

<#
.SYNOPSIS
    Bulk-assign dedicated VMware Horizon 8.12.1 desktops to users defined in a CSV file.

.DESCRIPTION
    Reads a CSV mapping (MachineName, User[, Domain]) and assigns each Horizon desktop VM
    to the specified user. Supports only dedicated-assignment pools.

    Starting with Horizon 2312 the internal .NET proxy classes were renamed from
    **VMware.Hv.*** to **Omnissa.Horizon.***. Horizon 8.12.1 still uses the old
    VMware.Hv namespace. We therefore first look for *VMware.Hv.QueryServiceService* and
    fall back to the Omnissa namespace if running a newer build. When neither proxy type
    exists, we attempt to import **Omnissa.Horizon.Helper** which can generate them.

.PARAMETER CsvPath       Path to CSV file (headers: MachineName, User [,Domain]).
.PARAMETER ConnectionServer  Horizon Connection Server FQDN (or $env:HVConnectionServer).
.PARAMETER Credential    PSCredential for Horizon login (prompted if omitted).

.EXAMPLE
    .\Assign-HorizonDesktops.ps1 -CsvPath .\Assignments.csv -ConnectionServer view01.iprod.local -Verbose

.NOTES
    Author  : Marinus van Deventer
    Version : 2.6
    Date    : 30-May-2025
    Changelog
      2.6 - Detect VMware.Hv namespace first (as used by Horizon 8.12.1); import helper
            only when neither VMware.Hv nor Omnissa.Horizon proxies are present.
      2.5 - Added Omnissa fallback logic.
      2.4 - Removed helper dependency.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)][ValidateScript({ Test-Path $_ })][string]$CsvPath,
    [Parameter()][ValidateNotNullOrEmpty()][string]$ConnectionServer = $env:HVConnectionServer,
    [Parameter()][System.Management.Automation.PSCredential]$Credential
)
if (-not $ConnectionServer) { throw 'ConnectionServer is mandatory when $env:HVConnectionServer is not set.' }
$ErrorActionPreference = 'Stop'

#region Logging
$workDir = 'C:\temp\scripts'
$logDir  = Join-Path $workDir 'logs'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$logFile = Join-Path $logDir ("Assign-HorizonDesktop_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $logFile -Append | Out-Null
function Write-Log { param($Message,[ValidateSet('INFO','WARN','ERROR','DEBUG')]$Level='INFO'); Write-Information "$(Get-Date -UFormat '%Y-%m-%d %H:%M:%S') - $Level : $Message" -InformationAction Continue }
#endregion

try {
    #region Load Horizon module and connect
    Write-Log 'Importing VMware PowerCLI Horizon module...'
    Import-Module VMware.VimAutomation.HorizonView -ErrorAction Stop
    if (-not $Credential) { $Credential = Get-Credential -Message "Credentials for $ConnectionServer" }
    Write-Log "Connecting to $ConnectionServer..."
    $hvServer = Connect-HVServer -Server $ConnectionServer -Credential $Credential
    #endregion

    #region Detect proxy namespace
    $vmwareNs  = [type]::GetType('VMware.Hv.QueryServiceService', $false)
    $omnissaNs = [type]::GetType('Omnissa.Horizon.QueryServiceService', $false)
    if (-not ($vmwareNs -or $omnissaNs)) {
        Write-Log 'Horizon proxy types missing – importing Omnissa.Horizon.Helper...' 'WARN'
        try {
            Import-Module Omnissa.Horizon.Helper -ErrorAction Stop
            $vmwareNs  = [type]::GetType('VMware.Hv.QueryServiceService', $false)
            $omnissaNs = [type]::GetType('Omnissa.Horizon.QueryServiceService', $false)
            if (-not ($vmwareNs -or $omnissaNs)) { throw 'Proxy types still absent after helper load' }
            Write-Log 'Proxy types generated via helper.' 'INFO'
        } catch { throw "Unable to load Horizon API proxy types: $($_.Exception.Message)" }
    }
    $nsPrefix = if ($vmwareNs) { 'VMware.Hv' } else { 'Omnissa.Horizon' }
    Write-Log "Using proxy namespace: $nsPrefix" 'DEBUG'
    #endregion

    #region Helper functions (namespace‑agnostic)
    function New-ServiceInstance([string]$class) {
        return [Activator]::CreateInstance([type]::GetType("$nsPrefix.$class", $true))
    }

    function Get-HvUserObject {
        param([string]$User,[string]$Domain,$Conn)
        $ uname = $User -replace '^(.*\\)|@.*$',''
        $dom = if ($Domain) { $Domain } elseif ($User -match '\\') { ($User -split '\\')[0] } elseif ($User -match '@') { ($User -split '@')[1] } else { ($env:USERDNSDOMAIN) }
        $qs = New-ServiceInstance 'QueryServiceService'
        $def = New-ServiceInstance 'QueryDefinition'
        $def.queryEntityType = 'ADUserOrGroupSummaryView'
        $f1  = New-ServiceInstance 'QueryFilterEquals'
        $f1.memberName='base.loginName';$f1.value=$uname
        $f2  = New-ServiceInstance 'QueryFilterEquals'
        $f2.memberName='base.domain';$f2.value=$dom
        $and = New-ServiceInstance 'QueryFilterAnd'
        $and.Filters=@($f1,$f2); $def.Filter=$and
        $res = ($qs.queryService_create($Conn.ExtensionData,$def)).results
        $qs.QueryService_DeleteAll($Conn.ExtensionData)
        return $res[0]
    }

    function Set-HvMachineUser {
        param($Machine,$User,$Conn)
        $svc   = New-ServiceInstance 'MachineService'
        $helper= $svc.read($Conn.ExtensionData,$Machine.id)
        $helper.getbasehelper().setuser($User.id)
        $svc.update($Conn.ExtensionData,$helper)
    }
    #endregion

    #region CSV processing
    Write-Log "Reading assignments from $CsvPath..."
    $rows = Import-Csv -Path $CsvPath
    if (-not $rows) { throw 'CSV has no data.' }
    foreach ($col in 'MachineName','User') { if (-not ($rows[0].psobject.Properties.Name -contains $col)) { throw "CSV missing $col column" } }
    $total=$rows.Count; $i=0; $out=@()
    foreach ($row in $rows) {
        $i++
        Write-Progress -Activity 'Assigning desktops' -Status "Processing $i of $total" -PercentComplete ([int]($i/$total*100))
        if ([string]::IsNullOrWhiteSpace($row.MachineName) -or [string]::IsNullOrWhiteSpace($row.User)) { Write-Log "Row $i incomplete – skipped." 'WARN'; continue }
        try {
            $machine = Get-HVMachine -MachineName $row.MachineName -HvServer $hvServer
            if (-not $machine) { throw 'Machine not found' }
            $pool = Get-HVPoolSummary -PoolName $machine.base.desktopName -HvServer $hvServer
            if ($pool.userAssignment -ne 'DEDICATED') { throw 'Pool not dedicated' }
            $user = Get-HvUserObject -User $row.User -Domain $row.Domain -Conn $hvServer
            if (-not $user) { throw 'User not found' }
            if ($PSCmdlet.ShouldProcess($row.MachineName,"assign to $($row.User)")) { Set-HvMachineUser $machine $user $hvServer }
            $status='Assigned'
        } catch { $status="Failed - $($_.Exception.Message)"; Write-Log "$($row.MachineName)->$($row.User): $status" 'ERROR' }
        $out += [pscustomobject]@{Machine=$row.MachineName;User=$row.User;Status=$status}
    }
    Write-Progress -Activity 'Assigning desktops' -Completed -Status 'Done'
    Write-Log 'Processing finished.'
    $out | Sort-Object Status,Machine | Format-Table -AutoSize
    #endregion
}
finally { if ($hvServer) { Disconnect-HVServer -Server $hvServer -Confirm:$false } Stop-Transcript -ErrorAction SilentlyContinue | Out-Null }

# Signed-off-by: Marinus van Deventer
# End of script
