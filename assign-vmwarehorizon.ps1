#Requires -Version 5.1
#Requires -Modules VMware.VimAutomation.HorizonView, Omnissa.Horizon.Helper, ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, HelpMessage = 'Full path to CSV mapping file')]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(HelpMessage = 'Horizon Connection Server FQDN')]
    [ValidateNotNullOrEmpty()]
    [string]$ConnectionServer = $env:HVConnectionServer,

    [Parameter(HelpMessage = 'Credential for Horizon authentication')]
    [System.Management.Automation.PSCredential]$Credential
)

$ErrorActionPreference = 'Stop'

<#
.SYNOPSIS
    Bulk‑assign dedicated VMware Horizon 8.12.1 desktops to users defined in a CSV file.

.DESCRIPTION
    Reads a mapping file (CSV) and binds each Horizon desktop machine to a user account.
    Only *dedicated‑assignment* pools are supported – floating pools do not maintain a
    persistent machine‑to‑user relationship.

    The script relies solely on official Omnissa/VMware modules and does **not** require
    ControlUp.

.EXAMPLE
    PS C:\temp\scripts> .\Assign-HorizonDesktops.ps1 -CsvPath .\Mappings.csv -ConnectionServer view01.acme.local -Verbose

.NOTES
    Author  : Marinus van Deventer
    Created : 30‑May‑2025
    Version : 1.3
    Logs    : C:\temp\scripts\logs\Assign-HorizonDesktop_yyyyMMdd_HHmmss.log
    Change‑log:
        1.3 – Moved #Requires statements to the top and reordered $ErrorActionPreference per VS Code/PSScriptAnalyzer.
#>

#region Initialisation
$workingDir = 'C:\temp\scripts'
$logDir     = Join-Path $workingDir 'logs'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

$timeStamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile    = Join-Path $logDir "Assign-HorizonDesktop_$timeStamp.log"
Start-Transcript -Path $logFile -Append | Out-Null

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','DEBUG','ERROR','WARN')][string]$Level = 'INFO'
    )
    $prefix = "[$((Get-Date).ToString('u'))] [$Level]"
    Write-Information "$prefix $Message" -InformationAction Continue
}
#endregion

#region Module loading
Write-Log 'Loading Omnissa PowerCLI modules…'
Import-Module VMware.VimAutomation.HorizonView -ErrorAction Stop
if (-not (Get-Module -ListAvailable Omnissa.Horizon.Helper)) {
    Install-Module Omnissa.Horizon.Helper -Scope AllUsers -Force -AllowClobber
}
Import-Module Omnissa.Horizon.Helper -ErrorAction Stop
#endregion

#region Helper functions
function Get-HvUserObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter()][string]$Domain,
        [Parameter(Mandatory)]$HvServer
    )
    $sam  = $User -replace '^(.*\\)|@.*$',''
    $dom  = if ($Domain) { $Domain } elseif ($User -match '\\') { ($User -split '\\')[0] } elseif ($User -match '@') { ($User -split '@')[1] } else { (Get-ADDomain).DNSRoot }
    Get-HVUser -HVUserLoginName $sam -HVDomain $dom -HVConnectionServer $HvServer
}

function Set-HvMachineUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Machine,
        [Parameter(Mandatory)]$User,
        [Parameter(Mandatory)]$HvServer
    )
    $machineService = [VMware.Hv.MachineService]::new()
    $machineHelper  = $machineService.read($HvServer.ExtensionData, $Machine.id)
    $machineHelper.getbasehelper().setuser($User.id)
    $machineService.update($HvServer.ExtensionData, $machineHelper)
}
#endregion

try {
    #region Connect to Horizon
    if (-not $Credential) { $Credential = Get-Credential -Message "Credentials for $ConnectionServer" }
    Write-Log "Connecting to $ConnectionServer…"
    $hvServer = Connect-HVServer -Server $ConnectionServer -Credential $Credential
    #endregion

    #region Import & validate CSV
    Write-Log "Importing CSV from $CsvPath…"
    $mapping   = Import-Csv -Path $CsvPath
    if ($mapping.Count -eq 0) { throw 'CSV contains no rows.' }
    $rowIndex  = 0; $totalRows = $mapping.Count; $results = @()
    #endregion

    foreach ($row in $mapping) {
        $rowIndex++
        Write-Progress -Activity 'Assigning desktops' -Status "Processing $rowIndex of $totalRows" -PercentComplete (($rowIndex/$totalRows)*100)
        $machineName = $row.MachineName.Trim(); $userString = $row.User.Trim(); $domain = $row.Domain
        try {
            Write-Log "Mapping $machineName → $userString" 'DEBUG'
            $machine = Get-HVMachine -MachineName $machineName -HvServer $hvServer
            if (-not $machine) { throw 'Machine not found' }
            $pool = Get-HVPoolSummary -PoolName $machine.base.desktopName -HvServer $hvServer
            if ($pool.userAssignment -ne 'DEDICATED') { throw "Pool is $($pool.userAssignment); only DEDICATED supported." }
            $user = Get-HvUserObject -User $userString -Domain $domain -HvServer $hvServer
            if (-not $user) { throw 'User not found' }
            if ($PSCmdlet.ShouldProcess($machineName, "assign to $userString")) { Set-HvMachineUser -Machine $machine -User $user -HvServer $hvServer }
            $status = 'Assigned'
        } catch { $status = "Failed – $($_.Exception.Message)"; Write-Log "$machineName → $userString : $status" 'ERROR' }
        $results += [pscustomobject]@{Machine=$machineName;User=$userString;Status=$status}
    }

    Write-Progress -Activity 'Assigning desktops' -Completed -Status 'Complete'
    Write-Log 'Assignment run complete.'
    $results | Sort-Object Status, Machine | Format-Table -AutoSize
}
finally {
    if ($hvServer) { Disconnect-HVServer -Server $hvServer -Confirm:$false }
    Stop-Transcript | Out-Null
}

# Signed‑off‑by: Marinus van Deventer
# End of script
