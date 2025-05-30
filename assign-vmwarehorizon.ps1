#Requires -Version 5.1
#Requires -Modules VMware.VimAutomation.HorizonView

<#
.SYNOPSIS
    Bulk-assign dedicated VMware Horizon 8.12.1 desktops to users defined in a CSV file.

.DESCRIPTION
    Reads a CSV mapping (MachineName, User[, Domain]) and assigns each Horizon desktop VM
    to the specified user. Only dedicated‑assignment pools are supported. The script uses
    the raw Horizon View API classes exposed by the PowerCLI module. In Horizon 8.12 the
    namespaces were renamed from **VMware.Hv.*** to **Omnissa.Horizon.***. If those API
    proxy classes are not present, the script can dynamically import the community module
    **Omnissa.Horizon.Helper** which generates them at runtime.

.PARAMETER CsvPath
    Path to the CSV file. Required columns: MachineName, User. Optional: Domain.

.PARAMETER ConnectionServer
    FQDN of a Horizon Connection Server. Falls back to $env:HVConnectionServer.

.PARAMETER Credential
    PSCredential for authenticating to Horizon. Prompted if omitted.

.EXAMPLE
    .\Assign-HorizonDesktops.ps1 -CsvPath .\Assignments.csv -ConnectionServer view01.iprod.local -Verbose

.NOTES
    Author      : Marinus van Deventer
    Version     : 2.5
    Created     : 30‑May‑2025
    Change‑log  :
        2.5 – Added dynamic fallback to Omnissa.Horizon.Helper when core API proxy types
              are missing (fixes QueryServiceService type error).
        2.4 – Dependency trimmed to core PowerCLI; helper removed.
        2.3 – Fixed divide‑by‑zero progress calculation.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)][ValidateScript({ Test-Path $_ })][string]$CsvPath,
    [Parameter()][ValidateNotNullOrEmpty()][string]$ConnectionServer = $env:HVConnectionServer,
    [Parameter()][System.Management.Automation.PSCredential]$Credential
)
if (-not $ConnectionServer) { throw 'ConnectionServer is mandatory when $env:HVConnectionServer is not set.' }
$ErrorActionPreference = 'Stop'

#region Logging setup
$workingDir = 'C:\temp\scripts'
$logDir = Join-Path $workingDir 'logs'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$timeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Start-Transcript -Path (Join-Path $logDir "Assign-HorizonDesktop_$timeStamp.log") -Append | Out-Null
#endregion

#region Helper functions
function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level='INFO')
    Write-Information "$(Get-Date -UFormat '%Y-%m-%d %H:%M:%S') - $Level : $Message" -InformationAction Continue
}
#endregion

try {
    #region Module loading
    Write-Log 'Importing VMware PowerCLI Horizon module...'
    Import-Module VMware.VimAutomation.HorizonView -ErrorAction Stop
    #endregion

    #region Horizon connection
    if (-not $Credential) { $Credential = Get-Credential -Message "Credentials for $ConnectionServer" }
    Write-Log "Connecting to $ConnectionServer..."
    $hvServer = Connect-HVServer -Server $ConnectionServer -Credential $Credential
    #endregion

    #region Ensure API proxy types are present
    $proxyType = [type]::GetType('Omnissa.Horizon.QueryServiceService', $false)
    if (-not $proxyType) {
        Write-Log 'Omnissa.Horizon proxy types not found – attempting to load Omnissa.Horizon.Helper...' 'WARN'
        try {
            Import-Module Omnissa.Horizon.Helper -ErrorAction Stop
            $proxyType = [type]::GetType('Omnissa.Horizon.QueryServiceService', $false)
            if ($proxyType) {
                Write-Log 'Omnissa.Horizon.Helper imported successfully.' 'INFO'
            } else {
                Write-Log 'Proxy types still missing after helper load.' 'ERROR'
                throw 'Required Horizon API proxy types are not available. Ensure PowerCLI 13.3+ or install Omnissa.Horizon.Helper.'
            }
        }
        catch {
            throw "Failed to import Omnissa.Horizon.Helper: $($_.Exception.Message)"
        }
    }
    #endregion

    #region CSV import & validation (unchanged core logic)
    Write-Log "Reading CSV $CsvPath..."
    $rows = Import-Csv -Path $CsvPath
    if (-not $rows) { throw 'CSV is empty or header only.' }
    foreach ($col in 'MachineName','User') {
        if (-not ($rows[0].psobject.Properties.Name -contains $col)) { throw "CSV missing column: $col" }
    }
    #endregion

    #region Assignment helpers using namespace‑agnostic resolution
    function Get-HvUserObject {
        param([string]$User,[string]$Domain,$HvConn)
        $userName = $User -replace '^(.*\\)|@.*$',''
        $resolvedDomain = if ($Domain) { $Domain } elseif ($User -match '\\') {
            ($User -split '\\')[0]
        } elseif ($User -match '@') {
            ($User -split '@')[1]
        } else {
            if (Get-Command Get-ADDomain -ErrorAction SilentlyContinue) { (Get-ADDomain).DNSRoot } else { $env:USERDNSDOMAIN }
        }
        Get-HVUser -HVUserLoginName $userName -HVDomain $resolvedDomain -HVConnectionServer $HvConn
    }
    function Set-HvMachineUser {
        param($Machine,$User,$HvConn)
        # Resolve namespace dynamically
        $ns = ([type]::GetType('Omnissa.Horizon.MachineService',$false) -as [type])
        if (-not $ns) { $ns = [type]::GetType('VMware.Hv.MachineService',$true) }
        $svc   = [Activator]::CreateInstance($ns)
        $helper = $svc.read($HvConn.ExtensionData,$Machine.id)
        $helper.getbasehelper().setuser($User.id)
        $svc.update($HvConn.ExtensionData,$helper)
    }
    #endregion

    #region Processing loop (same as previous logic but using functions above) ...
    $total = $rows.Count; $i=0; $result=@()
    foreach ($r in $rows) {
        $i++
        Write-Progress -Activity 'Assigning desktops' -Status "Processing $i of $total" -PercentComplete ([math]::Round($i/$total*100))
        if ([string]::IsNullOrWhiteSpace($r.MachineName) -or [string]::IsNullOrWhiteSpace($r.User)) {
            Write-Log "Row $i incomplete – skipped." 'WARN'; continue
        }
        try {
            $machine = Get-HVMachine -MachineName $r.MachineName -HvServer $hvServer
            if (-not $machine) { throw 'Machine not found' }
            $pool = Get-HVPoolSummary -PoolName $machine.base.desktopName -HvServer $hvServer
            if ($pool.userAssignment -ne 'DEDICATED') { throw 'Pool not dedicated' }
            $user = Get-HvUserObject -User $r.User -Domain $r.Domain -HvConn $hvServer
            if (-not $user) { throw 'User not found' }
            if ($PSCmdlet.ShouldProcess($r.MachineName,"assign to $($r.User)")) { Set-HvMachineUser $machine $user $hvServer }
            $status='Assigned'
        } catch { $status = "Failed - $($_.Exception.Message)"; Write-Log "$($r.MachineName)->$($r.User): $status" 'ERROR' }
        $result += [pscustomobject]@{Machine=$r.MachineName;User=$r.User;Status=$status}
    }
    Write-Progress -Activity 'Assigning desktops' -Completed -Status 'Done'
    Write-Log 'Processing finished.'
    $result | Sort-Object Status,Machine | Format-Table -AutoSize
    #endregion
}
finally {
    if ($hvServer) { Disconnect-HVServer -Server $hvServer -Confirm:$false }
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

# Signed-off-by: Marinus van Deventer
# End of script
