#Requires -Version 5.1
#Requires -Modules VMware.VimAutomation.HorizonView, Omnissa.Horizon.Helper, ActiveDirectory

<#!
.SYNOPSIS
    Bulk‑assign dedicated VMware Horizon 8.12.1 desktops to users specified in a CSV file.

.DESCRIPTION
    Reads a CSV that maps Horizon desktop machines to Active Directory users and assigns
    each machine accordingly using the Horizon REST / View API.  The script supports only
    *dedicated‑assignment* pools because floating pools cannot persist a user→VM binding.

    It relies solely on the official Omnissa / VMware PowerCLI modules – no ControlUp
    components are required.

.PARAMETER CsvPath
    Path to a CSV containing at least the columns MachineName and User.  An optional
    Domain column may be supplied when User is not fully qualified.

.PARAMETER ConnectionServer
    FQDN of a Horizon Connection Server.  If omitted the script tries the environment
    variable HVConnectionServer.

.PARAMETER Credential
    PSCredential for authenticating to Horizon.  If omitted you will be prompted.

.EXAMPLE
    PS C:\temp\scripts> .\Assign‑HorizonDesktops.ps1 `
                          ‑CsvPath .\Assignments.csv `
                          ‑ConnectionServer view01.iprod.local ‑Verbose

.NOTES
    Author   : Marinus van Deventer
    Created  : 30‑May‑2025
    Version  : 2.0
    Log File : C:\temp\scripts\logs\Assign‑HorizonDesktop_yyyyMMdd_HHmmss.log

    Change‑log
        2.0 – Rewritten from scratch to meet coding standards v2025‑05.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, HelpMessage = 'Full path to the CSV mapping file')]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter(HelpMessage = 'Horizon Connection Server FQDN')]
    [ValidateNotNullOrEmpty()]
    [string]$ConnectionServer = $env:HVConnectionServer,

    [Parameter(HelpMessage = 'Credential used to authenticate to Horizon')]
    [System.Management.Automation.PSCredential]$Credential
)

# Abort early if ConnectionServer could not be resolved
if (-not $ConnectionServer) {
    throw 'ConnectionServer is mandatory when the HVConnectionServer environment variable is not set.'
}

$ErrorActionPreference = 'Stop'

#region Paths & logging setup
$workingDir = 'C:\temp\scripts'
$logDir     = Join-Path $workingDir 'logs'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$timeStamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile    = Join-Path $logDir "Assign‑HorizonDesktop_$timeStamp.log"
Start-Transcript -Path $logFile -Append | Out-Null
#endregion Paths

#region Utility functions
function Write‑Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $prefix = "$(Get-Date -UFormat '%Y-%m-%d %H:%M:%S') - $Level :"
    Write-Information "$prefix $Message" -InformationAction Continue
}

function Get‑HvUserObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter()][string]$Domain,
        [Parameter(Mandatory)]$HvServer
    )
    $userName = $User -replace '^(.*\\)|@.*$',''  # strip any domain
    $resolvedDomain = if ($Domain) {
        $Domain
    } elseif ($User -match '\\') {
        ($User -split '\\')[0]
    } elseif ($User -match '@') {
        ($User -split '@')[1]
    } else {
        (Get‑ADDomain).DNSRoot
    }
    Get‑HVUser -HVUserLoginName $userName -HVDomain $resolvedDomain -HVConnectionServer $HvServer
}

function Set‑HvMachineUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Machine,
        [Parameter(Mandatory)][object]$User,
        [Parameter(Mandatory)]$HvServer
    )
    $svc   = [VMware.Hv.MachineService]::new()
    $helper = $svc.read($HvServer.ExtensionData, $Machine.id)
    $helper.getbasehelper().setuser($User.id)
    $svc.update($HvServer.ExtensionData, $helper)
}
#endregion Utility

try {
    #region Module loading
    Write‑Log 'Importing VMware PowerCLI modules…'
    Import‑Module VMware.VimAutomation.HorizonView -ErrorAction Stop
    if (-not (Get‑Module -ListAvailable Omnissa.Horizon.Helper)) {
        Install‑Module Omnissa.Horizon.Helper -Scope AllUsers -Force -AllowClobber
    }
    Import‑Module Omnissa.Horizon.Helper -ErrorAction Stop
    #endregion

    #region Horizon connection
    if (-not $Credential) {
        $Credential = Get‑Credential -Message "Credentials for $ConnectionServer"
    }
    Write‑Log "Connecting to $ConnectionServer…"
    $hvServer = Connect‑HVServer -Server $ConnectionServer -Credential $Credential
    #endregion

    #region CSV import
    Write‑Log "Reading assignments from $CsvPath…"
    $csv = Import‑Csv -Path $CsvPath | ForEach‑Object {
        # Trim whitespace in every field
        $_ | Select‑Object @{N='MachineName';E={$_.MachineName.Trim()}},
                           @{N='User';E={$_.User.Trim()}},
                           @{N='Domain';E={($_.Domain -as [string]).Trim()}}
    }

    if ($csv.Count -eq 0) { throw 'CSV is empty.' }

    foreach ($required in 'MachineName','User') {
        if (-not ($csv[0].psobject.Properties.Name -contains $required)) {
            throw "CSV missing required column: $required"
        }
    }
    #endregion

    #region Assignment loop
    $index = 0
    $total = $csv.Count
    $results = @()

    foreach ($row in $csv) {
        $index++
        Write‑Progress -Activity 'Assigning desktops' -Status "Processing $index of $total" -PercentComplete (($index/$total)*100)

        if ([string]::IsNullOrWhiteSpace($row.MachineName) -or [string]::IsNullOrWhiteSpace($row.User)) {
            Write‑Warning "Incomplete row detected (index $index). Skipping."
            continue
        }

        Write‑Log "Attempting assignment: $($row.MachineName) → $($row.User)" 'DEBUG'
        try {
            $machine = Get‑HVMachine -MachineName $row.MachineName -HvServer $hvServer
            if (-not $machine) { throw 'Machine not found in Horizon' }

            $poolSummary = Get‑HVPoolSummary -PoolName $machine.base.desktopName -HvServer $hvServer
            if ($poolSummary.userAssignment -ne 'DEDICATED') {
                throw "Pool $($poolSummary.id) is $($poolSummary.userAssignment) – dedicated pools only."
            }

            $user = Get‑HvUserObject -User $row.User -Domain $row.Domain -HvServer $hvServer
            if (-not $user) { throw 'User not found in Horizon' }

            if ($PSCmdlet.ShouldProcess($row.MachineName,"assign to $($row.User)")) {
                Set‑HvMachineUser -Machine $machine -User $user -HvServer $hvServer
            }
            $state = 'Assigned'
        }
        catch {
            $state = "Failed – $($_.Exception.Message)"
            Write‑Log "$($row.MachineName) → $($row.User) : $state" 'ERROR'
        }

        $results += [pscustomobject]@{
            Machine = $row.MachineName
            User    = $row.User
            Status  = $state
        }
    }

    Write‑Progress -Activity 'Assigning desktops' -Completed -Status 'Done'
    Write‑Log 'Assignment process finished.'
    $results | Sort‑Object Status, Machine | Format‑Table -AutoSize
    #endregion Assignment
}
finally {
    if ($hvServer) { Disconnect‑HVServer -Server $hvServer -Confirm:$false }
    Stop‑Transcript | Out‑Null
}

# Signed‑off‑by: Marinus van Deventer
# End of script
