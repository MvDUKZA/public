$ErrorActionPreference = 'Stop'
<#
.SYNOPSIS
    Bulk‑assign dedicated VMware Horizon 8.12.1 desktops to users listed in a CSV file.

.DESCRIPTION
    Reads a CSV mapping file and assigns each Horizon "machine" (desktop VM) to the
    corresponding user.  Only dedicated‑assignment desktop pools are supported because
    user‑to‑machine binding is meaningless in floating pools.

    The script uses the official VMware PowerCLI Horizon modules (Omnissa.VimAutomation.HorizonView)
    together with Omnissa.Horizon.Helper, which wraps the View API for high‑level functions.

.PARAMETER CsvPath
    Full path of the CSV file containing the mappings.  Columns are:
        MachineName – Horizon machine name (as shown in the console)
        User        – sAMAccountName or user@domain format
        Domain      – (optional) NetBIOS or DNS domain name.  If omitted and the User field
                      contains a domain qualifier (e.g. ACME\jsmith or jsmith@acme.local)
                      that value is used; otherwise the script attempts to detect the
                      current AD forest’s DNS root.

.PARAMETER ConnectionServer
    FQDN of a Horizon Connection Server.  Defaults to the value of the HVConnectionServer
    environment variable, if present, otherwise the parameter is mandatory.

.PARAMETER Credential
    PSCredential used for Horizon API authentication.  If omitted you will be prompted.

.EXAMPLE
    PS C:\temp\scripts> .\Assign-HorizonDesktops.ps1 -CsvPath .\Mappings.csv -ConnectionServer view01.acme.local

.NOTES
    Author      : Powershell‑expert (ChatGPT)
    Created     : 30‑May‑2025
    Requirements:  ▪ PowerShell 5.1+ (Core also fine)
                   ▪ Omnissa PowerCLI 13.3+   →  Install‑Module VMware.PowerCLI ‑Scope AllUsers
                   ▪ Omnissa.Horizon.Helper   →  Install‑Module Omnissa.Horizon.Helper ‑Scope AllUsers
    Working dir : C:\temp\scripts
    Log file    : C:\temp\scripts\logs\Assign‑HorizonDesktop_yyyyMMdd_HHmmss.log
#>

#region Parameters & Validation
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({Test‑Path $_})]
    [string]$CsvPath,

    [Parameter()]
    [string]$ConnectionServer = $env:HVConnectionServer,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential
)
if (-not $ConnectionServer) { throw 'ConnectionServer is mandatory when $env:HVConnectionServer is not set.' }
#endregion

#region Initialisation
$script:WorkingDir = 'C:\temp\scripts'
$script:LogDir     = Join‑Path $script:WorkingDir 'logs'
if (-not (Test‑Path $LogDir)) { New‑Item -Path $LogDir -ItemType Directory -Force | Out‑Null }
$TimeStamp = Get‑Date -Format 'yyyyMMdd_HHmmss'
$LogFile   = Join‑Path $LogDir "Assign‑HorizonDesktop_$TimeStamp.log"
Start‑Transcript -Path $LogFile -Append | Out‑Null

function Write‑Log {
    param([string]$Message,[string]$Level = 'INFO')
    Write‑Verbose "[$((Get‑Date).ToString('u'))] [$Level] $Message"
}
#endregion

#region Module loading
Write‑Log 'Importing Omnissa PowerCLI modules…'
Import‑Module -Name VMware.VimAutomation.HorizonView -ErrorAction Stop
if (-not (Get‑Module -ListAvailable -Name Omnissa.Horizon.Helper)) {
    Install‑Module -Name Omnissa.Horizon.Helper -Force -AllowClobber -Scope AllUsers
}
Import‑Module -Name Omnissa.Horizon.Helper -ErrorAction Stop
#endregion

#region Helper functions
function Get‑HvUserObject {
    param(
        [string]$SamOrUpn,
        [string]$Domain,
        [object]$HvSrv
    )
    $login  = $SamOrUpn -replace '^(.*\\)|@.*$',''   # strip domain
    $domain = if ($Domain) { $Domain } elseif ($SamOrUpn -match '\\') {
        ($SamOrUpn -split '\\')[0]
    } elseif ($SamOrUpn -match '@') {
        ($SamOrUpn -split '@')[1]
    } else {
        (Get‑ADDomain).DNSRoot
    }
    Get‑HVUser -HVUserLoginName $login -HVDomain $domain -HVConnectionServer $HvSrv
}

function New‑HvDesktopAssignment {
    param(
        [object]$MachineObj,
        [object]$UserObj,
        [object]$HvSrv
    )
    $machineService  = New‑Object vmware.hv.machineservice
    $machineHelper   = $machineService.read($HvSrv.ExtensionData, $MachineObj.id)
    $machineHelper.getbasehelper().setuser($UserObj.id)
    $machineService.update($HvSrv.ExtensionData, $machineHelper)
}
#endregion

try {
    #region Connect to Horizon
    if (-not $Credential) { $Credential = Get‑Credential -Message "Credentials for $ConnectionServer" }
    Write‑Log "Connecting to $ConnectionServer…"
    $HvServer = Connect‑HVServer -Server $ConnectionServer -Credential $Credential
    #endregion

    #region Process CSV
    $Mapping = Import‑Csv -Path $CsvPath
    $Result  = foreach ($row in $Mapping) {
        $machineName = $row.MachineName
        $userString  = $row.User
        $domain      = $row.Domain

        Write‑Log "Processing mapping: $machineName ⇨ $userString" 'DEBUG'
        try {
            $machineObj = Get‑HVMachine -MachineName $machineName -HvServer $HvServer
            if (-not $machineObj) { throw "Machine not found" }

            $poolSummary = Get‑HVPoolSummary -PoolName $machineObj.base.desktopName -HvServer $HvServer
            if ($poolSummary.userAssignment -ne 'DEDICATED') { throw "Pool user assignment is $($poolSummary.userAssignment) – only DEDICATED supported" }

            $userObj = Get‑HvUserObject -SamOrUpn $userString -Domain $domain -HvSrv $HvServer
            if (-not $userObj) { throw "User not found" }

            if ($PSCmdlet.ShouldProcess("$machineName","assign to $userString")) {
                New‑HvDesktopAssignment -MachineObj $machineObj -UserObj $userObj -HvSrv $HvServer
            }
            [pscustomobject]@{Machine=$machineName;User=$userString;Status='Assigned'}
        }
        catch {
            Write‑Log "❌ $machineName ⇨ $userString : $_" 'ERROR'
            [pscustomobject]@{Machine=$machineName;User=$userString;Status="Failed – $_"}
        }
    }
    #endregion

    Write‑Log 'Assignment run complete.'
    $Result | Sort‑Object Status,Machine | Format‑Table -AutoSize
}
finally {
    if ($HvServer) { Disconnect‑HVServer -Server $HvServer -Confirm:$false }
    Stop‑Transcript | Out‑Null
}
