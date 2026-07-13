<#
    Audit-E1000.ps1
    ---------------
    Finds every VM with an E1000/E1000e vNIC and reports who is on it,
    using Horizon as the source of truth for sessions (fast - no guest
    credentials, no quser, no per-VM VIX call).

    Run it:
        .\Audit-E1000.ps1
        .\Audit-E1000.ps1 -vCenter vc01.iprod.local -ConnectionServer cs01.iprod.local

    Prompts once for credentials (used for both vCenter and Horizon).
    If Horizon is unreachable it still runs and returns the vCenter data
    with the session columns blank.
#>

param(
    [string] $vCenter          = 'vcenter.iprod.local',
    [string] $ConnectionServer = 'horizon-cs01.iprod.local',
    [string] $OutCsv           = "C:\Temp\E1000_Audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv",
    [switch] $SkipHorizon
)

Import-Module VMware.PowerCLI -ErrorAction Stop
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null

$cred = Get-Credential -Message "Credentials for vCenter and Horizon (DOMAIN\user)"

# ---------------------------------------------------------------- vCenter
Write-Host "Connecting to $vCenter ..." -ForegroundColor Cyan
Connect-VIServer -Server $vCenter -Credential $cred -ErrorAction Stop | Out-Null

# ---------------------------------------------------------------- Horizon
$hvSessions  = @{}
$hvConnected = $false

if (-not $SkipHorizon) {
    try {
        Write-Host "Connecting to $ConnectionServer ..." -ForegroundColor Cyan
        $hv  = Connect-HVServer -Server $ConnectionServer -Credential $cred -ErrorAction Stop
        $api = $hv.ExtensionData
        $hvConnected = $true

        Write-Host "Querying Horizon machines ..." -ForegroundColor Cyan
        $qs = New-Object VMware.Hv.QueryServiceService
        $qd = New-Object VMware.Hv.QueryDefinition
        $qd.QueryEntityType = 'MachineNamesView'
        $qd.Limit = 1000

        $r = $qs.QueryService_Create($api, $qd)
        while ($r.Results) {
            foreach ($m in $r.Results) {
                $key  = "$($m.Base.Name)".ToUpper()
                $pool = ''; $user = ''
                try { $pool = "$($m.NamesData.DesktopName)" } catch {}
                try { $user = "$($m.NamesData.UserName)"    } catch {}
                $hvSessions[$key] = [pscustomobject]@{
                    Pool         = $pool
                    HorizonState = "$($m.Base.BasicState)"
                    AssignedUser = $user
                }
            }
            if (-not $r.Id) { break }
            $r = $qs.QueryService_GetNext($api, $r.Id)
        }
        if ($r.Id) { $qs.QueryService_Delete($api, $r.Id) }

        Write-Host "Horizon machines: $($hvSessions.Count)" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Horizon unavailable - session columns will be blank. ($($_.Exception.Message))"
    }
}

# ---------------------------------------------------------------- scan
Write-Host "Caching portgroups ..." -ForegroundColor Cyan
$dvpg = @{}
Get-View -ViewType DistributedVirtualPortgroup -Property Name,Key | ForEach-Object { $dvpg[$_.Key] = $_.Name }

Write-Host "Scanning VMs ..." -ForegroundColor Cyan
$views = Get-View -ViewType VirtualMachine -Property `
    Name,Config.Hardware.Device,Config.Version,Config.GuestFullName,Config.Template,
    Runtime.PowerState,Guest.ToolsRunningStatus,Guest.ToolsVersionStatus,Guest.IpAddress,Guest.HostName

$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($v in $views) {

    $i++
    if ($i % 100 -eq 0) {
        Write-Progress -Activity 'Scanning' -Status "$i / $($views.Count)" -PercentComplete (($i / $views.Count) * 100)
    }

    # never touch templates or Horizon internals
    if ($v.Config.Template) { continue }
    if ($v.Name -match '^(cp-parent-|cp-replica-|cp-template-|ClonePrep)') { continue }

    $allNics = $v.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualEthernetCard] }
    $badNics = $allNics | Where-Object { $_ -is [VMware.Vim.VirtualE1000] -or $_ -is [VMware.Vim.VirtualE1000e] }
    if (-not $badNics) { continue }

    $hz = $hvSessions["$($v.Name)".ToUpper()]

    # DISCONNECTED still has a live session - only AVAILABLE/MAINTENANCE are free
    $safe = if     (-not $hvConnected)                   { 'UNKNOWN' }
            elseif (-not $hz)                            { 'NOT-IN-HORIZON' }
            elseif ($hz.HorizonState -eq 'AVAILABLE')    { 'YES' }
            elseif ($hz.HorizonState -eq 'MAINTENANCE')  { 'YES' }
            else                                         { 'NO' }

    foreach ($nic in $badNics) {

        if ($nic.Backing -is [VMware.Vim.VirtualEthernetCardDistributedVirtualPortBackingInfo]) {
            $pgName = $dvpg[$nic.Backing.Port.PortgroupKey]; $pgType = 'DVS'
        } else {
            $pgName = $nic.Backing.DeviceName;               $pgType = 'Standard'
        }

        $results.Add([pscustomobject]@{
            VMName         = $v.Name
            Pool           = if ($hz) { $hz.Pool }         else { '' }
            HorizonState   = if ($hz) { $hz.HorizonState } else { '' }
            AssignedUser   = if ($hz) { $hz.AssignedUser } else { '' }
            SafeToConvert  = $safe
            PowerState     = $v.Runtime.PowerState
            GuestOS        = $v.Config.GuestFullName
            GuestIP        = $v.Guest.IpAddress
            HWVersion      = $v.Config.Version
            ToolsRunning   = $v.Guest.ToolsRunningStatus
            ToolsVersion   = $v.Guest.ToolsVersionStatus
            NicLabel       = $nic.DeviceInfo.Label
            NicType        = ($nic.GetType().Name -replace '^Virtual','')
            MacAddress     = $nic.MacAddress
            PortGroup      = $pgName
            PortGroupType  = $pgType
            Connected      = $nic.Connectable.Connected
            TotalNicsOnVM  = $allNics.Count
            MoRef          = $v.MoRef.ToString()
        })
    }
}
Write-Progress -Activity 'Scanning' -Completed

# ---------------------------------------------------------------- output
$dir = Split-Path $OutCsv -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$results | Sort-Object Pool,VMName | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8

$vms = $results | Select-Object -Unique VMName

Write-Host ""
Write-Host "VMs with E1000/E1000e : $($vms.Count)" -ForegroundColor Yellow
Write-Host "Adapters              : $($results.Count)" -ForegroundColor Yellow
Write-Host "Safe to convert now   : $(($results | Where-Object SafeToConvert -eq 'YES' | Select-Object -Unique VMName).Count)" -ForegroundColor Green
Write-Host "In use (skip)         : $(($results | Where-Object SafeToConvert -eq 'NO'  | Select-Object -Unique VMName).Count)" -ForegroundColor Red
Write-Host "Multi-NIC (manual)    : $(($results | Where-Object TotalNicsOnVM -gt 1     | Select-Object -Unique VMName).Count)" -ForegroundColor Magenta
Write-Host ""
Write-Host "By pool:" -ForegroundColor Cyan
$results | Group-Object Pool | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("  {0,-32} {1}" -f $(if ($_.Name) { $_.Name } else { '<not in Horizon>' }), $_.Count)
}
Write-Host ""
Write-Host "CSV: $OutCsv" -ForegroundColor Green

if ($hvConnected) { Disconnect-HVServer -Server $ConnectionServer -Confirm:$false }
