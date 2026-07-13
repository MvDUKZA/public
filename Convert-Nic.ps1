<#
    Convert-Nic.ps1
    ---------------
    Swaps E1000/E1000e -> VMXNET3 on Horizon full-clone desktops.

        .\Convert-Nic.ps1 -VMName VDI-W11-0042
        .\Convert-Nic.ps1 -VMName VDI-W11-0042 -Execute
        .\Convert-Nic.ps1 -InputCsv C:\Temp\batch01.csv -Execute

    No guest credentials. No rollback machinery. Verification is Horizon's own
    health check.

    WHY THERE IS NO ROLLBACK
    The VMXNET3 is added BEFORE the E1000 is removed. If the add fails, nothing
    changed. If the remove fails, the VM has two adapters and still boots and
    works - it is logged as BOTH-NICS and you clean it up later. There is no
    sequence that leaves a VM without a network adapter, so there is nothing to
    roll back from.

    WHY THERE IS NO GUEST SCRIPT
    A desktop whose NIC did not come up properly does not reach AVAILABLE in
    Horizon - it sits at AGENT_UNREACHABLE. That is the verification, and the
    broker does it for us. The leftover ghost E1000 in the guest is cosmetic on
    DHCP; sweep it fleet-wide from MECM afterwards, which is the right tool for it.

    STATUSES
    CONVERTED        swapped, booted, Horizon says AVAILABLE, back in the pool
    HELD             swapped but Horizon never said AVAILABLE - confirmed back in maintenance
    FAILED-NOT-HELD  unhealthy AND could not be put back in maintenance - STILL BOOKABLE, fix now
    BOTH-NICS        add worked, remove did not - VM powered on but health not verified; left in maintenance
    FAILED           see Detail - LEFT IN MAINTENANCE
    SKIPPED          in use, or not a single-E1000 VM
    ALREADY          already a single VMXNET3
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory)] [string[]] $VMName,
    [Parameter(ParameterSetName = 'ByCsv',  Mandatory)] [string]   $InputCsv,

    [string] $vCenter          = 'vcenter.iprod.local',
    [string] $ConnectionServer = 'horizon-cs01.iprod.local',
    [string] $CsvColumn        = 'VMName',
    [int]    $Limit            = 0,
    [switch] $Execute,

    [int]    $MaintWaitSec     = 90,
    [int]    $PowerWaitSec     = 180,
    [int]    $ShutdownWaitSec  = 180,
    [int]    $HealthyWaitSec   = 420    # how long to give Horizon to say AVAILABLE
)

$ErrorActionPreference = 'Stop'
$DryRun = -not $Execute
$LogCsv = "C:\Temp\NicSwap_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# targets
if ($PSCmdlet.ParameterSetName -eq 'ByCsv') {
    if (-not (Test-Path $InputCsv)) { throw "CSV not found: $InputCsv" }
    $csv = @(Import-Csv $InputCsv)
    if ($csv.Count -eq 0) { throw "CSV is empty: $InputCsv" }
    if ($csv[0].PSObject.Properties.Name -notcontains $CsvColumn) {
        throw "CSV has no '$CsvColumn' column. Columns present: $(($csv[0].PSObject.Properties.Name) -join ', ')"
    }
    $targets = @($csv.$CsvColumn | Where-Object { $_ } | Select-Object -Unique)
} else {
    $targets = @($VMName | Where-Object { $_ } | Select-Object -Unique)
}
if ($targets.Count -eq 0) { throw 'No targets to work on.' }
if ($Limit -gt 0) { $targets = $targets | Select-Object -First $Limit }

Import-Module VMware.PowerCLI -ErrorAction Stop
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null

Write-Host ""
if ($DryRun) { Write-Host "DRY RUN - add -Execute to commit" -ForegroundColor Magenta }
else         { Write-Host "EXECUTE - changes will be made" -ForegroundColor Red }
Write-Host "$($targets.Count) target(s):" -ForegroundColor Gray
$targets | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
Write-Host ""

if (-not $DryRun -and (Read-Host "Type YES to proceed") -ne 'YES') { return }

$cred   = Get-Credential -Message "vCenter and Horizon credentials (DOMAIN\user)"
$viConn = $null; $hvConn = $null

try {
    $viConn = Connect-VIServer  -Server $vCenter          -Credential $cred
    $hvConn = Connect-HVServer  -Server $ConnectionServer -Credential $cred
    $api    = $hvConn.ExtensionData

    New-Item -ItemType Directory -Path (Split-Path $LogCsv) -Force -EA SilentlyContinue | Out-Null

    function Log {
        param($Machine,$Status,$Detail)
        [pscustomobject]@{ Time=(Get-Date -f 's'); Machine=$Machine; Status=$Status; Detail=$Detail } |
            Export-Csv $LogCsv -NoTypeInformation -Append
    }
    function HvState { param($Id) try { "$($api.Machine.Machine_Get($Id).Base.BasicState)" } catch { 'UNKNOWN' } }
    function WaitPower {
        param($Name,$Want,$Secs)
        $end = (Get-Date).AddSeconds($Secs)
        while ((Get-Date) -lt $end) {
            if ((Get-VM $Name).PowerState -eq $Want) { return $true }
            Start-Sleep 5
        }
        return ((Get-VM $Name).PowerState -eq $Want)
    }

    # index Horizon machines
    $qs = New-Object VMware.Hv.QueryServiceService
    $qd = New-Object VMware.Hv.QueryDefinition
    $qd.QueryEntityType = 'MachineNamesView'; $qd.Limit = 1000
    $hvIndex = @{}
    $r = $qs.QueryService_Create($api, $qd)
    while ($r.Results) {
        foreach ($m in $r.Results) { $hvIndex["$($m.Base.Name)".ToUpper()] = $m.Id }
        if (-not $r.Id) { break }
        $r = $qs.QueryService_GetNext($api, $r.Id)
    }
    if ($r.Id) { $qs.QueryService_Delete($api, $r.Id) }

    foreach ($name in $targets) {

        $id = $hvIndex["$name".ToUpper()]
        if (-not $id) { Write-Host "$name : not in Horizon" -F Yellow; Log $name 'SKIPPED' 'Not in Horizon'; continue }

        $vm   = @(Get-VM -Name $name -EA SilentlyContinue)
        if ($vm.Count -ne 1) { Log $name 'SKIPPED' "vCenter returned $($vm.Count) VMs for this name"; continue }
        $vm   = $vm[0]

        $nics = @(Get-NetworkAdapter -VM $vm)
        $old  = @($nics | Where-Object { $_.Type -in 'e1000','e1000e' })

        if ($nics.Count -eq 1 -and $nics[0].Type -eq 'Vmxnet3') {
            Write-Host "$name : already vmxnet3" -F DarkGray; Log $name 'ALREADY' ''; continue
        }
        if ($nics.Count -ne 1 -or $old.Count -ne 1) {
            $t = ($nics | % Type) -join ';'
            Write-Host "$name : SKIP ($($nics.Count) adapters: $t)" -F Yellow
            Log $name 'SKIPPED' "$($nics.Count) adapters: $t"; continue
        }

        $pg = $old[0].NetworkName

        if ($DryRun) {
            Write-Host "$name : would swap $($old[0].Type) -> vmxnet3 on '$pg'  [state $(HvState $id)]" -F Magenta
            Log $name 'DRYRUN' "$($old[0].Type) on '$pg'"; continue
        }

        Write-Host "=== $name  $($old[0].Type) on '$pg'" -F Cyan
        $maint = $false

        try {
            # nobody on it?
            $st = HvState $id
            if ($st -ne 'AVAILABLE') {
                Write-Host "    SKIP - state $st" -F Yellow; Log $name 'SKIPPED' "State $st"; continue
            }

            # lock it, and confirm the lock actually took
            $api.Machine.Machine_EnterMaintenanceMode($id)
            $maint = $true
            $end = (Get-Date).AddSeconds($MaintWaitSec)
            do { Start-Sleep 5; $st = HvState $id } until ($st -eq 'MAINTENANCE' -or (Get-Date) -gt $end)
            if ($st -ne 'MAINTENANCE') {
                $api.Machine.Machine_ExitMaintenanceMode($id)
                Write-Host "    SKIP - maintenance not confirmed ($st)" -F Yellow
                Log $name 'SKIPPED' "Maintenance not confirmed ($st)"; continue
            }
            Write-Host "    maintenance on" -F DarkGray

            # power off, confirmed
            if ((Get-VM $name).PowerState -eq 'PoweredOn') {
                Shutdown-VMGuest -VM (Get-VM $name) -Confirm:$false -EA SilentlyContinue | Out-Null
                if (-not (WaitPower $name 'PoweredOff' $ShutdownWaitSec)) {
                    Stop-VM -VM (Get-VM $name) -Confirm:$false | Out-Null
                    if (-not (WaitPower $name 'PoweredOff' $PowerWaitSec)) {
                        throw 'Would not power off - not touching the hardware'
                    }
                }
            }
            Write-Host "    powered off" -F DarkGray

            # ADD first. If this throws, nothing has changed.
            New-NetworkAdapter -VM (Get-VM $name) -NetworkName $pg -Type Vmxnet3 `
                -StartConnected:$true -Confirm:$false | Out-Null
            Write-Host "    vmxnet3 added" -F DarkGray

            # then remove. If THIS throws, the VM has both and still works.
            try {
                Get-NetworkAdapter -VM (Get-VM $name) |
                    Where-Object { $_.Type -in 'e1000','e1000e' } |
                    Remove-NetworkAdapter -Confirm:$false
            }
            catch {
                $rmErr = $_.Exception.Message -replace "`r?`n",' '
                Start-VM -VM (Get-VM $name) -Confirm:$false -EA SilentlyContinue | Out-Null
                if (WaitPower $name 'PoweredOn' $PowerWaitSec) {
                    Write-Host "    BOTH-NICS - VM powered on with both adapters. Left in maintenance." -F Yellow
                    Log $name 'BOTH-NICS' "E1000 removal failed: $rmErr | VM powered on but health not verified"
                } else {
                    Write-Host "    FAILED - E1000 removal failed and the VM did not power on" -F Red
                    Log $name 'FAILED' "E1000 removal failed: $rmErr | VM did not power back on"
                }
                continue    # left in maintenance either way
            }
            Write-Host "    e1000 removed" -F DarkGray

            # verify what we actually ended up with before powering on
            $after = @(Get-NetworkAdapter -VM (Get-VM $name))
            if ($after.Count -ne 1)                              { throw "After swap: $($after.Count) adapters, expected 1" }
            if ($after[0].Type -ne 'Vmxnet3')                    { throw "After swap: adapter is $($after[0].Type), expected Vmxnet3" }
            if ($after[0].NetworkName -ne $pg)                   { throw "After swap: on '$($after[0].NetworkName)', expected '$pg'" }
            if (-not $after[0].ConnectionState.StartConnected)   { throw 'After swap: adapter is not set to connect at power on' }

            # power on, confirmed
            Start-VM -VM (Get-VM $name) -Confirm:$false | Out-Null
            if (-not (WaitPower $name 'PoweredOn' $PowerWaitSec)) { throw 'Did not power back on' }

            # Horizon is the health check. A desktop with a broken NIC never
            # reaches AVAILABLE - it sits at AGENT_UNREACHABLE.
            $api.Machine.Machine_ExitMaintenanceMode($id)
            $maint = $false

            $end = (Get-Date).AddSeconds($HealthyWaitSec)
            do { Start-Sleep 15; $st = HvState $id } until ($st -eq 'AVAILABLE' -or (Get-Date) -gt $end)

            if ($st -eq 'AVAILABLE') {
                Write-Host "    CONVERTED (Horizon: AVAILABLE)" -F Green
                Log $name 'CONVERTED' 'Horizon reports AVAILABLE'
            }
            else {
                # Put it back in maintenance - and PROVE it went back, otherwise a
                # broken desktop is sitting in the pool waiting for a user.
                $api.Machine.Machine_EnterMaintenanceMode($id)
                $end = (Get-Date).AddSeconds($MaintWaitSec)
                do { Start-Sleep 5; $m = HvState $id } until ($m -eq 'MAINTENANCE' -or (Get-Date) -gt $end)

                if ($m -eq 'MAINTENANCE') {
                    Write-Host "    HELD - Horizon says $st, put back in maintenance" -F Red
                    Log $name 'HELD' "Horizon state $st after swap - agent not healthy. Held in maintenance."
                } else {
                    Write-Host "    FAILED-NOT-HELD - unhealthy AND STILL BOOKABLE. Fix now." -F Red
                    Log $name 'FAILED-NOT-HELD' "Horizon state $st after swap and could not re-enter maintenance (state=$m). THIS DESKTOP IS STILL IN THE POOL."
                }
            }
        }
        catch {
            $e = $_.Exception.Message -replace "`r?`n",' '
            Write-Host "    FAILED: $e" -F Red

            # FAILED must mean "confirmed parked". Verify - and if it is not in
            # maintenance, try to put it there and verify again. Only if that
            # also fails is it FAILED-NOT-HELD: broken and still bookable.
            $cur = HvState $id
            if ($cur -ne 'MAINTENANCE') {
                try {
                    $api.Machine.Machine_EnterMaintenanceMode($id)
                    $end = (Get-Date).AddSeconds($MaintWaitSec)
                    do { Start-Sleep 5; $cur = HvState $id } until ($cur -eq 'MAINTENANCE' -or (Get-Date) -gt $end)
                } catch { }
            }

            if ($cur -eq 'MAINTENANCE') {
                Log $name 'FAILED' "$e | Confirmed in maintenance"
            } else {
                Write-Host "    FAILED-NOT-HELD - failed AND STILL BOOKABLE (state=$cur). Fix now." -F Red
                Log $name 'FAILED-NOT-HELD' "$e | Could not confirm maintenance (state=$cur). THIS DESKTOP IS STILL IN THE POOL."
            }
        }
    }

    Write-Host ""
    Import-Csv $LogCsv | Group-Object Status | Sort-Object Name |
        ForEach-Object { Write-Host ("{0,-12} {1}" -f $_.Name, $_.Count) }
    Write-Host ""
    Write-Host "Log: $LogCsv" -F Green

    $stuck = Import-Csv $LogCsv | Where-Object Status -in 'FAILED','HELD','BOTH-NICS','FAILED-NOT-HELD'
    if ($stuck) {
        Write-Host ""
        Write-Host "NEEDS ATTENTION:" -F Red
        $stuck | ForEach-Object { Write-Host ("  {0,-20} {1,-16} {2}" -f $_.Machine, $_.Status, $_.Detail) }
    }

    $loose = Import-Csv $LogCsv | Where-Object Status -eq 'FAILED-NOT-HELD'
    if ($loose) {
        Write-Host ""
        Write-Host "*** UNHEALTHY AND STILL BOOKABLE - USERS CAN GET THESE ***" -F Red
        $loose | ForEach-Object { Write-Host "  $($_.Machine)" -F Red }
    }
}
finally {
    if ($hvConn) { Disconnect-HVServer -Server $ConnectionServer -Confirm:$false -EA SilentlyContinue }
    if ($viConn) { Disconnect-VIServer -Server $vCenter -Confirm:$false -EA SilentlyContinue }
}
