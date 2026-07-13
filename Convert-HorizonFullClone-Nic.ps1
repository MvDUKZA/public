<#
    Convert-HorizonFullClone-Nic.ps1
    --------------------------------
    Replaces the E1000/E1000e vNIC with a VMXNET3 on Horizon FULL CLONE desktops.
    That is the whole job.

        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042 -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\batch01.csv -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\audit.csv -Limit 25 -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\halfswapped.csv -RepairHalfSwapped -Execute

    FULL CLONES ONLY
    ----------------
    Every machine's pool is checked for full-clone provisioning (Horizon source
    VIRTUAL_CENTER). Instant-clone and linked-clone pools are REJECTED, not
    converted - their vNIC comes from the golden image and any change to a clone
    is destroyed on the next refresh or push.

    WHAT IT TOUCHES IN THE GUEST
    ----------------------------
    1. Removes the non-present E1000 device, matched on its Intel PCI hardware ID
       (VEN_8086 DEV_100F / DEV_10D3). It does NOT bulk-remove non-present network
       devices - Zscaler, AnyConnect, GlobalProtect and similar VPN/ZTNA adapters
       legitimately appear as non-present. If MORE THAN ONE Intel ghost is found,
       nothing is removed and the VM is held for review.
    2. Renames the new adapter back to 'Ethernet' if that name is free.
    No NetworkList purge. No NlaSvc restart. Both reach well beyond this job.

    SAFETY DESIGN
    -------------
    * FULL-CLONE VALIDATION before anything is touched.
    * MAINTENANCE MODE IS PROVEN - polled until the broker reports MAINTENANCE.
      An AVAILABLE reading proves nothing.
    * POWER STATE IS PROVEN - polled until PoweredOff before any hardware change,
      and until PoweredOn after. Never assumed from a sleep.
    * ADD BEFORE REMOVE - a failed create changes nothing.
    * SURGICAL, VERIFIED ROLLBACK - only the adapter this script created is removed
      (tracked by ID). The E1000 is restored with an auto MAC first, the VMXNET3 is
      then withdrawn, and only then is the original MAC applied - so the original MAC
      is never claimed by two adapters at once, and the VM is never adapterless.
      The rolled-back config is then VERIFIED (type, MAC, portgroup, start-connected,
      power state) before ROLLED-BACK is logged.
    * ONCE COMMITTED, NO ROLLBACK - after the adapter set verifies correct, a later
      failure is reported, not reversed.
    * ONLY RELEASE WHAT WE LOCKED.
    * STARTED row written before each machine, so an interrupted run tells you
      exactly which VM was in flight.

    STATUSES
    --------
    STARTED           work began on this machine (paired with a later outcome row)
    CONVERTED         done, verified, reachable, back in the pool
    ROLLED-BACK       swap failed, original adapter restored AND verified
    HELD              converted but not verified - LEFT IN MAINTENANCE MODE
    FAILED            see Detail - LEFT IN MAINTENANCE MODE
    SKIPPED           session present / maintenance not confirmed / multi-NIC
    NOT-FULL-CLONE    instant or linked clone pool - fix the golden image instead
    ALREADY-VMXNET3   exactly one VMXNET3, start-connected
    VMXNET3-REVIEW    one VMXNET3 but not start-connected - look at it
    HALF-SWAPPED      exactly 1 VMXNET3 + 1 legacy, same portgroup - use -RepairHalfSwapped
    HALF-SWAPPED-ODD  mixed adapters that do not fit the clean repair case - by hand
    NO-NIC            zero adapters - wreckage, needs a human
    OTHER-NIC-TYPE    vmxnet2 / pcnet32 / something unexpected
    NOT-IN-HORIZON    broker does not know it
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory)]
    [string[]] $VMName,

    [Parameter(ParameterSetName = 'ByCsv', Mandatory)]
    [string]   $InputCsv,

    [string]   $CsvColumn          = 'VMName',
    [int]      $Limit              = 0,

    [string]   $vCenter            = 'vcenter.iprod.local',
    [string]   $ConnectionServer   = 'horizon-cs01.iprod.local',
    [string]   $LogCsv             = '',

    [switch]   $Execute,
    [switch]   $PreserveMac,
    [switch]   $NoDhcpRelease,
    [switch]   $NoGuestCleanup,
    [switch]   $RepairHalfSwapped,

    [int]      $VerifyPort         = 22443,     # Horizon agent Blast. 3389 if you prefer RDP.
    [int]      $MaintWaitSec       = 90,
    [int]      $PowerWaitSec       = 180,
    [int]      $ShutdownWaitSec    = 180,
    [int]      $ToolsWaitSec       = 300
)

$ErrorActionPreference = 'Stop'
$DryRun = -not $Execute
$RunId  = [guid]::NewGuid().ToString('N').Substring(0,8)

if (-not $LogCsv) {
    $LogCsv = "C:\Temp\NIC_Swap_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$RunId.csv"
}

# ---------------------------------------------------------------- targets
if ($PSCmdlet.ParameterSetName -eq 'ByCsv') {
    if (-not (Test-Path $InputCsv)) { throw "CSV not found: $InputCsv" }
    $csv = @(Import-Csv $InputCsv)
    if ($csv.Count -eq 0) { throw "CSV is empty: $InputCsv" }
    if ($csv[0].PSObject.Properties.Name -notcontains $CsvColumn) {
        throw "CSV has no '$CsvColumn' column. Columns: $(($csv[0].PSObject.Properties.Name) -join ', ')"
    }
    $targets = @($csv.$CsvColumn | Where-Object { $_ } | Select-Object -Unique)
} else {
    $targets = @($VMName | Where-Object { $_ } | Select-Object -Unique)
}

if ($Limit -gt 0 -and $targets.Count -gt $Limit) {
    Write-Host "Capping $($targets.Count) targets to $Limit." -ForegroundColor Yellow
    $targets = $targets | Select-Object -First $Limit
}

Import-Module VMware.PowerCLI -ErrorAction Stop
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null

Write-Host ""
Write-Host "Run ID: $RunId" -ForegroundColor Gray
if ($DryRun) { Write-Host "*** DRY RUN - no changes will be made. Add -Execute to commit. ***" -ForegroundColor Magenta }
else         { Write-Host "*** EXECUTE MODE - changes WILL be made. ***" -ForegroundColor Red }
Write-Host "Targets: $($targets.Count)   PreserveMac: $($PreserveMac.IsPresent)   GuestCleanup: $(-not $NoGuestCleanup)   Repair: $($RepairHalfSwapped.IsPresent)   VerifyPort: $VerifyPort" -ForegroundColor Gray
$targets | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

if ($NoGuestCleanup -and -not $DryRun) {
    Write-Host ""
    Write-Host "WARNING: -NoGuestCleanup - nothing is verified inside the guest and the ghost" -ForegroundColor Yellow
    Write-Host "         E1000 is left in place. Converted machines will be HELD." -ForegroundColor Yellow
}

Write-Host ""
if (-not $DryRun) {
    $ok = Read-Host "Convert the $($targets.Count) machine(s) above? Type YES to proceed"
    if ($ok -ne 'YES') { Write-Host "Aborted." -ForegroundColor Yellow; return }
}

# ---------------------------------------------------------------- connect
$cred = Get-Credential -Message "Credentials for vCenter and Horizon (DOMAIN\user)"

# Guest credentials are only ever used to change something. Never prompt on a dry run.
$guestCred = $null
if (-not $NoGuestCleanup -and -not $DryRun) {
    $guestCred = Get-Credential -Message "Guest OS credentials (local admin on the desktops)"
}

$viConn = $null
$hvConn = $null

try {
    $viConn = Connect-VIServer -Server $vCenter -Credential $cred -ErrorAction Stop
    $hvConn = Connect-HVServer -Server $ConnectionServer -Credential $cred -ErrorAction Stop
    $api    = $hvConn.ExtensionData

    # ------------------------------------------------------------ logging
    $dir = Split-Path $LogCsv -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $counts = @{}

    function Add-Log {
        param($Name,$Pool,$State,$Status,$Detail,$OldMac,$NewMac)
        [pscustomobject]@{
            Timestamp = (Get-Date -Format 's'); RunId = $RunId
            Machine = $Name; Pool = $Pool; HorizonState = $State
            Status = $Status; Detail = $Detail; OldMac = $OldMac; NewMac = $NewMac
        } | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8 -Append
        if ($Status -ne 'STARTED') {
            if (-not $counts.ContainsKey($Status)) { $counts[$Status] = 0 }
            $counts[$Status]++
        }
    }

    function Get-HvState { param($Id)
        try { return "$($api.Machine.Machine_Get($Id).Base.BasicState)" } catch { return 'UNKNOWN' }
    }
    function Enter-Maint { param($Id)
        try { $api.Machine.Machine_EnterMaintenanceMode($Id) } catch { $api.Machine.Machine_EnterMaintenanceModes(@($Id)) }
    }
    function Exit-Maint { param($Id)
        try { $api.Machine.Machine_ExitMaintenanceMode($Id) } catch { $api.Machine.Machine_ExitMaintenanceModes(@($Id)) }
    }

    # Power state is polled, never assumed from a sleep.
    function Wait-PowerState {
        param([string]$Name, [ValidateSet('PoweredOn','PoweredOff')][string]$Want, [int]$Seconds)
        $deadline = (Get-Date).AddSeconds($Seconds)
        do {
            if ((Get-VM -Name $Name).PowerState -eq $Want) { return $true }
            Start-Sleep -Seconds 5
        } while ((Get-Date) -lt $deadline)
        return ((Get-VM -Name $Name).PowerState -eq $Want)
    }

    # The guest's IPv4 for the specific adapter, by MAC. Guest.IpAddress alone can be
    # IPv6, link-local, or a stale entry from the adapter we just removed.
    function Get-GuestIpv4ByMac {
        param([string]$Name, [string]$Mac)
        $nets = @((Get-VM -Name $Name).ExtensionData.Guest.Net | Where-Object { $_.MacAddress -eq $Mac })
        foreach ($n in $nets) {
            foreach ($a in @($n.IpAddress)) {
                if ($a -match '^\d{1,3}(\.\d{1,3}){3}$' -and $a -notlike '169.254.*' -and $a -ne '0.0.0.0') { return $a }
            }
        }
        return $null
    }

    # ------------------------------------------------------------ Horizon index
    Write-Host "Indexing Horizon pools and machines ..." -ForegroundColor Cyan
    $qs = New-Object VMware.Hv.QueryServiceService

    # pools first - we need the provisioning source to prove full clone
    $poolIndex = @{}
    $qd = New-Object VMware.Hv.QueryDefinition
    $qd.QueryEntityType = 'DesktopSummaryView'
    $qd.Limit = 1000
    $r = $qs.QueryService_Create($api, $qd)
    while ($r.Results) {
        foreach ($d in $r.Results) {
            $poolIndex[$d.Id.Id] = [pscustomobject]@{
                Name   = "$($d.DesktopSummaryData.Name)"
                Source = "$($d.DesktopSummaryData.Source)"   # VIRTUAL_CENTER | VIEW_COMPOSER | INSTANT_CLONE_ENGINE | UNMANAGED
                Type   = "$($d.DesktopSummaryData.Type)"
            }
        }
        if (-not $r.Id) { break }
        $r = $qs.QueryService_GetNext($api, $r.Id)
    }
    if ($r.Id) { $qs.QueryService_Delete($api, $r.Id) }

    $hvIndex = @{}
    $qd = New-Object VMware.Hv.QueryDefinition
    $qd.QueryEntityType = 'MachineNamesView'
    $qd.Limit = 1000
    $r = $qs.QueryService_Create($api, $qd)
    while ($r.Results) {
        foreach ($mm in $r.Results) {
            $did = ''
            try { $did = "$($mm.Base.Desktop.Id)" } catch {}
            $hvIndex["$($mm.Base.Name)".ToUpper()] = [pscustomobject]@{ Id = $mm.Id; DesktopId = $did }
        }
        if (-not $r.Id) { break }
        $r = $qs.QueryService_GetNext($api, $r.Id)
    }
    if ($r.Id) { $qs.QueryService_Delete($api, $r.Id) }

    Write-Host "Pools: $($poolIndex.Count)   Machines: $($hvIndex.Count)" -ForegroundColor Cyan
    Write-Host ""

    # ------------------------------------------------------------ guest script
    # Returns JSON. Nothing here is parsed out of free text.
    $cleanupScript = @'
$ErrorActionPreference = 'SilentlyContinue'
$r = [ordered]@{
    ok = $false; vmxnet3 = $false; ifAlias = ''; renamed = $false
    profile = ''; ghostsFound = 0; ghostRemoved = ''; ghostAmbiguous = $false; note = ''
}

$nic = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
if (-not $nic) { $r.note = 'no vmxnet3 adapter in guest'; ($r | ConvertTo-Json -Compress); exit }
$r.vmxnet3 = $true

# Intel E1000 (82545EM) and E1000e (82574L) only. Nothing else.
$ghosts = @(Get-PnpDevice -Class Net -Status Unknown | Where-Object {
    $_.InstanceId -like 'PCI\VEN_8086&DEV_100F*' -or $_.InstanceId -like 'PCI\VEN_8086&DEV_10D3*'
})
$r.ghostsFound = $ghosts.Count

if ($ghosts.Count -gt 1) {
    # more than one match - do not guess, remove nothing
    $r.ghostAmbiguous = $true
    $r.note = 'multiple Intel ghost NICs found - none removed, review by hand'
}
elseif ($ghosts.Count -eq 1) {
    & pnputil /remove-device $ghosts[0].InstanceId 2>&1 | Out-Null
    $still = Get-PnpDevice -Class Net -Status Unknown | Where-Object { $_.InstanceId -eq $ghosts[0].InstanceId }
    if (-not $still) { $r.ghostRemoved = $ghosts[0].InstanceId }
    else             { $r.note = 'ghost removal did not take - cosmetic, not fatal' }
}

if ($nic.Name -ne 'Ethernet' -and -not (Get-NetAdapter -Name 'Ethernet')) {
    Rename-NetAdapter -Name $nic.Name -NewName 'Ethernet'
    $r.renamed = $true
}
$nic = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
$r.ifAlias = $nic.Name

# wait for NLA. do not force it - restarting NlaSvc makes ZTNA clients re-evaluate.
for ($i = 0; $i -lt 24; $i++) {
    $r.profile = "$((Get-NetConnectionProfile -InterfaceIndex $nic.ifIndex).NetworkCategory)"
    if ($r.profile -eq 'DomainAuthenticated') { break }
    Start-Sleep -Seconds 5
}

ipconfig /registerdns | Out-Null

# ghost removal is best effort and does not gate success. ambiguity does.
$r.ok = ($r.profile -eq 'DomainAuthenticated') -and (-not $r.ghostAmbiguous)
($r | ConvertTo-Json -Compress)
'@

    # ============================================================ work
    foreach ($name in $targets) {

        $hz = $hvIndex["$name".ToUpper()]
        if (-not $hz) {
            Write-Host "$name : not known to Horizon" -ForegroundColor Yellow
            Add-Log $name '' '' 'NOT-IN-HORIZON' 'Broker does not know this machine' '' ''
            continue
        }

        # ---- FULL CLONE OR NOTHING ----
        $pi     = $poolIndex[$hz.DesktopId]
        $pool   = if ($pi) { $pi.Name }   else { '' }
        $source = if ($pi) { $pi.Source } else { 'UNKNOWN' }

        if ($source -ne 'VIRTUAL_CENTER') {
            Write-Host "$name : NOT A FULL CLONE (pool source $source) - skipped" -ForegroundColor Yellow
            Add-Log $name $pool '' 'NOT-FULL-CLONE' "Pool provisioning source is $source. Fix the golden image and push, do not touch clones." '' ''
            continue
        }

        # ---- exactly one VM by that name ----
        $vmAll = @(Get-VM -Name $name -ErrorAction SilentlyContinue)
        if ($vmAll.Count -eq 0) { Add-Log $name $pool '' 'FAILED' 'Not found in vCenter' '' ''; continue }
        if ($vmAll.Count -gt 1) {
            Write-Host "$name : AMBIGUOUS - $($vmAll.Count) VMs share this name" -ForegroundColor Red
            Add-Log $name $pool '' 'SKIPPED' "Ambiguous - $($vmAll.Count) VMs share this name" '' ''
            continue
        }
        $vm = $vmAll[0]

        # ---- classify the adapter set ----
        $nics   = @(Get-NetworkAdapter -VM $vm)
        $legacy = @($nics | Where-Object { $_.Type -in @('e1000','e1000e') })
        $vmx    = @($nics | Where-Object { $_.Type -eq 'Vmxnet3' })
        $other  = @($nics | Where-Object { $_.Type -notin @('e1000','e1000e','Vmxnet3') })

        if ($nics.Count -eq 0) {
            Write-Host "$name : NO NETWORK ADAPTER AT ALL" -ForegroundColor Red
            Add-Log $name $pool '' 'NO-NIC' 'Zero network adapters - wreckage. Fix by hand.' '' ''
            continue
        }

        # already done? only if it is exactly one vmxnet3 AND it will connect at boot.
        if ($legacy.Count -eq 0 -and $other.Count -eq 0 -and $vmx.Count -eq 1 -and $nics.Count -eq 1) {
            if ($vmx[0].ConnectionState.StartConnected) {
                Write-Host "$name : already vmxnet3 on '$($vmx[0].NetworkName)'" -ForegroundColor DarkGray
                Add-Log $name $pool '' 'ALREADY-VMXNET3' "portgroup=$($vmx[0].NetworkName)" '' $vmx[0].MacAddress
            } else {
                Write-Host "$name : vmxnet3 present but NOT start-connected" -ForegroundColor Yellow
                Add-Log $name $pool '' 'VMXNET3-REVIEW' "vmxnet3 not start-connected. portgroup=$($vmx[0].NetworkName)" '' $vmx[0].MacAddress
            }
            continue
        }

        if ($legacy.Count -eq 0) {
            $types = ($nics | ForEach-Object { $_.Type }) -join ';'
            $st = if ($other.Count -gt 0) { 'OTHER-NIC-TYPE' } else { 'SKIPPED' }
            Write-Host "$name : SKIP - $($nics.Count) adapter(s): $types" -ForegroundColor Yellow
            Add-Log $name $pool '' $st "Adapters: $types" '' ''
            continue
        }

        # half-swapped: ONLY the clean case is auto-repairable.
        # exactly two adapters, exactly one vmxnet3, exactly one legacy, nothing else,
        # and both on the same portgroup. Anything else is a human's problem.
        $isCleanHalf = ($nics.Count -eq 2 -and $vmx.Count -eq 1 -and $legacy.Count -eq 1 -and
                        $other.Count -eq 0 -and $vmx[0].NetworkName -eq $legacy[0].NetworkName)
        $isAnyHalf   = ($vmx.Count -ge 1 -and $legacy.Count -ge 1)

        if ($isAnyHalf -and -not $isCleanHalf) {
            $types = ($nics | ForEach-Object { "$($_.Type):$($_.NetworkName)" }) -join ';'
            Write-Host "$name : HALF-SWAPPED-ODD - $types" -ForegroundColor Red
            Add-Log $name $pool '' 'HALF-SWAPPED-ODD' "Does not fit the clean repair case: $types" $legacy[0].MacAddress $vmx[0].MacAddress
            continue
        }
        if ($isCleanHalf -and -not $RepairHalfSwapped) {
            Write-Host "$name : HALF-SWAPPED - re-run with -RepairHalfSwapped" -ForegroundColor Red
            Add-Log $name $pool '' 'HALF-SWAPPED' "1 vmxnet3 + 1 $($legacy[0].Type) on '$($legacy[0].NetworkName)'. Use -RepairHalfSwapped." $legacy[0].MacAddress $vmx[0].MacAddress
            continue
        }

        # normal path must be exactly one legacy adapter and nothing else.
        # this guarantees there is only ever ONE adapter to remove, so a partial
        # removal cannot hide behind a boolean.
        if (-not $isCleanHalf -and ($nics.Count -ne 1 -or $legacy.Count -ne 1)) {
            $types = ($nics | ForEach-Object { $_.Type }) -join ';'
            Write-Host "$name : SKIP (multi-NIC: $types)" -ForegroundColor Yellow
            Add-Log $name $pool '' 'SKIPPED' "Multiple vNICs: $types" $legacy[0].MacAddress ''
            continue
        }

        $origMac  = $legacy[0].MacAddress
        $origType = $legacy[0].Type
        $origPg   = $legacy[0].NetworkName

        if ($DryRun) {
            $st   = Get-HvState $hz.Id
            $what = if ($isCleanHalf) { "repair half-swap: drop the $origType" } else { "swap $origType -> vmxnet3" }
            Write-Host "=== $name  [$pool]  state=$st" -ForegroundColor Cyan
            Write-Host "    would $what  (portgroup '$origPg')" -ForegroundColor Magenta
            Add-Log $name $pool $st 'DRYRUN' "Would $what on '$origPg'" $origMac ''
            continue
        }

        Write-Host "=== $name  [$pool]  $origType on '$origPg'" -ForegroundColor Cyan
        Add-Log $name $pool '' 'STARTED' "$origType on '$origPg'" $origMac ''

        $maintOwned = $false
        $addedNicId = $null
        $legacyGone = $false
        $committed  = $false
        $state      = 'UNKNOWN'
        $newMac     = ''

        try {
            # ---- maintenance mode, proven ----
            $preState = Get-HvState $hz.Id

            if ($preState -eq 'MAINTENANCE') {
                $state = 'MAINTENANCE'
                Write-Host "    already in maintenance (not ours - will not release it)" -ForegroundColor DarkGray
            }
            elseif ($preState -ne 'AVAILABLE') {
                Write-Host "    SKIP - state is $preState" -ForegroundColor Yellow
                Add-Log $name $pool $preState 'SKIPPED' 'Session present or machine in transition' $origMac ''
                continue
            }
            else {
                Enter-Maint $hz.Id
                $maintOwned = $true

                $confirmed = $false
                $seen      = $preState
                $deadline  = (Get-Date).AddSeconds($MaintWaitSec)
                do {
                    Start-Sleep -Seconds 5
                    $seen = Get-HvState $hz.Id
                    if ($seen -eq 'MAINTENANCE') { $confirmed = $true; break }
                    if ($seen -in @('CONNECTED','DISCONNECTED')) { break }
                } while ((Get-Date) -lt $deadline)

                if (-not $confirmed) {
                    Write-Host "    SKIP - maintenance not confirmed (state=$seen), backing out" -ForegroundColor Yellow
                    Exit-Maint $hz.Id
                    Add-Log $name $pool $seen 'SKIPPED' "Maintenance mode never confirmed (state=$seen) - backed out" $origMac ''
                    continue
                }
                $state = 'MAINTENANCE'
                Write-Host "    maintenance mode confirmed" -ForegroundColor DarkGray
            }

            $vm = Get-VM -Name $name

            # ---- release the lease while the NIC still exists ----
            if ($guestCred -and -not $NoDhcpRelease -and $vm.PowerState -eq 'PoweredOn' -and
                $vm.ExtensionData.Guest.ToolsRunningStatus -eq 'guestToolsRunning') {
                try {
                    Invoke-VMScript -VM $vm -ScriptType Bat -ScriptText 'ipconfig /release' `
                        -GuestCredential $guestCred -ErrorAction Stop | Out-Null
                    Write-Host "    dhcp lease released" -ForegroundColor DarkGray
                }
                catch { Write-Host "    WARN: lease release failed - it will age out" -ForegroundColor Yellow }
            }

            # ---- shut down, and PROVE it ----
            if ((Get-VM -Name $name).PowerState -eq 'PoweredOn') {
                Shutdown-VMGuest -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                if (-not (Wait-PowerState -Name $name -Want 'PoweredOff' -Seconds $ShutdownWaitSec)) {
                    Write-Host "    graceful shutdown timed out - forcing" -ForegroundColor Yellow
                    Stop-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null
                    if (-not (Wait-PowerState -Name $name -Want 'PoweredOff' -Seconds $PowerWaitSec)) {
                        throw 'VM did not reach PoweredOff - refusing to change virtual hardware'
                    }
                }
            }
            Write-Host "    powered off (confirmed)" -ForegroundColor DarkGray

            # ---- ADD (unless the vmxnet3 already exists from an interrupted run) ----
            if ($isCleanHalf) {
                Write-Host "    repairing - vmxnet3 already present" -ForegroundColor Yellow
            }
            else {
                $new = New-NetworkAdapter -VM (Get-VM -Name $name) -NetworkName $origPg -Type Vmxnet3 `
                        -StartConnected:$true -Confirm:$false -ErrorAction Stop
                $addedNicId = $new.Id
                Write-Host "    vmxnet3 added ($($new.MacAddress))" -ForegroundColor DarkGray
            }

            # ---- REMOVE the single legacy adapter ----
            $toGo = @(Get-NetworkAdapter -VM (Get-VM -Name $name) | Where-Object { $_.Type -in @('e1000','e1000e') })
            if ($toGo.Count -ne 1) { throw "Expected exactly 1 legacy adapter to remove, found $($toGo.Count)" }
            Remove-NetworkAdapter -NetworkAdapter $toGo[0] -Confirm:$false -ErrorAction Stop
            $legacyGone = $true
            Write-Host "    $origType removed" -ForegroundColor DarkGray

            # ---- only now is the original MAC free ----
            if ($PreserveMac) {
                $t = @(Get-NetworkAdapter -VM (Get-VM -Name $name) | Where-Object Type -eq 'Vmxnet3')[0]
                Set-NetworkAdapter -NetworkAdapter $t -MacAddress $origMac -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Host "    mac preserved ($origMac)" -ForegroundColor DarkGray
            }

            # ---- verify ----
            $after = @(Get-NetworkAdapter -VM (Get-VM -Name $name))
            $vmxA  = @($after | Where-Object Type -eq 'Vmxnet3')
            $legA  = @($after | Where-Object { $_.Type -in @('e1000','e1000e') })

            if ($after.Count -ne 1 -or $vmxA.Count -ne 1 -or $legA.Count -ne 0) {
                throw "Adapter set wrong after swap: $($after.Count) total, $($vmxA.Count) vmxnet3, $($legA.Count) legacy"
            }
            if ($vmxA[0].NetworkName -ne $origPg) {
                throw "vmxnet3 landed on '$($vmxA[0].NetworkName)', expected '$origPg'"
            }
            if (-not $vmxA[0].ConnectionState.StartConnected) {
                throw 'vmxnet3 is not set to connect at power on'
            }

            $newMac    = $vmxA[0].MacAddress
            $committed = $true

            # ---- power on, and PROVE it ----
            Start-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null
            if (-not (Wait-PowerState -Name $name -Want 'PoweredOn' -Seconds $PowerWaitSec)) {
                throw 'VM did not reach PoweredOn'
            }

            $deadline = (Get-Date).AddSeconds($ToolsWaitSec)
            do { Start-Sleep -Seconds 10; $g = (Get-VM -Name $name).ExtensionData.Guest.ToolsRunningStatus }
            while ($g -ne 'guestToolsRunning' -and (Get-Date) -lt $deadline)

            if ($g -ne 'guestToolsRunning') {
                Write-Host "    HELD - powered on but Tools never started" -ForegroundColor Red
                Add-Log $name $pool $state 'HELD' 'Tools did not start after power on' $origMac $newMac
                continue
            }

            # ---- the IPv4 of THAT adapter, by MAC ----
            $ip = $null
            $deadline = (Get-Date).AddSeconds(120)
            do {
                $ip = Get-GuestIpv4ByMac -Name $name -Mac $newMac
                if ($ip) { break }
                Start-Sleep -Seconds 10
            } while ((Get-Date) -lt $deadline)

            if (-not $ip) {
                Write-Host "    HELD - no IPv4 on the new adapter" -ForegroundColor Red
                Add-Log $name $pool $state 'HELD' 'No IPv4 bound to the vmxnet3 MAC - check DHCP' $origMac $newMac
                continue
            }
            Write-Host "    up - $ip" -ForegroundColor Green

            # ---- guest cleanup, structured ----
            $guest = $null
            $clean = 'SKIPPED-NO-CREDS'
            if ($guestCred) {
                try {
                    $res = Invoke-VMScript -VM (Get-VM -Name $name) -ScriptType Powershell `
                            -ScriptText $cleanupScript -GuestCredential $guestCred -ErrorAction Stop
                    $json = ($res.ScriptOutput -split "`r?`n" | Where-Object { $_ -match '^\s*\{' } | Select-Object -First 1)
                    if ($json) { $guest = $json | ConvertFrom-Json }
                    $clean = if ($guest) {
                        "ok=$($guest.ok) profile=$($guest.profile) ghosts=$($guest.ghostsFound) removed=$($guest.ghostRemoved) $($guest.note)"
                    } else { 'CLEANUP-NO-JSON' }
                    Write-Host "    cleanup: $clean" -ForegroundColor DarkGray
                }
                catch { $clean = "CLEANUP-FAILED $($_.Exception.Message -replace "`r?`n",' ')" }
            }

            # ---- independent reachability check ----
            $reach = $false
            for ($t = 0; $t -lt 6 -and -not $reach; $t++) {
                Start-Sleep -Seconds 10
                $reach = Test-NetConnection -ComputerName $ip -Port $VerifyPort -InformationLevel Quiet -WarningAction SilentlyContinue
            }
            Write-Host "    port $VerifyPort reachable: $reach" -ForegroundColor DarkGray

            $guestOk = ($guest -and $guest.ok)
            if (-not $guestOk -or -not $reach) {
                Write-Host "    HELD - not verified" -ForegroundColor Red
                Add-Log $name $pool $state 'HELD' "ip=$ip reachable=$reach cleanup=$clean" $origMac $newMac
                continue
            }

            # ---- back into service ----
            if ($maintOwned) {
                try {
                    Exit-Maint $hz.Id
                    Write-Host "    CONVERTED" -ForegroundColor Green
                    Add-Log $name $pool $state 'CONVERTED' "ip=$ip $clean" $origMac $newMac
                }
                catch {
                    Write-Host "    CONVERTED but could not leave maintenance" -ForegroundColor Yellow
                    Add-Log $name $pool $state 'HELD' "Converted OK (ip=$ip) but ExitMaintenanceMode failed: $($_.Exception.Message -replace "`r?`n",' ')" $origMac $newMac
                }
            }
            else {
                Write-Host "    CONVERTED - left in maintenance (was already there)" -ForegroundColor Green
                Add-Log $name $pool $state 'CONVERTED' "ip=$ip $clean (was already in maintenance)" $origMac $newMac
            }
        }
        catch {
            $err = ($_.Exception.Message -replace "`r?`n",' ')
            Write-Host "    ERROR: $err" -ForegroundColor Red

            if ($committed) {
                Add-Log $name $pool $state 'FAILED' "Config correct but a later step failed: $err" $origMac $newMac
            }
            elseif ($addedNicId -or $legacyGone) {
                Write-Host "    rolling back ..." -ForegroundColor Yellow
                try {
                    if ((Get-VM -Name $name).PowerState -eq 'PoweredOn') {
                        Stop-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null
                        if (-not (Wait-PowerState -Name $name -Want 'PoweredOff' -Seconds $PowerWaitSec)) {
                            throw 'VM would not power off for rollback'
                        }
                    }

                    # 1. put the original type back with an AUTO mac. The vmxnet3 may still
                    #    be holding $origMac (if -PreserveMac ran), so we must not claim it yet.
                    if ($legacyGone) {
                        New-NetworkAdapter -VM (Get-VM -Name $name) -NetworkName $origPg -Type $origType `
                            -StartConnected:$true -Confirm:$false -ErrorAction Stop | Out-Null
                    }

                    # 2. withdraw only the adapter WE created, by id. This frees $origMac.
                    if ($addedNicId) {
                        $mine = Get-NetworkAdapter -VM (Get-VM -Name $name) | Where-Object { $_.Id -eq $addedNicId }
                        if ($mine) { Remove-NetworkAdapter -NetworkAdapter $mine -Confirm:$false -ErrorAction Stop }
                    }

                    # 3. now the original MAC is free - restore it
                    $back = @(Get-NetworkAdapter -VM (Get-VM -Name $name))
                    if ($back.Count -ne 1) { throw "Rollback left $($back.Count) adapters" }
                    if ($back[0].MacAddress -ne $origMac) {
                        Set-NetworkAdapter -NetworkAdapter $back[0] -MacAddress $origMac -Confirm:$false -ErrorAction Stop | Out-Null
                    }

                    # 4. VERIFY the rollback before claiming it worked
                    $back = @(Get-NetworkAdapter -VM (Get-VM -Name $name))
                    if ($back.Count -ne 1)                                { throw "Rollback verify: $($back.Count) adapters" }
                    if ($back[0].Type       -ne $origType)                { throw "Rollback verify: type is $($back[0].Type), expected $origType" }
                    if ($back[0].MacAddress -ne $origMac)                 { throw "Rollback verify: mac is $($back[0].MacAddress), expected $origMac" }
                    if ($back[0].NetworkName -ne $origPg)                 { throw "Rollback verify: portgroup is $($back[0].NetworkName), expected $origPg" }
                    if (-not $back[0].ConnectionState.StartConnected)     { throw 'Rollback verify: not start-connected' }

                    Start-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null
                    if (-not (Wait-PowerState -Name $name -Want 'PoweredOn' -Seconds $PowerWaitSec)) {
                        throw 'Rollback verify: VM did not power back on'
                    }

                    Write-Host "    ROLLED BACK - original $origType restored and verified" -ForegroundColor Yellow
                    Add-Log $name $pool $state 'ROLLED-BACK' "Reversed and verified. Cause: $err" $origMac ''
                }
                catch {
                    $rb = ($_.Exception.Message -replace "`r?`n",' ')
                    Write-Host "    ROLLBACK FAILED - NEEDS MANUAL ATTENTION" -ForegroundColor Red
                    Add-Log $name $pool $state 'FAILED' "Swap failed: $err | ROLLBACK FAILED: $rb" $origMac ''
                }
            }
            else {
                Add-Log $name $pool $state 'FAILED' $err $origMac ''
            }
            # maintenance mode deliberately not released on any failure path
        }
    }

    # ------------------------------------------------------------ summary
    Write-Host ""
    $counts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("{0,-20} {1}" -f $_.Key, $_.Value) }
    Write-Host ""
    Write-Host "Log: $LogCsv" -ForegroundColor Green

    if (Test-Path $LogCsv) {
        $rows = @(Import-Csv $LogCsv | Where-Object { $_.RunId -eq $RunId })

        # a STARTED with no outcome = the run was interrupted on that machine
        $started  = $rows | Where-Object Status -eq 'STARTED' | Select-Object -ExpandProperty Machine
        $finished = $rows | Where-Object Status -ne 'STARTED' | Select-Object -ExpandProperty Machine
        $inflight = $started | Where-Object { $_ -notin $finished }
        if ($inflight) {
            Write-Host ""
            Write-Host "INTERRUPTED MID-MACHINE - check these first:" -ForegroundColor Red
            $inflight | ForEach-Object { Write-Host "  $_" }
        }

        $stuck = $rows | Where-Object { $_.Status -in @('FAILED','HELD','ROLLED-BACK') }
        if ($stuck) {
            Write-Host ""
            Write-Host "IN MAINTENANCE MODE - clear before users log in:" -ForegroundColor Red
            $stuck | ForEach-Object { Write-Host ("  {0,-20} {1,-16} {2}" -f $_.Machine, $_.Status, $_.Detail) }
        }

        $bad = $rows | Where-Object { $_.Status -in @('NO-NIC','HALF-SWAPPED','HALF-SWAPPED-ODD','VMXNET3-REVIEW') }
        if ($bad) {
            Write-Host ""
            Write-Host "NEEDS A LOOK:" -ForegroundColor Yellow
            $bad | ForEach-Object { Write-Host ("  {0,-20} {1,-16} {2}" -f $_.Machine, $_.Status, $_.Detail) }
        }
    }
}
finally {
    if ($hvConn) { Disconnect-HVServer -Server $ConnectionServer -Confirm:$false -ErrorAction SilentlyContinue }
    if ($viConn) { Disconnect-VIServer -Server $vCenter -Confirm:$false -ErrorAction SilentlyContinue }
}
