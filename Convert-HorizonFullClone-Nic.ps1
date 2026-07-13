<#
    Convert-HorizonFullClone-Nic.ps1
    --------------------------------
    Replaces the E1000/E1000e vNIC with a VMXNET3 on Horizon FULL CLONE desktops.
    That is the whole job. It does nothing else to the guest.

        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042 -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\batch01.csv -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\audit.csv -Limit 25 -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\halfswapped.csv -RepairHalfSwapped -Execute

    WHAT IT TOUCHES IN THE GUEST
    ----------------------------
    Only two things, both scoped as tightly as possible:
      1. Removes the non-present E1000 device, matched on its Intel PCI hardware ID.
         It does NOT bulk-remove non-present network devices - Zscaler, AnyConnect,
         GlobalProtect and similar VPN adapters legitimately appear as non-present
         and must be left alone.
      2. Renames the new adapter back to 'Ethernet' if that name is free.
    It does NOT purge NetworkList profiles and does NOT restart NlaSvc. Both have a
    blast radius well beyond this job and can disturb VPN/ZTNA clients.

    SAFETY DESIGN
    -------------
    * MAINTENANCE MODE IS PROVEN, NOT ASSUMED. After requesting it, the broker is
      polled until it actually reports MAINTENANCE. An AVAILABLE reading proves
      nothing. If it never confirms, the request is backed out and the VM skipped.
    * ADD BEFORE REMOVE. The VMXNET3 is created before the E1000 is deleted, so a
      failed create changes nothing and the VM is never without an adapter.
    * SURGICAL ROLLBACK. Only what this script created or deleted is undone, tracked
      by adapter ID - never a blanket "remove all adapters". The E1000 is restored
      BEFORE the VMXNET3 is withdrawn, so there is no adapterless moment there either.
    * ONCE COMMITTED, NO ROLLBACK. After the adapter set has been verified correct,
      a later failure (power-on, Tools) is reported, not reversed. Reverting a
      correct config helps nobody.
    * ONLY RELEASE WHAT WE LOCKED. A VM already in maintenance when we found it is
      left in maintenance.
    * INCREMENTAL LOG, written per machine.

    STATUSES
    --------
    CONVERTED        done, verified, reachable, back in the pool
    ROLLED-BACK      swap failed and was reversed - original E1000 restored
    HELD             converted but not verified - LEFT IN MAINTENANCE MODE
    FAILED           see Detail - LEFT IN MAINTENANCE MODE
    SKIPPED          session present / maintenance not confirmed / multi-NIC
    ALREADY-VMXNET3  exactly one VMXNET3 and nothing else
    HALF-SWAPPED     has BOTH a VMXNET3 and an E1000 - re-run with -RepairHalfSwapped
    NO-NIC           zero adapters - wreckage, needs a human. NOT treated as done.
    OTHER-NIC-TYPE   vmxnet2 / pcnet32 / something unexpected
    NOT-IN-HORIZON   broker does not know it
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory)]
    [string[]] $VMName,

    [Parameter(ParameterSetName = 'ByCsv', Mandatory)]
    [string]   $InputCsv,

    [string]   $CsvColumn          = 'VMName',
    [int]      $Limit              = 0,          # 0 = no cap

    [string]   $vCenter            = 'vcenter.iprod.local',
    [string]   $ConnectionServer   = 'horizon-cs01.iprod.local',
    [string]   $LogCsv             = "C:\Temp\NIC_Swap_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv",

    [switch]   $Execute,           # without this it is a DRY RUN
    [switch]   $PreserveMac,       # only for DHCP reservations / MAC-based port auth
    [switch]   $NoDhcpRelease,     # skip ipconfig /release (lease ages out instead)
    [switch]   $NoGuestCleanup,    # no guest creds - converted machines will be HELD
    [switch]   $RepairHalfSwapped, # finish off VMs left with BOTH a VMXNET3 and an E1000

    [int]      $VerifyPort         = 22443,      # Horizon agent Blast. 3389 if you prefer RDP.
    [int]      $MaintWaitSec       = 90,
    [int]      $ShutdownWaitSec    = 180,
    [int]      $ToolsWaitSec       = 300
)

$DryRun = -not $Execute

# ---------------------------------------------------------------- targets
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

if ($Limit -gt 0 -and $targets.Count -gt $Limit) {
    Write-Host "Capping $($targets.Count) targets to $Limit." -ForegroundColor Yellow
    $targets = $targets | Select-Object -First $Limit
}

Import-Module VMware.PowerCLI -ErrorAction Stop
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null

Write-Host ""
if ($DryRun) { Write-Host "*** DRY RUN - no changes will be made. Add -Execute to commit. ***" -ForegroundColor Magenta }
else         { Write-Host "*** EXECUTE MODE - changes WILL be made. ***" -ForegroundColor Red }
Write-Host "Targets: $($targets.Count)   PreserveMac: $($PreserveMac.IsPresent)   GuestCleanup: $(-not $NoGuestCleanup)   Repair: $($RepairHalfSwapped.IsPresent)   VerifyPort: $VerifyPort" -ForegroundColor Gray
$targets | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

if ($NoGuestCleanup -and -not $DryRun) {
    Write-Host ""
    Write-Host "WARNING: -NoGuestCleanup - the ghost E1000 will not be removed and nothing" -ForegroundColor Yellow
    Write-Host "         will be verified inside the guest. Converted machines will be HELD." -ForegroundColor Yellow
}

Write-Host ""
if (-not $DryRun) {
    $ok = Read-Host "Convert the $($targets.Count) machine(s) above? Type YES to proceed"
    if ($ok -ne 'YES') { Write-Host "Aborted." -ForegroundColor Yellow; return }
}

# ---------------------------------------------------------------- connect
$cred = Get-Credential -Message "Credentials for vCenter and Horizon (DOMAIN\user)"

$guestCred = $null
if (-not $NoGuestCleanup) {
    $guestCred = Get-Credential -Message "Guest OS credentials (local admin on the desktops)"
}

Connect-VIServer -Server $vCenter -Credential $cred -ErrorAction Stop | Out-Null
$hv  = Connect-HVServer -Server $ConnectionServer -Credential $cred -ErrorAction Stop
$api = $hv.ExtensionData

# ---------------------------------------------------------------- helpers
$dir = Split-Path $LogCsv -Parent
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$counts = @{}

function Add-Log {
    param($Name,$Pool,$State,$Status,$Detail,$OldMac,$NewMac)
    [pscustomobject]@{
        Timestamp = (Get-Date -Format 's'); Machine = $Name; Pool = $Pool
        HorizonState = $State; Status = $Status; Detail = $Detail
        OldMac = $OldMac; NewMac = $NewMac
    } | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8 -Append
    if (-not $counts.ContainsKey($Status)) { $counts[$Status] = 0 }
    $counts[$Status]++
}

function Get-HvState { param($Id)
    try { return "$($api.Machine.Machine_Get($Id).Base.BasicState)" } catch { return 'UNKNOWN' }
}
function Enter-Maint { param($Id)
    try   { $api.Machine.Machine_EnterMaintenanceMode($Id) }
    catch { $api.Machine.Machine_EnterMaintenanceModes(@($Id)) }
}
function Exit-Maint { param($Id)
    try   { $api.Machine.Machine_ExitMaintenanceMode($Id) }
    catch { $api.Machine.Machine_ExitMaintenanceModes(@($Id)) }
}

# ---------------------------------------------------------------- Horizon index
Write-Host "Indexing Horizon machines ..." -ForegroundColor Cyan
$qs = New-Object VMware.Hv.QueryServiceService
$qd = New-Object VMware.Hv.QueryDefinition
$qd.QueryEntityType = 'MachineNamesView'
$qd.Limit = 1000

$hvIndex = @{}
$r = $qs.QueryService_Create($api, $qd)
while ($r.Results) {
    foreach ($mm in $r.Results) {
        $p = ''; try { $p = "$($mm.NamesData.DesktopName)" } catch {}
        $hvIndex["$($mm.Base.Name)".ToUpper()] = [pscustomobject]@{ Id = $mm.Id; Pool = $p }
    }
    if (-not $r.Id) { break }
    $r = $qs.QueryService_GetNext($api, $r.Id)
}
if ($r.Id) { $qs.QueryService_Delete($api, $r.Id) }
Write-Host "Horizon machines indexed: $($hvIndex.Count)" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------- in-guest cleanup
# Deliberately minimal. Two actions, both narrowly scoped.
#
# The ghost removal matches on Intel PCI hardware IDs, NOT on "any non-present
# network device". Zscaler, AnyConnect, GlobalProtect and other VPN/ZTNA miniports
# routinely show as non-present and must not be removed.
#   E1000  = Intel 82545EM  ->  PCI\VEN_8086&DEV_100F
#   E1000e = Intel 82574L   ->  PCI\VEN_8086&DEV_10D3
#
# No NetworkList profile purge, no NlaSvc restart. NLA reclassifies the adapter on
# its own once a DC is reachable; we wait for that rather than kicking the service,
# which would make ZTNA clients re-evaluate their trusted network.
$cleanupScript = @'
$ErrorActionPreference = 'SilentlyContinue'
$out = @()

$nic = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
if (-not $nic) { 'FAIL no vmxnet3 adapter present in guest'; exit 1 }

$ghosts = @(Get-PnpDevice -Class Net -Status Unknown | Where-Object {
    $_.InstanceId -like 'PCI\VEN_8086&DEV_100F*' -or $_.InstanceId -like 'PCI\VEN_8086&DEV_10D3*'
})
foreach ($g in $ghosts) { & pnputil /remove-device $g.InstanceId 2>&1 | Out-Null }
$out += "e1000_ghosts_removed=$($ghosts.Count)"

if ($nic.Name -ne 'Ethernet' -and -not (Get-NetAdapter -Name 'Ethernet')) {
    Rename-NetAdapter -Name $nic.Name -NewName 'Ethernet'
    $out += 'renamed=Ethernet'
}

# wait for NLA to classify - do not force it
$nic  = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
$prof = ''
for ($i = 0; $i -lt 24; $i++) {
    $prof = (Get-NetConnectionProfile -InterfaceIndex $nic.ifIndex).NetworkCategory
    if ($prof -eq 'DomainAuthenticated') { break }
    Start-Sleep -Seconds 5
}
$out += "profile=$prof"

ipconfig /registerdns | Out-Null

if ($prof -eq 'DomainAuthenticated') { 'OK ' + ($out -join ' ') }
else                                 { 'WARN-PROFILE ' + ($out -join ' ') }
'@

# ---------------------------------------------------------------- work
foreach ($name in $targets) {

    $hz = $hvIndex["$name".ToUpper()]
    if (-not $hz) {
        Write-Host "$name : not known to Horizon" -ForegroundColor Yellow
        Add-Log $name '' '' 'NOT-IN-HORIZON' 'Broker does not know this machine' '' ''
        continue
    }
    $pool = $hz.Pool

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

    $halfSwapped = ($vmx.Count -ge 1 -and $legacy.Count -ge 1)

    if ($nics.Count -eq 0) {
        Write-Host "$name : NO NETWORK ADAPTER AT ALL" -ForegroundColor Red
        Add-Log $name $pool '' 'NO-NIC' 'Zero network adapters - wreckage from an interrupted run. Fix by hand.' '' ''
        continue
    }
    if ($legacy.Count -eq 0 -and $vmx.Count -eq 1 -and $nics.Count -eq 1) {
        Write-Host "$name : already vmxnet3" -ForegroundColor DarkGray
        Add-Log $name $pool '' 'ALREADY-VMXNET3' '' '' $vmx[0].MacAddress
        continue
    }
    if ($legacy.Count -eq 0 -and $other.Count -gt 0) {
        $types = ($nics | ForEach-Object { $_.Type }) -join ';'
        Write-Host "$name : SKIP - unexpected adapter type(s): $types" -ForegroundColor Yellow
        Add-Log $name $pool '' 'OTHER-NIC-TYPE' "Adapter types: $types" '' ''
        continue
    }
    if ($legacy.Count -eq 0) {
        $types = ($nics | ForEach-Object { $_.Type }) -join ';'
        Write-Host "$name : SKIP - $($nics.Count) adapters ($types)" -ForegroundColor Yellow
        Add-Log $name $pool '' 'SKIPPED' "$($nics.Count) adapters: $types" '' ''
        continue
    }
    if ($halfSwapped -and -not $RepairHalfSwapped) {
        Write-Host "$name : HALF-SWAPPED ($($vmx.Count) vmxnet3 + $($legacy.Count) legacy)" -ForegroundColor Red
        Add-Log $name $pool '' 'HALF-SWAPPED' "$($vmx.Count) vmxnet3 + $($legacy.Count) legacy. Re-run with -RepairHalfSwapped." $legacy[0].MacAddress $vmx[0].MacAddress
        continue
    }
    if (-not $halfSwapped -and $nics.Count -gt 1) {
        Write-Host "$name : SKIP (multi-NIC)" -ForegroundColor Yellow
        Add-Log $name $pool '' 'SKIPPED' "Multiple vNICs ($($nics.Count))" $legacy[0].MacAddress ''
        continue
    }

    $origMac  = $legacy[0].MacAddress
    $origType = $legacy[0].Type
    $origPg   = $legacy[0].NetworkName

    if ($DryRun) {
        $st = Get-HvState $hz.Id
        $what = if ($halfSwapped) { "repair half-swapped: drop $($legacy.Count) legacy" } else { "swap $origType -> vmxnet3" }
        Write-Host "=== $name  [$pool]  state=$st" -ForegroundColor Cyan
        Write-Host "    would $what  (portgroup '$origPg')" -ForegroundColor Magenta
        Add-Log $name $pool $st 'DRYRUN' "Would $what on '$origPg'" $origMac ''
        continue
    }

    Write-Host "=== $name  [$pool]  $origType on '$origPg'" -ForegroundColor Cyan

    # ---- precise record of what we have actually done, for rollback ----
    $maintOwned  = $false   # WE put it into maintenance, so WE release it
    $addedNicId  = $null    # id of the adapter we created, if any
    $legacyGone  = $false   # the E1000 has actually been deleted
    $committed   = $false   # adapter set verified correct - past the point of reverting
    $state       = 'UNKNOWN'
    $newMac      = ''

    try {
        # ---- maintenance mode, and PROVE it took ----
        # Requesting it is asynchronous and can be refused. Reading back AVAILABLE
        # proves nothing, so poll until the broker actually says MAINTENANCE.
        $preState = Get-HvState $hz.Id

        if ($preState -eq 'MAINTENANCE') {
            # already parked by someone else - work on it, but do not release it after
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
                if ($seen -in @('CONNECTED','DISCONNECTED')) { break }   # a user beat us to it
            } while ((Get-Date) -lt $deadline)

            if (-not $confirmed) {
                Write-Host "    SKIP - maintenance mode not confirmed (state=$seen), backing out" -ForegroundColor Yellow
                Exit-Maint $hz.Id
                Add-Log $name $pool $seen 'SKIPPED' "Maintenance mode never confirmed (state=$seen) - request backed out" $origMac ''
                continue
            }
            $state = 'MAINTENANCE'
            Write-Host "    maintenance mode confirmed" -ForegroundColor DarkGray
        }

        $vm = Get-VM -Name $name

        # ---- release the DHCP lease while the NIC still exists ----
        if ($guestCred -and -not $NoDhcpRelease -and $vm.PowerState -eq 'PoweredOn' -and
            $vm.ExtensionData.Guest.ToolsRunningStatus -eq 'guestToolsRunning') {
            try {
                Invoke-VMScript -VM $vm -ScriptType Bat -ScriptText 'ipconfig /release' `
                    -GuestCredential $guestCred -ErrorAction Stop | Out-Null
                Write-Host "    dhcp lease released" -ForegroundColor DarkGray
            }
            catch { Write-Host "    WARN: lease release failed - it will age out" -ForegroundColor Yellow }
        }

        # ---- shut down ----
        if ($vm.PowerState -eq 'PoweredOn') {
            Shutdown-VMGuest -VM $vm -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            $deadline = (Get-Date).AddSeconds($ShutdownWaitSec)
            do { Start-Sleep -Seconds 5; $vm = Get-VM -Name $name }
            while ($vm.PowerState -eq 'PoweredOn' -and (Get-Date) -lt $deadline)

            if ($vm.PowerState -eq 'PoweredOn') {
                Write-Host "    graceful shutdown timed out - forcing" -ForegroundColor Yellow
                Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
                Start-Sleep -Seconds 5
            }
        }
        Write-Host "    powered off" -ForegroundColor DarkGray

        # ---- ADD the vmxnet3 (unless one already exists from an interrupted run) ----
        # If this throws, nothing has been changed: no adapter was added, none removed.
        if ($halfSwapped) {
            Write-Host "    repairing - vmxnet3 already present, only the legacy adapter to drop" -ForegroundColor Yellow
        }
        else {
            $new = New-NetworkAdapter -VM (Get-VM -Name $name) -NetworkName $origPg -Type Vmxnet3 `
                    -StartConnected:$true -Confirm:$false -ErrorAction Stop
            $addedNicId = $new.Id      # set ONLY after the create actually succeeded
            Write-Host "    vmxnet3 added ($($new.MacAddress))" -ForegroundColor DarkGray
        }

        # ---- REMOVE the legacy adapter(s), and nothing else ----
        foreach ($n in @(Get-NetworkAdapter -VM (Get-VM -Name $name) | Where-Object { $_.Type -in @('e1000','e1000e') })) {
            Remove-NetworkAdapter -NetworkAdapter $n -Confirm:$false -ErrorAction Stop
        }
        $legacyGone = $true
        Write-Host "    $origType removed" -ForegroundColor DarkGray

        # ---- only now is the old MAC free to reuse ----
        # It could not have been set at creation time: the E1000 still held it and
        # vCenter would have rejected the duplicate.
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

        $newMac    = $vmxA[0].MacAddress
        $committed = $true      # config is correct. From here, failures are reported, not reverted.

        # ---- power on ----
        Start-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null
        $deadline = (Get-Date).AddSeconds($ToolsWaitSec)
        do { Start-Sleep -Seconds 10; $g = (Get-VM -Name $name).ExtensionData.Guest.ToolsRunningStatus }
        while ($g -ne 'guestToolsRunning' -and (Get-Date) -lt $deadline)

        if ($g -ne 'guestToolsRunning') {
            Write-Host "    HELD - powered on but Tools never started" -ForegroundColor Red
            Add-Log $name $pool $state 'HELD' 'Tools did not start after power on - check the console' $origMac $newMac
            continue
        }

        $ip = (Get-VM -Name $name).ExtensionData.Guest.IpAddress
        Write-Host "    up - $ip" -ForegroundColor Green

        # ---- in-guest cleanup ----
        $clean = 'SKIPPED-NO-CREDS'
        if ($guestCred) {
            try {
                $res = Invoke-VMScript -VM (Get-VM -Name $name) -ScriptType Powershell `
                        -ScriptText $cleanupScript -GuestCredential $guestCred -ErrorAction Stop
                $clean = ($res.ScriptOutput -split "`r?`n" | Where-Object { $_ -match '\S' } | Select-Object -Last 1)
                Write-Host "    cleanup: $clean" -ForegroundColor DarkGray
            }
            catch { $clean = "CLEANUP-FAILED $($_.Exception.Message -replace "`r?`n",' ')" }
        }

        # ---- independent reachability check - no guest creds needed ----
        $reach = $false
        if ($ip) {
            for ($t = 0; $t -lt 6 -and -not $reach; $t++) {
                Start-Sleep -Seconds 10
                $reach = Test-NetConnection -ComputerName $ip -Port $VerifyPort -InformationLevel Quiet -WarningAction SilentlyContinue
            }
        }
        Write-Host "    port $VerifyPort reachable: $reach" -ForegroundColor DarkGray

        if ($clean -notlike 'OK*' -or -not $reach) {
            Write-Host "    HELD - not verified" -ForegroundColor Red
            Add-Log $name $pool $state 'HELD' "ip=$ip reachable=$reach cleanup=$clean" $origMac $newMac
            continue
        }

        # ---- back into service, but only if we were the ones who locked it ----
        if ($maintOwned) {
            try {
                Exit-Maint $hz.Id
                Write-Host "    CONVERTED" -ForegroundColor Green
                Add-Log $name $pool $state 'CONVERTED' "ip=$ip $clean" $origMac $newMac
            }
            catch {
                Write-Host "    CONVERTED but could not leave maintenance - release it manually" -ForegroundColor Yellow
                Add-Log $name $pool $state 'HELD' "Converted OK (ip=$ip) but ExitMaintenanceMode failed: $($_.Exception.Message -replace "`r?`n",' ')" $origMac $newMac
            }
        }
        else {
            Write-Host "    CONVERTED - left in maintenance (it was already there when we found it)" -ForegroundColor Green
            Add-Log $name $pool $state 'CONVERTED' "ip=$ip $clean (was already in maintenance - not released)" $origMac $newMac
        }
    }
    catch {
        $err = ($_.Exception.Message -replace "`r?`n",' ')
        Write-Host "    ERROR: $err" -ForegroundColor Red

        # ---- rollback ----
        # Undo exactly what we did and nothing else. Never a blanket adapter purge.
        # Restore the E1000 FIRST, then withdraw the vmxnet3 - so the VM is never
        # left without an adapter even if the second step fails.
        if ($committed) {
            # config was already verified correct - power-on or Tools failed. Reverting
            # a good config achieves nothing. Report and leave it in maintenance.
            Add-Log $name $pool $state 'FAILED' "Config correct but post-swap step failed: $err" $origMac $newMac
        }
        elseif ($legacyGone -or $addedNicId) {
            Write-Host "    rolling back ..." -ForegroundColor Yellow
            try {
                $vm = Get-VM -Name $name
                if ($vm.PowerState -eq 'PoweredOn') {
                    Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
                    Start-Sleep -Seconds 5
                }

                # 1. restore the original adapter if we deleted it
                if ($legacyGone) {
                    New-NetworkAdapter -VM (Get-VM -Name $name) -NetworkName $origPg -Type $origType `
                        -MacAddress $origMac -StartConnected:$true -Confirm:$false -ErrorAction Stop | Out-Null
                    Write-Host "    original $origType restored" -ForegroundColor Yellow
                }

                # 2. withdraw only the adapter WE created, by id
                if ($addedNicId) {
                    $mine = Get-NetworkAdapter -VM (Get-VM -Name $name) | Where-Object { $_.Id -eq $addedNicId }
                    if ($mine) { Remove-NetworkAdapter -NetworkAdapter $mine -Confirm:$false -ErrorAction Stop }
                }

                Start-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null

                Write-Host "    ROLLED BACK - powered on" -ForegroundColor Yellow
                Add-Log $name $pool $state 'ROLLED-BACK' "Swap failed and was reversed: $err" $origMac ''
                # left in maintenance on purpose - eyes on it before a user gets it
            }
            catch {
                $rb = ($_.Exception.Message -replace "`r?`n",' ')
                Write-Host "    ROLLBACK FAILED - NEEDS MANUAL ATTENTION" -ForegroundColor Red
                Add-Log $name $pool $state 'FAILED' "Swap failed: $err | ROLLBACK ALSO FAILED: $rb" $origMac ''
            }
        }
        else {
            # nothing was changed
            Add-Log $name $pool $state 'FAILED' $err $origMac ''
        }
        # maintenance mode deliberately not released on any failure path
    }
}

# ---------------------------------------------------------------- summary
Write-Host ""
$counts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("{0,-18} {1}" -f $_.Key, $_.Value) }
Write-Host ""
Write-Host "Log: $LogCsv" -ForegroundColor Green

if (Test-Path $LogCsv) {
    $rows = Import-Csv $LogCsv

    $stuck = $rows | Where-Object { $_.Status -in @('FAILED','HELD','ROLLED-BACK') }
    if ($stuck) {
        Write-Host ""
        Write-Host "IN MAINTENANCE MODE - clear these before users log in:" -ForegroundColor Red
        $stuck | ForEach-Object { Write-Host ("  {0,-20} {1,-14} {2}" -f $_.Machine, $_.Status, $_.Detail) }
    }

    $broken = $rows | Where-Object { $_.Status -in @('NO-NIC','HALF-SWAPPED') }
    if ($broken) {
        Write-Host ""
        Write-Host "BAD STATE FROM A PREVIOUS RUN:" -ForegroundColor Red
        $broken | ForEach-Object { Write-Host ("  {0,-20} {1,-14} {2}" -f $_.Machine, $_.Status, $_.Detail) }
        Write-Host "  HALF-SWAPPED can be finished off with -RepairHalfSwapped" -ForegroundColor Yellow
    }
}

Disconnect-HVServer -Server $ConnectionServer -Confirm:$false
