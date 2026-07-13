<#
    Convert-HorizonFullClone-Nic.ps1
    --------------------------------
    Cold-swaps E1000/E1000e -> VMXNET3 on Horizon FULL CLONE desktops.

    You tell it WHICH machines to do. It will not go looking for work on its own.

        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042 -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042,VDI-W11-0043 -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\batch01.csv -Execute
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\E1000_Audit.csv -Limit 25 -Execute

    SAFETY DESIGN
    -------------
    * ADD THEN REMOVE. The VMXNET3 is added before the E1000 is removed, so a
      failure to add changes nothing. There is no window in which the VM has
      no network adapter.
    * ROLLBACK. If the remove fails, or the post-swap adapter set is not exactly
      one VMXNET3, the original E1000 is restored with its original MAC and
      portgroup and the VM is powered back on.
    * MAINTENANCE MODE FIRST. Entered before the session check, then the state is
      re-read. If a user connected in the gap, maintenance mode is backed out and
      the machine is skipped untouched.
    * INCREMENTAL LOG. Every machine is appended to the CSV as it completes. Kill
      the script at any point and the log still tells you exactly where you got to.
    * VERIFY BEFORE RETURNING TO SERVICE. A machine only leaves maintenance mode
      once it is powered on, Tools is up, and it answers on the Horizon agent port.
      Anything else is HELD in maintenance for a human.
    * NO SILENT DOWNGRADE. Without guest credentials the in-guest cleanup cannot
      run, so machines are HELD rather than returned to the broker unverified.
      -NoGuestCleanup makes that choice explicit.

    STATUSES
    --------
    CONVERTED        done, verified, back in the pool
    ROLLED-BACK      swap failed, original E1000 restored, machine powered on
    HELD             converted but not verified - STILL IN MAINTENANCE MODE
    FAILED           see Detail - STILL IN MAINTENANCE MODE, check by hand
    SKIPPED          session present, multi-NIC, or state not safe - untouched
    ALREADY-VMXNET3  nothing to do
    NOT-IN-HORIZON   broker does not know it - untouched
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory)]
    [string[]] $VMName,

    [Parameter(ParameterSetName = 'ByCsv', Mandatory)]
    [string]   $InputCsv,

    [string]   $CsvColumn         = 'VMName',
    [int]      $Limit             = 0,          # 0 = no cap

    [string]   $vCenter           = 'vcenter.iprod.local',
    [string]   $ConnectionServer  = 'horizon-cs01.iprod.local',
    [string]   $LogCsv            = "C:\Temp\NIC_Swap_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv",

    [switch]   $Execute,          # without this it is a DRY RUN
    [switch]   $PreserveMac,      # only for DHCP reservations / MAC-based port auth
    [switch]   $NoDhcpRelease,    # skip ipconfig /release (lease ages out instead)
    [switch]   $NoGuestCleanup,   # run without guest creds - machines will be HELD, not returned to service

    [int]      $VerifyPort        = 22443,      # Horizon agent Blast. 3389 if you prefer RDP.
    [int]      $ShutdownWaitSec   = 180,
    [int]      $ToolsWaitSec      = 300
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
Write-Host "Targets: $($targets.Count)   PreserveMac: $($PreserveMac.IsPresent)   GuestCleanup: $(-not $NoGuestCleanup)   VerifyPort: $VerifyPort" -ForegroundColor Gray
$targets | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

if ($NoGuestCleanup -and -not $DryRun) {
    Write-Host ""
    Write-Host "WARNING: -NoGuestCleanup is set. The ghost NIC will not be purged and the" -ForegroundColor Yellow
    Write-Host "         firewall profile will not be checked. Every converted machine will be" -ForegroundColor Yellow
    Write-Host "         HELD in maintenance mode for you to verify by hand." -ForegroundColor Yellow
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

# Only these mean nobody is on the machine. DISCONNECTED still has a live session.
$SafeStates = @('AVAILABLE','MAINTENANCE')

# ---------------------------------------------------------------- logging
# Written per machine, not at the end. If this script is killed the log is still
# an accurate record of what was done and what is still in maintenance mode.
$dir = Split-Path $LogCsv -Parent
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$counts = @{}

function Add-Log {
    param($Name,$Pool,$State,$Status,$Detail,$OldMac,$NewMac)
    $row = [pscustomobject]@{
        Timestamp = (Get-Date -Format 's'); Machine = $Name; Pool = $Pool
        HorizonState = $State; Status = $Status; Detail = $Detail
        OldMac = $OldMac; NewMac = $NewMac
    }
    $row | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8 -Append
    if (-not $counts.ContainsKey($Status)) { $counts[$Status] = 0 }
    $counts[$Status]++
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
$cleanupScript = @'
$ErrorActionPreference = 'SilentlyContinue'
$out = @()

$nic = Get-NetAdapter -Physical | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
if (-not $nic) { 'FAIL no vmxnet3 adapter present'; exit 1 }

$ghosts = Get-PnpDevice -Class Net -Status Unknown
foreach ($g in $ghosts) { & pnputil /remove-device $g.InstanceId 2>&1 | Out-Null }
$out += "ghosts=$($ghosts.Count)"

if ($nic.Name -ne 'Ethernet' -and -not (Get-NetAdapter -Name 'Ethernet')) {
    Rename-NetAdapter -Name $nic.Name -NewName 'Ethernet'
    $out += "renamed"
}

$nic  = Get-NetAdapter -Physical | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
$prof = (Get-NetConnectionProfile -InterfaceIndex $nic.ifIndex).NetworkCategory
if ($prof -ne 'DomainAuthenticated') {
    Restart-Service NlaSvc -Force
    Start-Sleep -Seconds 20
    $prof = (Get-NetConnectionProfile -InterfaceIndex $nic.ifIndex).NetworkCategory
}
$out += "profile=$prof"

$keep = (Get-NetConnectionProfile).Name
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles' | ForEach-Object {
    $n = (Get-ItemProperty $_.PSPath).ProfileName
    if ($n -and $n -notin $keep) { Remove-Item $_.PSPath -Recurse -Force; $out += "stale_removed" }
}

ipconfig /registerdns | Out-Null

if ($prof -eq 'DomainAuthenticated') { "OK " + ($out -join ' ') }
else                                 { "WARN-PROFILE " + ($out -join ' ') }
'@

# ---------------------------------------------------------------- work
foreach ($name in $targets) {

    $hz = $hvIndex["$name".ToUpper()]
    if (-not $hz) {
        Write-Host "$name : not known to Horizon - skipped" -ForegroundColor Yellow
        Add-Log $name '' '' 'NOT-IN-HORIZON' 'Broker does not know this machine' '' ''
        continue
    }
    $pool = $hz.Pool

    # ---- one and only one VM by that name ----
    $vmAll = @(Get-VM -Name $name -ErrorAction SilentlyContinue)
    if ($vmAll.Count -eq 0) { Add-Log $name $pool '' 'FAILED' 'Not found in vCenter' '' ''; continue }
    if ($vmAll.Count -gt 1) {
        Write-Host "$name : AMBIGUOUS - $($vmAll.Count) VMs with this name" -ForegroundColor Red
        Add-Log $name $pool '' 'SKIPPED' "Ambiguous - $($vmAll.Count) VMs share this name" '' ''
        continue
    }
    $vm = $vmAll[0]

    # ---- is there anything to do ----
    $nics = @(Get-NetworkAdapter -VM $vm)
    $old  = $nics | Where-Object { $_.Type -in @('e1000','e1000e') } | Select-Object -First 1

    if (-not $old) {
        Write-Host "$name : already vmxnet3" -ForegroundColor DarkGray
        Add-Log $name $pool '' 'ALREADY-VMXNET3' '' '' ''
        continue
    }
    if ($nics.Count -gt 1) {
        Write-Host "$name : SKIP (multi-NIC)" -ForegroundColor Yellow
        Add-Log $name $pool '' 'SKIPPED' 'Multiple vNICs - handle manually' $old.MacAddress ''
        continue
    }

    $oldMac  = $old.MacAddress
    $oldType = $old.Type
    $pg      = $old.NetworkName

    if ($DryRun) {
        $st = 'UNKNOWN'; try { $st = "$($api.Machine.Machine_Get($hz.Id).Base.BasicState)" } catch {}
        Write-Host "=== $name  [$pool]  $oldType on '$pg'  state=$st" -ForegroundColor Cyan
        Write-Host "    would swap to vmxnet3" -ForegroundColor Magenta
        Add-Log $name $pool $st 'DRYRUN' "Would swap $oldType on '$pg'" $oldMac ''
        continue
    }

    Write-Host "=== $name  [$pool]  $oldType on '$pg'" -ForegroundColor Cyan

    $maintOwned = $false     # did WE put it into maintenance
    $swapped    = $false     # has the adapter set been changed

    try {
        # ---- maintenance mode FIRST, then re-read state ----
        # Entering first closes the race where a user connects between the check
        # and the lock. If it turns out someone was already on, we back out.
        $preState = 'UNKNOWN'
        try { $preState = "$($api.Machine.Machine_Get($hz.Id).Base.BasicState)" } catch {}

        if ($preState -ne 'MAINTENANCE') {
            Enter-Maint $hz.Id
            $maintOwned = $true
            Start-Sleep -Seconds 5
        }

        $state = 'UNKNOWN'
        try { $state = "$($api.Machine.Machine_Get($hz.Id).Base.BasicState)" } catch {}

        if ($state -notin $SafeStates) {
            Write-Host "    SKIP - state is $state, backing out" -ForegroundColor Yellow
            if ($maintOwned) { Exit-Maint $hz.Id }
            Add-Log $name $pool $state 'SKIPPED' 'Session appeared or machine in transition - backed out' $oldMac ''
            continue
        }
        Write-Host "    maintenance mode on" -ForegroundColor DarkGray

        $vm = Get-VM -Name $name

        # ---- release the DHCP lease while we still have a NIC ----
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

        # ---- SWAP: add first, remove second ----
        # If the add fails, nothing has changed and the VM still has its E1000.
        # There is never a moment where the VM has no adapter.
        $vm = Get-VM -Name $name

        $p = @{ VM = $vm; NetworkName = $pg; Type = 'Vmxnet3'; StartConnected = $true
                Confirm = $false; ErrorAction = 'Stop' }
        if ($PreserveMac) { $p['MacAddress'] = $oldMac }

        $new = New-NetworkAdapter @p
        $swapped = $true
        Write-Host "    vmxnet3 added ($($new.MacAddress))" -ForegroundColor DarkGray

        $old = Get-NetworkAdapter -VM $vm | Where-Object { $_.Type -in @('e1000','e1000e') } | Select-Object -First 1
        Remove-NetworkAdapter -NetworkAdapter $old -Confirm:$false -ErrorAction Stop

        # ---- verify the adapter set is exactly what we want ----
        $after = @(Get-NetworkAdapter -VM (Get-VM -Name $name))
        $vmx   = @($after | Where-Object Type -eq 'Vmxnet3')
        $leg   = @($after | Where-Object { $_.Type -in @('e1000','e1000e') })

        if ($vmx.Count -ne 1 -or $leg.Count -ne 0 -or $after.Count -ne 1) {
            throw "Adapter set wrong after swap: $($after.Count) total, $($vmx.Count) vmxnet3, $($leg.Count) legacy"
        }

        $newMac = $vmx[0].MacAddress

        # ---- power on ----
        Start-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null
        $deadline = (Get-Date).AddSeconds($ToolsWaitSec)
        do { Start-Sleep -Seconds 10; $g = (Get-VM -Name $name).ExtensionData.Guest.ToolsRunningStatus }
        while ($g -ne 'guestToolsRunning' -and (Get-Date) -lt $deadline)

        if ($g -ne 'guestToolsRunning') {
            Write-Host "    HELD - powered on but Tools never started" -ForegroundColor Red
            Add-Log $name $pool $state 'HELD' 'Tools did not start after power on - check the console' $oldMac $newMac
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

        # ---- independent reachability check (does not need guest creds) ----
        # This is what actually proves the firewall profile came back right.
        $reach = $false
        if ($ip) {
            for ($t = 0; $t -lt 6 -and -not $reach; $t++) {
                Start-Sleep -Seconds 10
                $reach = (Test-NetConnection -ComputerName $ip -Port $VerifyPort -InformationLevel Quiet -WarningAction SilentlyContinue)
            }
        }
        Write-Host "    port $VerifyPort reachable: $reach" -ForegroundColor DarkGray

        # ---- gate: only a clean, reachable machine goes back into the pool ----
        $goodClean = ($clean -like 'OK*')
        if (-not $goodClean -or -not $reach) {
            Write-Host "    HELD IN MAINTENANCE - not verified" -ForegroundColor Red
            Add-Log $name $pool $state 'HELD' "ip=$ip reachable=$reach cleanup=$clean" $oldMac $newMac
            continue
        }

        # ---- back into service ----
        try {
            Exit-Maint $hz.Id
            Add-Log $name $pool $state 'CONVERTED' "ip=$ip $clean" $oldMac $newMac
            Write-Host "    CONVERTED" -ForegroundColor Green
        }
        catch {
            # The desktop is fine, we just could not release the lock.
            Add-Log $name $pool $state 'HELD' "Converted OK (ip=$ip) but ExitMaintenanceMode failed: $($_.Exception.Message -replace "`r?`n",' ')" $oldMac $newMac
            Write-Host "    CONVERTED but still in maintenance - release it manually" -ForegroundColor Yellow
        }
    }
    catch {
        $err = ($_.Exception.Message -replace "`r?`n",' ')
        Write-Host "    ERROR: $err" -ForegroundColor Red

        # ---- rollback: get the machine back to a working E1000 and power it on ----
        if ($swapped) {
            Write-Host "    rolling back to $oldType ..." -ForegroundColor Yellow
            try {
                $vm = Get-VM -Name $name
                if ($vm.PowerState -eq 'PoweredOn') { Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null; Start-Sleep -Seconds 5 }

                # strip whatever is there
                Get-NetworkAdapter -VM (Get-VM -Name $name) | Remove-NetworkAdapter -Confirm:$false -ErrorAction Stop

                # restore the original, MAC and all
                New-NetworkAdapter -VM (Get-VM -Name $name) -NetworkName $pg -Type $oldType `
                    -MacAddress $oldMac -StartConnected:$true -Confirm:$false -ErrorAction Stop | Out-Null

                Start-VM -VM (Get-VM -Name $name) -Confirm:$false -ErrorAction Stop | Out-Null

                Write-Host "    ROLLED BACK - original $oldType restored, powered on" -ForegroundColor Yellow
                Add-Log $name $pool $state 'ROLLED-BACK' "Swap failed and was reversed: $err" $oldMac ''
                # left in maintenance deliberately - you want eyes on it before a user gets it
                continue
            }
            catch {
                $rb = ($_.Exception.Message -replace "`r?`n",' ')
                Write-Host "    ROLLBACK FAILED - MACHINE NEEDS MANUAL ATTENTION" -ForegroundColor Red
                Add-Log $name $pool $state 'FAILED' "Swap failed: $err | ROLLBACK ALSO FAILED: $rb" $oldMac ''
                continue
            }
        }

        Add-Log $name $pool $state 'FAILED' $err $oldMac ''
        # not exiting maintenance mode - a machine we do not understand stays out of the broker
    }
}

# ---------------------------------------------------------------- summary
Write-Host ""
$counts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("{0,-18} {1}" -f $_.Key, $_.Value) }
Write-Host ""
Write-Host "Log: $LogCsv" -ForegroundColor Green

if (Test-Path $LogCsv) {
    $stuck = Import-Csv $LogCsv | Where-Object { $_.Status -in @('FAILED','HELD','ROLLED-BACK') }
    if ($stuck) {
        Write-Host ""
        Write-Host "STILL IN MAINTENANCE MODE - deal with these before users log in:" -ForegroundColor Red
        $stuck | ForEach-Object { Write-Host ("  {0,-20} {1,-14} {2}" -f $_.Machine, $_.Status, $_.Detail) }
    }
}

Disconnect-HVServer -Server $ConnectionServer -Confirm:$false
