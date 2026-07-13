<#
    Convert-HorizonFullClone-Nic.ps1
    --------------------------------
    Cold-swaps E1000/E1000e -> VMXNET3 on Horizon FULL CLONE desktops.

    You tell it WHICH machines to do - either by naming them, or by feeding it
    a CSV. It will not go looking for work on its own.

    Test one:
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042 -Execute

    Test two:
        .\Convert-HorizonFullClone-Nic.ps1 -VMName VDI-W11-0042,VDI-W11-0043 -Execute

    From the audit CSV (any CSV with a VMName column - filter it in Excel first):
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\E1000_Audit.csv
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\batch01.csv -Execute

    Belt and braces - cap the run regardless of how many rows are in the CSV:
        .\Convert-HorizonFullClone-Nic.ps1 -InputCsv C:\Temp\E1000_Audit.csv -Limit 25 -Execute

    Per machine:
      1. Re-check Horizon state LIVE (the audit CSV goes stale in minutes)
      2. Enter maintenance mode - stops brokering AND stops Horizon power-managing
         the VM back on underneath you
      3. ipconfig /release - Windows does NOT release the lease on shutdown
      4. Graceful guest shutdown, hard stop after timeout
      5. Remove E1000, add VMXNET3 on the same portgroup
      6. Power on, wait for Tools
      7. In-guest cleanup: purge ghost NIC, restore adapter name, verify the
         firewall profile came back as Domain
      8. Exit maintenance mode - but ONLY if step 7 came back clean

    Resumable: anything already on VMXNET3 is logged and skipped.
    Failures are LEFT IN MAINTENANCE MODE deliberately and listed at the end.
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory)]
    [string[]] $VMName,

    [Parameter(ParameterSetName = 'ByCsv', Mandatory)]
    [string]   $InputCsv,

    [string]   $CsvColumn         = 'VMName',   # column in the CSV holding the machine name
    [int]      $Limit             = 0,          # 0 = no cap, do everything supplied

    [string]   $vCenter           = 'vcenter.iprod.local',
    [string]   $ConnectionServer  = 'horizon-cs01.iprod.local',
    [string]   $LogCsv            = "C:\Temp\NIC_Swap_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv",

    [switch]   $Execute,          # without this it is a DRY RUN
    [switch]   $PreserveMac,      # only if you have DHCP reservations or MAC-based port auth
    [switch]   $NoDhcpRelease,    # also disables the in-guest cleanup (both need guest creds)

    [int]      $ShutdownWaitSec   = 180,
    [int]      $ToolsWaitSec      = 300
)

$DryRun = -not $Execute

# ---------------------------------------------------------------- targets
if ($PSCmdlet.ParameterSetName -eq 'ByCsv') {
    if (-not (Test-Path $InputCsv)) { throw "CSV not found: $InputCsv" }
    $csv = Import-Csv $InputCsv
    if ($csv.Count -eq 0) { throw "CSV is empty: $InputCsv" }
    if ($csv[0].PSObject.Properties.Name -notcontains $CsvColumn) {
        throw "CSV has no '$CsvColumn' column. Columns present: $(($csv[0].PSObject.Properties.Name) -join ', ')"
    }
    $targets = $csv.$CsvColumn | Where-Object { $_ } | Select-Object -Unique
} else {
    $targets = $VMName | Where-Object { $_ } | Select-Object -Unique
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
Write-Host "Targets: $($targets.Count)   PreserveMac: $($PreserveMac.IsPresent)   DhcpRelease/Cleanup: $(-not $NoDhcpRelease)" -ForegroundColor Gray
$targets | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
Write-Host ""

if (-not $DryRun) {
    $ok = Read-Host "Convert the $($targets.Count) machine(s) above? Type YES to proceed"
    if ($ok -ne 'YES') { Write-Host "Aborted." -ForegroundColor Yellow; return }
}

# ---------------------------------------------------------------- connect
$cred = Get-Credential -Message "Credentials for vCenter and Horizon (DOMAIN\user)"

$guestCred = $null
if (-not $NoDhcpRelease) {
    $guestCred = Get-Credential -Message "Guest OS credentials (local admin on the desktops)"
}

Connect-VIServer -Server $vCenter -Credential $cred -ErrorAction Stop | Out-Null
$hv  = Connect-HVServer -Server $ConnectionServer -Credential $cred -ErrorAction Stop
$api = $hv.ExtensionData

# Only these mean nobody is on the machine. DISCONNECTED still has a live
# session with the user's apps open - it just has no client attached.
$SafeStates = @('AVAILABLE','MAINTENANCE')

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
        $pool = ''; try { $pool = "$($mm.NamesData.DesktopName)" } catch {}
        $hvIndex["$($mm.Base.Name)".ToUpper()] = [pscustomobject]@{
            Id   = $mm.Id
            Pool = $pool
        }
    }
    if (-not $r.Id) { break }
    $r = $qs.QueryService_GetNext($api, $r.Id)
}
if ($r.Id) { $qs.QueryService_Delete($api, $r.Id) }
Write-Host "Horizon machines indexed: $($hvIndex.Count)" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------- in-guest cleanup
# Purges the ghost E1000, restores the adapter name, and - the important bit -
# makes sure the new NIC landed on the Domain firewall profile. If NLA races the
# DC on first boot the profile lands as Public, Windows Firewall blocks
# Blast/PCoIP/RDP/MECM, and the desktop is up but unreachable.
$cleanupScript = @'
$ErrorActionPreference = 'SilentlyContinue'
$out = @()

$nic = Get-NetAdapter -Physical | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
if (-not $nic) { 'FAIL: no vmxnet3 adapter present'; exit 1 }

# purge ghost E1000 (clears Enum, Class and Tcpip interface keys)
$ghosts = Get-PnpDevice -Class Net -Status Unknown
foreach ($g in $ghosts) { & pnputil /remove-device $g.InstanceId 2>&1 | Out-Null }
$out += "ghosts_removed=$($ghosts.Count)"

# restore the connection name so anything keyed on 'Ethernet' still matches
if ($nic.Name -ne 'Ethernet') {
    if (-not (Get-NetAdapter -Name 'Ethernet')) {
        Rename-NetAdapter -Name $nic.Name -NewName 'Ethernet'
        $out += "renamed=$($nic.Name)->Ethernet"
    } else {
        $out += "rename_skipped=Ethernet_in_use"
    }
}

# firewall profile - the one that actually matters
$nic  = Get-NetAdapter -Physical | Where-Object { $_.InterfaceDescription -match 'vmxnet3' } | Select-Object -First 1
$prof = (Get-NetConnectionProfile -InterfaceIndex $nic.ifIndex).NetworkCategory
if ($prof -ne 'DomainAuthenticated') {
    Restart-Service NlaSvc -Force
    Start-Sleep -Seconds 20
    $prof = (Get-NetConnectionProfile -InterfaceIndex $nic.ifIndex).NetworkCategory
}
$out += "profile=$prof"

# purge stale network profiles left by the old adapter
$keep = (Get-NetConnectionProfile).Name
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles' | ForEach-Object {
    $n = (Get-ItemProperty $_.PSPath).ProfileName
    if ($n -and $n -notin $keep) { Remove-Item $_.PSPath -Recurse -Force; $out += "stale_profile_removed=$n" }
}

ipconfig /registerdns | Out-Null

if ($prof -eq 'DomainAuthenticated') { "OK " + ($out -join ' ') }
else                                 { "WARN-FIREWALL-PROFILE " + ($out -join ' ') }
'@

$log = New-Object System.Collections.Generic.List[object]

function Add-Log {
    param($Name,$Pool,$State,$Status,$Detail,$OldMac,$NewMac)
    $log.Add([pscustomobject]@{
        Machine = $Name; Pool = $Pool; HorizonState = $State; Status = $Status
        Detail  = $Detail; OldMac = $OldMac; NewMac = $NewMac; Timestamp = (Get-Date)
    })
}

# ---------------------------------------------------------------- work
foreach ($name in $targets) {

    $hz = $hvIndex["$name".ToUpper()]
    if (-not $hz) {
        Write-Host "$name : NOT IN HORIZON - skipped" -ForegroundColor Yellow
        Add-Log $name '' '' 'NOT-IN-HORIZON' 'Broker does not know this machine - convert manually if intended' '' ''
        continue
    }
    $pool = $hz.Pool

    # ---- live state re-check ----
    $state = 'UNKNOWN'
    try { $state = "$($api.Machine.Machine_Get($hz.Id).Base.BasicState)" } catch { }

    if ($state -notin $SafeStates) {
        Write-Host "$name : SKIP (state $state)" -ForegroundColor Yellow
        Add-Log $name $pool $state 'SKIPPED' 'Session present or machine in transition' '' ''
        continue
    }

    # ---- does it need doing ----
    try   { $vm = Get-VM -Name $name -ErrorAction Stop }
    catch { Add-Log $name $pool $state 'FAILED' 'Not found in vCenter' '' ''; continue }

    $nics = @(Get-NetworkAdapter -VM $vm)
    $old  = $nics | Where-Object { $_.Type -in @('e1000','e1000e') } | Select-Object -First 1

    if (-not $old) {
        Write-Host "$name : already vmxnet3" -ForegroundColor DarkGray
        Add-Log $name $pool $state 'ALREADY-VMXNET3' '' '' ''
        continue
    }
    if ($nics.Count -gt 1) {
        Write-Host "$name : SKIP (multi-NIC)" -ForegroundColor Yellow
        Add-Log $name $pool $state 'SKIPPED' 'Multiple vNICs - handle manually' $old.MacAddress ''
        continue
    }

    $oldMac = $old.MacAddress
    $pg     = $old.NetworkName

    Write-Host "=== $name  [$pool]  $($old.Type) on '$pg'" -ForegroundColor Cyan

    if ($DryRun) {
        Write-Host "    would swap to vmxnet3" -ForegroundColor Magenta
        Add-Log $name $pool $state 'DRYRUN' "Would swap $($old.Type) on '$pg'" $oldMac ''
        continue
    }

    try {
        # ---- maintenance mode ----
        try   { $api.Machine.Machine_EnterMaintenanceMode($hz.Id) }
        catch { $api.Machine.Machine_EnterMaintenanceModes(@($hz.Id)) }
        Start-Sleep -Seconds 3
        Write-Host "    maintenance mode on" -ForegroundColor DarkGray

        $vm = Get-VM -Name $name

        # ---- release the DHCP lease before we lose the NIC ----
        if ($guestCred -and $vm.PowerState -eq 'PoweredOn' -and
            $vm.ExtensionData.Guest.ToolsRunningStatus -eq 'guestToolsRunning') {
            try {
                Invoke-VMScript -VM $vm -ScriptType Bat -ScriptText 'ipconfig /release' `
                    -GuestCredential $guestCred -ErrorAction Stop | Out-Null
                Write-Host "    dhcp lease released" -ForegroundColor DarkGray
            }
            catch { Write-Host "    WARN: lease release failed - lease will age out" -ForegroundColor Yellow }
        }

        # ---- shut down ----
        if ($vm.PowerState -eq 'PoweredOn') {
            Shutdown-VMGuest -VM $vm -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            $deadline = (Get-Date).AddSeconds($ShutdownWaitSec)
            do {
                Start-Sleep -Seconds 5
                $vm = Get-VM -Name $name
            } while ($vm.PowerState -eq 'PoweredOn' -and (Get-Date) -lt $deadline)

            if ($vm.PowerState -eq 'PoweredOn') {
                Write-Host "    graceful shutdown timed out - forcing" -ForegroundColor Yellow
                Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
                Start-Sleep -Seconds 5
            }
        }
        Write-Host "    powered off" -ForegroundColor DarkGray

        # ---- swap ----
        $vm  = Get-VM -Name $name
        $old = Get-NetworkAdapter -VM $vm | Where-Object { $_.Type -in @('e1000','e1000e') } | Select-Object -First 1
        Remove-NetworkAdapter -NetworkAdapter $old -Confirm:$false -ErrorAction Stop

        $p = @{
            VM = $vm; NetworkName = $pg; Type = 'Vmxnet3'
            StartConnected = $true; Confirm = $false; ErrorAction = 'Stop'
        }
        if ($PreserveMac) { $p['MacAddress'] = $oldMac }

        $new    = New-NetworkAdapter @p
        $newMac = $new.MacAddress
        Write-Host "    vmxnet3 added ($newMac)" -ForegroundColor DarkGray

        # ---- power on ----
        Start-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
        $deadline = (Get-Date).AddSeconds($ToolsWaitSec)
        do {
            Start-Sleep -Seconds 10
            $g = (Get-VM -Name $name).ExtensionData.Guest.ToolsRunningStatus
        } while ($g -ne 'guestToolsRunning' -and (Get-Date) -lt $deadline)

        if ($g -ne 'guestToolsRunning') { throw 'Powered on but VMware Tools never started - check the console' }

        $ip = (Get-VM -Name $name).ExtensionData.Guest.IpAddress
        Write-Host "    up - $ip" -ForegroundColor Green

        # ---- in-guest remnant cleanup ----
        $clean = 'NOT-RUN'
        if ($guestCred) {
            try {
                $vm  = Get-VM -Name $name
                $res = Invoke-VMScript -VM $vm -ScriptType Powershell -ScriptText $cleanupScript `
                        -GuestCredential $guestCred -ErrorAction Stop
                $clean = ($res.ScriptOutput -split "`r?`n" | Where-Object { $_ -match '\S' } | Select-Object -Last 1)
                Write-Host "    cleanup: $clean" -ForegroundColor DarkGray
            }
            catch {
                $clean = "CLEANUP-FAILED: $($_.Exception.Message -replace "`r?`n",' ')"
                Write-Host "    $clean" -ForegroundColor Yellow
            }
        }

        # A desktop on the wrong firewall profile is unreachable - do NOT put it
        # back in the broker. Leave it in maintenance for a human.
        if ($clean -like 'WARN-FIREWALL-PROFILE*' -or $clean -like 'CLEANUP-FAILED*' -or $clean -like 'FAIL*') {
            Write-Host "    HELD IN MAINTENANCE - cleanup did not come back clean" -ForegroundColor Red
            Add-Log $name $pool $state 'HELD' "ip=$ip $clean" $oldMac $newMac
            continue
        }

        # ---- back into service ----
        try   { $api.Machine.Machine_ExitMaintenanceMode($hz.Id) }
        catch { $api.Machine.Machine_ExitMaintenanceModes(@($hz.Id)) }

        Add-Log $name $pool $state 'CONVERTED' "ip=$ip $clean" $oldMac $newMac
    }
    catch {
        $err = ($_.Exception.Message -replace "`r?`n",' ')
        Write-Host "    FAILED: $err" -ForegroundColor Red
        Write-Host "    LEFT IN MAINTENANCE MODE" -ForegroundColor Red
        Add-Log $name $pool $state 'FAILED' $err $oldMac ''
        # not auto-exiting maintenance mode - a broken desktop should stay out of the broker
    }
}

# ---------------------------------------------------------------- summary
$dir = Split-Path $LogCsv -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$log | Export-Csv -Path $LogCsv -NoTypeInformation -Encoding UTF8

Write-Host ""
$log | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("{0,-18} {1}" -f $_.Name, $_.Count) }
Write-Host ""
Write-Host "Log: $LogCsv" -ForegroundColor Green

$held = $log | Where-Object Status -in @('FAILED','HELD')
if ($held) {
    Write-Host ""
    Write-Host "STILL IN MAINTENANCE MODE - fix and release manually:" -ForegroundColor Red
    $held | ForEach-Object { Write-Host "  $($_.Machine)  -  $($_.Detail)" }
}

Disconnect-HVServer -Server $ConnectionServer -Confirm:$false
