#Requires -Modules VMware.PowerCLI
<#
.SYNOPSIS
    Stamps the NVIDIA vGPU Manager (VIB) version from each ESXi host onto its
    guest VMs as a GuestInfo variable, so guests can self-detect whether their
    NVIDIA guest driver needs updating.

.DESCRIPTION
    Connects to vCenter, iterates every powered-on VM, finds its ESXi host,
    reads the NVIDIA VIB version from that host, and writes it into the VM's
    guestinfo.nvidia.vgpumanager.version advanced setting.

    Run this script:
      - After any vGPU Manager (VIB) upgrade on the ESXi hosts
      - On a schedule (e.g. daily via Task Scheduler / SCCM baseline) to keep
        the GuestInfo values current

.PARAMETER vCenterServer
    FQDN or IP of your vCenter server.

.PARAMETER Credential
    PSCredential with read/write access to vCenter (Get-Credential).

.PARAMETER ClusterName
    Optional. Restrict to a specific cluster name. Omit to process all clusters.

.PARAMETER WhatIf
    Show what would be written without making any changes.

.EXAMPLE
    .\Set-NvidiaGuestInfo.ps1 -vCenterServer vcenter.corp.local -Credential (Get-Credential)

.EXAMPLE
    .\Set-NvidiaGuestInfo.ps1 -vCenterServer vcenter.corp.local -Credential (Get-Credential) -ClusterName "VDI-Cluster-01"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$vCenterServer,

    [Parameter(Mandatory)]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [string]$ClusterName,

    [Parameter()]
    [string]$GuestInfoKey = "guestinfo.nvidia.vgpumanager.version"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Logging helper ────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colour = switch ($Level) {
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        "OK"    { "Green" }
        default { "Cyan" }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $colour
}

# ── Connect to vCenter ────────────────────────────────────────────────────────
Write-Log "Connecting to vCenter: $vCenterServer"
try {
    $null = Connect-VIServer -Server $vCenterServer -Credential $Credential -ErrorAction Stop
    Write-Log "Connected successfully." "OK"
} catch {
    Write-Log "Failed to connect to vCenter: $_" "ERROR"
    exit 1
}

try {
    # ── Build host → NVIDIA VIB version map ──────────────────────────────────
    Write-Log "Enumerating ESXi hosts and NVIDIA VIB versions..."
    $hostVibMap = @{}

    $esxHosts = if ($ClusterName) {
        Get-Cluster -Name $ClusterName -ErrorAction Stop | Get-VMHost
    } else {
        Get-VMHost
    }

    foreach ($esxHost in $esxHosts) {
        Write-Log "  Checking host: $($esxHost.Name)"
        try {
            $esxcli   = Get-EsxCli -VMHost $esxHost -V2 -ErrorAction Stop
            $nvidiaVib = $esxcli.software.vib.list.Invoke() |
                         Where-Object { $_.Name -match "nvidia" } |
                         Select-Object -First 1

            if ($nvidiaVib) {
                $hostVibMap[$esxHost.Name] = $nvidiaVib.Version
                Write-Log "    NVIDIA VIB version: $($nvidiaVib.Version)" "OK"
            } else {
                $hostVibMap[$esxHost.Name] = "NOT_FOUND"
                Write-Log "    No NVIDIA VIB found on this host." "WARN"
            }
        } catch {
            $hostVibMap[$esxHost.Name] = "ERROR"
            Write-Log "    Error querying VIBs: $_" "ERROR"
        }
    }

    # ── Stamp GuestInfo on each powered-on VM ────────────────────────────────
    Write-Log ""
    Write-Log "Stamping GuestInfo on powered-on VMs..."

    $vms = if ($ClusterName) {
        Get-Cluster -Name $ClusterName | Get-VM | Where-Object { $_.PowerState -eq "PoweredOn" }
    } else {
        Get-VM | Where-Object { $_.PowerState -eq "PoweredOn" }
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($vm in $vms) {
        $hostName   = $vm.VMHost.Name
        $vibVersion = $hostVibMap[$hostName]

        if (-not $vibVersion -or $vibVersion -in @("NOT_FOUND", "ERROR")) {
            Write-Log "  Skipping $($vm.Name) — host $hostName has no valid NVIDIA VIB." "WARN"
            $results.Add([PSCustomObject]@{
                VM          = $vm.Name
                Host        = $hostName
                VIBVersion  = $vibVersion
                Result      = "SKIPPED"
            })
            continue
        }

        Write-Log "  $($vm.Name) → setting $GuestInfoKey = $vibVersion"

        if ($PSCmdlet.ShouldProcess($vm.Name, "Set GuestInfo $GuestInfoKey = $vibVersion")) {
            try {
                $existing = Get-AdvancedSetting -Entity $vm -Name $GuestInfoKey -ErrorAction SilentlyContinue

                if ($existing) {
                    $existing | Set-AdvancedSetting -Value $vibVersion -Confirm:$false | Out-Null
                } else {
                    New-AdvancedSetting -Entity $vm -Name $GuestInfoKey -Value $vibVersion -Confirm:$false | Out-Null
                }

                $results.Add([PSCustomObject]@{
                    VM         = $vm.Name
                    Host       = $hostName
                    VIBVersion = $vibVersion
                    Result     = "OK"
                })
            } catch {
                Write-Log "    Error setting GuestInfo on $($vm.Name): $_" "ERROR"
                $results.Add([PSCustomObject]@{
                    VM         = $vm.Name
                    Host       = $hostName
                    VIBVersion = $vibVersion
                    Result     = "ERROR: $_"
                })
            }
        }
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    Write-Log ""
    Write-Log "── Summary ──────────────────────────────────────────────────"
    $results | Format-Table -AutoSize
    Write-Log "Done. $($results.Where({$_.Result -eq 'OK'}).Count) VMs updated, $($results.Where({$_.Result -eq 'SKIPPED'}).Count) skipped, $($results.Where({$_.Result -like 'ERROR*'}).Count) errors."

} finally {
    Write-Log "Disconnecting from vCenter."
    Disconnect-VIServer -Server $vCenterServer -Confirm:$false -ErrorAction SilentlyContinue
}
