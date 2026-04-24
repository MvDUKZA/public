<#
.SYNOPSIS
    Removes \\iprod.local\NETLOGON and \\iprod.local\SYSVOL values from
    HKLM\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths
    and triggers a Qualys on-demand Policy Compliance scan after cleanup.
    Does not depend on the Remote Registry service.
#>

param(
    [string]$ComputerListPath = ".\machines.txt",
    [string]$LogPath = ".\HardenedPaths_Cleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [int]$ThrottleLimit = 25
)

$computers = Get-Content $ComputerListPath | Where-Object { $_ -and $_ -notmatch '^\s*#' }

Write-Host "Processing $($computers.Count) machines with throttle $ThrottleLimit..." -ForegroundColor Cyan

$results = $computers | ForEach-Object -Parallel {
    $computer = $_
    $result = [PSCustomObject]@{
        ComputerName     = $computer
        Online           = $false
        NetlogonHad      = $null
        NetlogonNow      = $null
        SysvolHad        = $null
        SysvolNow        = $null
        RegCleanupStatus = ''
        QualysAgentFound = $null
        QualysScanTrigger= ''
        Error            = ''
        Timestamp        = Get-Date -Format 's'
    }

    # Reachability check
    if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet -TimeoutSeconds 2)) {
        $result.RegCleanupStatus = 'Offline'
        return $result
    }
    $result.Online = $true

    try {
        $scriptResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
            $out = @{}

            # --- 1. HardenedPaths cleanup ---
            $hpPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths"
            $targets = @('\\iprod.local\NETLOGON', '\\iprod.local\SYSVOL')

            foreach ($name in $targets) {
                $before = (Get-ItemProperty -Path $hpPath -Name $name -ErrorAction SilentlyContinue).$name
                $out["$name`_Before"] = [bool]($null -ne $before)

                if ($null -ne $before) {
                    Remove-ItemProperty -Path $hpPath -Name $name -Force -ErrorAction Stop
                }

                $after = (Get-ItemProperty -Path $hpPath -Name $name -ErrorAction SilentlyContinue).$name
                $out["$name`_After"] = [bool]($null -ne $after)
            }

            # --- 2. Trigger Qualys on-demand Policy Compliance scan ---
            $qPath   = "HKLM:\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\PolicyCompliance"
            $qParent = "HKLM:\SOFTWARE\Qualys\QualysAgent"

            if (Test-Path $qParent) {
                $out['QualysAgentFound'] = $true
                try {
                    if (-not (Test-Path $qPath)) {
                        New-Item -Path $qPath -Force | Out-Null
                    }
                    # Standard Qualys on-demand scan values
                    New-ItemProperty -Path $qPath -Name 'ScanOnDemand'      -Value 1 -PropertyType DWord -Force | Out-Null
                    New-ItemProperty -Path $qPath -Name 'ScanOnStartup'     -Value 0 -PropertyType DWord -Force | Out-Null
                    New-ItemProperty -Path $qPath -Name 'CpuLimit'          -Value 50 -PropertyType DWord -Force | Out-Null
                    New-ItemProperty -Path $qPath -Name 'ScanOnDemandTimeout' -Value 0 -PropertyType DWord -Force | Out-Null
                    $out['QualysScanTrigger'] = 'Triggered'
                }
                catch {
                    $out['QualysScanTrigger'] = "Failed: $($_.Exception.Message)"
                }
            }
            else {
                $out['QualysAgentFound'] = $false
                $out['QualysScanTrigger'] = 'AgentNotInstalled'
            }

            return $out
        }

        $result.NetlogonHad       = $scriptResult['\\iprod.local\NETLOGON_Before']
        $result.NetlogonNow       = $scriptResult['\\iprod.local\NETLOGON_After']
        $result.SysvolHad         = $scriptResult['\\iprod.local\SYSVOL_Before']
        $result.SysvolNow         = $scriptResult['\\iprod.local\SYSVOL_After']
        $result.QualysAgentFound  = $scriptResult['QualysAgentFound']
        $result.QualysScanTrigger = $scriptResult['QualysScanTrigger']

        if (-not $result.NetlogonNow -and -not $result.SysvolNow) {
            $result.RegCleanupStatus = 'Success'
        } else {
            $result.RegCleanupStatus = 'PartialFailure'
        }
    }
    catch {
        $result.RegCleanupStatus = 'Error'
        $result.Error            = $_.Exception.Message
    }

    return $result

} -ThrottleLimit $ThrottleLimit

$results | Export-Csv -Path $LogPath -NoTypeInformation

# Summary
Write-Host "`n=== Registry Cleanup Summary ===" -ForegroundColor Green
$results | Group-Object RegCleanupStatus | Select-Object Name, Count | Format-Table -AutoSize

Write-Host "=== Qualys Scan Trigger Summary ===" -ForegroundColor Green
$results | Group-Object QualysScanTrigger | Select-Object Name, Count | Format-Table -AutoSize

Write-Host "Log: $LogPath" -ForegroundColor Cyan
