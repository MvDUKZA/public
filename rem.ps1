<#
.SYNOPSIS
    Removes \\KAK.local\NETLOGON and \\KAK.local\SYSVOL values from 
    HKLM\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths
    across a list of machines, without using Remote Registry service.
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
        ComputerName = $computer
        Online       = $false
        NetlogonHad  = $null
        NetlogonNow  = $null
        SysvolHad    = $null
        SysvolNow    = $null
        Status       = ''
        Error        = ''
        Timestamp    = Get-Date -Format 's'
    }

    # Quick reachability test — avoids long timeouts on dead machines
    if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet -TimeoutSeconds 2)) {
        $result.Status = 'Offline'
        return $result
    }
    $result.Online = $true

    try {
        # Invoke-Command uses WinRM, NOT Remote Registry — works as long as WinRM is up
        $scriptResult = Invoke-Command -ComputerName $computer -ErrorAction Stop -ScriptBlock {
            $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkProvider\HardenedPaths"
            $targets = @('\\KAK.local\NETLOGON', '\\KAK.local\SYSVOL')
            $out = @{}

            foreach ($name in $targets) {
                $before = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue).$name
                $out["$name`_Before"] = if ($null -ne $before) { $true } else { $false }

                if ($null -ne $before) {
                    Remove-ItemProperty -Path $path -Name $name -Force -ErrorAction Stop
                }

                $after = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue).$name
                $out["$name`_After"] = if ($null -ne $after) { $true } else { $false }
            }
            return $out
        }

        $result.NetlogonHad = $scriptResult['\\KAK.local\NETLOGON_Before']
        $result.NetlogonNow = $scriptResult['\\KAK.local\NETLOGON_After']
        $result.SysvolHad   = $scriptResult['\\KAK.local\SYSVOL_Before']
        $result.SysvolNow   = $scriptResult['\\KAK.local\SYSVOL_After']

        if (-not $result.NetlogonNow -and -not $result.SysvolNow) {
            $result.Status = 'Success'
        } else {
            $result.Status = 'PartialFailure'
        }
    }
    catch {
        $result.Status = 'Error'
        $result.Error  = $_.Exception.Message
    }

    return $result

} -ThrottleLimit $ThrottleLimit

$results | Export-Csv -Path $LogPath -NoTypeInformation

# Summary
$summary = $results | Group-Object Status | Select-Object Name, Count
Write-Host "`n=== Summary ===" -ForegroundColor Green
$summary | Format-Table -AutoSize
Write-Host "Log: $LogPath" -ForegroundColor Cyan
