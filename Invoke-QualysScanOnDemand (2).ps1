<#
.SYNOPSIS
    Triggers a Qualys Cloud Agent vulnerability scan on one or more computers.

.DESCRIPTION
    Writes the ScanOnDemand registry values that the Qualys Cloud Agent monitors.
    Uses PowerShell Remoting because Remote Registry is disabled.

    Reference:
    https://docs.qualys.com/en/ca/install-guide/windows/configuration/configure_scan_on_demand.htm

.PARAMETER ComputerName
    One or more target computers.

.PARAMETER InputFile
    Path to a text file containing one computer name per line.

.EXAMPLE
    .\Invoke-QualysScanOnDemand.ps1 -ComputerName PC001

.EXAMPLE
    .\Invoke-QualysScanOnDemand.ps1 -ComputerName PC001,PC002,PC003

.EXAMPLE
    .\Invoke-QualysScanOnDemand.ps1 -InputFile .\hosts.txt
#>

[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory, Position = 0)]
    [string[]] $ComputerName,

    [Parameter(ParameterSetName = 'ByFile', Mandatory)]
    [string] $InputFile
)

if ($PSCmdlet.ParameterSetName -eq 'ByFile') {
    $ComputerName = Get-Content -LiteralPath $InputFile |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }
}

Invoke-Command -ComputerName $ComputerName -ScriptBlock {
    $key = 'HKLM:\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability'

    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }

    New-ItemProperty -Path $key -Name 'CpuLimit'     -PropertyType DWord -Value 100 -Force | Out-Null
    New-ItemProperty -Path $key -Name 'ScanOnDemand' -PropertyType DWord -Value 1   -Force | Out-Null

    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        Triggered    = $true
    }
}
