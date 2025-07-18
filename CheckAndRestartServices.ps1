<#
.SYNOPSIS
    Retrieves Windows 10 and 11 workstation computers from Active Directory, checks if they are alive, 
    verifies specified services, restarts them if stopped, and logs stop events. Outputs results to CSV.

.DESCRIPTION
    This script uses PowerShell 7 to query Active Directory for computers running Windows 10 or 11.
    It performs a quick ping check to determine if each computer is alive. For alive computers, it checks 
    the status of specified services. If a service is stopped, it attempts to start it, queries the System 
    event log for the most recent stop event, and records the details. Results are exported to a CSV file 
    with columns: ComputerName, OS, IsAlive, ServiceName, ServiceStatus, Service Stopped on, 
    Service stopped By or reason, Service Successfully restarted.
    
    The script uses parallel processing for efficiency with large numbers of computers (2000-2500).
    Logging occurs to C:\temp\scripts\logs\servicecheck.log for errors.
    
    Reference: https://learn.microsoft.com/en-us/powershell/module/addsadministration/get-adcomputer?view=win10-ps
    Reference: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-connection?view=powershell-7.4
    Reference: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/get-service?view=powershell-7.4
    Reference: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-service?view=powershell-7.4
    Reference: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershelldiagnostics/get-winevent?view=powershell-7.4

.PARAMETER ServiceNames
    An array of service names to check and manage on each computer.

.PARAMETER OutputPath
    The path to the output CSV file. Defaults to C:\temp\scripts\service_report.csv.

.EXAMPLE
    .\CheckAndRestartServices.ps1 -ServiceNames 'wuauserv', 'bits' -OutputPath 'C:\temp\scripts\report.csv'
    Checks the 'wuauserv' and 'bits' services on all Windows 10/11 workstations and outputs to the specified CSV.

.NOTES
    Requires administrative privileges to install RSAT-AD-PowerShell if not present.
    Assumes PS Remoting is enabled on target computers for service and event log queries.
    Runs in PowerShell 7.5 or later for ForEach-Object -Parallel support.
    Uses environment variables where applicable, e.g., $env:COMPUTERNAME for local context if needed.
    Change Log:
    - Version 1.0: Initial creation.
    - Version 1.1: Added OS column to output. Set N/A explicitly for non-stopped services. Added handling for no stop event found.
    - Version 1.2: Updated service operations to use Invoke-Command for compatibility with PowerShell 7, as -ComputerName is not supported in Get-Service and Start-Service.
    - Version 1.3: Replaced $using:service with -ArgumentList and param in Invoke-Command ScriptBlocks to resolve variable scoping issues in parallel execution.
    - Version 1.4: Implemented concurrent logging using ConcurrentBag to fix logging issues in parallel execution. Sanitized reason field to replace newlines and prevent empty lines in CSV. Connection failures are handled in try-catch.

#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ServiceNames,

    [Parameter()]
    [ValidateScript({ if (Test-Path (Split-Path $_ -Parent) -PathType Container) { $true } else { throw "Parent directory does not exist." } })]
    [string]$OutputPath = 'C:\temp\scripts\service_report.csv'
)

# Error handling preferences
$ErrorActionPreference = 'Stop'

#region Initialization
# Working directories
$scriptDir = 'C:\temp\scripts'
$logDir = "$scriptDir\logs"
$logPath = "$logDir\servicecheck.log"

# Create directories if not exist
if (-not (Test-Path $scriptDir -PathType Container)) {
    New-Item -Path $scriptDir -ItemType Directory | Out-Null
}
if (-not (Test-Path $logDir -PathType Container)) {
    New-Item -Path $logDir -ItemType Directory | Out-Null
}

# Check and install ActiveDirectory module if not available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    try {
        Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeManagementTools
    } catch {
        Add-Content -Path $logPath -Value "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Failed to install RSAT-AD-PowerShell: $($_.Exception.Message)"
        throw "Failed to install Active Directory module: $($_.Exception.Message)"
    }
}

# Import module
Import-Module ActiveDirectory

# Retrieve computers
try {
    $computers = Get-ADComputer -Filter '(OperatingSystem -like "Windows 10*") -or (OperatingSystem -like "Windows 11*")' -Properties OperatingSystem |
                 Select-Object Name, OperatingSystem |
                 Sort-Object Name
    Write-Verbose "Retrieved $($computers.Count) computers from Active Directory."
} catch {
    Add-Content -Path $logPath -Value "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Failed to retrieve AD computers: $($_.Exception.Message)"
    throw "Failed to retrieve AD computers: $($_.Exception.Message)"
}
#endregion

#region Processing
# Create a concurrent bag for log messages
$logMessages = New-Object System.Collections.Concurrent.ConcurrentBag[string]

# Process computers in parallel
$results = $computers | ForEach-Object -Parallel {
    $computer = $_.Name
    $os = $_.OperatingSystem
    $serviceNames = $using:ServiceNames
    $logMessages = $using:logMessages

    # Local results collection
    $localResults = @()

    # Check if alive (quick ping)
    $isAlive = Test-Connection -ComputerName $computer -Count 1 -Quiet -ErrorAction SilentlyContinue

    foreach ($service in $serviceNames) {
        if (-not $isAlive) {
            $localResults += [PSCustomObject]@{
                ComputerName                  = $computer
                OS                            = $os
                IsAlive                       = $false
                ServiceName                   = $service
                ServiceStatus                 = 'N/A'
                'Service Stopped on'          = 'N/A'
                'Service stopped By or reason' = 'N/A'
                'Service Successfully restarted' = 'N/A'
            }
            continue
        }

        try {
            # Get service using Invoke-Command
            $serv = Invoke-Command -ComputerName $computer -ScriptBlock { param($svcName) Get-Service -Name $svcName -ErrorAction Stop } -ArgumentList $service
            $initialStatus = $serv.Status
            $wasStopped = $initialStatus -eq 'Stopped'
            $stoppedOn = 'N/A'
            $reason = 'N/A'
            $success = 'N/A'

            if ($wasStopped) {
                $stoppedOn = ''
                $reason = ''

                # Query most recent stop event
                $eventFilter = @{
                    LogName      = 'System'
                    ProviderName = 'Microsoft-Windows-Service Control Manager'
                    ID           = 7036
                }
                $events = Get-WinEvent -ComputerName $computer -FilterHashtable $eventFilter -MaxEvents 100 -ErrorAction Stop |
                          Where-Object { $_.Message -match $service -and $_.Message -match 'stopped state' } |
                          Select-Object -First 1

                if ($events) {
                    $stoppedOn = $events.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
                    $reason = $events.Message.Trim() -replace "`r`n|`n|`r", ' '
                } else {
                    $stoppedOn = 'No event found'
                    $reason = 'No stop event found in logs'
                }

                # Attempt to start service using Invoke-Command
                try {
                    Invoke-Command -ComputerName $computer -ScriptBlock { param($svcName) Start-Service -Name $svcName -ErrorAction Stop } -ArgumentList $service
                    $success = $true
                } catch {
                    $success = $false
                    $logMessages.Add("[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Failed to start ${service} on ${computer}: $($_.Exception.Message)")
                }
            }

            # Get final status using Invoke-Command
            $finalStatus = (Invoke-Command -ComputerName $computer -ScriptBlock { param($svcName) Get-Service -Name $svcName -ErrorAction Stop } -ArgumentList $service).Status

        } catch {
            $finalStatus = 'Error'
            $stoppedOn = 'N/A'
            $reason = $_.Exception.Message -replace "`r`n|`n|`r", ' '
            $success = 'N/A'
            $logMessages.Add("[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Error processing ${service} on ${computer}: $($_.Exception.Message)")
        }

        $localResults += [PSCustomObject]@{
            ComputerName                  = $computer
            OS                            = $os
            IsAlive                       = $isAlive
            ServiceName                   = $service
            ServiceStatus                 = $finalStatus
            'Service Stopped on'          = $stoppedOn
            'Service stopped By or reason' = $reason
            'Service Successfully restarted' = $success
        }
    }

    # Return local results
    $localResults

} -ThrottleLimit 50

# Write collected log messages to file
$logMessages | ForEach-Object { Add-Content -Path $logPath -Value $_ }
#endregion

#region Output
# Export to CSV
try {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Verbose "Exported results to $OutputPath"
} catch {
    Add-Content -Path $logPath -Value "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] Failed to export CSV: $($_.Exception.Message)"
    throw "Failed to export CSV: $($_.Exception.Message)"
}
#endregion
