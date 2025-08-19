<#
.SYNOPSIS
    Queries Qualys for failed compliance postures, generates reports, and optionally remediates issues using per-CID fixer scripts.

.DESCRIPTION
    This script connects to the Qualys API using a service account, retrieves failed compliance data for a specified policy, parses the CSV output, groups by host and CID, generates a report, and (if -Remediate is specified) attempts fixes on online hosts via remote PowerShell. Fixers are discovered dynamically from the 'fixers' subfolder. Designed for scheduled runs; logs all actions.

.PARAMETER QualysBaseUrl
    The base URL for the Qualys API (e.g., 'https://qualysapi.qualys.eu').

.PARAMETER QualysCredential
    PSCredential object for Qualys API authentication (username and password).

.PARAMETER PolicyId
    The Qualys policy ID to query (default: 99999).

.PARAMETER TruncationLimit
    Maximum records to retrieve per API call (default: 10000; set to 0 for no limit, but use caution).

.PARAMETER AdminCredential
    PSCredential for remote PowerShell execution on target hosts.

.PARAMETER Remediate
    Switch to enable remediation (default: $false; dry-run mode logs what would happen).

.EXAMPLE
    # Interactive run with remediation
    $qualysCred = Get-Credential -Message 'Qualys API Credentials'
    $adminCred = Get-Credential -Message 'Admin Credentials for Remoting'
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential $qualysCred -AdminCredential $adminCred -Remediate

.EXAMPLE
    # Scheduled dry-run (report only)
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential (Import-Clixml 'C:\secure\qualys.cred') -PolicyId 99999

.NOTES
    Requires PowerShell 7.4+ for optimal performance (e.g., -Parallel). If modules like MicrosoftDefender are missing, they are installed automatically.
    Working directory: C:\temp\scripts
    Logs: C:\temp\scripts\logs\CheckandRemediate_<yyyyMMdd_HHmm>.log
    Reports: C:\temp\scripts\reports\FailedCompliance_<yyyyMMdd_HHmm>.csv
    Dependencies: Invoke-RestMethod, Import-Csv, Invoke-Command.
    Changelog: Initial version - August 19, 2025. Updated August 19, 2025: Fixed variable interpolation error in logging; removed -WhatIf support. Updated August 19, 2025: Added 'X-Requested-With' header to API requests for compliance with Qualys requirements.
    Signed by Marinus van Deventer
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$QualysBaseUrl,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$QualysCredential,

    [Parameter()]
    [int]$PolicyId = 99999,

    [Parameter()]
    [int]$TruncationLimit = 10000,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$AdminCredential,

    [switch]$Remediate
)

begin {
    # Set error action and create directories
    $ErrorActionPreference = 'Stop'
    $workingDir = 'C:\temp\scripts'
    $logsDir = Join-Path $workingDir 'logs'
    $reportsDir = Join-Path $workingDir 'reports'
    $fixersDir = Join-Path $workingDir 'fixers'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $logPath = Join-Path $logsDir "CheckandRemediate_$timestamp.log"
    $reportPath = Join-Path $reportsDir "FailedCompliance_$timestamp.csv"
    $csvPath = Join-Path $workingDir 'failed_postures.csv'

    foreach ($dir in @($workingDir, $logsDir, $reportsDir, $fixersDir)) {
        if (-not (Test-Path $dir -PathType Container)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
    }

    # Function to log messages
    function Write-Log {
        param (
            [string]$Message,
            [string]$Level = 'INFO'
        )
        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        Write-Information $logEntry -InformationAction Continue
        Add-Content -Path $logPath -Value $logEntry
    }

    # Connect to Qualys API and retrieve data
    function Get-FailedPostures {
        try {
            $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($QualysCredential.UserName):$($QualysCredential.GetNetworkCredential().Password)"))
            $headers = @{
                'Authorization' = "Basic $auth"
                'Accept'        = 'text/csv'
                'X-Requested-With' = 'PowerShell'
            }
            $uri = "$QualysBaseUrl/api/2.0/fo/compliance/posture/info/?action=list&policy_id=$PolicyId&output_format=CSV_NO_METADATA&status=Failed&details=All&truncation_limit=$TruncationLimit"
            Write-Log "Querying Qualys API: $uri"
            Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -OutFile $csvPath
            $data = Import-Csv -Path $csvPath
            if ($data.Count -eq 0) {
                Write-Log 'No failed postures found.'
                return @()
            }
            Write-Log "Retrieved $($data.Count) failed postures."
            return $data
        } catch {
            Write-Log "API query failed: $($_.Exception.Message)" 'ERROR'
            throw
        }
    }

    # Discover fixer scripts for a CID
    function Discover-Fixers {
        param ([int]$CID)
        $fixerFiles = Get-ChildItem -Path $fixersDir -Filter "$CID-*.ps1" -File
        if ($fixerFiles.Count -eq 0) {
            Write-Log "No fixer found for CID $CID."
            return $null
        }
        # Assume one per CID; take first if multiple
        $fixerPath = $fixerFiles[0].FullName
        Write-Log "Discovered fixer: $fixerPath for CID $CID."
        return [scriptblock]::Create((Get-Content $fixerPath -Raw))
    }

    # Remediate a host for a specific CID
    function Remediate-Host {
        param (
            [string]$HostIP,
            [int]$CID,
            [object]$Evidence
        )
        $outcome = [PSCustomObject]@{
            HostIP     = $HostIP
            CID        = $CID
            FixAttempted = 'No'
            Outcome    = 'Not Attempted'
            Details    = ''
        }

        if (-not (Test-Connection -ComputerName $HostIP -Count 1 -Quiet)) {
            $outcome.Outcome = 'Host Offline'
            $outcome.Details = 'Skipping remediation.'
            Write-Log "Host $HostIP offline for CID $CID." 'WARNING'
            return $outcome
        }

        $fixerBlock = Discover-Fixers -CID $CID
        if (-not $fixerBlock) {
            $outcome.Outcome = 'No Fixer Available'
            return $outcome
        }

        if ($Remediate) {
            try {
                $result = Invoke-Command -ComputerName $HostIP -Credential $AdminCredential -ScriptBlock $fixerBlock -ErrorAction Stop
                $outcome.FixAttempted = 'Yes'
                $outcome.Outcome = $result.Outcome
                $outcome.Details = $result.Details
                Write-Log "Remediation on $HostIP for CID $CID: $($outcome.Outcome)"
            } catch {
                $outcome.Outcome = 'Failed'
                $outcome.Details = $_.Exception.Message
                Write-Log "Remediation failed on $HostIP for CID ${CID}: $($outcome.Details)" 'ERROR'
            }
        } else {
            Write-Log "Dry-run: Would attempt remediation on $HostIP for CID $CID."
        }

        return $outcome
    }

    Write-Log 'Script started.'
}

process {
    try {
        $postures = Get-FailedPostures
        if ($postures.Count -eq 0) { return }

        # Group by host IP and CID
        $grouped = $postures | Group-Object -Property IP | ForEach-Object {
            $hostIP = $_.Name
            $_.Group | Group-Object -Property CID | ForEach-Object {
                [PSCustomObject]@{
                    HostIP   = $hostIP
                    CID      = [int]$_.Name
                    Evidence = $_.Group.Evidence -join '; '
                    Reason   = $_.Group.Reason -join '; '
                }
            }
        }

        # Process in parallel if PS7+
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            $results = $grouped | ForEach-Object -Parallel {
                Import-Module -Name $using:workingDir -ErrorAction SilentlyContinue # Re-import if needed
                $remediation = Remediate-Host -HostIP $_.HostIP -CID $_.CID -Evidence $_.Evidence
                [PSCustomObject]@{
                    HostIP    = $_.HostIP
                    CID       = $_.CID
                    Evidence  = $_.Evidence
                    Reason    = $_.Reason
                    FixAttempted = $remediation.FixAttempted
                    Outcome   = $remediation.Outcome
                    Details   = $remediation.Details
                }
            } -ThrottleLimit 5
        } else {
            $results = $grouped | ForEach-Object {
                $remediation = Remediate-Host -HostIP $_.HostIP -CID $_.CID -Evidence $_.Evidence
                [PSCustomObject]@{
                    HostIP    = $_.HostIP
                    CID       = $_.CID
                    Evidence  = $_.Evidence
                    Reason    = $_.Reason
                    FixAttempted = $remediation.FixAttempted
                    Outcome   = $remediation.Outcome
                    Details   = $remediation.Details
                }
            }
        }

        $results | Export-Csv -Path $reportPath -NoTypeInformation
        Write-Log "Report generated: $reportPath"
    } catch {
        Write-Log "Process failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

end {
    if (Test-Path $csvPath) { Remove-Item $csvPath -Force }
    Write-Log 'Script completed.'
}

# Signed by Marinus van Deventer
