<#
.SYNOPSIS
    Queries Qualys for failed compliance postures, generates reports, and optionally remediates issues using per-CID fixer scripts.
.DESCRIPTION
    Connects to the Qualys API using a service account, retrieves failed compliance data for a specified policy, parses the CSV output with hardcoded columns, adds remediation columns to the original data, generates a report, and (if -Remediate is specified) attempts fixes on online hosts via remote PowerShell. Fixers are discovered dynamically from the 'fixers' subfolder.
.PARAMETER QualysBaseUrl
    The base URL for the Qualys API (e.g., 'https://qualysapi.qualys.eu').
.PARAMETER QualysCredential
    PSCredential object for Qualys API authentication (username and password).
.PARAMETER PolicyId
    The Qualys policy ID to query (default: 99999).
.PARAMETER TruncationLimit
    Maximum records to retrieve per API call (default: 10000; set to 0 for no limit).
.PARAMETER AdminCredential
    PSCredential for remote PowerShell execution on target hosts. Mandatory only when -Remediate is used.
.PARAMETER Remediate
    Switch to enable remediation (default: $false; report-only mode).
.EXAMPLE
    $qualysCredential = Get-Credential -Message 'Qualys API Credentials'
    $adminCredential = Get-Credential -Message 'Admin Credentials for Remoting'
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential $qualysCredential -AdminCredential $adminCredential -Remediate
.EXAMPLE
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential (Import-Clixml 'C:\secure\qualys.cred') -PolicyId 99999
.NOTES
    Requires PowerShell 7.4+ for optimal performance. If modules are not available install them.
    Working directory: C:\temp\scripts
    Logs: C:\temp\scripts\logs\CheckandRemediate_<yyyyMMdd_HHmm>.log
    Reports: C:\temp\scripts\reports\FailedCompliance_<yyyyMMdd_HHmm>.csv
    Dependencies: Invoke-RestMethod, Import-Csv, Invoke-Command.
    Changelog: Initial version - 2025-08-19. Updated 2025-08-19: Added UTF8 encoding for logs and reports, hardcoded column names from guide, fixed variable interpolation in logs. Updated 2025-08-19: Changed to modify original data by adding columns instead of regrouping for report, group only for remediation to avoid duplicates.
    Signed by Marinus van Deventer
#>

[CmdletBinding(DefaultParameterSetName = 'Report')]
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

    [Parameter(Mandatory = $true, ParameterSetName = 'Remediate')]
    [System.Management.Automation.PSCredential]$AdminCredential,

    [Parameter(ParameterSetName = 'Remediate')]
    [switch]$Remediate
)

begin {
    #region Initialisation
    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $workingDirectory = 'C:\temp\scripts'
    $logsDirectory = Join-Path $workingDirectory 'logs'
    $reportsDirectory = Join-Path $workingDirectory 'reports'
    $fixersDirectory = Join-Path $workingDirectory 'fixers'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $logPath = Join-Path $logsDirectory "CheckandRemediate_$timestamp.log"
    $reportPath = Join-Path $reportsDirectory "FailedCompliance_$timestamp.csv"
    $csvPath = Join-Path $workingDirectory 'failed_postures.csv'

    foreach ($directory in @($workingDirectory, $logsDirectory, $reportsDirectory, $fixersDirectory)) {
        if (-not (Test-Path $directory -PathType Container)) {
            New-Item -Path $directory -ItemType Directory -Force | Out-Null
        }
    }

    function Write-Log {
        param (
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [ValidateSet('INFO', 'WARNING', 'ERROR')]
            [string]$Level = 'INFO'
        )
        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        Write-Information $logEntry -InformationAction Continue
        Add-Content -Path $logPath -Value $logEntry -Encoding UTF8
    }

    Write-Log "Script started. ParameterSetName=$($PSCmdlet.ParameterSetName)"
    #endregion

    #region Helpers
    function Invoke-QualysRequest {
        param (
            [Parameter(Mandatory = $true)]
            [hashtable]$Body,

            [string]$OutFile = $csvPath
        )
        try {
            $authenticationString = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($QualysCredential.UserName):$($QualysCredential.GetNetworkCredential().Password)"))
            $headers = @{
                'Authorization' = "Basic $authenticationString"
                'Accept'        = 'text/csv'
                'X-Requested-With' = 'PowerShell'
            }
            $uri = "$QualysBaseUrl/api/2.0/fo/compliance/posture/info/"
            Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $Body -OutFile $OutFile
        } catch {
            Write-Log "Qualys request failed: $($_.Exception.Message)" 'ERROR'
            throw
        }
    }

    function Get-FailedPostures {
        $body = @{
            action = 'list'
            policy_id = $PolicyId
            output_format = 'csv_no_metadata'
            status = 'Failed'
            details = 'All'
        }
        if ($TruncationLimit -gt 0) {
            $body.truncation_limit = $TruncationLimit
        }
        Invoke-QualysRequest -Body $body
        $data = Import-Csv -Path $csvPath
        if ($data.Count -eq 0) {
            Write-Log "No failed postures found."
            return @()
        }
        # Validate required columns
        $requiredColumns = @('IP', 'Control ID', 'Posture Evidence', 'Reference')
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $data[0].PSObject.Properties.Name }
        if ($missingColumns) {
            Write-Log "Missing required columns: $($missingColumns -join ', ')" 'WARNING'
        }
        Write-Log "Retrieved $($data.Count) failed postures."
        return $data
    }

    function Discover-Fixer {
        param (
            [Parameter(Mandatory = $true)]
            [int]$ControlId
        )
        $fixerPath = Get-ChildItem -Path $fixersDirectory -Filter "$ControlId-*.ps1" -File | Select-Object -First 1 -ExpandProperty FullName
        if (-not $fixerPath) {
            Write-Log "No fixer found for Control ID $ControlId."
            return $null
        }
        Write-Log "Discovered fixer: $fixerPath for Control ID $ControlId."
        $scriptContent = Get-Content -Path $fixerPath -Raw
        return [scriptblock]::Create($scriptContent)
    }

    function Remediate-Host {
        param (
            [Parameter(Mandatory = $true)]
            [string]$HostIp,

            [Parameter(Mandatory = $true)]
            [int]$ControlId,

            [Parameter(Mandatory = $true)]
            [string]$Evidence
        )
        $outcomeObject = [PSCustomObject]@{
            HostIp = $HostIp
            ControlId = $ControlId
            FixAttempted = 'No'
            Outcome = 'Not Attempted'
            Details = ''
        }

        if (-not (Test-Connection -ComputerName $HostIp -Count 1 -Quiet)) {
            $outcomeObject.Outcome = 'Host Offline'
            $outcomeObject.Details = 'Skipping remediation.'
            Write-Log "Host $HostIp offline for Control ID $ControlId." 'WARNING'
            return $outcomeObject
        }

        $fixerBlock = Discover-Fixer -ControlId $ControlId
        if (-not $fixerBlock) {
            $outcomeObject.Outcome = 'No Fixer Available'
            return $outcomeObject
        }

        if ($Remediate.IsPresent) {
            try {
                $result = Invoke-Command -ComputerName $HostIp -Credential $AdminCredential -ScriptBlock $fixerBlock -ArgumentList $Evidence -ErrorAction Stop
                $outcomeObject.FixAttempted = 'Yes'
                $outcomeObject.Outcome = $result.Outcome
                $outcomeObject.Details = $result.Details
                Write-Log "Remediation on $HostIp for Control ID $ControlId: $($outcomeObject.Outcome)"
            } catch {
                $outcomeObject.Outcome = 'Failed'
                $outcomeObject.Details = $_.Exception.Message
                Write-Log "Remediation failed on $HostIp for Control ID $ControlId: $($outcomeObject.Details)" 'ERROR'
            }
        } else {
            Write-Log "Dry-run: Would attempt remediation on $HostIp for Control ID $ControlId."
        }

        return $outcomeObject
    }
    #endregion
}

process {
    try {
        $postures = Get-FailedPostures
        if ($postures.Count -eq 0) { return }

        # Filter invalid rows
        $postures = $postures | Where-Object { -not [string]::IsNullOrEmpty($_.IP) -and $_. 'Control ID' -match '^\d+$' }

        # Group by IP and Control ID for remediation (to avoid duplicates)
        $grouped = $postures | Group-Object -Property IP | ForEach-Object {
            $hostIp = $_.Name
            $_.Group | Group-Object -Property 'Control ID' | ForEach-Object {
                [PSCustomObject]@{
                    HostIp = $hostIp
                    ControlId = [int]$_.Name
                    Evidence = $_.Group.'Posture Evidence' -join '; '
                    Reason = $_.Group.Reference -join '; '
                }
            }
        }

        $results = @()
        foreach ($item in $grouped) {
            $remediation = Remediate-Host -HostIp $item.HostIp -ControlId $item.ControlId -Evidence $item.Evidence
            $postures | Where-Object { $_.IP -eq $item.HostIp -and $_. 'Control ID' -eq $item.ControlId } | ForEach-Object {
                $_ | Add-Member -NotePropertyName FixAttempted -NotePropertyValue $remediation.FixAttempted
                $_ | Add-Member -NotePropertyName Outcome -NotePropertyValue $remediation.Outcome
                $_ | Add-Member -NotePropertyName Details -NotePropertyValue $remediation.Details
                $results += $_
            }
        }

        $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Report generated at $reportPath."
    } catch {
        Write-Log "Process failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

end {
    if (Test-Path $csvPath) { Remove-Item $csvPath -Force }
    Write-Log "Script completed."
}

# Signed by Marinus van Deventer
