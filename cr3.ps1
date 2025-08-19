<#
.SYNOPSIS
    Queries Qualys for failed compliance postures, generates reports, and optionally remediates issues using per-CID fixer scripts.

.DESCRIPTION
    Connects to the Qualys API using a service account, retrieves failed compliance data for a specified policy, parses the CSV output,
    normalizes columns to canonical names, groups by host and CID, generates a report, and (if -Remediate is specified) attempts fixes
    on online hosts via remote PowerShell. Fixers are discovered dynamically from the 'fixers' subfolder and must be Authenticode-signed.
    Optionally, upon successful remediation, can request a Qualys compliance rescan for the host when -LaunchRescan is provided.

.PARAMETER QualysBaseUrl
    The base URL for the Qualys API (e.g., 'https://qualysapi.qualys.eu').

.PARAMETER QualysCredential
    PSCredential object for Qualys API authentication (username and password).

.PARAMETER PolicyId
    The Qualys policy ID to query (default: 99999).

.PARAMETER TruncationLimit
    Maximum records to retrieve per API call (default: 10000; set to 0 to omit the parameter).

.PARAMETER AdminCredential
    PSCredential for remote PowerShell execution on target hosts. Mandatory only when -Remediate is used.

.PARAMETER Remediate
    Switch to enable remediation (default: off; report-only). Requires -AdminCredential.

.PARAMETER LaunchRescan
    When used with -Remediate, requests a Qualys compliance rescan for hosts where a fixer reported Success.

.PARAMETER WorkingDirectory
    Base working directory for scripts, logs, reports, and fixers.

.EXAMPLE
    $qualysCred = Get-Credential -Message 'Qualys API Credentials'
    $adminCred = Get-Credential -Message 'Admin Credentials for Remoting'
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential $qualysCred -AdminCredential $adminCred -Remediate -LaunchRescan

.EXAMPLE
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential (Import-Clixml 'C:\secure\qualys.cred') -PolicyId 99999

.NOTES
    Requires PowerShell 7.4+ (tested 7.5.2).
    Logs and reports are stored in timestamped files in their respective subdirectories.
    Fixers must be named in the format: {CID}-{DescriptiveName}.ps1 (e.g., 12345-DisableWeakCipher.ps1)
    Fixers must be Authenticode signed and return an object with Outcome and Details properties.

    Changelog (latest first):
      - 2025-08-19: Removed ZIP extraction logic as requested
      - 2025-08-19: Enhanced header resolution based on actual Qualys CSV format
      - 2025-08-19: Added configurable working directory parameter
      - 2025-08-19: Improved error handling and retry logic
      - 2025-08-19: Enhanced logging with rotation and configurable levels
      - 2025-08-19: Added input validation and parameter checking
      - 2025-08-19: FIX PowerShell 7.5.2 encoding error. Improved UTF-8 text handling.
      - 2025-08-19: CSV wrapper handling retained; header resolver requires IP+ControlID only; Evidence/Reason optional.
      - 2025-08-19: Better diagnostics if schema cannot be resolved; PS 7.5.2 WSMan import retained.

    Signed by Marinus van Deventer
#>

[CmdletBinding(DefaultParameterSetName = 'Report', SupportsShouldProcess = $false)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$QualysBaseUrl,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$QualysCredential,

    [Parameter()]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$PolicyId = 99999,

    [Parameter()]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$TruncationLimit = 10000,

    [Parameter(Mandatory = $true, ParameterSetName = 'Remediate')]
    [System.Management.Automation.PSCredential]$AdminCredential,

    [Parameter(ParameterSetName = 'Remediate')]
    [switch]$Remediate,

    [Parameter(ParameterSetName = 'Remediate')]
    [switch]$LaunchRescan,

    [Parameter()]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$WorkingDirectory = 'C:\temp\scripts'
)

begin {
    #region Initialisation
    $ErrorActionPreference = 'Stop'
    $InformationPreference = 'Continue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    try { 
        Import-Module Microsoft.WSMan.Management -ErrorAction SilentlyContinue 
    } catch {
        Write-Warning "Failed to import Microsoft.WSMan.Management module: $($_.Exception.Message)"
    }

    # Validate parameters
    if ($Remediate -and -not $AdminCredential) {
        throw "AdminCredential parameter is required when using Remediate switch"
    }

    if ($LaunchRescan -and -not $Remediate) {
        throw "LaunchRescan can only be used with the Remediate switch"
    }

    # Directory setup
    $logsDir = Join-Path $WorkingDirectory 'logs'
    $reportsDir = Join-Path $WorkingDirectory 'reports'
    $fixersDir = Join-Path $WorkingDirectory 'fixers'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $logPath = Join-Path $logsDir "CheckandRemediate_$timestamp.log"
    $reportPath = Join-Path $reportsDir "FailedCompliance_$timestamp.csv"
    $csvRawPath = Join-Path $env:TEMP "failed_postures_raw_$([guid]::NewGuid().Guid).csv"
    $csvPath = Join-Path $env:TEMP "failed_postures_$([guid]::NewGuid().Guid).csv"

    foreach ($dir in @($WorkingDirectory, $logsDir, $reportsDir, $fixersDir)) {
        if (-not (Test-Path $dir -PathType Container)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
            Write-Information "Created directory: $dir"
        }
    }

    # Logging function
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)][string]$Message,
            [ValidateSet('INFO','WARNING','ERROR','DEBUG')][string]$Level = 'INFO'
        )
        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        
        switch ($Level) {
            'ERROR' { Write-Error $Message }
            'WARNING' { Write-Warning $Message }
            'DEBUG' { Write-Debug $Message }
            default { Write-Information $Message }
        }
        
        try {
            Add-Content -Path $logPath -Value $logEntry -Encoding UTF8 -ErrorAction Stop
        } catch {
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }

    Write-Log "Script started. PSVersion=$($PSVersionTable.PSVersion) ParamSet=$($PSCmdlet.ParameterSetName)"
    Write-Log "Working directory: $WorkingDirectory"
    Write-Log "Log file: $logPath"
    Write-Log "Report file: $reportPath"
    #endregion

    #region Helpers: Qualys, CSV extraction/normalizer, Fixers and Rescans
    function Invoke-QualysRequest {
        param(
            [Parameter(Mandatory = $true)][string]$EndpointPath,
            [Parameter(Mandatory = $true)][hashtable]$Body,
            [Parameter()][string]$Accept = 'text/csv',
            [Parameter()][string]$OutFile
        )
        
        $uri = ($QualysBaseUrl.TrimEnd('/')) + $EndpointPath
        $auth = [Convert]::ToBase64String(
            [Text.Encoding]::ASCII.GetBytes(
                "$($QualysCredential.UserName):$($QualysCredential.GetNetworkCredential().Password)"
            )
        )
        
        $headers = @{
            'Authorization'    = "Basic $auth"
            'Accept'           = $Accept
            'X-Requested-With' = 'PowerShell'
        }

        $maxAttempts = 3
        $attempt = 0
        $delaySeconds = 5
        
        do {
            $attempt++
            try {
                Write-Log "Qualys POST $EndpointPath (attempt $attempt of $maxAttempts)" 'DEBUG'
                
                $irmParams = @{
                    Uri = $uri
                    Headers = $headers
                    Method = 'Post'
                    ContentType = 'application/x-www-form-urlencoded'
                    Body = $Body
                    ErrorAction = 'Stop'
                }
                
                if ($OutFile) {
                    $irmParams['OutFile'] = $OutFile
                    Invoke-RestMethod @irmParams
                } else {
                    return Invoke-RestMethod @irmParams
                }
                
                Write-Log "Qualys request successful" 'DEBUG'
                break
            } catch {
                if ($attempt -ge $maxAttempts) {
                    Write-Log "Qualys request failed after $maxAttempts attempts: $($_.Exception.Message)" 'ERROR'
                    throw
                }
                
                Write-Log "Qualys request failed: $($_.Exception.Message). Retrying in ${delaySeconds}s..." 'WARNING'
                Start-Sleep -Seconds $delaySeconds
                $delaySeconds = $delaySeconds * 2 # Exponential backoff
            }
        } while ($true)
    }

    function Unwrap-QualysCsv {
        param(
            [Parameter(Mandatory=$true)][string]$RawPath,
            [Parameter(Mandatory=$true)][string]$CsvPath
        )
        
        try {
            # Read the raw response
            $text = Get-Content -Path $RawPath -Raw -Encoding UTF8
            
            if ([string]::IsNullOrWhiteSpace($text)) {
                throw "Empty response body"
            }

            # Strip Qualys wrapper lines and extract CSV content
            $lines = $text -split "`r?`n"
            $cleanLines = $lines | Where-Object { 
                $_ -and ($_ -notmatch '^----BEGIN_') -and ($_ -notmatch '^----END_') 
            }

            # Find the header row (first line that looks like a CSV header)
            $headerIndex = -1
            for ($i = 0; $i -lt $cleanLines.Count; $i++) {
                if ($cleanLines[$i] -like "*,*") {
                    $headerIndex = $i
                    break
                }
            }

            if ($headerIndex -eq -1) {
                throw "No CSV header found in response"
            }

            # Extract from header to end
            $csvContent = $cleanLines[$headerIndex..($cleanLines.Count-1)] -join [Environment]::NewLine
            
            # Write to file
            Set-Content -Path $CsvPath -Value $csvContent -Encoding UTF8 -Force
            Write-Log "Extracted CSV with $($cleanLines.Count - $headerIndex) lines" 'DEBUG'
            
        } catch {
            Write-Log "Failed to process Qualys response: $($_.Exception.Message)" 'ERROR'
            throw
        }
    }

    function Get-FailedPostures {
        try {
            $body = @{
                action        = 'list'
                policy_id     = $PolicyId
                output_format = 'csv_no_metadata'
                status        = 'Failed'
                details       = 'All'
            }
            if ($TruncationLimit -gt 0) { $body.truncation_limit = $TruncationLimit }

            Write-Log "Requesting failed postures for policy $PolicyId"
            Invoke-QualysRequest -EndpointPath '/api/2.0/fo/compliance/posture/info/' -Body $body -OutFile $csvRawPath | Out-Null
            Unwrap-QualysCsv -RawPath $csvRawPath -CsvPath $csvPath

            $raw = Import-Csv -Path $csvPath
            if (-not $raw -or $raw.Count -eq 0) {
                Write-Log 'No failed postures found.'
                return @()
            }

            # Trim BOM and whitespace on both headers and values
            $data = $raw | ForEach-Object {
                $o = [ordered]@{}
                foreach ($p in $_.PSObject.Properties) {
                    $name = ($p.Name -replace "^\uFEFF","").Trim()
                    $val  = if ($null -ne $p.Value) { $p.Value.ToString().Trim() } else { $null }
                    $o[$name] = $val
                }
                [pscustomobject]$o
            }

            Write-Log "Retrieved $($data.Count) failed postures."
            return $data
        } catch {
            Write-Log "API query failed: $($_.Exception.Message)" 'ERROR'
            throw
        }
    }

    function Resolve-Header {
        param([Parameter(Mandatory=$true)][string[]]$Available)

        # Enhanced mapping based on the sample CSV format
        $map = @{
            IP         = @('IP', 'IP Address', 'IP Address(es)', 'Host IP', 'Host IP Address', 'IP Address(es) List')
            ControlID  = @('Control ID', 'CID', 'Control ID (CID)', 'CID (Control ID)', 'Control Identifier', 'Control', 'ControlID')
            Evidence   = @('Evidence', 'Evidence/Results', 'Instance Evidence', 'Evidence Value', 'Actual Value', 'Value', 'Finding', 'Observed Value', 'Posture Evidence')
            Reason     = @('Reason', 'Reason for Failure', 'Failure Reason', 'Reason/Recommendation', 'Rationale', 'Recommendation', 'Expected Value', 'Expected', 'Control Statement')
        }
        
        $resolved = @{}
        foreach ($k in $map.Keys) {
            $alias = $map[$k] | Where-Object { $_ -in $Available }
            if ($alias) { 
                $resolved[$k] = $alias[0] 
                Write-Log "Resolved header '$k' to '$($alias[0])'" 'DEBUG'
            } else {
                Write-Log "No header found for '$k' in available headers: $($Available -join ', ')" 'DEBUG'
            }
        }
        return $resolved
    }

    function Normalize-QualysRows {
        param([Parameter(Mandatory=$true)][object[]]$Rows)

        if (-not $Rows -or $Rows.Count -eq 0) { return @() }
        $headers = $Rows[0].PSObject.Properties.Name
        $resolved = Resolve-Header -Available $headers

        # Only IP and ControlID are required to act; Evidence/Reason are optional.
        $requiredCanon = @('IP','ControlID')
        $missingCanon = $requiredCanon | Where-Object { $_ -notin $resolved.Keys }
        if ($missingCanon) {
            Write-Log ("Detected CSV headers: {0}" -f ($headers -join ', ')) 'ERROR'
            throw "CSV missing required columns (canonical): $($missingCanon -join ', ')."
        }

        $ipH   = $resolved.IP
        $cidH  = $resolved.ControlID
        $evH   = if ($resolved.ContainsKey('Evidence')) { $resolved.Evidence } else { $null }
        $rsH   = if ($resolved.ContainsKey('Reason'))   { $resolved.Reason }   else { $null }

        $Rows | ForEach-Object {
            [pscustomobject]@{
                IP        = $_.$ipH
                ControlID = $_.$cidH
                Evidence  = if ($evH) { $_.$evH } else { '' }
                Reason    = if ($rsH) { $_.$rsH } else { '' }
            }
        }
    }

    function Discover-Fixers {
        param([Parameter(Mandatory = $true)][int]$CID)
        $fixerFiles = Get-ChildItem -Path $fixersDir -Filter "$CID-*.ps1" -File | Sort-Object Name
        if (-not $fixerFiles) {
            Write-Log "No fixer found for CID $CID." 'DEBUG'
            return $null
        }
        
        foreach ($fixerFile in $fixerFiles) {
            $candidate = $fixerFile.FullName
            try {
                $sig = Get-AuthenticodeSignature -FilePath $candidate
                if ($sig.Status -ne 'Valid') {
                    Write-Log "Fixer $candidate has invalid or missing signature: $($sig.Status). Skipping." 'WARNING'
                    continue
                }
                
                Write-Log "Discovered signed fixer: $candidate for CID $CID." 'DEBUG'
                $scriptText = Get-Content $candidate -Raw
                return [scriptblock]::Create($scriptText)
            } catch {
                Write-Log "Error examining fixer $candidate: $($_.Exception.Message)" 'WARNING'
            }
        }
        
        Write-Log "No valid signed fixer found for CID $CID." 'WARNING'
        return $null
    }

    function Invoke-HostRescan {
        param([Parameter(Mandatory = $true)][string]$HostIP)
        try {
            $body = @{
                action        = 'launch'
                scan_title    = "AutoRescan_$($HostIP)_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss'))"
                ip            = $HostIP
                priority      = 'Normal'
            }
            Invoke-QualysRequest -EndpointPath '/api/2.0/fo/scan/compliance/' -Body $body -Accept 'application/xml' | Out-Null
            Write-Log "Requested Qualys compliance rescan for $HostIP."
            return $true
        } catch {
            Write-Log "Failed to request Qualys rescan for $HostIP: $($_.Exception.Message)" 'WARNING'
            return $false
        }
    }

    function Test-HostForRemoting {
        param([Parameter(Mandatory = $true)][string]$HostIP)
        try {
            if (-not (Test-Connection -ComputerName $HostIP -Count 1 -Quiet -ErrorAction Stop)) {
                return 'Host Offline'
            }
            
            Test-WSMan -ComputerName $HostIP -Authentication Default -ErrorAction Stop | Out-Null
            return 'OK'
        } catch {
            return 'WinRM Unreachable'
        }
    }

    function Remediate-Host {
        param (
            [Parameter(Mandatory = $true)][string]$HostIP,
            [Parameter(Mandatory = $true)][int]$CID,
            [Parameter()][object]$Evidence
        )
        $outcome = [PSCustomObject]@{
            HostIP       = $HostIP
            CID          = $CID
            FixAttempted = 'No'
            Outcome      = 'Not Attempted'
            Details      = ''
        }

        $reach = Test-HostForRemoting -HostIP $HostIP
        if ($reach -ne 'OK') {
            $outcome.Outcome = $reach
            $outcome.Details = 'Skipping remediation.'
            Write-Log "Remediation skipped for $HostIP CID $CID: $reach" 'WARNING'
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
                
                # Handle different return types from fixers
                if ($null -eq $result) { 
                    $result = [pscustomobject]@{ Outcome='Succeeded'; Details='No details provided.' } 
                } elseif ($result -is [string]) {
                    $result = [pscustomobject]@{ Outcome=$result; Details='' }
                } elseif (-not $result.PSObject.Properties.Match('Outcome')) {
                    $result | Add-Member -NotePropertyName Outcome -NotePropertyValue 'Unknown' -Force
                }
                if (-not $result.PSObject.Properties.Match('Details')) { 
                    $result | Add-Member -NotePropertyName Details -NotePropertyValue '' -Force
                }
                
                $outcome.Outcome = [string]$result.Outcome
                $outcome.Details = [string]$result.Details
                Write-Log "Remediation on $HostIP for CID $CID: $($outcome.Outcome) - $($outcome.Details)"
            } catch {
                $outcome.Outcome = 'Failed'
                $outcome.Details = $_.Exception.Message
                Write-Log "Remediation failed on $HostIP for CID $CID: $($outcome.Details)" 'ERROR'
            }
        } else {
            $outcome.Outcome = 'Dry Run'
            $outcome.Details = 'Would attempt remediation if -Remediate was specified'
            Write-Log "Dry-run: Would attempt remediation on $HostIP for CID $CID." 'DEBUG'
        }

        return $outcome
    }
    #endregion
}

process {
    try {
        $posturesRaw = Get-FailedPostures
        if (-not $posturesRaw -or $posturesRaw.Count -eq 0) { 
            Write-Log "No failed compliance postures found for policy $PolicyId"
            return 
        }

        # Normalize to canonical columns
        $postures = Normalize-QualysRows -Rows $posturesRaw

        # Filter out rows with null/empty IP and non-numeric ControlIDs
        $postures = $postures |
            Where-Object { -not [string]::IsNullOrEmpty($_.IP) } |
            Where-Object { $_.ControlID -match '^\d+$' }

        if ($postures.Count -eq 0) {
            Write-Log "No valid posture records after filtering (IP and ControlID validation)"
            return
        }

        # Group by host IP and CID
        $grouped = $postures |
            Group-Object -Property IP | ForEach-Object {
                $hostIP = $_.Name
                $_.Group | Group-Object -Property ControlID | ForEach-Object {
                    [PSCustomObject]@{
                        HostIP   = $hostIP
                        CID      = [int]$_.Name
                        Evidence = ($_.Group | ForEach-Object { $_.Evidence } | Where-Object { $_ } | Select-Object -Unique) -join '; '
                        Reason   = ($_.Group | ForEach-Object { $_.Reason }   | Where-Object { $_ } | Select-Object -Unique) -join '; '
                    }
                }
            }

        $results = New-Object System.Collections.Generic.List[object]
        $total = ($grouped | Measure-Object).Count
        $i = 0

        foreach ($item in $grouped) {
            $i++
            Write-Progress -Activity 'Processing Compliance Issues' -Status "Processing $($item.HostIP) CID $($item.CID)" -PercentComplete ([int](($i/$total)*100))
            
            $remediation = Remediate-Host -HostIP $item.HostIP -CID $item.CID -Evidence $item.Evidence
            
            if ($Remediate -and $LaunchRescan -and $remediation.Outcome -match '^(Succeeded|Success|Fixed)$') {
                $rescanRequested = Invoke-HostRescan -HostIP $item.HostIP
                $remediation | Add-Member -NotePropertyName RescanRequested -NotePropertyValue $rescanRequested -Force
            }
            
            $results.Add([PSCustomObject]@{
                HostIP       = $item.HostIP
                CID          = $item.CID
                Evidence     = $item.Evidence
                Reason       = $item.Reason
                FixAttempted = $remediation.FixAttempted
                Outcome      = $remediation.Outcome
                Details      = $remediation.Details
                Timestamp    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            })
        }
        Write-Progress -Activity 'Processing Compliance Issues' -Completed -Status 'Done'

        $mode = if ($Remediate) { 'Remediate' } else { 'Report' }
        $results |
            Select-Object @{n='Mode';e={$mode}}, HostIP, CID, Evidence, Reason, FixAttempted, Outcome, Details, Timestamp |
            Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
            
        Write-Log "Report generated: $reportPath with $($results.Count) records"
        
    } catch {
        Write-Log "Process failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

end {
    # Cleanup temporary files
    foreach ($p in @($csvRawPath, $csvPath)) { 
        if (Test-Path $p) { 
            try {
                Remove-Item $p -Force -ErrorAction SilentlyContinue
                Write-Log "Cleaned up temporary file: $p" 'DEBUG'
            } catch {
                Write-Log "Failed to clean up temporary file $p: $($_.Exception.Message)" 'WARNING'
            }
        } 
    }
    
    Write-Log 'Script completed.'
}
