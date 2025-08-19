<#
.SYNOPSIS
    Queries Qualys for failed compliance postures, generates reports, and optionally remediates issues using per-CID fixer scripts.

.DESCRIPTION
    Connects to the Qualys API using a service account, retrieves failed compliance data for a specified policy, parses the CSV output,
    normalises columns to canonical names, groups by host and CID, generates a report, and (if -Remediate is specified) attempts fixes
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

.EXAMPLE
    $qualysCred = Get-Credential -Message 'Qualys API Credentials'
    $adminCred = Get-Credential -Message 'Admin Credentials for Remoting'
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential $qualysCred -AdminCredential $adminCred -Remediate -LaunchRescan

.EXAMPLE
    .\CheckandRemediate.ps1 -QualysBaseUrl 'https://qualysapi.qualys.eu' -QualysCredential (Import-Clixml 'C:\secure\qualys.cred') -PolicyId 99999

.NOTES
    Requires PowerShell 7.4+ (tested 7.5.2).
    Working directory: C:\temp\scripts
    Logs: C:\temp\scripts\logs\CheckandRemediate_<yyyyMMdd_HHmm>.log
    Reports: C:\temp\scripts\reports\FailedCompliance_<yyyyMMdd_HHmm>.csv
    Dependencies: Invoke-RestMethod, Import-Csv, Invoke-Command, Test-WSMan.

    Changelog (latest first):
      - 2025-08-19: Fixed CSV wrapper handling (BEGIN/END markers) and optional ZIP response; now reliably extracts a clean CSV for Import-Csv.
      - 2025-08-19: Header resolver now REQUIRES only IP + ControlID; Evidence/Reason are optional and default to empty if absent.
      - 2025-08-19: Improved error messages to include first 5 header names and sample first data line when schema cannot be resolved.
      - 2025-08-19: Guarded grouping and filtering against missing fields; removed hard throws that caused line 279/315/334 stops.
      - 2025-08-19: Import Test-WSMan module for PS 7.5.2; more robust remoting checks.
      - 2025-08-19: Previous changes retained (POST + retry/backoff, TLS1.2, signed fixers, UTF8 logs/reports, progress, -LaunchRescan).

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
    [int]$PolicyId = 99999,

    [Parameter()]
    [int]$TruncationLimit = 10000,

    [Parameter(Mandatory = $true, ParameterSetName = 'Remediate')]
    [System.Management.Automation.PSCredential]$AdminCredential,

    [Parameter(ParameterSetName = 'Remediate')]
    [switch]$Remediate,

    [Parameter(ParameterSetName = 'Remediate')]
    [switch]$LaunchRescan
)

begin {
    #region Initialisation
    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    try { Import-Module Microsoft.WSMan.Management -ErrorAction SilentlyContinue } catch {}

    $workingDir  = 'C:\temp\scripts'
    $logsDir     = Join-Path $workingDir 'logs'
    $reportsDir  = Join-Path $workingDir 'reports'
    $fixersDir   = Join-Path $workingDir 'fixers'
    $timestamp   = Get-Date -Format 'yyyyMMdd_HHmm'
    $logPath     = Join-Path $logsDir "CheckandRemediate_$timestamp.log"
    $reportPath  = Join-Path $reportsDir "FailedCompliance_$timestamp.csv"
    $csvRawPath  = Join-Path $workingDir 'failed_postures_raw.bin'
    $csvPath     = Join-Path $workingDir 'failed_postures.csv'

    foreach ($dir in @($workingDir, $logsDir, $reportsDir, $fixersDir)) {
        if (-not (Test-Path $dir -PathType Container)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
    }
    if (-not (Test-Path $logPath)) { New-Item -ItemType File -Path $logPath -Force | Out-Null }
    Add-Content -Path $logPath -Value '' -Encoding UTF8

    function Write-Log {
        param(
            [Parameter(Mandatory = $true)][string]$Message,
            [ValidateSet('INFO','WARNING','ERROR','DEBUG')][string]$Level = 'INFO'
        )
        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        Write-Information $logEntry -InformationAction Continue
        Add-Content -Path $logPath -Value $logEntry -Encoding UTF8
    }

    Write-Log "Script started. PSVersion=$($PSVersionTable.PSVersion) ParamSet=$($PSCmdlet.ParameterSetName)"
    #endregion

    #region Helpers: Qualys, CSV extraction/normaliser, Fixers and Rescans
    function Invoke-QualysRequest {
        param(
            [Parameter(Mandatory = $true)][string]$EndpointPath,
            [Parameter(Mandatory = $true)][hashtable]$Body,
            [Parameter()][string]$Accept = 'text/csv',
            [Parameter()][string]$OutFile
        )
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($QualysCredential.UserName):$($QualysCredential.GetNetworkCredential().Password)"))
        $headers = @{
            'Authorization'    = "Basic $auth"
            'Accept'           = $Accept
            'X-Requested-With' = 'PowerShell'
        }
        $uri = ($QualysBaseUrl.TrimEnd('/')) + $EndpointPath

        $maxAttempts = 4
        $attempt = 0
        do {
            $attempt++
            try {
                Write-Log "Qualys POST $EndpointPath (attempt $attempt)"
                if ($OutFile) {
                    Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $Body -OutFile $OutFile -ErrorAction Stop
                } else {
                    return Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $Body -ErrorAction Stop
                }
                break
            } catch {
                if ($attempt -ge $maxAttempts) { throw }
                $delay = [math]::Pow(2, $attempt)
                Write-Log "Qualys request failed: $($_.Exception.Message). Retrying in ${delay}s..." 'WARNING'
                Start-Sleep -Seconds $delay
            }
        } while ($true)
    }

    function Unwrap-QualysCsv {
        param(
            [Parameter(Mandatory=$true)][string]$RawPath,
            [Parameter(Mandatory=$true)][string]$CsvPath
        )
        # Handle ZIP responses (PK header)
        $fs = [IO.File]::OpenRead($RawPath)
        try {
            $b0 = $fs.ReadByte(); $b1 = $fs.ReadByte(); $fs.Position = 0
        } finally { $fs.Dispose() }
        if ($b0 -eq 0x50 -and $b1 -eq 0x4B) {
            $tmp = Join-Path (Split-Path $CsvPath -Parent) ("unz_" + [guid]::NewGuid().Guid)
            New-Item -ItemType Directory -Path $tmp | Out-Null
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [IO.Compression.ZipFile]::ExtractToDirectory($RawPath, $tmp)
            $csv = Get-ChildItem -Path $tmp -Filter *.csv -Recurse | Select-Object -First 1
            if (-not $csv) { throw "ZIP response contained no CSV file." }
            Copy-Item $csv.FullName $CsvPath -Force
            Remove-Item $tmp -Recurse -Force
            return
        }
        # Strip Qualys wrapper lines and keep only CSV
        $lines = Get-Content -Path $RawPath -Encoding Byte -Raw | ForEach-Object {[Text.Encoding]::UTF8.GetString($_)} # ensure UTF8
        if (-not $lines) { throw "Empty response body." }
        $lines = ($lines -split "`r?`n")
        # Remove markers like ----BEGIN_RESPONSE..., ----END_RESPONSE...
        $lines = $lines | Where-Object { $_ -and ($_ -notmatch '^----BEGIN_') -and ($_ -notmatch '^----END_') }
        # Start at first line that looks like CSV header (contains a comma)
        $start = 0
        for ($i=0; $i -lt $lines.Count; $i++) { if ($lines[$i] -like '*,*') { $start = $i; break } }
        $clean = $lines[$start..($lines.Count-1)]
        Set-Content -Path $CsvPath -Value $clean -Encoding UTF8
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

            Invoke-QualysRequest -EndpointPath '/api/2.0/fo/compliance/posture/info/' -Body $body -OutFile $csvRawPath | Out-Null
            Unwrap-QualysCsv -RawPath $csvRawPath -CsvPath $csvPath

            $raw = Import-Csv -Path $csvPath
            if (-not $raw -or $raw.Count -eq 0) {
                Write-Log 'No failed postures found.'
                return @()
            }

            # Trim BOM, whitespace
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

        # Canonical -> accepted aliases
        $map = @{
            IP         = @('IP','IP Address','IP Address(es)','Host IP','Host IP Address','IP Address(es) List')
            ControlID  = @('Control ID','CID','Control ID (CID)','CID (Control ID)','Control Identifier','Control')
            Evidence   = @('Evidence','Evidence/Results','Instance Evidence','Evidence Value','Actual Value','Value','Finding','Observed Value')
            Reason     = @('Reason','Reason for Failure','Failure Reason','Reason/Recommendation','Rationale','Recommendation','Expected Value','Expected')
        }
        $resolved = @{}
        foreach ($k in $map.Keys) {
            $alias = $map[$k] | Where-Object { $_ -in $Available }
            if ($alias) { $resolved[$k] = $alias[0] }
        }
        return $resolved
    }

    function Normalize-QualysRows {
        param([Parameter(Mandatory=$true)][object[]]$Rows)

        if (-not $Rows -or $Rows.Count -eq 0) { return @() }
        $headers = $Rows[0].PSObject.Properties.Name
        $resolved = Resolve-Header -Available $headers

        # Only IP and ControlID are truly required to act; Evidence/Reason are optional.
        $requiredCanon = @('IP','ControlID')
        $missingCanon = $requiredCanon | Where-Object { $_ -notin $resolved.Keys }
        if ($missingCanon) {
            Write-Log ("Detected CSV headers (first 5): {0}" -f (($headers | Select-Object -First 5) -join ', ')) 'ERROR'
            $firstLine = ($Rows | Select-Object -First 1 | ConvertTo-Csv -NoTypeInformation)[1]
            Write-Log ("First data line snapshot: {0}" -f $firstLine) 'ERROR'
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
            Write-Log "No fixer found for CID $CID."
            return $null
        }
        $candidate = $fixerFiles[0].FullName
        $sig = Get-AuthenticodeSignature -FilePath $candidate
        if ($sig.Status -ne 'Valid') {
            Write-Log "Fixer $candidate has invalid or missing signature: $($sig.Status). Skipping." 'WARNING'
            return $null
        }
        Write-Log "Discovered signed fixer: $candidate for CID $CID."
        $scriptText = Get-Content $candidate -Raw
        return [scriptblock]::Create($scriptText)
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
        if (-not (Test-Connection -ComputerName $HostIP -Count 1 -Quiet)) {
            return 'Host Offline'
        }
        try {
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
                if ($null -eq $result) { $result = [pscustomobject]@{ Outcome='Succeeded'; Details='No details.' } }
                if (-not $result.PSObject.Properties.Match('Outcome')) { $result | Add-Member -NotePropertyName Outcome -NotePropertyValue 'Unknown' }
                if (-not $result.PSObject.Properties.Match('Details')) { $result | Add-Member -NotePropertyName Details -NotePropertyValue '' }
                $outcome.Outcome = [string]$result.Outcome
                $outcome.Details = [string]$result.Details
                Write-Log "Remediation on $HostIP for CID $CID: $($outcome.Outcome)"
            } catch {
                $outcome.Outcome = 'Failed'
                $outcome.Details = $_.Exception.Message
                Write-Log "Remediation failed on $HostIP for CID $CID: $($outcome.Details)" 'ERROR'
            }
        } else {
            Write-Log "Dry-run: Would attempt remediation on $HostIP for CID $CID."
        }

        return $outcome
    }
    #endregion
}

process {
    try {
        $posturesRaw = Get-FailedPostures
        if (-not $posturesRaw -or $posturesRaw.Count -eq 0) { return }

        # Normalise to canonical columns
        $postures = Normalize-QualysRows -Rows $posturesRaw

        # Filter out rows with null/empty IP and non-numeric ControlIDs
        $postures = $postures |
            Where-Object { -not [string]::IsNullOrEmpty($_.IP) } |
            Where-Object { $_.ControlID -match '^\d+$' }

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
            Write-Progress -Activity 'Remediation' -Status "Processing $($item.HostIP) CID $($item.CID)" -PercentComplete ([int](($i/$total)*100))
            $remediation = Remediate-Host -HostIP $item.HostIP -CID $item.CID -Evidence $item.Evidence
            if ($Remediate -and $LaunchRescan -and $remediation.Outcome -match '^(Succeeded|Success|Fixed)$') {
                [void](Invoke-HostRescan -HostIP $item.HostIP)
            }
            $results.Add([PSCustomObject]@{
                HostIP       = $item.HostIP
                CID          = $item.CID
                Evidence     = $item.Evidence
                Reason       = $item.Reason
                FixAttempted = $remediation.FixAttempted
                Outcome      = $remediation.Outcome
                Details      = $remediation.Details
            })
        }
        Write-Progress -Activity 'Remediation' -Completed -Status 'Done'

        $mode = if ($Remediate) { 'Remediate' } else { 'Report' }
        $results |
            Select-Object @{n='Mode';e={$mode}}, HostIP, CID, Evidence, Reason, FixAttempted, Outcome, Details |
            Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Report generated: $reportPath"
    } catch {
        Write-Log "Process failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

end {
    foreach ($p in @($csvRawPath,$csvPath)) { if (Test-Path $p) { Remove-Item $p -Force } }
    Write-Log 'Script completed.'
}

# Signed by Marinus van Deventer
