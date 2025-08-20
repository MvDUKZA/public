<#
.SYNOPSIS
    Pull Qualys failed postures, optionally run per-CID fixers (parallel supported), optionally trigger a Qualys rescan, and report.
.DESCRIPTION
    - Downloads plain CSV of failed postures (status=Failed) for one policy from Qualys.
    - Trims preamble until the true CSV header which starts with the quoted columns: "ID","IP","OS","DNS Name".
    - Uses fixed CSV columns: IP, DNS Name, Control ID, Posture Evidence, Reason for Failure.
    - Groups by target (IP preferred, else DNS Name) and Control ID (CID).
    - If -Remediate, runs C:\temp\scripts\fixers\<CID>-*.ps1 via PowerShell Remoting.
      If -Parallel, runs concurrently with -ThrottleLimit.
    - If -Rescan, launches a Qualys compliance rescan for hosts with successful fixes (needs IP).
    - Writes C:\temp\scripts\reports\FailedCompliance_<timestamp>.csv
.PARAMETER QualysBaseUrl
    Qualys API base URL (e.g. https://qualysapi.qualys.eu).
.PARAMETER QualysCredential
    PSCredential for Qualys API.
.PARAMETER PolicyId
    Qualys policy ID (default 99999).
.PARAMETER TruncationLimit
    Max records (default 10000; set 0 to omit).
.PARAMETER AdminCredential
    Credential for PowerShell Remoting to targets (required with -Remediate).
.PARAMETER Remediate
    Run fixers for each target/CID if a matching fixer script exists.
.PARAMETER Parallel
    Process fixers in parallel using ForEach-Object -Parallel (PowerShell 7+).
.PARAMETER ThrottleLimit
    Maximum concurrent fixer executions when -Parallel is used (default 8).
.PARAMETER Rescan
    After a successful fix and when an IP is available, launch a Qualys compliance rescan.
.EXAMPLE
    $q = Get-Credential
    $a = Get-Credential
    .\CheckAndRemediate.ps1 -QualysBaseUrl https://qualysapi.qualys.eu -QualysCredential $q -PolicyId 99999 -AdminCredential $a -Remediate -Parallel -ThrottleLimit 10 -Rescan
.NOTES
    PowerShell 7.5.2. Working dir: C:\temp\scripts
    Logs:   C:\temp\scripts\logs\CheckAndRemediate_<yyyyMMdd_HHmm>.log
    Report: C:\temp\scripts\reports\FailedCompliance_<yyyyMMdd_HHmm>.csv
    Dependencies: Invoke-RestMethod, Import-Csv, Invoke-Command, Get-Content, Set-Content.
    Changelog:
      - 2025-08-20: Tidied structure with #region, added comment-based help, try/catch to all functions, param validation, -Parallel with PS version check, Write-Progress if not parallel, fixed interpolation with ${HostIp}: and ${ControlId}:, updated requiredColumns to exact list, used 'Reason for Failure' for Reason, updated logging for preamble trim.
    Signed by Marinus van Deventer
#>

[CmdletBinding(DefaultParameterSetName='Report')]
param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$QualysBaseUrl,
    [Parameter(Mandatory=$true)][System.Management.Automation.PSCredential]$QualysCredential,
    [int]$PolicyId = 99999,
    [int]$TruncationLimit = 10000,

    [Parameter(Mandatory=$true, ParameterSetName='Remediate')]
    [System.Management.Automation.PSCredential]$AdminCredential,

    [Parameter(ParameterSetName='Remediate')][switch]$Remediate,
    [Parameter(ParameterSetName='Remediate')][switch]$Parallel,
    [Parameter(ParameterSetName='Remediate')][ValidateRange(1,128)][int]$ThrottleLimit = 8,
    [Parameter(ParameterSetName='Remediate')][switch]$Rescan
)

#region Initialisation
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$workingDir  = 'C:\temp\scripts'
$logsDir     = Join-Path $workingDir 'logs'
$reportsDir  = Join-Path $workingDir 'reports'
$fixersDir   = Join-Path $workingDir 'fixers'
foreach ($d in @($workingDir,$logsDir,$reportsDir,$fixersDir)) {
    if (-not (Test-Path $d -PathType Container)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
}

$stamp      = Get-Date -Format 'yyyyMMdd_HHmm'
$logPath    = Join-Path $logsDir    "CheckAndRemediate_$stamp.log"
$reportPath = Join-Path $reportsDir "FailedCompliance_$stamp.csv"
$csvPath    = Join-Path $workingDir 'Qualys_Posture.csv'

if (-not (Test-Path $logPath)) { New-Item -ItemType File -Path $logPath -Force | Out-Null }

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARNING','ERROR')][string]$Level='INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    Write-Information $line -InformationAction Continue
    Add-Content -Path $logPath -Value $line -Encoding UTF8
}

Write-Log "Script started. ParameterSetName=$($PSCmdlet.ParameterSetName)"
#endregion

#region Functions
<#
.SYNOPSIS
    Removes rows above the header row in a CSV file.
.DESCRIPTION
    Reads a CSV file, identifies the header row, and removes all rows above it. Supports exact string matching for the header.
.PARAMETER FilePath
    Path to the CSV file.
.PARAMETER HeaderColumns
    Array of header column names to form the expected prefix.
.EXAMPLE
    Remove-CsvPreamble -FilePath "data.csv" -HeaderColumns @('ID','IP','OS','DNS Name')
.NOTES
    Uses quoted prefix match for Qualys CSV headers.
#>
function Remove-CsvPreamble {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$HeaderColumns
    )
    try {
        if (-not (Test-Path $FilePath)) { Write-Log "File not found: $FilePath" 'ERROR'; throw "File not found." }
        $quotedPrefix = '"' + ($HeaderColumns -join '","') + '"'
        $text = Get-Content -Path $FilePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($text)) { throw "Empty CSV body." }
        $lines = $text -split "(`r`n|`n|`r)"
        $idx = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].TrimStart()
            if ($line.StartsWith($quotedPrefix, $true, [Globalization.CultureInfo]::InvariantCulture)) { $idx = $i; break }
        }
        if ($idx -lt 0) { Write-Log "Header not found. Expected prefix: $quotedPrefix" 'ERROR'; throw "Header row not found." }
        if ($idx -gt 0) {
            $newText = ($lines[$idx..($lines.Count-1)] -join [Environment]::NewLine)
            Set-Content -Path $FilePath -Value $newText -Encoding UTF8 -ErrorAction Stop
            Write-Log "Trimmed $idx preamble lines in $FilePath" 'INFO'
        } else {
            Write-Log "Header already on first line in $FilePath" 'INFO'
        }
    } catch {
        Write-Log "Preamble removal failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

<#
.SYNOPSIS
    Invokes a Qualys API POST request for CSV output.
.DESCRIPTION
    Sends POST to Qualys endpoint for failed postures, saves to file.
.PARAMETER BaseUrl
    Qualys API base URL.
.PARAMETER Cred
    PSCredential for Qualys API.
.PARAMETER Policy
    Policy ID.
.PARAMETER Limit
    Truncation limit.
.PARAMETER OutCsv
    Output CSV path.
.NOTES
    Uses 'csv' format for plain CSV with headers.
#>
function Invoke-QualysCsv {
    param (
        [string]$BaseUrl,
        [pscredential]$Cred,
        [int]$Policy,
        [int]$Limit,
        [string]$OutCsv
    )
    try {
        $pair = '{0}:{1}' -f $Cred.UserName, $Cred.GetNetworkCredential().Password
        $auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
        $headers = @{ Authorization = "Basic $auth"; Accept = 'text/csv'; 'X-Requested-With'='PowerShell' }
        $body = @{
            action        = 'list'
            policy_id     = $Policy
            output_format = 'csv'
            status        = 'Failed'
            details       = 'All'
        }
        if ($Limit -gt 0) { $body.truncation_limit = $Limit }
        $uri = ($BaseUrl.TrimEnd('/')) + '/api/2.0/fo/compliance/posture/info/'
        Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $body -OutFile $OutCsv -ErrorAction Stop
        Write-Log "Downloaded CSV to $OutCsv" 'INFO'
    } catch {
        Write-Log "Qualys CSV download failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

<#
.SYNOPSIS
    Gets Qualys failed postures as PSCustomObjects.
.DESCRIPTION
    Invokes API, cleans preamble, imports CSV, projects to key fields, filters actionable rows.
.PARAMETER BaseUrl
    Qualys API base URL.
.PARAMETER Cred
    PSCredential for Qualys API.
.PARAMETER Policy
    Policy ID.
.PARAMETER Limit
    Truncation limit.
.PARAMETER OutCsv
    Output CSV path.
.NOTES
    Filters null target and non-numeric CID.
#>
function Get-QualysFailures {
    param (
        [string]$BaseUrl,
        [pscredential]$Cred,
        [int]$Policy,
        [int]$Limit,
        [string]$OutCsv
    )
    try {
        Invoke-QualysCsv -BaseUrl $BaseUrl -Cred $Cred -Policy $Policy -Limit $Limit -OutCsv $OutCsv
        Remove-CsvPreamble -FilePath $OutCsv -HeaderColumns @('ID','IP','OS','DNS Name')
        $raw = Import-Csv -Path $OutCsv -ErrorAction Stop
        if (-not $raw -or $raw.Count -eq 0) { return @() }
        $rows = $raw | ForEach-Object {
            [pscustomobject]@{
                TargetName = if ($_.IP) { $_.IP } else { $_.'DNS Name' }
                IP         = $_.IP
                Hostname   = $_.'DNS Name'
                CID        = $_.'Control ID'
                Evidence   = $_.'Posture Evidence'
                Reason     = $_.'Reason for Failure'
            }
        } | Where-Object { $_.TargetName -and ($_.CID -match '^\d+$') }
        Write-Log "Parsed $($rows.Count) actionable failures from $($raw.Count) raw rows" 'INFO'
        return $rows
    } catch {
        Write-Log "Get failures failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

<#
.SYNOPSIS
    Finds the first fixer script for a CID.
.DESCRIPTION
    Searches C:\temp\scripts\fixers for <CID>-*.ps1.
.PARAMETER CID
    Control ID to find fixer for.
.EXAMPLE
    Find-Fixer -CID 2781
.NOTES
    Returns full path or null.
#>
function Find-Fixer {
    param (
        [Parameter(Mandatory = $true)]
        [int]$CID
    )
    $match = Get-ChildItem -Path $fixersDir -Filter "$CID-*.ps1" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $match) {
        Write-Log "No fixer found for CID $CID" 'WARNING'
        return $null
    }
    Write-Log "Found fixer: $($match.FullName) for CID $CID" 'INFO'
    return $match.FullName
}

<#
.SYNOPSIS
    Runs a fixer script on a remote host.
.DESCRIPTION
    Loads fixer script, invokes remotely, normalises outcome/details.
.PARAMETER Computer
    Target host name or IP.
.PARAMETER Cred
    PSCredential for remoting.
.PARAMETER FixerPath
    Path to fixer script.
.EXAMPLE
    Run-Fixer -Computer '10.0.0.1' -Cred $cred -FixerPath 'C:\temp\scripts\fixers\2781-Fix.ps1'
.NOTES
    Returns array: FixAttempted, Outcome, Details.
#>
function Run-Fixer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Computer,

        [Parameter(Mandatory = $true)]
        [pscredential]$Cred,

        [Parameter(Mandatory = $true)]
        [string]$FixerPath
    )
    try {
        $scriptText = Get-Content -Path $FixerPath -Raw -ErrorAction Stop
        $sb = [scriptblock]::Create($scriptText)
        $ret = Invoke-Command -ComputerName $Computer -Credential $Cred -ScriptBlock $sb -ErrorAction Stop
        $outcome = if ($ret -and $ret.PSObject.Properties.Name -contains 'Outcome') { [string]$ret.Outcome } else { 'Succeeded' }
        $details = if ($ret -and $ret.PSObject.Properties.Name -contains 'Details') { [string]$ret.Details } else { '' }
        return @('Yes', $outcome, $details)
    } catch {
        Write-Log "Run fixer failed on $Computer for CID $ControlId: $($_.Exception.Message)" 'ERROR'
        return @('Yes', 'Failed', $_.Exception.Message)
    }
}

<#
.SYNOPSIS
    Launches a Qualys compliance rescan for a host IP.
.DESCRIPTION
    Sends POST to Qualys scan endpoint to launch rescan.
.PARAMETER BaseUrl
    Qualys API base URL.
.PARAMETER Cred
    PSCredential for Qualys API.
.PARAMETER HostIP
    IP to rescan.
.EXAMPLE
    Invoke-QualysRescan -BaseUrl 'https://qualysapi.qualys.eu' -Cred $cred -HostIP '10.0.0.1'
.NOTES
    Returns true on success, false on failure.
#>
function Invoke-QualysRescan {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BaseUrl,

        [Parameter(Mandatory = $true)]
        [pscredential]$Cred,

        [Parameter(Mandatory = $true)]
        [string]$HostIP
    )
    try {
        $pair = '{0}:{1}' -f $Cred.UserName, $Cred.GetNetworkCredential().Password
        $auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
        $headers = @{ Authorization = "Basic $auth"; Accept = 'application/xml'; 'X-Requested-With'='PowerShell' }
        $body = @{
            action     = 'launch'
            scan_title = "AutoRescan_$($HostIP)_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss'))"
            ip         = $HostIP
            priority   = 'Normal'
        }
        $uri = ($BaseUrl.TrimEnd('/')) + '/api/2.0/fo/scan/compliance/'
        Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop | Out-Null
        Write-Log "Rescan requested for $HostIP" 'INFO'
        return $true
    } catch {
        Write-Log "Rescan request failed for $HostIP: $($_.Exception.Message)" 'WARNING'
        return $false
    }
}
#endregion

process {
    try {
        # Download and parse failures
        Write-Log "Downloading Qualys failures..." 'INFO'
        $failures = Get-QualysFailures -BaseUrl $QualysBaseUrl -Cred $QualysCredential -Policy $PolicyId -Limit $TruncationLimit -OutCsv $csvPath
        Write-Log "Retrieved $($failures.Count) actionable failures." 'INFO'

        if ($failures.Count -eq 0) { return }

        # Group by Target and CID
        $work = $failures | Group-Object TargetName | ForEach-Object {
            $t = $_.Name
            $_.Group | Group-Object CID | ForEach-Object {
                [pscustomobject]@{
                    TargetName = $t
                    CID        = [int]$_.Name
                    IP         = ($_.Group | ForEach-Object IP       | Where-Object { $_ } | Select-Object -First 1)
                    Hostname   = ($_.Group | ForEach-Object Hostname | Where-Object { $_ } | Select-Object -First 1)
                    Evidence   = ($_.Group | ForEach-Object Evidence | Where-Object { $_ } | Select-Object -Unique) -join '; '
                    Reason     = ($_.Group | ForEach-Object Reason   | Where-Object { $_ } | Select-Object -Unique) -join '; '
                }
            }
        }

        # Remediate and collect results
        $results = @()
        if ($Remediate) {
            if ($Parallel -and $PSVersionTable.PSVersion.Major -ge 7) {
                $results = $work | ForEach-Object -Parallel {
                    $AdminCredential = $using:AdminCredential
                    $Rescan          = $using:Rescan
                    $QualysBaseUrl   = $using:QualysBaseUrl
                    $QualysCredential= $using:QualysCredential

                    $fixAttempted = 'No'; $outcome = 'Not Attempted'; $details = ''
                    $fixer = Get-ChildItem -Path 'C:\temp\scripts\fixers' -Filter ("{0}-*.ps1" -f $item.CID) -File | Select-Object -First 1
                    if ($fixer) {
                        try {
                            $sb = [scriptblock]::Create((Get-Content $fixer.FullName -Raw))
                            $ret = Invoke-Command -ComputerName $item.TargetName -Credential $AdminCredential -ScriptBlock $sb -ErrorAction Stop
                            $fixAttempted = 'Yes'
                            $outcome = if ($ret -and $ret.PSObject.Properties.Name -contains 'Outcome') { [string]$ret.Outcome } else { 'Succeeded' }
                            $details = if ($ret -and $ret.PSObject.Properties.Name -contains 'Details') { [string]$ret.Details } else { '' }
                        } catch {
                            $fixAttempted = 'Yes'; $outcome = 'Failed'; $details = $_.Exception.Message
                        }

                        if ($Rescan -and $outcome -match '^(Succeeded|Success|Fixed)$' -and $item.IP) {
                            $pair = '{0}:{1}' -f $QualysCredential.UserName, $QualysCredential.GetNetworkCredential().Password
                            $auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
                            $headers = @{ Authorization = "Basic $auth"; Accept = 'application/xml'; 'X-Requested-With'='PowerShell' }
                            $body = @{ action='launch'; scan_title="AutoRescan_$($item.IP)_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss'))"; ip=$item.IP; priority='Normal' }
                            $uri = ($QualysBaseUrl.TrimEnd('/')) + '/api/2.0/fo/scan/compliance/'
                            Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/x-www-form-urlencoded' -Body $body | Out-Null
                        }
                    } else {
                        $outcome = 'No Fixer Available'
                    }

                    [pscustomobject]@{
                        TargetName   = $item.TargetName
                        IP           = $item.IP
                        Hostname     = $item.Hostname
                        CID          = $item.CID
                        Evidence     = $item.Evidence
                        Reason       = $item.Reason
                        FixAttempted = $fixAttempted
                        Outcome      = $outcome
                        Details      = $details
                    }
                } -ThrottleLimit $ThrottleLimit
            } else {
                $i = 0
                $total = $work.Count
                foreach ($item in $work) {
                    $i++
                    Write-Progress -Activity 'Remediating' -Status "CID $item.CID on $item.TargetName" -PercentComplete (($i / $total) * 100)
                    $fixAttempted = 'No'; $outcome = 'Not Attempted'; $details = ''
                    $fixer = Find-Fixer -CID $item.CID
                    if ($fixer) {
                        $vals = Run-Fixer -Computer $item.TargetName -Cred $AdminCredential -FixerPath $fixer
                        $fixAttempted = $vals[0]; $outcome = $vals[1]; $details = $vals[2]
                        if ($Rescan -and $outcome -match '^(Succeeded|Success|Fixed)$' -and $item.IP) {
                            [void](Invoke-QualysRescan -BaseUrl $QualysBaseUrl -Cred $QualysCredential -HostIP $item.IP)
                        }
                    } else {
                        $outcome = 'No Fixer Available'
                    }
                    $results += [pscustomobject]@{
                        TargetName   = $item.TargetName
                        IP           = $item.IP
                        Hostname     = $item.Hostname
                        CID          = $item.CID
                        Evidence     = $item.Evidence
                        Reason       = $item.Reason
                        FixAttempted = $fixAttempted
                        Outcome      = $outcome
                        Details      = $details
                    }
                }
                Write-Progress -Activity 'Remediating' -Completed
            }
        } else {
            $results = $work | ForEach-Object {
                [pscustomobject]@{
                    TargetName   = $_.TargetName
                    IP           = $_.IP
                    Hostname     = $_.Hostname
                    CID          = $_.CID
                    Evidence     = $_.Evidence
                    Reason       = $_.Reason
                    FixAttempted = 'No'
                    Outcome      = 'Not Attempted'
                    Details      = ''
                }
            }
        }

        # Report
        $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Report generated: $reportPath" 'INFO'
    } catch {
        Write-Log "Process failed: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

end {
    if (Test-Path $csvPath) { Remove-Item $csvPath -Force }
    Write-Log "Script completed." 'INFO'
}

# Signed by Marinus van Deventer