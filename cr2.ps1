<#
.SYNOPSIS
    Queries Qualys for failed compliance postures, generates reports, and optionally remediates issues using per-CID fixer scripts.

.DESCRIPTION
    Connects to the Qualys API using a service account, retrieves failed compliance data for a specified policy, parses plain CSV output,
    normalises columns to canonical names, groups by target and CID, generates a report, and (if -Remediate is specified) attempts fixes
    on online hosts via remote PowerShell. Fixers are discovered dynamically from the 'fixers' subfolder and must be Authenticode-signed.
    Optionally, upon successful remediation, can request a Qualys compliance rescan for the host when -LaunchRescan is provided.

.PARAMETER QualysBaseUrl
    The base URL for the Qualys API (e.g., 'https://qualysapi.qualys.eu').

.PARAMETER QualysCredential
    PSCredential for Qualys API.

.PARAMETER PolicyId
    Qualys policy ID (default 99999).

.PARAMETER TruncationLimit
    API truncation limit (default 10000; set 0 to omit).

.PARAMETER AdminCredential
    PSCredential used for remoting to target machines. Mandatory only with -Remediate.

.PARAMETER Remediate
    Attempt remediation using signed fixer scripts in C:\temp\scripts\fixers\<CID>-*.ps1

.PARAMETER LaunchRescan
    After a successful fix, launch a Qualys compliance rescan for the host.

.NOTES
    Tested on PowerShell 7.5.2.
    Working dir: C:\temp\scripts
    Logs:       C:\temp\scripts\logs\CheckandRemediate_<yyyyMMdd_HHmm>.log
    Reports:    C:\temp\scripts\reports\FailedCompliance_<yyyyMMdd_HHmm>.csv

    Changelog (latest)
      - 2025-08-19: Simplified to CSV-only flow per tenant behaviour; removed ZIP/wrapper handling.
      - 2025-08-19: Header detector skips metadata banner and locks onto posture table.
      - 2025-08-19: Evidence aliases include 'Posture Evidence'; Reason aliases include 'Control Statement'.
      - 2025-08-19: All literal '$CID:'/$variable: artefacts removed; logging interpolates variables correctly.
      - 2025-08-19: Prior improvements retained (POST+retry/backoff, UTF-8 logs/reports, WSMan probe, signed fixers, progress, optional rescan).

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
    [Parameter(ParameterSetName='Remediate')][switch]$LaunchRescan
)

begin {
    #region Init
    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try { Import-Module Microsoft.WSMan.Management -ErrorAction SilentlyContinue } catch {}

    $workingDir = 'C:\temp\scripts'
    $logsDir    = Join-Path $workingDir 'logs'
    $reportsDir = Join-Path $workingDir 'reports'
    $fixersDir  = Join-Path $workingDir 'fixers'
    foreach ($d in @($workingDir,$logsDir,$reportsDir,$fixersDir)) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

    $stamp       = Get-Date -Format 'yyyyMMdd_HHmm'
    $logPath     = Join-Path $logsDir    "CheckandRemediate_$stamp.log"
    $reportPath  = Join-Path $reportsDir "FailedCompliance_$stamp.csv"
    $csvPath     = Join-Path $workingDir 'failed_postures.csv'

    if (-not (Test-Path $logPath)) { New-Item -ItemType File -Path $logPath -Force | Out-Null }
    Add-Content -Path $logPath -Value '' -Encoding UTF8

    function Write-Log {
        param([string]$Message,[ValidateSet('INFO','WARNING','ERROR','DEBUG')]$Level='INFO')
        $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        Write-Information $line -InformationAction Continue
        Add-Content -Path $logPath -Value $line -Encoding UTF8
    }

    Write-Log "Script started. PSVersion=$($PSVersionTable.PSVersion) ParamSet=$($PSCmdlet.ParameterSetName)"
    #endregion

    #region Qualys helpers (CSV-only)
    function Invoke-QualysRequest {
        param([string]$EndpointPath,[hashtable]$Body,[string]$Accept='text/csv',[string]$OutFile)
        $auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($QualysCredential.UserName):$($QualysCredential.GetNetworkCredential().Password)"))
        $headers = @{ Authorization="Basic $auth"; Accept=$Accept; 'X-Requested-With'='PowerShell' }
        $uri = ($QualysBaseUrl.TrimEnd('/')) + $EndpointPath

        $attempt=0; $max=4
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
                if ($attempt -ge $max) { throw }
                $delay = [math]::Pow(2,$attempt)
                Write-Log "Qualys request failed: $($_.Exception.Message). Retrying in ${delay}s..." 'WARNING'
                Start-Sleep -Seconds $delay
            }
        } while ($true)
    }

    function Get-FailedPostures {
        try {
            $body = @{ action='list'; policy_id=$PolicyId; output_format='csv_no_metadata'; status='Failed'; details='All' }
            if ($TruncationLimit -gt 0) { $body.truncation_limit = $TruncationLimit }

            Invoke-QualysRequest -EndpointPath '/api/2.0/fo/compliance/posture/info/' -Body $body -OutFile $csvPath | Out-Null

            # Some tenants prepend a small metadata table (e.g., POLICY ID,DATETIME). We must start at the real header row.
            $text  = [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($csvPath))
            if ([string]::IsNullOrWhiteSpace($text)) { Write-Log 'Empty CSV body from Qualys.' 'WARNING'; return @() }
            $lines = $text -split "`r?`n"
            $headerIndex = $null
            $headerPattern = '(?i)(\b(Control\s*ID|CID)\b).*(\b(IP|IP Address|Host|Hostname|DNS Name)\b)'
            for ($i=0; $i -lt $lines.Count; $i++) {
                $l = $lines[$i]
                if ($l -like '*,*' -and ($l -match $headerPattern)) { $headerIndex = $i; break }
            }
            if ($null -eq $headerIndex) {
                # Fall back to first non-empty, comma-containing line
                for ($i=0; $i -lt $lines.Count; $i++) { if ($lines[$i] -like '*,*') { $headerIndex = $i; break } }
            }
            if ($null -eq $headerIndex) { Write-Log 'Could not locate CSV header.' 'ERROR'; return @() }

            $csvClean = ($lines[$headerIndex..($lines.Count-1)] -join [Environment]::NewLine)
            $raw = $csvClean | ConvertFrom-Csv

            if (-not $raw -or $raw.Count -eq 0) {
                Write-Log 'No failed postures found.'
                return @()
            }

            # Trim BOM/whitespace
            $data = $raw | ForEach-Object {
                $o=[ordered]@{}
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
    #endregion

    #region CSV normalisation
    function Resolve-Header {
        param([string[]]$Available)

        $map = @{
            IP          = @('IP','IP Address','IP Address(es)','Host IP','Host IP Address')
            Hostname    = @('Hostname','Host Name','DNS Name','DNS Hostname','FQDN','NetBIOS','Computer Name')
            ControlID   = @('Control ID','CID','Control ID (CID)','CID (Control ID)','Control Identifier','Control')
            Evidence    = @('Posture Evidence','Evidence','Evidence/Results','Instance Evidence','Evidence Value','Actual Value','Value','Finding','Observed Value')
            Reason      = @('Control Statement','Reason','Reason for Failure','Failure Reason','Reason/Recommendation','Rationale','Recommendation','Expected Value','Expected')
        }

        $resolved=@{}
        foreach ($k in $map.Keys) {
            $alias = $map[$k] | Where-Object { $_ -in $Available }
            if ($alias) { $resolved[$k] = $alias[0] }
        }
        return $resolved
    }

    function Normalize-QualysRows {
        param([object[]]$Rows)

        if (-not $Rows -or $Rows.Count -eq 0) { return @() }
        $headers  = $Rows[0].PSObject.Properties.Name
        $resolved = Resolve-Header -Available $headers

        # Must have ControlID and at least one of IP/Hostname
        $missing = @()
        if ('ControlID' -notin $resolved.Keys) { $missing += 'ControlID' }
        if (('IP' -notin $resolved.Keys) -and ('Hostname' -notin $resolved.Keys)) { $missing += 'IP or Hostname' }

        if ($missing.Count -gt 0) {
            Write-Log ("Detected CSV headers (first 5): {0}" -f (($headers | Select-Object -First 5) -join ', ')) 'ERROR'
            $snap = ($Rows | Select-Object -First 1 | ConvertTo-Csv -NoTypeInformation)[1]
            Write-Log ("First data line snapshot: {0}" -f $snap) 'ERROR'
            throw "CSV missing required columns (canonical): $($missing -join ', ')."
        }

        $ipH  = $resolved['IP']
        $hnH  = $resolved['Hostname']
        $cidH = $resolved['ControlID']
        $evH  = $resolved['Evidence']
        $rsH  = $resolved['Reason']

        $Rows | ForEach-Object {
            $ip  = if ($ipH) { $_.$ipH } else { $null }
            $hn  = if ($hnH) { $_.$hnH } else { $null }
            $tgt = if ($ip) { $ip } else { $hn }
            [pscustomobject]@{
                TargetName = $tgt
                IP         = $ip
                Hostname   = $hn
                ControlID  = $_.$cidH
                Evidence   = if ($evH) { $_.$evH } else { '' }
                Reason     = if ($rsH) { $_.$rsH } else { '' }
            }
        }
    }
    #endregion

    #region Fixers, rescan, remoting
    function Discover-Fixers { param([int]$CID)
        $fixer = Get-ChildItem -Path $fixersDir -Filter "$CID-*.ps1" -File | Sort-Object Name | Select-Object -First 1
        if (-not $fixer) { Write-Log "No fixer found for CID $CID."; return $null }
        $sig = Get-AuthenticodeSignature -FilePath $fixer.FullName
        if ($sig.Status -ne 'Valid') { Write-Log "Fixer $($fixer.FullName) signature: $($sig.Status). Skipping." 'WARNING'; return $null }
        Write-Log "Discovered signed fixer: $($fixer.FullName) for CID $CID."
        return [scriptblock]::Create((Get-Content $fixer.FullName -Raw))
    }

    function Invoke-HostRescan { param([string]$HostIP)
        try {
            $body = @{ action='launch'; scan_title="AutoRescan_${HostIP}_$([DateTime]::UtcNow.ToString('yyyyMMdd_HHmmss'))"; ip=$HostIP; priority='Normal' }
            Invoke-QualysRequest -EndpointPath '/api/2.0/fo/scan/compliance/' -Body $body -Accept 'application/xml' | Out-Null
            Write-Log "Requested Qualys compliance rescan for $HostIP."
            return $true
        } catch { Write-Log "Failed to request Qualys rescan for $HostIP: $($_.Exception.Message)" 'WARNING'; return $false }
    }

    function Test-TargetForRemoting { param([string]$TargetName)
        if (-not (Test-Connection -ComputerName $TargetName -Count 1 -Quiet)) { return 'Host Offline' }
        try { Test-WSMan -ComputerName $TargetName -Authentication Default -ErrorAction Stop | Out-Null; 'OK' }
        catch { 'WinRM Unreachable' }
    }

    function Remediate-Target {
        param([string]$TargetName,[int]$CID,[object]$Evidence)
        $o = [pscustomobject]@{ TargetName=$TargetName; CID=$CID; FixAttempted='No'; Outcome='Not Attempted'; Details='' }

        $reach = Test-TargetForRemoting -TargetName $TargetName
        if ($reach -ne 'OK') {
            $o.Outcome=$reach; $o.Details='Skipping remediation.'
            Write-Log "Remediation skipped for $TargetName CID $CID: $reach" 'WARNING'
            return $o
        }

        $fixer = Discover-Fixers -CID $CID
        if (-not $fixer) { $o.Outcome='No Fixer Available'; return $o }

        if ($Remediate) {
            try {
                $ret = Invoke-Command -ComputerName $TargetName -Credential $AdminCredential -ScriptBlock $fixer -ErrorAction Stop
                $o.FixAttempted='Yes'
                if ($null -eq $ret) { $ret=[pscustomobject]@{Outcome='Succeeded';Details='No details.'} }
                if (-not $ret.PSObject.Properties.Match('Outcome')) { $ret | Add-Member Outcome 'Unknown' }
                if (-not $ret.PSObject.Properties.Match('Details')) { $ret | Add-Member Details '' }
                $o.Outcome = [string]$ret.Outcome
                $o.Details = [string]$ret.Details
                Write-Log "Remediation on $TargetName for CID $CID: $($o.Outcome)"
            } catch {
                $o.Outcome='Failed'; $o.Details=$_.Exception.Message
                Write-Log "Remediation failed on $TargetName for CID $CID: $($o.Details)" 'ERROR'
            }
        } else {
            Write-Log "Dry-run: Would attempt remediation on $TargetName for CID $CID."
        }
        return $o
    }
    #endregion
}

process {
    try {
        $raw = Get-FailedPostures
        if (-not $raw -or $raw.Count -eq 0) { return }

        $rows = Normalize-QualysRows -Rows $raw

        # Filter: must have TargetName and numeric ControlID
        $rows = $rows | Where-Object { $_.TargetName } | Where-Object { $_.ControlID -match '^\d+$' }

        # Group by target + CID
        $grouped = $rows |
            Group-Object -Property TargetName | ForEach-Object {
                $target = $_.Name
                $_.Group | Group-Object -Property ControlID | ForEach-Object {
                    [pscustomobject]@{
                        TargetName = $target
                        CID        = [int]$_.Name
                        Evidence   = ($_.Group | ForEach-Object Evidence | Where-Object { $_ } | Select-Object -Unique) -join '; '
                        Reason     = ($_.Group | ForEach-Object Reason   | Where-Object { $_ } | Select-Object -Unique) -join '; '
                        IP         = ($_.Group | ForEach-Object IP       | Where-Object { $_ } | Select-Object -First 1)
                        Hostname   = ($_.Group | ForEach-Object Hostname | Where-Object { $_ } | Select-Object -First 1)
                    }
                }
            }

        $results = New-Object System.Collections.Generic.List[object]
        $total = ($grouped | Measure-Object).Count; $i=0
        foreach ($g in $grouped) {
            $i++
            Write-Progress -Activity 'Remediation' -Status "Processing $($g.TargetName) CID $($g.CID)" -PercentComplete ([int](($i/$total)*100))
            $rem = Remediate-Target -TargetName $g.TargetName -CID $g.CID -Evidence $g.Evidence
            if ($Remediate -and $LaunchRescan -and $rem.Outcome -match '^(Succeeded|Success|Fixed)$' -and $g.IP) {
                [void](Invoke-HostRescan -HostIP $g.IP)
            }
            $results.Add([pscustomobject]@{
                TargetName  = $g.TargetName
                IP          = $g.IP
                Hostname    = $g.Hostname
                CID         = $g.CID
                Evidence    = $g.Evidence
                Reason      = $g.Reason
                FixAttempted= $rem.FixAttempted
                Outcome     = $rem.Outcome
                Details     = $rem.Details
            })
        }
        Write-Progress -Activity 'Remediation' -Completed

        $mode = if ($Remediate) { 'Remediate' } else { 'Report' }
        $results | Select-Object @{n='Mode';e={$mode}}, TargetName, IP, Hostname, CID, Evidence, Reason, FixAttempted, Outcome, Details |
            Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
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
