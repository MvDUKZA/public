#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Windows Update Deployment Log Analyser
    Server: XXXXX | Site: PRD

.DESCRIPTION
    1. Connects to the SCCM site server and retrieves ADR-based deployments.
    2. Prompts the operator to select a deployment via Out-GridView.
    3. Enumerates all devices targeted by that deployment.
    4. For each device, reads the relevant client-side log files from the
       default CCM log path (\\<device>\C$\Windows\CCM\Logs\).
    5. Parses download start/finish, install start/finish, and reboot events
       per KB / update article.
    6. Exports a CSV: MachineName, KBArticleID, DownloadStart, DownloadEnd,
       InstallStart, InstallEnd, RebootRequired, RebootTime

.NOTES
    Run from a machine with:
      - ConfigurationManager PowerShell module (or SCCM console installed)
      - Admin rights to the SCCM site server
      - Read access to \\<client>\C$\Windows\CCM\Logs\ on target devices

    Log files examined (current + rollover variants):
      CAS.log, ContentTransferManager.log, DataTransferService.log,
      PolicyAgent.log, RebootCoordinator.log, UpdatesDeployment.log,
      UpdatesHandler.log, WUAHandler.log
#>

[CmdletBinding()]
param (
    [string]$SiteServer   = 'XXXXX',
    [string]$SiteCode     = 'PRD',
    [string]$CCMLogPath   = 'C$\Windows\CCM\Logs',
    [string]$OutputCSV    = ".\SCCMUpdateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#region ── Helper functions ──────────────────────────────────────────────────

function Import-SCCMModule {
    # Try loading from console install path first, then module path
    $consolePath = "${env:SMS_ADMIN_UI_PATH}\..\ConfigurationManager.psd1"
    if (Test-Path $consolePath) {
        Import-Module $consolePath -ErrorAction Stop
        return
    }
    # Fallback: hope it is already on the PSModulePath
    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        Import-Module ConfigurationManager -ErrorAction Stop
    }
}

function Get-LogFiles {
    <#
    Returns an ordered list of log file paths (current + rolled-over) from
    a remote CCM log directory, for a given base log name.
    Rolled-over files match the pattern  BaseName-yyyyMMdd-HHmmss.log
    #>
    param(
        [string]$UNCLogDir,
        [string]$BaseName
    )
    $pattern  = "$BaseName-*.log"
    $rollover = @(Get-ChildItem -Path $UNCLogDir -Filter $pattern -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime)
    $current  = Join-Path $UNCLogDir "$BaseName.log"
    $files    = @()
    if ($rollover) { $files += $rollover.FullName }
    if (Test-Path $current) { $files += $current }
    return $files
}

function Read-LogContent {
    param([string[]]$Paths)
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $Paths) {
        if (Test-Path $p) {
            try {
                $content = Get-Content $p -Encoding Default -ErrorAction Stop
                $lines.AddRange([string[]]$content)
            } catch {
                Write-Warning "Could not read $p : $_"
            }
        }
    }
    return $lines.ToArray()
}

function Parse-SCCMTimestamp {
    <#
    SCCM log timestamp examples:
      date="03-16-2026" time="12:02:55.123+000"
      <time="12:02:55.123+000" date="03-16-2026" ...>
    Returns a [datetime] or $null
    #>
    param([string]$Line)
    if ($Line -match 'date="(\d{2}-\d{2}-\d{4})"\s+time="(\d{2}:\d{2}:\d{2})') {
        try { return [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'MM-dd-yyyy HH:mm:ss', $null) }
        catch {}
    }
    if ($Line -match 'time="(\d{2}:\d{2}:\d{2})[^"]*"\s+date="(\d{2}-\d{2}-\d{4})"') {
        try { return [datetime]::ParseExact("$($Matches[2]) $($Matches[1])", 'MM-dd-yyyy HH:mm:ss', $null) }
        catch {}
    }
    return $null
}

function Extract-KBFromLine {
    param([string]$Line)
    if ($Line -match 'KB(\d{6,8})') { return "KB$($Matches[1])" }
    if ($Line -match 'Article[:\s]+(\d{6,8})') { return "KB$($Matches[1])" }
    return $null
}

#endregion

#region ── 1. Connect to SCCM ────────────────────────────────────────────────

Write-Host "`n[1/4] Loading ConfigurationManager module..." -ForegroundColor Cyan
try {
    Import-SCCMModule
} catch {
    Write-Error "Failed to load ConfigurationManager module. Ensure the SCCM console is installed. $_"
    exit 1
}

$originalLocation = Get-Location
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
}
Set-Location "$SiteCode`:\"

#endregion

#region ── 2. Select deployment via Out-GridView ─────────────────────────────

Write-Host "[2/4] Retrieving ADR deployments from $SiteServer ($SiteCode)..." -ForegroundColor Cyan

# Get Software Update Deployments (AssignmentName often contains ADR name)
$deployments = Get-CMSoftwareUpdateDeployment -ErrorAction Stop |
    Select-Object AssignmentName, AssignmentID, CollectionName, CollectionID,
                  CreationTime, EnforcementDeadline, Description |
    Sort-Object CreationTime -Descending

if (-not $deployments) {
    Write-Error "No software update deployments found."
    Set-Location $originalLocation; exit 1
}

$selected = $deployments |
    Out-GridView -Title "Select a Windows Update Deployment — then click OK" -PassThru

if (-not $selected) {
    Write-Warning "No deployment selected. Exiting."
    Set-Location $originalLocation; exit 0
}

# Support multi-select: take the first if more than one
if ($selected -is [array]) { $selected = $selected[0] }

Write-Host "  Selected: $($selected.AssignmentName)  (Collection: $($selected.CollectionName))" -ForegroundColor Green

#endregion

#region ── 3. Get all machines in the deployment's collection ─────────────────

Write-Host "[3/4] Enumerating devices in collection '$($selected.CollectionName)'..." -ForegroundColor Cyan

$collectionMembers = Get-CMCollectionMember -CollectionId $selected.CollectionID -ErrorAction Stop
$machines = $collectionMembers | Select-Object -ExpandProperty Name | Sort-Object

Write-Host "  Found $($machines.Count) device(s)." -ForegroundColor Green
Set-Location $originalLocation  # leave SCCM PSDrive before doing file I/O

#endregion

#region ── 4. Parse logs per machine ─────────────────────────────────────────

Write-Host "[4/4] Parsing client logs on each device..." -ForegroundColor Cyan

# Log base names we care about (without .log extension)
$logBases = @(
    'CAS',
    'ContentTransferManager',
    'DataTransferService',
    'PolicyAgent',
    'RebootCoordinator',
    'UpdatesDeployment',
    'UpdatesHandler',
    'WUAHandler'
)

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($machine in $machines) {
    Write-Host "  Processing: $machine" -ForegroundColor DarkCyan
    $uncLogDir = "\\$machine\$CCMLogPath"

    if (-not (Test-Path $uncLogDir)) {
        Write-Warning "    Log path unreachable: $uncLogDir — skipping."
        continue
    }

    # Read all relevant logs into a single chronological array
    $allLines = @()
    foreach ($base in $logBases) {
        $files = Get-LogFiles -UNCLogDir $uncLogDir -BaseName $base
        if ($files) {
            $allLines += Read-LogContent -Paths $files
        }
    }

    if (-not $allLines) {
        Write-Warning "    No log content found for $machine."
        continue
    }

    # ── Parse events per KB ──────────────────────────────────────────────────
    # Key patterns (case-insensitive):
    #   Download start  : "Starting download" / "Initiating download" / "CAS: Requesting content"
    #   Download end    : "Download succeeded" / "Successfully downloaded" / "Content is available"
    #   Install start   : "Starting install" / "Initiating install" / "WUA: Installing update"
    #   Install end     : "Successfully installed" / "Install succeeded" / "Installation job completed"
    #   Reboot required : "Reboot is required" / "A restart is required" / "RebootCoordinator"
    #   Reboot time     : "Restart initiated" / "Machine is rebooting"

    # We build a hashtable keyed by KB article ID
    $kbData = @{}

    foreach ($line in $allLines) {
        $ts = Parse-SCCMTimestamp -Line $line
        $kb = Extract-KBFromLine  -Line $line
        if (-not $kb) { continue }

        if (-not $kbData.ContainsKey($kb)) {
            $kbData[$kb] = [PSCustomObject]@{
                MachineName     = $machine
                KBArticleID     = $kb
                DownloadStart   = $null
                DownloadEnd     = $null
                InstallStart    = $null
                InstallEnd      = $null
                RebootRequired  = $false
                RebootTime      = $null
            }
        }

        $rec = $kbData[$kb]
        $l   = $line.ToLower()

        # Download start
        if ($ts -and (-not $rec.DownloadStart) -and
            ($l -match 'start.*download|initiating.*download|requesting content|download.*started')) {
            $rec.DownloadStart = $ts
        }
        # Download end
        if ($ts -and (-not $rec.DownloadEnd) -and
            ($l -match 'download succeeded|successfully downloaded|content is available|download.*completed|download.*success')) {
            $rec.DownloadEnd = $ts
        }
        # Install start
        if ($ts -and (-not $rec.InstallStart) -and
            ($l -match 'start.*install|initiating.*install|install.*started|wua.*install|beginning install')) {
            $rec.InstallStart = $ts
        }
        # Install end
        if ($ts -and (-not $rec.InstallEnd) -and
            ($l -match 'install.*succeed|successfully installed|installation.*completed|install.*success|update.*installed')) {
            $rec.InstallEnd = $ts
        }
        # Reboot required
        if ($l -match 'reboot.*required|restart.*required|pending reboot|requires.*restart') {
            $rec.RebootRequired = $true
        }
        # Reboot time
        if ($ts -and (-not $rec.RebootTime) -and
            ($l -match 'restart initiated|machine.*reboot|system.*restart|initiating restart|rebooting')) {
            $rec.RebootTime = $ts
            $rec.RebootRequired = $true
        }
    }

    foreach ($kb in $kbData.Keys | Sort-Object) {
        $results.Add($kbData[$kb])
    }
}

#endregion

#region ── 5. Export CSV ──────────────────────────────────────────────────────

if ($results.Count -eq 0) {
    Write-Warning "No update events extracted from any machine logs. CSV not written."
} else {
    $results |
        Select-Object MachineName, KBArticleID,
                      DownloadStart, DownloadEnd,
                      InstallStart,  InstallEnd,
                      RebootRequired, RebootTime |
        Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding UTF8

    Write-Host "`n✔  Report saved to: $OutputCSV" -ForegroundColor Green
    Write-Host   "   Rows: $($results.Count)" -ForegroundColor Green

    # Open in Excel/GridView for a quick preview
    $results | Out-GridView -Title "SCCM Update Deployment Report — $($selected.AssignmentName)"
}

#endregion
