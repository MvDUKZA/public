#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Windows Update Deployment Log Analyser
    Server: XXXXX | Site: PRD

.DESCRIPTION
    1. Connects to the SCCM site server and retrieves ADR-based deployments.
    2. Prompts the operator to select a deployment via Out-GridView.
    3. Enumerates all devices targeted by that deployment.
    4. For each device, reads the relevant client-side log files — both the
       current log AND all rollover/dated backup logs (e.g. UpdatesHandler-20260316-035718.log)
       in strict chronological order (oldest rollover first, then current).
    5. Parses each log type with its own specific event patterns to extract
       download start/finish, install start/finish, and reboot events per KB.
    6. Exports a CSV: MachineName, KBArticleID, DownloadStart, DownloadEnd,
       InstallStart, InstallEnd, RebootRequired, RebootTime

.NOTES
    Run from a machine with:
      - ConfigurationManager PowerShell module (or SCCM console installed)
      - Admin rights to the SCCM site server
      - Read access to \\<client>\C$\Windows\CCM\Logs\ on target devices

    Rollover log naming convention handled:
      <BaseName>-yyyyMMdd-HHmmss.log   e.g. UpdatesHandler-20260316-035718.log
      <BaseName>-yyyyMMdd-HHmmss-1.log (additional rollover suffix)
    Files are sorted by the date/time embedded in their filename, so log
    history is always processed in the correct chronological order regardless
    of file system timestamps.
#>

[CmdletBinding()]
param (
    [string]$SiteServer = 'XXXXX',
    [string]$SiteCode   = 'PRD',
    [string]$CCMLogPath = 'C$\Windows\CCM\Logs',
    [string]$OutputCSV  = ".\SCCMUpdateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [switch]$VerboseLogs   # Print every log file path as it is opened
)

Set-StrictMode -Off   # Allow unset hash keys without error

#region ── Helper: SCCM module ───────────────────────────────────────────────

function Import-SCCMModule {
    $consolePath = "${env:SMS_ADMIN_UI_PATH}\..\ConfigurationManager.psd1"
    if (Test-Path $consolePath) {
        Import-Module $consolePath -ErrorAction Stop
        return
    }
    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        Import-Module ConfigurationManager -ErrorAction Stop
    }
}

#endregion

#region ── Helper: Log file discovery ────────────────────────────────────────

function Get-LogFilesOrdered {
    <#
    .SYNOPSIS
        Returns log file paths for a given base name in strict chronological
        order: oldest rollover first, then the current (un-dated) log last.

    .DESCRIPTION
        SCCM rolls logs by renaming the current file to:
          <Base>-yyyyMMdd-HHmmss.log          (standard)
          <Base>-yyyyMMdd-HHmmss-1.log        (when more than one rollover same second)

        We sort by the datetime embedded in the filename — NOT by file system
        LastWriteTime, which can be unreliable across network shares.
    #>
    param(
        [string]$UNCLogDir,
        [string]$BaseName
    )

    # Discover all rollover files: match BaseName-<digits>-<digits>[optional-suffix].log
    $rollovers = Get-ChildItem -Path $UNCLogDir -Filter "$BaseName-*.log" `
                               -ErrorAction SilentlyContinue

    $sorted = @(
        $rollovers | ForEach-Object {
            # Extract the yyyyMMdd-HHmmss portion from the filename
            $fn = $_.BaseName   # e.g. UpdatesHandler-20260316-035718 or CAS-20260311-210354
            if ($fn -match '-(\d{8})-(\d{6})') {
                $dtStr = "$($Matches[1])$($Matches[2])"   # 20260316035718
                try {
                    $dt = [datetime]::ParseExact($dtStr, 'yyyyMMddHHmmss', $null)
                } catch {
                    $dt = $_.LastWriteTime   # fallback
                }
            } else {
                $dt = $_.LastWriteTime
            }
            [PSCustomObject]@{ Path = $_.FullName; SortKey = $dt }
        } | Sort-Object SortKey
    )

    $files = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $sorted) { $files.Add($r.Path) }

    $current = Join-Path $UNCLogDir "$BaseName.log"
    if (Test-Path $current) { $files.Add($current) }

    return $files.ToArray()
}

#endregion

#region ── Helper: Log reading ───────────────────────────────────────────────

function Read-LogLines {
    <#
    Reads lines from one or more log files. Returns an array of
    [PSCustomObject]@{Line; Source} so we can report which file each line came from.
    #>
    param(
        [string[]]$Paths,
        [switch]$Verbose
    )
    $out = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($p in $Paths) {
        if (-not (Test-Path $p)) { continue }
        if ($Verbose) { Write-Host "      Reading: $(Split-Path $p -Leaf)" -ForegroundColor DarkGray }
        try {
            # FileShare.ReadWrite lets us read files the CCM agent currently has open.
            # ReadAllLines / Get-Content both request exclusive access and will fail
            # on live log files.
            $fs = [System.IO.File]::Open(
                      $p,
                      [System.IO.FileMode]::Open,
                      [System.IO.FileAccess]::Read,
                      [System.IO.FileShare]::ReadWrite)
            $reader = [System.IO.StreamReader]::new(
                          $fs,
                          [System.Text.Encoding]::Default,
                          $true)  # detectEncodingFromByteOrderMarks
            try {
                while (-not $reader.EndOfStream) {
                    $ln = $reader.ReadLine()
                    $out.Add([PSCustomObject]@{ Line = $ln; Source = $p })
                }
            } finally {
                $reader.Dispose()
                $fs.Dispose()
            }
        } catch {
            Write-Warning "      Could not read $p : $_"
        }
    }
    return $out.ToArray()
}

#endregion

#region ── Helper: Timestamp parsing ─────────────────────────────────────────

function Parse-SCCMTimestamp {
    <#
    Handles all SCCM CMTrace log timestamp formats:

      Format A (most common):
        <![LOG[message]LOG]!><time="HH:mm:ss.fff+UUU" date="MM-dd-yyyy" ...>
        e.g. time="12:02:55.123+000" date="03-16-2026"

      Format B (some older logs):
        date="MM-dd-yyyy" time="HH:mm:ss.fff-UUU"

      UTC offset (+UUU or -UUU) is stripped — we report local machine time
      as written in the log (the offset is the client's UTC bias, not ours).

    Returns [datetime] or $null.
    #>
    param([string]$Line)

    # Try Format A: time= before date=
    if ($Line -match 'time="(\d{2}:\d{2}:\d{2})[\.\d]*[+\-]\d+"[^>]*date="(\d{2}-\d{2}-\d{4})"') {
        try { return [datetime]::ParseExact("$($Matches[2]) $($Matches[1])", 'MM-dd-yyyy HH:mm:ss', $null) }
        catch {}
    }
    # Try Format B: date= before time=
    if ($Line -match 'date="(\d{2}-\d{2}-\d{4})"[^>]*time="(\d{2}:\d{2}:\d{2})[\.\d]*[+\-]\d+"') {
        try { return [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'MM-dd-yyyy HH:mm:ss', $null) }
        catch {}
    }
    return $null
}

#endregion

#region ── Helper: KB extraction ─────────────────────────────────────────────

function Get-KBsFromLine {
    <#
    Returns ALL KB article IDs found in a log line (there can be more than one).
    Handles:
      KB5079473          direct KB mention
      Article: 5079473   WUA/UpdatesHandler style
      ArticleID=5079473  some XML blobs in logs
    #>
    param([string]$Line)
    $kbs = [System.Collections.Generic.List[string]]::new()

    # Match explicit KB prefix
    $m = [regex]::Matches($Line, '(?i)\bKB(\d{6,8})\b')
    foreach ($hit in $m) { $kbs.Add("KB$($hit.Groups[1].Value)") }

    # Match bare article numbers after keyword
    $m2 = [regex]::Matches($Line, '(?i)(?:Article(?:ID)?[\s:=]+)(\d{6,8})\b')
    foreach ($hit in $m2) {
        $candidate = "KB$($hit.Groups[1].Value)"
        if (-not $kbs.Contains($candidate)) { $kbs.Add($candidate) }
    }

    return $kbs.ToArray()
}

#endregion

#region ── 1. Connect to SCCM ────────────────────────────────────────────────

Write-Host "`n[1/4] Loading ConfigurationManager module..." -ForegroundColor Cyan
try {
    Import-SCCMModule
} catch {
    Write-Error "Failed to load ConfigurationManager module. Ensure the SCCM console is installed.`n$_"
    exit 1
}

$originalLocation = Get-Location
if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
}
Set-Location "$SiteCode`:\"

#endregion

#region ── 2. Select deployment ──────────────────────────────────────────────

Write-Host "[2/4] Retrieving software update deployments from $SiteServer ($SiteCode)..." -ForegroundColor Cyan

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
if ($selected -is [array]) { $selected = $selected[0] }

Write-Host "  Selected : $($selected.AssignmentName)" -ForegroundColor Green
Write-Host "  Collection: $($selected.CollectionName)" -ForegroundColor Green

#endregion

#region ── 3. Get machines ────────────────────────────────────────────────────

Write-Host "[3/4] Enumerating devices in collection '$($selected.CollectionName)'..." -ForegroundColor Cyan

$members  = Get-CMCollectionMember -CollectionId $selected.CollectionID -ErrorAction Stop
$machines = $members | Select-Object -ExpandProperty Name | Sort-Object

Write-Host "  Found $($machines.Count) device(s)." -ForegroundColor Green
Set-Location $originalLocation

#endregion

#region ── 4. Parse logs per machine ─────────────────────────────────────────

Write-Host "[4/4] Parsing client logs on each device...`n" -ForegroundColor Cyan

$allLogBases = @('CAS','ContentTransferManager','DataTransferService',
                 'UpdatesHandler','UpdatesDeployment','WUAHandler','RebootCoordinator')

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($machine in $machines) {
    Write-Host "  ► $machine" -ForegroundColor Cyan
    $uncLogDir = "\\$machine\$CCMLogPath"

    if (-not (Test-Path $uncLogDir -ErrorAction SilentlyContinue)) {
        Write-Warning "    Log path unreachable: $uncLogDir — skipping."
        continue
    }

    # Discover and report log files found
    $logFileMap = @{}   # BaseName → string[] of ordered file paths
    $totalFiles = 0
    foreach ($base in $allLogBases) {
        $files = Get-LogFilesOrdered -UNCLogDir $uncLogDir -BaseName $base
        $logFileMap[$base] = $files
        if ($files.Count -gt 0) {
            Write-Host "    $base : $($files.Count) file(s)" -ForegroundColor DarkGray
            if ($VerboseLogs) {
                foreach ($f in $files) { Write-Host "      $(Split-Path $f -Leaf)" -ForegroundColor DarkGray }
            }
            $totalFiles += $files.Count
        }
    }

    if ($totalFiles -eq 0) {
        Write-Warning "    No log files found — skipping."
        continue
    }

    # ── Per-KB event hashtable ────────────────────────────────────────────────
    $kbData = @{}

    function Ensure-KBRecord ($kb) {
        if (-not $kbData.ContainsKey($kb)) {
            $kbData[$kb] = [PSCustomObject]@{
                MachineName    = $machine
                KBArticleID    = $kb
                DownloadStart  = $null
                DownloadEnd    = $null
                InstallStart   = $null
                InstallEnd     = $null
                RebootRequired = $false
                RebootTime     = $null
            }
        }
        return $kbData[$kb]
    }

    # ════════════════════════════════════════════════════════════════════════
    # PHASE 1 — Build GUID → KB map from UpdatesDeployment.log
    #
    # Real log line format observed:
    #   Update (Site_XXXX/SUM_<GUID>) Name (...KB5079473...) ArticleID (5079473)
    #   added to the targeted list of deployment ({collection-GUID})
    #
    # We extract the SUM_<GUID> and the ArticleID number from each such line
    # to build a lookup table used in Phase 2.
    # ════════════════════════════════════════════════════════════════════════

    $guidToKB = @{}   # SUM_<guid-string> (lower) → "KB######"

    $udLines = Read-LogLines -Paths $logFileMap['UpdatesDeployment'] -Verbose:$VerboseLogs
    foreach ($entry in $udLines) {
        $ln = $entry.Line
        # Extract SUM GUID:  SUM_xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
        $sumGuid = if ($ln -match '(?i)SUM_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
            $Matches[1].ToLower()
        } else { $null }

        # Extract ArticleID — format: ArticleID (5079473) or ArticleID=5079473
        $articleId = if ($ln -match '(?i)ArticleID[\s\(]+(\d{5,8})') { $Matches[1] } else { $null }

        if ($sumGuid -and $articleId) {
            $kbKey = "KB$articleId"
            if (-not $guidToKB.ContainsKey($sumGuid)) {
                $guidToKB[$sumGuid] = $kbKey
            }
            # Pre-create the KB record so it exists even if install events are missing
            Ensure-KBRecord $kbKey | Out-Null
        }

        # Also parse any direct KB mentions on the same line (belt-and-braces)
        $ts  = Parse-SCCMTimestamp -Line $ln
        $kbs = Get-KBsFromLine -Line $ln
        foreach ($kb in $kbs) { Ensure-KBRecord $kb | Out-Null }
    }

    Write-Host "    GUID→KB map: $($guidToKB.Count) entries" -ForegroundColor DarkGray

    # ════════════════════════════════════════════════════════════════════════
    # PHASE 2 — Parse UpdatesHandler.log using GUID→KB map
    #
    # UpdatesHandler uses SUM GUIDs, not KB numbers. Key patterns:
    #   "Update (Site_.../SUM_<guid>) - EnumeratingUpdates"       → we see this update
    #   "Update (Site_.../SUM_<guid>) - WaitForInstall"           → queued for install
    #   "Update (Site_.../SUM_<guid>) - Installing"               → install started
    #   "Update (Site_.../SUM_<guid>) - Installed"                → install complete
    #   "Update (Site_.../SUM_<guid>) - PendingReboot"            → reboot needed
    #   "Update (Site_.../SUM_<guid>) - Failed"                   → failed
    # ════════════════════════════════════════════════════════════════════════

    $uhLines = Read-LogLines -Paths $logFileMap['UpdatesHandler'] -Verbose:$VerboseLogs
    foreach ($entry in $uhLines) {
        $ln = $entry.Line
        $l  = $ln.ToLower()
        $ts = Parse-SCCMTimestamp -Line $ln

        # Extract SUM GUID from this line
        $sumGuid = if ($ln -match '(?i)SUM_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
            $Matches[1].ToLower()
        } else { $null }

        if (-not $sumGuid) { continue }

        # Resolve GUID → KB
        $kb = if ($guidToKB.ContainsKey($sumGuid)) { $guidToKB[$sumGuid] } else { $null }
        if (-not $kb) { continue }

        $rec = Ensure-KBRecord $kb

        # Install queued / starting
        if ($ts -and (-not $rec.InstallStart) -and
            ($l -match 'waitforinstall|- installing|- ciinstalling|cistate.*install|install.*action')) {
            $rec.InstallStart = $ts
        }
        # Install complete
        if ($ts -and (-not $rec.InstallEnd) -and
            ($l -match '- installed\b|- succeeded|successfully installed|install.*success|ciinstalled\b')) {
            $rec.InstallEnd = $ts
        }
        # Pending reboot
        if ($l -match 'pendingreboot|- pendingreboot|reboot.*required|pending.*reboot') {
            $rec.RebootRequired = $true
        }
    }

    # ════════════════════════════════════════════════════════════════════════
    # PHASE 3 — WUAHandler.log for install start/end (fallback if UpdatesHandler
    # didn't fire — WUA does log "Adding update (KBxxxxxxx) to the installation
    # list" and "Installation job completed" even when KB lines are sparse)
    # ════════════════════════════════════════════════════════════════════════

    $wuaLines = Read-LogLines -Paths $logFileMap['WUAHandler'] -Verbose:$VerboseLogs
    foreach ($entry in $wuaLines) {
        $ln  = $entry.Line
        $l   = $ln.ToLower()
        $ts  = Parse-SCCMTimestamp -Line $ln
        $kbs = Get-KBsFromLine -Line $ln

        foreach ($kb in $kbs) {
            $rec = Ensure-KBRecord $kb
            if ($ts -and (-not $rec.InstallStart) -and
                ($l -match 'adding update|install.*list|async.*install|wua.*install')) {
                $rec.InstallStart = $ts
            }
            if ($ts -and (-not $rec.InstallEnd) -and
                ($l -match 'successfully installed|install.*complet|installation.*job.*complet')) {
                $rec.InstallEnd = $ts
            }
            if ($l -match 'reboot.*required|pending.*reboot') {
                $rec.RebootRequired = $true
            }
        }

        # WUAHandler also logs "Installation job completed" without a KB on the line —
        # apply as InstallEnd to all KBs that have a start but no end yet
        if ($ts -and (-not $kbs) -and ($l -match 'installation.*job.*complet|async.*install.*complet')) {
            foreach ($kb in $kbData.Keys) {
                $rec = $kbData[$kb]
                if ($rec.InstallStart -and (-not $rec.InstallEnd)) {
                    $rec.InstallEnd = $ts
                }
            }
        }
    }

    # ════════════════════════════════════════════════════════════════════════
    # PHASE 4 — Download events from CAS / ContentTransferManager / DataTransferService
    # On VDI these are often absent (pre-staged content) but we still try.
    # We match on content GUIDs that appear in CAS alongside KB mentions in
    # UpdatesDeployment — or fall back to any KB number present on the line.
    # ════════════════════════════════════════════════════════════════════════

    foreach ($base in @('CAS','ContentTransferManager','DataTransferService')) {
        $lines = Read-LogLines -Paths $logFileMap[$base] -Verbose:$VerboseLogs
        foreach ($entry in $lines) {
            $ln  = $entry.Line
            $l   = $ln.ToLower()
            $ts  = Parse-SCCMTimestamp -Line $ln
            $kbs = Get-KBsFromLine -Line $ln

            foreach ($kb in $kbs) {
                $rec = Ensure-KBRecord $kb
                if ($ts -and (-not $rec.DownloadStart) -and
                    ($l -match 'requesting content|download.*start|initiating.*download|acquiring|job.*created')) {
                    $rec.DownloadStart = $ts
                }
                if ($ts -and (-not $rec.DownloadEnd) -and
                    ($l -match 'successfully retrieved|content.*available|transfer.*complet|transfer.*success|download.*complet')) {
                    $rec.DownloadEnd = $ts
                }
            }
        }
    }

    # ════════════════════════════════════════════════════════════════════════
    # PHASE 5 — RebootCoordinator: machine-level reboot timestamp
    # Applied to all KBs that have RebootRequired=true and no RebootTime yet.
    # ════════════════════════════════════════════════════════════════════════

    $rcLines = Read-LogLines -Paths $logFileMap['RebootCoordinator'] -Verbose:$VerboseLogs
    foreach ($entry in $rcLines) {
        $ln = $entry.Line
        $l  = $ln.ToLower()
        $ts = Parse-SCCMTimestamp -Line $ln

        $isRebootEvent = ($l -match 'scheduled reboot|rebootby|entered schedulerebootimpl|a reboot was requested|system.*restart.*initiated|initiating.*restart')

        if ($ts -and $isRebootEvent) {
            foreach ($kb in $kbData.Keys) {
                $rec = $kbData[$kb]
                if (-not $rec.RebootTime) {
                    $rec.RebootTime     = $ts
                    $rec.RebootRequired = $true
                }
            }
            break   # First reboot event per machine is enough
        }
    }

    # ── Collect results ───────────────────────────────────────────────────────
    $found = ($kbData.Keys | Measure-Object).Count
    Write-Host "    → $found KB record(s) extracted" -ForegroundColor $(if ($found -gt 0) { 'Green' } else { 'Yellow' })

    foreach ($kb in ($kbData.Keys | Sort-Object)) {
        $results.Add($kbData[$kb])
    }
}

#endregion

#region ── 5. Export CSV ──────────────────────────────────────────────────────

Write-Host ""
if ($results.Count -eq 0) {
    Write-Warning "No update events extracted from any machine. CSV not written."
    Write-Host "Tip: Run with -VerboseLogs to see exactly which files were opened." -ForegroundColor Yellow
} else {
    $results |
        Select-Object MachineName, KBArticleID,
                      DownloadStart, DownloadEnd,
                      InstallStart,  InstallEnd,
                      RebootRequired, RebootTime |
        Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding UTF8

    Write-Host "✔  CSV saved : $OutputCSV" -ForegroundColor Green
    Write-Host "   Rows      : $($results.Count)"   -ForegroundColor Green

    $results | Out-GridView -Title "SCCM Update Report — $($selected.AssignmentName)"
}

#endregion
