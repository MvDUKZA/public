#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Windows Update Deployment Log Analyser
    Server: appsmcm101fp.iprod.local | Site: PRD

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
    [string]$SiteServer = 'appsmcm101fp.iprod.local',
    [string]$SiteCode   = 'PRD',
    [string]$CCMLogPath = 'C$\Windows\CCM\Logs',
    [string]$OutputCSV  = ".\SCCMUpdateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    # Only record events on or after this datetime. Defaults to deployment CreationTime - 2h.
    [nullable[datetime]]$Since = $null,
    # By default, Microsoft Defender definition updates (KB2267602) and Edge updates
    # are excluded — they install continuously in the background and pollute the report.
    # Use these switches to include them.
    [switch]$IncludeDefender,
    [switch]$IncludeEdge,
    [switch]$VerboseLogs
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

# Establish event cutoff — ignore log entries before this point so we don't
# pick up stale historical installs (e.g. Defender definitions update daily).
# Use deployment CreationTime minus 2 hours as the floor, or the -Since override.
if (-not $Since) {
    $Since = $selected.CreationTime.AddHours(-2)
}
Write-Host "  Cutoff   : events before $($Since.ToString('dd/MM/yyyy HH:mm:ss')) will be ignored" -ForegroundColor DarkGray

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
    $kbData    = @{}
    $kbDescMap = @{}   # KB → friendly description (populated in Phase 1)

    function Ensure-KBRecord ($kb) {
        if (-not $kbData.ContainsKey($kb)) {
            $kbData[$kb] = [PSCustomObject]@{
                MachineName    = $machine
                KBArticleID    = $kb
                Description    = if ($kbDescMap.ContainsKey($kb)) { $kbDescMap[$kb] } else { '' }
                DownloadStart  = $null
                DownloadEnd    = $null
                InstallStart   = $null
                InstallEnd     = $null
                RebootRequired = $false
                RebootTime     = $null
            }
        } else {
            # Backfill description if it arrived after initial record creation
            if (-not $kbData[$kb].Description -and $kbDescMap.ContainsKey($kb)) {
                $kbData[$kb].Description = $kbDescMap[$kb]
            }
        }
        return $kbData[$kb]
    }

    # Returns $ts only if it falls on or after the deployment cutoff, else $null.
    function Valid-Ts ($ts) {
        if ($null -eq $ts) { return $null }
        if ($ts -ge $Since) { return $ts }
        return $null
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

        # Extract SUM GUID
        $sumGuid = if ($ln -match '(?i)SUM_([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
            $Matches[1].ToLower()
        } else { $null }

        # Extract ArticleID
        $articleId = if ($ln -match '(?i)ArticleID[\s\(]+(\d{5,8})') { $Matches[1] } else { $null }

        # Extract friendly name from:  Name (2026-03 Cumulative Update for Windows 11...)
        $friendlyName = if ($ln -match '(?i)\bName\s*\(([^)]+)\)') { $Matches[1].Trim() } else { '' }

        if ($sumGuid -and $articleId) {
            $kbKey = "KB$articleId"

            # ── Exclusion filter ──────────────────────────────────────────────
            # Defender definitions: KB2267602 by article ID, or name contains indicator
            $isDefender = ($kbKey -eq 'KB2267602') -or
                          ($friendlyName -match '(?i)Security Intelligence Update|Defender.*Antivirus.*Definition')
            # Edge updates
            $isEdge     = ($friendlyName -match '(?i)Microsoft Edge')

            if ($isDefender -and -not $IncludeDefender) { continue }
            if ($isEdge     -and -not $IncludeEdge)     { continue }
            # ─────────────────────────────────────────────────────────────────

            if (-not $guidToKB.ContainsKey($sumGuid)) {
                $guidToKB[$sumGuid] = $kbKey
            }
            # Store description (first non-empty name wins per KB)
            if ($friendlyName -and -not $kbDescMap.ContainsKey($kbKey)) {
                $kbDescMap[$kbKey] = $friendlyName
            }
            Ensure-KBRecord $kbKey | Out-Null
        }

        # Belt-and-braces: direct KB mentions
        $kbs = Get-KBsFromLine -Line $ln
        foreach ($kb in $kbs) {
            $isDefender = ($kb -eq 'KB2267602') -or ($ln -match '(?i)Security Intelligence Update|Defender.*Definition')
            $isEdge     = ($ln -match '(?i)Microsoft Edge')
            if ($isDefender -and -not $IncludeDefender) { continue }
            if ($isEdge     -and -not $IncludeEdge)     { continue }
            Ensure-KBRecord $kb | Out-Null
        }
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
    # PHASE 3 — WUAHandler.log
    #
    # ACTUAL log format observed:
    #
    #   CONTEXT line (has KB + WUA GUID):
    #     "1. Update (Missing): ...KB5079473... (90316cb0-..., 100)"
    #
    #   NEXT line — InstallStart (no KB on this line):
    #     "Async installation of updates started."
    #
    #   LATER line — InstallEnd (WUA GUID, no KB):
    #     "Update 1 (90316cb0-...) finished installing (0x00000000), Reboot Required? Yes"
    #
    #   DOWNLOAD lines (no KB):
    #     "Download progress callback: download result oPCode = 1"
    #     "Async download completed."
    #
    # Strategy:
    #   - Track last-seen KB(s) and WUA GUID from context lines
    #   - Apply them to the immediately following event lines
    #   - Build a wuaGUID→KB map for the "finished installing" lines
    # ════════════════════════════════════════════════════════════════════════

    $wuaLines   = Read-LogLines -Paths $logFileMap['WUAHandler'] -Verbose:$VerboseLogs
    $wuaGuidToKB = @{}   # WUA GUID (lower) → KB — built from context lines
    $lastKBs    = @()    # KB(s) seen on most recent context line
    $lastTs     = $null  # timestamp of that context line

    foreach ($entry in $wuaLines) {
        $ln  = $entry.Line
        $l   = $ln.ToLower()
        $ts  = Parse-SCCMTimestamp -Line $ln
        $kbs = Get-KBsFromLine -Line $ln

        # ── Context line: contains KB number(s) and optionally a WUA GUID ──
        # e.g. "1. Update (Missing): ...KB5079473... (90316cb0-9dfb-4e05-95df-3a29334d699f, 100)"
        if ($kbs.Count -gt 0) {

            # Apply exclusion filter — skip Defender/Edge context lines so they
            # never populate $lastKBs and can't bleed into subsequent event lines
            $filteredKBs = $kbs | Where-Object {
                $isDefender = ($_ -eq 'KB2267602') -or ($ln -match '(?i)Security Intelligence Update|Defender.*Definition')
                $isEdge     = ($ln -match '(?i)Microsoft Edge')
                if ($isDefender -and -not $IncludeDefender) { return $false }
                if ($isEdge     -and -not $IncludeEdge)     { return $false }
                return $true
            }

            if ($filteredKBs.Count -eq 0) {
                # This line only contains excluded KBs (Defender/Edge) — treat
                # it as if it doesn't exist so a prior valid $lastKBs context
                # (e.g. KB5079473) is preserved for subsequent event lines
                continue
            }

            $lastKBs = @($filteredKBs)
            $lastTs  = $ts

            # Extract WUA GUID from context line (bare GUID in parentheses, no SUM_ prefix)
            if ($ln -match '\(([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\s*,') {
                $wuaGuid = $Matches[1].ToLower()
                foreach ($kb in $kbs) {
                    if (-not $wuaGuidToKB.ContainsKey($wuaGuid)) {
                        $wuaGuidToKB[$wuaGuid] = $kb
                    }
                }
            }

            # Also handle "Reboot Required" on the same context line (rare but possible)
            if ($l -match 'reboot required\?\s*yes') {
                foreach ($kb in $kbs) { (Ensure-KBRecord $kb).RebootRequired = $true }
            }
            continue
        }

        # ── InstallStart: "Async installation of updates started." ──
        # Appears immediately after the context line(s)
        if ($ts -and $lastKBs.Count -gt 0 -and ($l -match 'async installation of updates started')) {
            foreach ($kb in $lastKBs) {
                $rec = Ensure-KBRecord $kb
                if (-not $rec.InstallStart) { $rec.InstallStart = $ts }
            }
            # Don't clear $lastKBs — we still need it for the finished line
        }

        # ── InstallEnd: "Update 1 (<wuaGUID>) finished installing (0x...), Reboot Required? Yes/No" ──
        if ($ts -and ($l -match 'finished installing')) {
            # Extract WUA GUID from this line
            $wuaGuid = if ($ln -match '\(([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\)') {
                $Matches[1].ToLower()
            } else { $null }

            # Resolve which KB this WUA GUID belongs to
            $targetKBs = if ($wuaGuid -and $wuaGuidToKB.ContainsKey($wuaGuid)) {
                @($wuaGuidToKB[$wuaGuid])
            } elseif ($lastKBs.Count -gt 0) {
                $lastKBs   # fallback to last seen KB
            } else { @() }

            $rebootRequired = ($l -match 'reboot required\?\s*yes')

            foreach ($kb in $targetKBs) {
                $rec = Ensure-KBRecord $kb
                if (-not $rec.InstallEnd) { $rec.InstallEnd = $ts }
                if ($rebootRequired)      { $rec.RebootRequired = $true }
            }
        }

        # ── DownloadStart: first "Download progress callback" after a context line ──
        if ($ts -and $lastKBs.Count -gt 0 -and ($l -match 'async download.*started|download progress callback')) {
            foreach ($kb in $lastKBs) {
                $rec = Ensure-KBRecord $kb
                if (-not $rec.DownloadStart) { $rec.DownloadStart = $ts }
            }
        }

        # ── DownloadEnd: "Async download completed." ──
        if ($ts -and $lastKBs.Count -gt 0 -and ($l -match 'async download completed')) {
            foreach ($kb in $lastKBs) {
                $rec = Ensure-KBRecord $kb
                if (-not $rec.DownloadEnd) { $rec.DownloadEnd = $ts }
            }
        }

        # ── Installation of updates completed (batch-level, all KBs in batch) ──
        if ($ts -and ($l -match '^installation of updates completed')) {
            foreach ($kb in $kbData.Keys) {
                $rec = $kbData[$kb]
                if ($rec.InstallStart -and (-not $rec.InstallEnd)) { $rec.InstallEnd = $ts }
            }
            $lastKBs = @()   # reset context after batch completes
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

    # ── Sanity-check timestamp ordering ──────────────────────────────────────
    foreach ($kb in ($kbData.Keys | Select-Object)) {
        $rec   = $kbData[$kb]
        $floor = if ($rec.DownloadEnd) { $rec.DownloadEnd } else { $Since }

        # InstallStart must be >= floor (DownloadEnd or $Since)
        if ($rec.InstallStart -and $rec.InstallStart -lt $floor) {
            $rec.InstallStart = $null
            $rec.InstallEnd   = $null
        }
        # InstallEnd without InstallStart is meaningless — clear it
        if ($rec.InstallEnd -and (-not $rec.InstallStart)) {
            $rec.InstallEnd = $null
        }
        # InstallEnd must be >= InstallStart
        if ($rec.InstallEnd -and $rec.InstallStart -and $rec.InstallEnd -lt $rec.InstallStart) {
            $rec.InstallEnd = $null
        }
        # DownloadEnd must be >= DownloadStart
        if ($rec.DownloadEnd -and $rec.DownloadStart -and $rec.DownloadEnd -lt $rec.DownloadStart) {
            $rec.DownloadEnd = $null
        }
    }

    # ── Final exclusion purge ─────────────────────────────────────────────────
    # Remove any KB records that slipped through (e.g. KB2267602 created from
    # WUAHandler context lines before the exclusion filter was applied)
    $toRemove = $kbData.Keys | Where-Object {
        $isDefender = ($_ -eq 'KB2267602')
        $isEdge     = ($kbData[$_].Description -match '(?i)Microsoft Edge')
        ($isDefender -and -not $IncludeDefender) -or ($isEdge -and -not $IncludeEdge)
    }
    foreach ($kb in $toRemove) { $kbData.Remove($kb) }

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
        Select-Object MachineName, KBArticleID, Description,
                      DownloadStart, DownloadEnd,
                      InstallStart,  InstallEnd,
                      RebootRequired, RebootTime |
        Export-Csv -Path $OutputCSV -NoTypeInformation -Encoding UTF8

    Write-Host "✔  CSV saved : $OutputCSV" -ForegroundColor Green
    Write-Host "   Rows      : $($results.Count)"   -ForegroundColor Green

    $results | Select-Object MachineName, KBArticleID, Description,
                             DownloadStart, DownloadEnd, InstallStart, InstallEnd,
                             RebootRequired, RebootTime |
        Out-GridView -Title "SCCM Update Report — $($selected.AssignmentName)"
}

#endregion
