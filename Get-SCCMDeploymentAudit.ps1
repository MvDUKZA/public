#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Deployment Audit - Parses client logs to capture precise start/end timestamps
    for update download, installation and reboot, plus overall success/fail per machine.

.DESCRIPTION
    For each machine in the selected collection the script reads these log files remotely
    over the admin share (\\machine\c$\Windows\CCM\Logs\):

      UpdatesDeployment.log  – download start/end, install start/end, reboot flag
      WUAHandler.log         – WUA-level download and install events
      UpdatesHandler.log     – handler-level install events

    Reboot time is correlated from the System Event Log (EventID 1074 / 6005 / 6006)
    via CIM, anchored to the first reboot AFTER the install completed.

    CSV columns:
      Machine | Online |
      DownloadStart | DownloadEnd | DownloadDurationS |
      InstallStart  | InstallEnd  | InstallDurationS  |
      RebootTime | RebootType |
      OverallResult | ErrorCode | Notes

.PARAMETER SiteServer
    SCCM Management Point / Site Server FQDN.

.PARAMETER SiteCode
    SCCM Site Code (e.g. PS1).

.PARAMETER CollectionName
    Device Collection – selected via Out-GridView if omitted.

.PARAMETER KBArticle
    KB number to audit (e.g. KB5034441) – selected via Out-GridView if omitted.

.PARAMETER OutputPath
    Folder to write the CSV. Defaults to current directory.

.PARAMETER MaxConcurrent
    Parallel runspaces. Default: 20.

.EXAMPLE
    .\Get-SCCMDeploymentAudit.ps1

.EXAMPLE
    .\Get-SCCMDeploymentAudit.ps1 -SiteServer SCCM-MP01.corp.local -SiteCode PS1 `
        -CollectionName "All Workstations - Prod" -KBArticle KB5034441

.NOTES
    Version : 3.0
    Requires: UNC admin share access (\\machine\c$) on all targets.
              WinRM or DCOM access for CIM (reboot time).
              Read rights on SCCM Site Server WMI namespace root\SMS\site_<code>.
              Run as an account with local admin on all target machines.
#>

[CmdletBinding()]
param(
    [string]$SiteServer,
    [string]$SiteCode,
    [string]$CollectionName,
    [string]$KBArticle,
    [string]$OutputPath    = (Get-Location).Path,
    [int]   $MaxConcurrent = 20
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

#region ── Banner ──────────────────────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║       SCCM Deployment Audit Tool  v3.1                  ║" -ForegroundColor Cyan
Write-Host "  ║  Download · Install · Reboot  —  precise log timestamps ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
#endregion

#region ── Helper: plain text prompt ──────────────────────────────────────────
function Read-Prompt {
    param([string]$Label, [string]$Default = "")
    $display = if ($Default) { "$Label [$Default]" } else { $Label }
    $val     = Read-Host "  $display"
    if ([string]::IsNullOrWhiteSpace($val) -and $Default) { return $Default }
    return $val.Trim()
}
#endregion

#region ── Step 1 – Site Server & Site Code ───────────────────────────────────
if (-not $SiteServer) { $SiteServer = Read-Prompt "SCCM Site Server / MP FQDN" "SCCM-MP01.corp.local" }
if (-not $SiteCode)   { $SiteCode   = Read-Prompt "Site Code" "PS1" }
$sccmNamespace = "root\SMS\site_$SiteCode"
#endregion

#region ── Step 2 – Collection picker (Out-GridView) ─────────────────────────
if (-not $CollectionName) {
    Write-Host ""
    Write-Host "  [*] Fetching device collections from $SiteServer ..." -ForegroundColor Cyan
    try {
        $allCollections = Get-WmiObject -ComputerName $SiteServer `
                                        -Namespace $sccmNamespace `
                                        -Class SMS_Collection `
                                        -Filter "CollectionType = 2" `
                                        -ErrorAction Stop |
                          Select-Object Name,
                                        CollectionID,
                                        @{N='Members';     E={ $_.MemberCount }},
                                        Comment,
                                        @{N='LastRefresh'; E={
                                            [Management.ManagementDateTimeConverter]::ToDateTime($_.LastRefreshTime)
                                        }} |
                          Sort-Object Name

        if (-not $allCollections) { throw "No device collections returned." }
        Write-Host "  [+] $($allCollections.Count) collections found – select one and click OK." -ForegroundColor Green

        $sel = $allCollections | Out-GridView `
                   -Title "Select Device Collection to Audit  (single-select → OK)" `
                   -OutputMode Single

        if (-not $sel) { Write-Error "No collection selected."; exit 1 }
        $CollectionName = $sel.Name
        $CollectionID   = $sel.CollectionID
        Write-Host "  [+] Collection : $CollectionName  ($CollectionID)" -ForegroundColor Green
    }
    catch {
        Write-Warning "  [!] WMI collection query failed: $_"
        $CollectionName = Read-Prompt "Enter Collection Name manually"
        $CollectionID   = $null
    }
}
#endregion

#region ── Step 3 – KB / Deployment picker (Out-GridView) ────────────────────
if (-not $KBArticle) {
    Write-Host ""
    Write-Host "  [*] Fetching software update deployments from $SiteServer ..." -ForegroundColor Cyan
    try {
        $deployments = Get-WmiObject -ComputerName $SiteServer `
                                     -Namespace $sccmNamespace `
                                     -Class SMS_UpdatesAssignment `
                                     -ErrorAction Stop |
                       Select-Object AssignmentName, AssignmentID,
                                     @{N='Created';  E={
                                         [Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime)
                                     }},
                                     @{N='Deadline'; E={
                                         if ($_.EnforcementDeadline) {
                                             [Management.ManagementDateTimeConverter]::ToDateTime($_.EnforcementDeadline)
                                         } else { 'None' }
                                     }} |
                       Sort-Object Created -Descending

        Write-Host "  [*] Fetching individual software updates (KB catalogue) ..." -ForegroundColor Cyan
        $swUpdates = Get-WmiObject -ComputerName $SiteServer `
                                   -Namespace $sccmNamespace `
                                   -Class SMS_SoftwareUpdate `
                                   -Filter "IsSuperseded = 0 AND IsExpired = 0" `
                                   -ErrorAction Stop |
                     Select-Object ArticleID, BulletinID,
                                   @{N='Title';    E={ $_.LocalizedDisplayName }},
                                   @{N='Severity'; E={ $_.SeverityName }},
                                   @{N='Released'; E={
                                       [Management.ManagementDateTimeConverter]::ToDateTime($_.DateRevised)
                                   }},
                                   NumMissing |
                     Sort-Object Released -Descending

        Write-Host ""
        Write-Host "   [1]  Pick from Deployments (SUG assignments)" -ForegroundColor White
        Write-Host "   [2]  Pick individual KB from update catalogue" -ForegroundColor White
        Write-Host ""
        $mode = Read-Host "  Choice [1/2]"

        if ($mode -eq '2') {
            $selUpd = $swUpdates | Out-GridView `
                          -Title "Select KB / Software Update  (single-select → OK)" `
                          -OutputMode Single
            if (-not $selUpd) { Write-Error "No update selected."; exit 1 }
            $KBArticle = "KB$($selUpd.ArticleID)"
            Write-Host "  [+] Update : $KBArticle — $($selUpd.Title)" -ForegroundColor Green
        }
        else {
            $selDep = $deployments | Out-GridView `
                          -Title "Select Deployment / SUG Assignment  (single-select → OK)" `
                          -OutputMode Single
            if (-not $selDep) { Write-Error "No deployment selected."; exit 1 }

            Write-Host "  [*] Resolving KBs in '$($selDep.AssignmentName)' ..." -ForegroundColor Cyan
            $depUpdates = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace `
                              -Query "SELECT ArticleID,LocalizedDisplayName,DateRevised FROM SMS_SoftwareUpdate WHERE CI_ID IN (SELECT UpdateCI_ID FROM SMS_UpdatesAssignment WHERE AssignmentID = $($selDep.AssignmentID))" `
                              -ErrorAction SilentlyContinue |
                          Select-Object ArticleID,
                                        @{N='Title';    E={ $_.LocalizedDisplayName }},
                                        @{N='Released'; E={
                                            [Management.ManagementDateTimeConverter]::ToDateTime($_.DateRevised)
                                        }} |
                          Sort-Object Released -Descending

            if (@($depUpdates).Count -gt 1) {
                $selUpd = $depUpdates | Out-GridView `
                              -Title "Select specific KB within deployment  (Cancel = audit all)" `
                              -OutputMode Single
                $KBArticle = if ($selUpd) { "KB$($selUpd.ArticleID)" } else { "AllInDeployment" }
            }
            elseif ($depUpdates) {
                $KBArticle = "KB$(@($depUpdates)[0].ArticleID)"
            }
            else {
                $KBArticle = Read-Prompt "Could not resolve KBs automatically – enter KB Article"
            }
            Write-Host "  [+] Deployment : $($selDep.AssignmentName)" -ForegroundColor Green
            Write-Host "  [+] KB Article : $KBArticle" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "  [!] WMI update query failed: $_"
        $KBArticle = Read-Prompt "Enter KB Article manually (e.g. KB5034441)"
    }
}

$KBNumber = $KBArticle -replace '[^0-9]', ''   # digits only for log pattern matching

Write-Host ""
Write-Host "  ── Parameters confirmed ───────────────────────────────────" -ForegroundColor DarkGray
Write-Host "    Site Server  : $SiteServer"
Write-Host "    Site Code    : $SiteCode"
Write-Host "    Collection   : $CollectionName"
Write-Host "    KB Article   : $KBArticle  (numeric: $KBNumber)"
Write-Host "    Output Path  : $OutputPath"
Write-Host ""
#endregion

#region ── Step 4 – Resolve collection members ────────────────────────────────
Write-Host "  [*] Resolving collection members ..." -ForegroundColor Cyan
$machines = @()
try {
    if (-not $CollectionID) {
        $cq = "SELECT CollectionID FROM SMS_Collection WHERE Name = '$($CollectionName -replace "'","''")'"
        $co = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace -Query $cq -ErrorAction Stop
        if (-not $co) { throw "Collection '$CollectionName' not found." }
        $CollectionID = $co.CollectionID
    }
    $mq       = "SELECT Name FROM SMS_FullCollectionMembership WHERE CollectionID = '$CollectionID'"
    $members  = Get-WmiObject -ComputerName $SiteServer -Namespace $sccmNamespace -Query $mq -ErrorAction Stop
    $machines = @($members | Select-Object -ExpandProperty Name | Sort-Object -Unique)
    Write-Host "  [+] $($machines.Count) machines in collection." -ForegroundColor Green
}
catch {
    Write-Warning "  [!] Could not retrieve members via WMI: $_"
    Write-Host "  [?] Enter machine names manually (one per line, blank line to finish):" -ForegroundColor Yellow
    while ($true) {
        $ln = Read-Host "  "
        if ([string]::IsNullOrWhiteSpace($ln)) { break }
        $machines += $ln.Trim()
    }
    $machines = $machines | Sort-Object -Unique
    Write-Host "  [+] Using $($machines.Count) manually entered machines." -ForegroundColor Yellow
}
if ($machines.Count -eq 0) { Write-Error "No machines to audit."; exit 1 }
Write-Host ""
#endregion

#region ── Log-parsing scriptblock ────────────────────────────────────────────
#
#  Data sources in priority order:
#
#  1. WUA COM API (IUpdateSearcher.QueryHistory) via Invoke-Command
#     → Persists in C:\Windows\SoftwareDistribution\DataStore\DataStore.edb
#     → Survives CCM log rotation. Gives InstallDate natively.
#     → Download timestamps come from the ServerSelection / hresult fields.
#
#  2. CBS.log  (C:\Windows\Logs\CBS\CBS.log)
#     → "Installing package" / "Installed package" lines with full timestamps
#     → Survives log rotation longer than CCM logs
#
#  3. CCM Logs (UpdatesDeployment.log, WUAHandler.log, UpdatesHandler.log)
#     → Most granular (download start/end, install start/end) when present
#     → Parsed if logs haven't rolled
#
#  4. Registry  HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing
#     → Package install timestamps stored as FILETIME values
#
#  5. Win32_QuickFixEngineering (last resort - date only, no time on modern OS)
#
#  Runs inside each runspace. All helpers must be self-contained (no outer scope).
#
$parseScriptBlock = {
    param(
        [string]$MachineName,
        [string]$KBNumber,    # digits only  e.g.  5034441
        [string]$KBArticle    # full string   e.g.  KB5034441
    )

    # ── Result object ──────────────────────────────────────────────────────
    $r = [PSCustomObject]@{
        Machine           = $MachineName
        Online            = $false
        DownloadStart     = ''
        DownloadEnd       = ''
        DownloadDurationS = ''
        InstallStart      = ''
        InstallEnd        = ''
        InstallDurationS  = ''
        RebootTime        = ''
        RebootType        = ''
        OverallResult     = 'Unknown'
        ErrorCode         = ''
        DataSources       = ''
        Notes             = ''
    }

    # ── Helpers ────────────────────────────────────────────────────────────
    function Parse-CMTime ([string]$Line) {
        if ($Line -match 'date="(\d{2}-\d{2}-\d{4})"' -and $Line -match 'time="(\d{2}:\d{2}:\d{2})') {
            try { return [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'MM-dd-yyyy HH:mm:ss', $null) } catch {}
        }
        if ($Line -match '(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})') {
            try { return [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'yyyy-MM-dd HH:mm:ss', $null) } catch {}
        }
        return $null
    }
    function First-Match ([string[]]$Lines, [string]$Pat) {
        foreach ($l in $Lines) { if ($l -match $Pat) { return $l } }; return $null
    }
    function Last-Match ([string[]]$Lines, [string]$Pat) {
        $h = $null; foreach ($l in $Lines) { if ($l -match $Pat) { $h = $l } }; return $h
    }
    function Ts ([string]$dt) {
        if ($dt) { return ([datetime]$dt).ToString('yyyy-MM-dd HH:mm:ss') } else { return '' }
    }

    # ── Ping ───────────────────────────────────────────────────────────────
    if (-not (Test-Connection -ComputerName $MachineName -Count 1 -Quiet -EA SilentlyContinue)) {
        $r.OverallResult = 'Offline'; $r.Notes = 'Did not respond to ping'; return $r
    }
    $r.Online = $true

    # ══════════════════════════════════════════════════════════════════════
    # SOURCE 1 — WUA COM API via Invoke-Command (IUpdateSearcher.QueryHistory)
    #
    # This queries DataStore.edb on the remote machine — persists indefinitely,
    # completely unaffected by CCM log rotation.
    # Returns: Title (contains KB number), Date (install completion time),
    #          HResult (0 = success), ResultCode (1-5 scale)
    #
    # ResultCode: 1=InProgress 2=Succeeded 3=SucceededWithErrors 4=Failed 5=Aborted
    # ══════════════════════════════════════════════════════════════════════
    $wuaData = $null
    try {
        $wuaData = Invoke-Command -ComputerName $MachineName -ErrorAction Stop -ScriptBlock {
            param($kb)
            try {
                $searcher   = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher()
                $totalCount = $searcher.GetTotalHistoryCount()
                if ($totalCount -eq 0) { return $null }
                # Query all history (cap at 1000 for performance)
                $count   = [math]::Min($totalCount, 1000)
                $history = $searcher.QueryHistory(0, $count)
                $match   = @()
                for ($i = 0; $i -lt $history.Count; $i++) {
                    $entry = $history.Item($i)
                    if ($entry.Title -match $kb -or $entry.Title -match "KB$kb") {
                        $match += [PSCustomObject]@{
                            Title        = $entry.Title
                            Date         = $entry.Date          # [datetime] – install completion
                            ResultCode   = $entry.ResultCode    # 2=Success 4=Failed
                            HResult      = '0x{0:X8}' -f [uint32]$entry.HResult
                            Operation    = $entry.Operation     # 1=Install 2=Uninstall
                            ClientAppID  = $entry.ClientApplicationID
                        }
                    }
                }
                return $match
            }
            catch { return $null }
        } -ArgumentList $KBNumber
    }
    catch {
        $r.Notes += "WUA COM invoke failed ($($_.Exception.Message)); "
    }

    $wuaEntry = $null
    if ($wuaData) {
        # Prefer a successful install entry; fall back to any entry for this KB
        $wuaEntry = @($wuaData | Where-Object { $_.Operation -eq 1 -and $_.ResultCode -eq 2 }) | Select-Object -Last 1
        if (-not $wuaEntry) {
            $wuaEntry = @($wuaData | Where-Object { $_.Operation -eq 1 }) | Select-Object -Last 1
        }
    }

    if ($wuaEntry) {
        # WUA Date = install COMPLETION time (most reliable single timestamp available)
        $r.InstallEnd   = Ts $wuaEntry.Date
        $r.ErrorCode    = $wuaEntry.HResult
        $r.DataSources += 'WUA-History; '

        $rcMap = @{ 1='InProgress'; 2='Succeeded'; 3='SucceededWithErrors'; 4='Failed'; 5='Aborted' }
        $rcText = if ($rcMap.ContainsKey([int]$wuaEntry.ResultCode)) { $rcMap[[int]$wuaEntry.ResultCode] } else { "RC$($wuaEntry.ResultCode)" }

        if ($wuaEntry.ResultCode -eq 4 -or $wuaEntry.ResultCode -eq 5) {
            $r.Notes += "WUA history: $rcText ($($wuaEntry.HResult)); "
        }
    }

    # ══════════════════════════════════════════════════════════════════════
    # SOURCE 2 — CBS.log  (C:\Windows\Logs\CBS\CBS.log)
    #
    # Format per line:
    #   YYYY-MM-DD HH:MM:SS, <type>,  CBS  <message>
    #
    # Key patterns:
    #   "Initiating changes to turn on update KB..."       ← install start
    #   "Installing package ... KB<n>"
    #   "Successfully installed package ... KB<n>"         ← install end (with timestamp)
    #   "Package ... KB<n> ... installation failed"        ← failure
    #
    # CBS.log is appended forever (rolled to CBS.persist.log when large) – far
    # outlasts CCM logs for older installs.
    # ══════════════════════════════════════════════════════════════════════
    $cbsLines = @()
    $cbsPath  = "\\$MachineName\c$\Windows\Logs\CBS\CBS.log"
    try {
        if (Test-Path $cbsPath -EA SilentlyContinue) {
            # Tail 80 000 lines – CBS.log can be very large
            $cbsLines = @(Get-Content $cbsPath -Tail 80000 -EA Stop)
            $r.DataSources += 'CBS.log; '
        }
    } catch { $r.Notes += "CBS.log unreadable; " }

    # Filter to lines mentioning this KB
    $cbsKBLines = @($cbsLines | Where-Object { $_ -match $KBNumber })

    if ($cbsKBLines.Count -gt 0) {
        # CBS install start
        $cbsStart = First-Match -Lines $cbsKBLines -Pat "(?i)(Initiating changes|Installing package|Begin.*installing|starting.*install)"
        if ($cbsStart -and -not $r.InstallStart) {
            $ts = Parse-CMTime $cbsStart
            # CBS lines begin with YYYY-MM-DD HH:MM:SS
            if (-not $ts -and $cbsStart -match '^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})') {
                try { $ts = [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'yyyy-MM-dd HH:mm:ss', $null) } catch {}
            }
            if ($ts) { $r.InstallStart = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        }

        # CBS install end (prefer success line)
        $cbsEnd = Last-Match -Lines $cbsKBLines -Pat "(?i)(Successfully installed|Install.*complet|Installed package)"
        if (-not $cbsEnd) { $cbsEnd = Last-Match -Lines $cbsKBLines -Pat "(?i)(package.*KB|KB.*package)" }
        if ($cbsEnd -and -not $r.InstallEnd) {
            $ts = $null
            if ($cbsEnd -match '^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})') {
                try { $ts = [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'yyyy-MM-dd HH:mm:ss', $null) } catch {}
            }
            if (-not $ts) { $ts = Parse-CMTime $cbsEnd }
            if ($ts) { $r.InstallEnd = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        }

        # CBS failure
        if (-not $r.ErrorCode -or $r.ErrorCode -eq '0x00000000') {
            $cbsFail = Last-Match -Lines $cbsKBLines -Pat "(?i)(failed|failure|error)"
            if ($cbsFail -and $cbsFail -match '(0x[0-9A-Fa-f]{8})') {
                $r.ErrorCode = $Matches[1].ToUpper()
            }
        }
    }

    # ══════════════════════════════════════════════════════════════════════
    # SOURCE 3 — CCM Logs (UpdatesDeployment, WUAHandler, UpdatesHandler)
    #
    # Best source for DOWNLOAD timestamps and granular install start/end.
    # Only useful when logs haven't rolled – but check anyway as they may
    # partially cover the event even if WUA/CBS already gave install time.
    # ══════════════════════════════════════════════════════════════════════
    $logRoot  = "\\$MachineName\c$\Windows\CCM\Logs"
    $udLines  = @(); $wuaLines = @(); $uhLines = @()
    $ccmReachable = Test-Path $logRoot -EA SilentlyContinue

    if ($ccmReachable) {
        try { $udLines  = @(Get-Content "$logRoot\UpdatesDeployment.log" -Tail 50000 -EA Stop); $r.DataSources += 'UpdatesDeployment.log; ' } catch { $r.Notes += "UpdatesDeployment.log unreadable; " }
        try { $wuaLines = @(Get-Content "$logRoot\WUAHandler.log"        -Tail 50000 -EA Stop); $r.DataSources += 'WUAHandler.log; '        } catch { $r.Notes += "WUAHandler.log unreadable; " }
        try { $uhLines  = @(Get-Content "$logRoot\UpdatesHandler.log"    -Tail 50000 -EA Stop); $r.DataSources += 'UpdatesHandler.log; '    } catch { $r.Notes += "UpdatesHandler.log unreadable; " }
    }
    else {
        $r.Notes += "CCM log share unreachable (\\$MachineName\c$\Windows\CCM\Logs); "
    }

    # ── Download Start (CCM logs are the only reliable source) ─────────────
    $dlStartLine = $null
    foreach ($pat in @(
        "(?i)($KBNumber.*(download)|(download).*$KBNumber)",
        "(?i)(Initiating download|Async installation.*started|Starting download for)",
        "(?i)(Download request GUID|Starting download for content)"
    )) {
        $dlStartLine = First-Match -Lines $wuaLines -Pat $pat
        if ($dlStartLine) { break }
    }
    if (-not $dlStartLine) {
        foreach ($pat in @(
            "(?i)(Starting download for content|Download request)",
            "(?i)(Initiating download)"
        )) {
            $dlStartLine = First-Match -Lines $udLines -Pat $pat
            if ($dlStartLine) { break }
        }
    }
    if ($dlStartLine) {
        $ts = Parse-CMTime $dlStartLine
        if ($ts) { $r.DownloadStart = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
    }

    # ── Download End ───────────────────────────────────────────────────────
    $dlEndLine = $null
    foreach ($pat in @(
        "(?i)(download.*success|Successfully.*download|Content download.*complet|WU client.*finish.*download)",
        "(?i)(Download.*complet|download.*succeeded)"
    )) {
        $dlEndLine = Last-Match -Lines $wuaLines -Pat $pat
        if ($dlEndLine) { break }
    }
    if (-not $dlEndLine) {
        foreach ($pat in @(
            "(?i)(Download.*succeeded|Content download.*complet)"
        )) {
            $dlEndLine = Last-Match -Lines $udLines -Pat $pat
            if ($dlEndLine) { break }
        }
    }
    if ($dlEndLine) {
        $ts = Parse-CMTime $dlEndLine
        if ($ts) { $r.DownloadEnd = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
    }

    # ── Install Start from CCM (only fill if not already set) ──────────────
    if (-not $r.InstallStart) {
        $isLine = $null
        $isLine = First-Match -Lines $udLines -Pat "(?i)(Installation job.*start|Performing install|Starting install)"
        if (-not $isLine) { $isLine = First-Match -Lines $uhLines  -Pat "(?i)(Installation job.*start|Processing update action)" }
        if (-not $isLine) { $isLine = First-Match -Lines $wuaLines -Pat "(?i)(Async installation.*started)"                      }
        if ($isLine) {
            $ts = Parse-CMTime $isLine
            if ($ts) { $r.InstallStart = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        }
    }

    # ── Install End from CCM (only fill if not already set by WUA/CBS) ─────
    if (-not $r.InstallEnd) {
        $ieLine = $null
        $ieLine = Last-Match -Lines $udLines -Pat "(?i)(Installation job.*complet|$KBNumber.*install.*success|install.*result 0x00000000)"
        if (-not $ieLine) { $ieLine = Last-Match -Lines $uhLines  -Pat "(?i)(Installation job.*complet|Update successfully installed)" }
        if (-not $ieLine) { $ieLine = Last-Match -Lines $wuaLines -Pat "(?i)(Successfully completed the installation|WU client finished installing)" }
        if ($ieLine) {
            $ts = Parse-CMTime $ieLine
            if ($ts) { $r.InstallEnd = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        }
    }

    # ── Error code from CCM logs (only if not set) ─────────────────────────
    if (-not $r.ErrorCode -or $r.ErrorCode -eq '') {
        foreach ($lines in @($udLines, $uhLines, $wuaLines)) {
            $el = Last-Match -Lines $lines -Pat "(?i)(result|error|fail).*(0x[89A-Fa-f][0-9A-Fa-f]{7})"
            if ($el -and $el -match '(0x[0-9A-Fa-f]{8})') { $r.ErrorCode = $Matches[1].ToUpper(); break }
        }
    }

    # ── Download timestamps from SoftwareDistribution\Download folder ──────
    # If CCM logs rolled, folder timestamps give a rough download window
    if (-not $r.DownloadStart -or -not $r.DownloadEnd) {
        $sdPath = "\\$MachineName\c$\Windows\SoftwareDistribution\Download"
        if (Test-Path $sdPath -EA SilentlyContinue) {
            # Each downloaded update gets its own subfolder; find folders modified
            # in a window around the install time if we have it
            $anchor = if ($r.InstallEnd) { ([datetime]$r.InstallEnd).AddDays(-7) } else { (Get-Date).AddDays(-90) }
            $dlFolders = @(Get-ChildItem $sdPath -Directory -EA SilentlyContinue |
                           Where-Object { $_.LastWriteTime -ge $anchor } |
                           Sort-Object LastWriteTime)
            if ($dlFolders.Count -gt 0) {
                if (-not $r.DownloadStart) {
                    $r.DownloadStart = $dlFolders[0].LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $r.Notes += "DownloadStart approx from SoftwareDistribution folder; "
                }
                if (-not $r.DownloadEnd) {
                    $r.DownloadEnd = $dlFolders[-1].LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                    $r.Notes += "DownloadEnd approx from SoftwareDistribution folder; "
                }
                $r.DataSources += 'SoftwareDistribution\Download; '
            }
        }
    }

    # ── Durations ──────────────────────────────────────────────────────────
    if ($r.DownloadStart -and $r.DownloadEnd) {
        $s = ([datetime]$r.DownloadEnd - [datetime]$r.DownloadStart).TotalSeconds
        $r.DownloadDurationS = [math]::Max(0, [math]::Round($s, 0))
    }
    if ($r.InstallStart -and $r.InstallEnd) {
        $s = ([datetime]$r.InstallEnd - [datetime]$r.InstallStart).TotalSeconds
        $r.InstallDurationS = [math]::Max(0, [math]::Round($s, 0))
    }

    # ══════════════════════════════════════════════════════════════════════
    # SOURCE 4 — Registry  (Component Based Servicing)
    #
    # HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages
    # Each installed package has an "InstallTimeEpoch" DWORD (seconds since 1970-01-01)
    # and/or "InstallTime" binary FILETIME.
    # Use only to fill InstallEnd if still empty.
    # ══════════════════════════════════════════════════════════════════════
    if (-not $r.InstallEnd) {
        try {
            $regData = Invoke-Command -ComputerName $MachineName -EA Stop -ScriptBlock {
                param($kb)
                $base = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages'
                $pkgs = Get-ChildItem $base -EA SilentlyContinue |
                        Where-Object { $_.PSChildName -match $kb }
                $result = @()
                foreach ($pkg in $pkgs) {
                    $epoch = (Get-ItemProperty $pkg.PSPath -Name 'InstallTimeEpoch' -EA SilentlyContinue).InstallTimeEpoch
                    if ($epoch) {
                        $result += [PSCustomObject]@{
                            Package     = $pkg.PSChildName
                            InstallTime = ([datetime]'1970-01-01').AddSeconds($epoch)
                        }
                    }
                }
                return $result
            } -ArgumentList $KBNumber

            if ($regData) {
                $latest = @($regData | Sort-Object InstallTime) | Select-Object -Last 1
                $r.InstallEnd   = $latest.InstallTime.ToString('yyyy-MM-dd HH:mm:ss')
                $r.DataSources += 'Registry-CBS; '
                $r.Notes       += "InstallEnd from CBS registry; "
            }
        }
        catch { $r.Notes += "Registry CBS query failed; " }
    }

    # ══════════════════════════════════════════════════════════════════════
    # REBOOT — System Event Log via CIM
    #
    # EventID 1074 = planned restart (best for patch reboots)
    # EventID 6005 = Event Log service start (system came up)
    # EventID 6006 = Event Log service stop (clean shutdown)
    #
    # Anchor: first qualifying event AFTER InstallEnd.
    # ══════════════════════════════════════════════════════════════════════
    $cimSess = $null
    try {
        $opt     = New-CimSessionOption -Protocol Wsman
        $cimSess = New-CimSession -ComputerName $MachineName -SessionOption $opt -OperationTimeoutSec 20 -EA Stop
    }
    catch {
        try {
            $opt     = New-CimSessionOption -Protocol Dcom
            $cimSess = New-CimSession -ComputerName $MachineName -SessionOption $opt -OperationTimeoutSec 20 -EA Stop
        }
        catch { $r.Notes += "CIM session failed; " }
    }

    if ($cimSess) {
        try {
            $bootEvts = Get-CimInstance -CimSession $cimSess -ClassName Win32_NTLogEvent `
                            -Filter "Logfile='System' AND (EventCode=1074 OR EventCode=6005 OR EventCode=6006)" `
                            -EA SilentlyContinue |
                        Sort-Object TimeGenerated

            if ($bootEvts) {
                $anchor   = if ($r.InstallEnd) { [datetime]$r.InstallEnd } else { [datetime]::MinValue }
                $postBoot = @($bootEvts | Where-Object { $_.TimeGenerated -gt $anchor })

                $chosen = if ($postBoot.Count -gt 0) {
                    $ev1074 = $postBoot | Where-Object { $_.EventCode -eq 1074 } | Select-Object -First 1
                    if ($ev1074) { $ev1074 } else { $postBoot[0] }
                }
                else { $bootEvts | Select-Object -Last 1 }

                if ($chosen) {
                    $r.RebootTime = $chosen.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss')
                    $r.RebootType = switch ($chosen.EventCode) {
                        1074    { 'Planned restart (EventID 1074)' }
                        6005    { 'System came online (EventID 6005)' }
                        6006    { 'Clean shutdown (EventID 6006)' }
                        default { "EventID $($chosen.EventCode)" }
                    }
                    if ($postBoot.Count -eq 0) { $r.RebootType += ' [pre-dates install – likely already done]' }
                }
            }
        }
        catch { $r.Notes += "Event log query error: $($_.Exception.Message); " }
        finally { Remove-CimSession -CimSession $cimSess -EA SilentlyContinue }
    }

    # ── Reboot fallback: UpdatesDeployment.log reboot-required line ────────
    if (-not $r.RebootTime -and $udLines.Count -gt 0) {
        $rbLine = Last-Match -Lines $udLines -Pat "(?i)(reboot.*required|pending.*reboot|restart.*required)"
        if ($rbLine) {
            $ts = Parse-CMTime $rbLine
            if ($ts) {
                $r.RebootTime = $ts.ToString('yyyy-MM-dd HH:mm:ss')
                $r.RebootType = 'Reboot required flagged in log (actual reboot time unavailable)'
            }
        }
    }

    # ══════════════════════════════════════════════════════════════════════
    # OVERALL RESULT
    # ══════════════════════════════════════════════════════════════════════
    $installFailed = $false
    if ($wuaEntry -and ($wuaEntry.ResultCode -eq 4 -or $wuaEntry.ResultCode -eq 5)) { $installFailed = $true }
    if (-not $installFailed -and ($r.ErrorCode -and $r.ErrorCode -ne '0x00000000')) {
        # Non-zero code from CBS/CCM logs is indicative but not definitive – mark warn
        $r.Notes += "Non-zero error code present; "
    }

    if ($installFailed) {
        $r.OverallResult = 'Install Failed'
    }
    elseif ($r.InstallEnd -and $r.RebootTime) {
        $r.OverallResult = 'Success – Rebooted'
        if (-not $r.ErrorCode) { $r.ErrorCode = '0x00000000' }
    }
    elseif ($r.InstallEnd) {
        # Check if reboot is still required
        $needsReboot = $false
        if ($udLines.Count -gt 0) {
            $needsReboot = $null -ne (First-Match -Lines $udLines -Pat "(?i)(reboot.*required|restart.*required)")
        }
        $r.OverallResult = if ($needsReboot) { 'Installed – Reboot Pending' } else { 'Installed – No Reboot Required' }
        if (-not $r.ErrorCode) { $r.ErrorCode = '0x00000000' }
    }
    elseif ($r.InstallStart) {
        $r.OverallResult = 'Install Started – No Completion Found'
    }
    elseif ($r.DownloadEnd) {
        $r.OverallResult = 'Downloaded – Install Not Started'
    }
    elseif ($r.DownloadStart) {
        $r.OverallResult = 'Download In Progress / Incomplete'
    }
    else {
        $r.OverallResult = 'No Activity Found in Any Source'
        $r.Notes        += "Checked: WUA history, CBS.log, CCM logs, SoftwareDistribution folder, CBS registry; "
    }

    $r.DataSources = $r.DataSources.TrimEnd('; ').Trim()
    $r.Notes       = $r.Notes.TrimEnd('; ').Trim()
    return $r
}
#endregion

#region ── Parallel execution via runspace pool ────────────────────────────────
Write-Host "  [*] Auditing $($machines.Count) machines (max $MaxConcurrent parallel) ..." -ForegroundColor Cyan
Write-Host ""

$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxConcurrent)
$pool.Open()

$jobs = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($machine in $machines) {
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($parseScriptBlock)
    [void]$ps.AddArgument($machine)
    [void]$ps.AddArgument($KBNumber)
    [void]$ps.AddArgument($KBArticle)
    $jobs.Add([PSCustomObject]@{ Machine = $machine; PS = $ps; Handle = $ps.BeginInvoke() })
}

$results   = [System.Collections.Generic.List[PSCustomObject]]::new()
$completed = 0
$total     = $jobs.Count

while ($jobs.Count -gt 0) {
    $done = @($jobs | Where-Object { $_.Handle.IsCompleted })
    foreach ($job in $done) {
        try   { $res = $job.PS.EndInvoke($job.Handle) }
        catch { $res = $null }
        if ($res) {
            $results.Add($res)
            $completed++
            $icon  = switch -Wildcard ($res.OverallResult) {
                'Success*'             { '[+]' }
                'Installed – No*'      { '[+]' }
                'Already*'             { '[+]' }
                'Offline'              { '[x]' }
                '*Failed*'             { '[x]' }
                '*Access Failed*'      { '[x]' }
                default                { '[~]' }
            }
            $color = switch -Wildcard ($res.OverallResult) {
                'Success*'             { 'Green'  }
                'Installed – No*'      { 'Green'  }
                'Already*'             { 'Green'  }
                'Offline'              { 'Red'    }
                '*Failed*'             { 'Red'    }
                '*Access Failed*'      { 'Red'    }
                default                { 'Yellow' }
            }
            Write-Host ("  {0} {1,-30} {2}" -f $icon, $res.Machine, $res.OverallResult) -ForegroundColor $color
        }
        $job.PS.Dispose()
        $jobs.Remove($job)
    }
    Write-Progress -Activity "SCCM Log Audit — $KBArticle" `
                   -Status "$completed of $total complete" `
                   -PercentComplete ([int](($completed / $total) * 100))
    Start-Sleep -Milliseconds 300
}

$pool.Close(); $pool.Dispose()
Write-Progress -Activity "SCCM Log Audit — $KBArticle" -Completed
#endregion

#region ── Summary ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ── Summary ────────────────────────────────────────────────" -ForegroundColor DarkGray
$results | Group-Object OverallResult | Sort-Object Count -Descending | ForEach-Object {
    $c = switch -Wildcard ($_.Name) {
        'Success*'   { 'Green'  } 'Installed*' { 'Green'  } 'Already*' { 'Green' }
        '*Failed*'   { 'Red'    } 'Offline'    { 'Red'    } '*Access*' { 'Red'   }
        default      { 'Yellow' }
    }
    Write-Host ("  {0,-42} : {1}" -f $_.Name, $_.Count) -ForegroundColor $c
}
Write-Host ""
#endregion

#region ── Export CSV ─────────────────────────────────────────────────────────
$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFile = Join-Path $OutputPath "SCCM_Audit_${KBArticle}_${stamp}.csv"

$results |
    Select-Object Machine, Online,
                  DownloadStart, DownloadEnd, DownloadDurationS,
                  InstallStart,  InstallEnd,  InstallDurationS,
                  RebootTime,    RebootType,
                  OverallResult, ErrorCode, DataSources, Notes |
    Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

if (Test-Path $csvFile) {
    Write-Host "  [+] CSV saved to:" -ForegroundColor Green
    Write-Host "      $csvFile" -ForegroundColor Cyan
    Write-Host ""
    $open = Read-Host "  Open CSV now? [Y/N]"
    if ($open -match '^[Yy]') { Start-Process $csvFile }
}
else {
    Write-Warning "  [!] CSV export failed – check path: $OutputPath"
}


Write-Host ""
Write-Host "  Audit complete." -ForegroundColor Green
Write-Host ""
#endregion
