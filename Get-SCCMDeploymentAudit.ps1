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
Write-Host "  ║       SCCM Deployment Audit Tool  v3.0                  ║" -ForegroundColor Cyan
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
#  Runs inside each runspace. All helpers must be self-contained (no outer scope).
#
$parseScriptBlock = {
    param(
        [string]$MachineName,
        [string]$KBNumber,    # digits only  e.g.  5034441
        [string]$KBArticle    # full string   e.g.  KB5034441
    )

    # ── Result object ───────────────────────────────────────────────────────
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
        Notes             = ''
    }

    # ── CMTrace timestamp parser ────────────────────────────────────────────
    # Log line format:  <![LOG[...message...]LOG]!><time="HH:mm:ss.fff+/-ofs" date="MM-DD-YYYY" ...>
    function Parse-CMTime ([string]$Line) {
        if ($Line -match 'date="(\d{2}-\d{2}-\d{4})"' -and $Line -match 'time="(\d{2}:\d{2}:\d{2})') {
            try { return [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'MM-dd-yyyy HH:mm:ss', $null) }
            catch {}
        }
        # Fallback: plain timestamp at start of line  YYYY-MM-DD HH:MM:SS
        if ($Line -match '(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})') {
            try { return [datetime]::ParseExact("$($Matches[1]) $($Matches[2])", 'yyyy-MM-dd HH:mm:ss', $null) }
            catch {}
        }
        return $null
    }

    function First-Match ([string[]]$Lines, [string]$Pat) {
        foreach ($l in $Lines) { if ($l -match $Pat) { return $l } }
        return $null
    }

    function Last-Match ([string[]]$Lines, [string]$Pat) {
        $hit = $null
        foreach ($l in $Lines) { if ($l -match $Pat) { $hit = $l } }
        return $hit
    }

    # ── Ping ────────────────────────────────────────────────────────────────
    if (-not (Test-Connection -ComputerName $MachineName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        $r.OverallResult = 'Offline'
        $r.Notes         = 'Did not respond to ping'
        return $r
    }
    $r.Online = $true

    # ── Admin share reachability ────────────────────────────────────────────
    $logRoot = "\\$MachineName\c$\Windows\CCM\Logs"
    if (-not (Test-Path $logRoot -ErrorAction SilentlyContinue)) {
        $r.OverallResult = 'Log Access Failed'
        $r.Notes         = "Cannot reach $logRoot – check admin share / firewall / permissions"
        return $r
    }

    # ── Read logs (tail 50 000 lines each) ──────────────────────────────────
    $udLines  = @(); $wuaLines = @(); $uhLines = @()
    try { $udLines  = @(Get-Content "$logRoot\UpdatesDeployment.log" -Tail 50000 -EA Stop) }   catch { $r.Notes += "UpdatesDeployment.log unreadable; " }
    try { $wuaLines = @(Get-Content "$logRoot\WUAHandler.log"        -Tail 50000 -EA Stop) }   catch { $r.Notes += "WUAHandler.log unreadable; " }
    try { $uhLines  = @(Get-Content "$logRoot\UpdatesHandler.log"    -Tail 50000 -EA Stop) }   catch { $r.Notes += "UpdatesHandler.log unreadable; " }

    # ══════════════════════════════════════════════════════════════════════
    # DOWNLOAD
    #
    # WUAHandler.log key lines (contains KB article in message text):
    #   Start : "Async installation of updates started"
    #           "Initiating download"
    #           "Starting download for update"
    #   End   : "Successfully completed the download"
    #           "Content download succeeded"
    #           "WU client finished downloading"
    #
    # UpdatesDeployment.log key lines:
    #   Start : "Starting download for content"
    #           "Download request GUID"
    #   End   : "Download succeeded for content"
    #           "Content download completed"
    # ══════════════════════════════════════════════════════════════════════

    # ── Download Start ──────────────────────────────────────────────────────
    # Try KB-specific match first, then generic download-start in WUAHandler
    $dlStartPatterns = @(
        "(?i)($KBNumber.*(download|install)|(download|install).*$KBNumber)",
        "(?i)(Initiating download|Async installation.*started|Starting download)",
        "(?i)(Download request|Starting download for content)"
    )
    $dlStartLine = $null
    foreach ($pat in $dlStartPatterns) {
        $dlStartLine = First-Match -Lines $wuaLines -Pat $pat
        if ($dlStartLine) { break }
    }
    if (-not $dlStartLine) {
        foreach ($pat in $dlStartPatterns[1..2]) {
            $dlStartLine = First-Match -Lines $udLines -Pat $pat
            if ($dlStartLine) { break }
        }
    }

    # ── Download End ────────────────────────────────────────────────────────
    $dlEndPatterns = @(
        "(?i)(download.*success|Successfully.*download|Content download.*complet|WU client.*finish.*download)",
        "(?i)(Download.*complet|download.*succeeded)"
    )
    $dlEndLine = $null
    foreach ($pat in $dlEndPatterns) {
        $dlEndLine = Last-Match -Lines $wuaLines -Pat $pat
        if ($dlEndLine) { break }
    }
    if (-not $dlEndLine) {
        foreach ($pat in $dlEndPatterns) {
            $dlEndLine = Last-Match -Lines $udLines -Pat $pat
            if ($dlEndLine) { break }
        }
    }

    if ($dlStartLine) {
        $ts = Parse-CMTime $dlStartLine
        if ($ts) { $r.DownloadStart = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        else     { $r.Notes += "Download start timestamp unparseable; " }
    }
    if ($dlEndLine) {
        $ts = Parse-CMTime $dlEndLine
        if ($ts) { $r.DownloadEnd = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        else     { $r.Notes += "Download end timestamp unparseable; " }
    }
    if ($r.DownloadStart -and $r.DownloadEnd) {
        $span = ([datetime]$r.DownloadEnd - [datetime]$r.DownloadStart).TotalSeconds
        $r.DownloadDurationS = [math]::Max(0, [math]::Round($span, 0))
    }

    # ══════════════════════════════════════════════════════════════════════
    # INSTALL
    #
    # UpdatesDeployment.log key lines:
    #   Start : "Installation job started"
    #           "Performing installation for"
    #           "Starting install for update"
    #   End   : "Installation job completed with result 0x0"   ← success
    #           "Update (KB...) installed successfully"
    #   Fail  : "Installation job failed"
    #           "Failed to install update"
    #           "result 0x8..." / error hex
    #
    # UpdatesHandler.log key lines:
    #   Start : "Processing update action"
    #           "Installation request sent"
    #   End   : "Update successfully installed"
    #           "Installation completed"
    # ══════════════════════════════════════════════════════════════════════

    # ── Install Start ───────────────────────────────────────────────────────
    $instStartPatterns = @(
        "(?i)(Installation job.*start|Performing install|Starting install.*$KBNumber|$KBNumber.*install.*start)",
        "(?i)(Installation job.*start|Processing update action|Installation request sent)"
    )
    $instStartLine = $null
    $instStartLine = First-Match -Lines $udLines -Pat $instStartPatterns[0]
    if (-not $instStartLine) { $instStartLine = First-Match -Lines $uhLines  -Pat $instStartPatterns[1] }
    if (-not $instStartLine) { $instStartLine = First-Match -Lines $wuaLines -Pat "(?i)(Async installation.*started)" }

    # ── Install End ─────────────────────────────────────────────────────────
    $instEndPatterns = @(
        "(?i)(Installation job.*complet|Update.*$KBNumber.*success|$KBNumber.*install.*success|successfully installed|install.*result 0x00000000)",
        "(?i)(Installation job.*complet|Update successfully installed|Installation completed|install.*result 0x0)"
    )
    $instEndLine = $null
    $instEndLine = Last-Match -Lines $udLines -Pat $instEndPatterns[0]
    if (-not $instEndLine) { $instEndLine = Last-Match -Lines $uhLines  -Pat $instEndPatterns[1] }
    if (-not $instEndLine) { $instEndLine = Last-Match -Lines $wuaLines -Pat "(?i)(Successfully completed the installation|WU client finished installing)" }

    # ── Install Failure ─────────────────────────────────────────────────────
    $instFailLine = $null
    $instFailLine = Last-Match -Lines $udLines -Pat "(?i)(Installation job.*fail|Failed.*install.*$KBNumber|$KBNumber.*install.*fail|install.*result 0x[^0])"
    if (-not $instFailLine) { $instFailLine = Last-Match -Lines $uhLines  -Pat "(?i)(Installation.*fail|Failed.*install)" }
    if (-not $instFailLine) { $instFailLine = Last-Match -Lines $wuaLines -Pat "(?i)(installation.*fail|Failed.*install)" }

    if ($instStartLine) {
        $ts = Parse-CMTime $instStartLine
        if ($ts) { $r.InstallStart = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        else     { $r.Notes += "Install start timestamp unparseable; " }
    }
    if ($instEndLine) {
        $ts = Parse-CMTime $instEndLine
        if ($ts) { $r.InstallEnd = $ts.ToString('yyyy-MM-dd HH:mm:ss') }
        else     { $r.Notes += "Install end timestamp unparseable; " }
    }
    if ($r.InstallStart -and $r.InstallEnd) {
        $span = ([datetime]$r.InstallEnd - [datetime]$r.InstallStart).TotalSeconds
        $r.InstallDurationS = [math]::Max(0, [math]::Round($span, 0))
    }

    # ── Error code ──────────────────────────────────────────────────────────
    # Extract the last non-zero hex error code from the relevant logs
    $errLine = $null
    if ($instFailLine) { $errLine = $instFailLine }
    else {
        foreach ($lines in @($udLines, $uhLines, $wuaLines)) {
            $candidate = Last-Match -Lines $lines -Pat "(?i)(result|error|fail).*(0x[89A-Fa-f][0-9A-Fa-f]{7})"
            if ($candidate) { $errLine = $candidate; break }
        }
    }
    if ($errLine -and $errLine -match '(0x[0-9A-Fa-f]{8})') {
        $r.ErrorCode = $Matches[1].ToUpper()
    }

    # ══════════════════════════════════════════════════════════════════════
    # REBOOT
    #
    # Primary: System Event Log via CIM
    #   EventID 1074 – Process/user initiated restart (most specific for patch reboot)
    #   EventID 6005 – Event Log service started (machine came back up)
    #   EventID 6006 – Event Log service stopped (clean shutdown)
    #
    # We anchor to the first reboot event AFTER InstallEnd.
    # Fallback: "Reboot required" / "Pending reboot" in UpdatesDeployment.log.
    # ══════════════════════════════════════════════════════════════════════

    $cimSess = $null
    try {
        $opt     = New-CimSessionOption -Protocol Wsman
        $cimSess = New-CimSession -ComputerName $MachineName -SessionOption $opt `
                                  -OperationTimeoutSec 20 -ErrorAction Stop
    }
    catch {
        try {
            $opt     = New-CimSessionOption -Protocol Dcom
            $cimSess = New-CimSession -ComputerName $MachineName -SessionOption $opt `
                                      -OperationTimeoutSec 20 -ErrorAction Stop
        }
        catch { $r.Notes += "CIM session failed – reboot time from log only; " }
    }

    if ($cimSess) {
        try {
            $evFilter = "Logfile='System' AND (EventCode=1074 OR EventCode=6005 OR EventCode=6006)"
            $bootEvts = Get-CimInstance -CimSession $cimSess `
                                        -ClassName Win32_NTLogEvent `
                                        -Filter $evFilter `
                                        -ErrorAction SilentlyContinue |
                        Sort-Object TimeGenerated

            if ($bootEvts) {
                # Prefer first event AFTER install completed
                $anchor   = if ($r.InstallEnd) { [datetime]$r.InstallEnd } else { [datetime]::MinValue }
                $postBoot = @($bootEvts | Where-Object { $_.TimeGenerated -gt $anchor })

                $chosen = if ($postBoot.Count -gt 0) {
                    # EventID 1074 (planned restart) is the most definitive
                    $ev1074 = $postBoot | Where-Object { $_.EventCode -eq 1074 } | Select-Object -First 1
                    if ($ev1074) { $ev1074 } else { $postBoot[0] }
                }
                else {
                    # Nothing after install – take the most recent boot event (already rebooted)
                    $bootEvts | Select-Object -Last 1
                }

                if ($chosen) {
                    $r.RebootTime = $chosen.TimeGenerated.ToString('yyyy-MM-dd HH:mm:ss')
                    $r.RebootType = switch ($chosen.EventCode) {
                        1074    { 'Planned restart (EventID 1074)' }
                        6005    { 'System came online (EventID 6005)' }
                        6006    { 'Clean shutdown (EventID 6006)' }
                        default { "EventID $($chosen.EventCode)" }
                    }
                    if ($anchor -eq [datetime]::MinValue) {
                        $r.RebootType += ' – install end unknown, may be unrelated'
                    }
                    elseif ($postBoot.Count -eq 0) {
                        $r.RebootType += ' – pre-dates install end, likely already done'
                    }
                }
            }
        }
        catch { $r.Notes += "Event log query error: $($_.Exception.Message); " }
        finally { Remove-CimSession -CimSession $cimSess -ErrorAction SilentlyContinue }
    }

    # Log-based reboot fallback (if CIM gave nothing)
    if (-not $r.RebootTime) {
        $rbLine = Last-Match -Lines $udLines -Pat "(?i)(reboot.*required|pending.*reboot|restart.*required|reboot.*pending)"
        if ($rbLine) {
            $ts = Parse-CMTime $rbLine
            if ($ts) {
                $r.RebootTime = $ts.ToString('yyyy-MM-dd HH:mm:ss')
                $r.RebootType = 'Reboot flagged in UpdatesDeployment.log (actual reboot time unavailable)'
            }
        }
    }

    # ══════════════════════════════════════════════════════════════════════
    # OVERALL RESULT
    # ══════════════════════════════════════════════════════════════════════

    # Determine if the fail line is more recent than the success line
    $failIsLatest = $false
    if ($instFailLine -and $instEndLine) {
        $tsFail = Parse-CMTime $instFailLine
        $tsEnd  = Parse-CMTime $instEndLine
        if ($tsFail -and $tsEnd -and ($tsFail -gt $tsEnd)) { $failIsLatest = $true }
    }
    elseif ($instFailLine -and -not $instEndLine) {
        $failIsLatest = $true
    }

    if ($failIsLatest) {
        $r.OverallResult = 'Install Failed'
        if (-not $r.ErrorCode) { $r.ErrorCode = 'See Notes / logs' }
        $r.Notes += "Install failure line found in log; "
    }
    elseif ($r.InstallEnd -and $r.RebootTime) {
        $r.OverallResult = 'Success – Rebooted'
        if (-not $r.ErrorCode) { $r.ErrorCode = '0x00000000' }
    }
    elseif ($r.InstallEnd) {
        $rbRequired = First-Match -Lines $udLines -Pat "(?i)(reboot.*required|restart.*required)"
        $r.OverallResult = if ($rbRequired) { 'Installed – Reboot Pending' } else { 'Installed – No Reboot Required' }
        if (-not $r.ErrorCode) { $r.ErrorCode = '0x00000000' }
    }
    elseif ($r.InstallStart) {
        $r.OverallResult = 'Install Started – No Completion Found'
    }
    elseif ($r.DownloadEnd) {
        $r.OverallResult = 'Downloaded – Install Not Started'
    }
    elseif ($r.DownloadStart) {
        $r.OverallResult = 'Download Started – Not Completed'
    }
    else {
        # Last resort: check Win32_QuickFixEngineering (already installed, logs may have rolled)
        $qfe = $null
        try {
            $optD = New-CimSessionOption -Protocol Dcom
            $cs2  = New-CimSession -ComputerName $MachineName -SessionOption $optD -OperationTimeoutSec 15 -ErrorAction Stop
            $qfe  = Get-CimInstance -CimSession $cs2 -ClassName Win32_QuickFixEngineering -ErrorAction SilentlyContinue |
                    Where-Object { $_.HotFixID -eq "KB$KBNumber" }
            Remove-CimSession $cs2 -ErrorAction SilentlyContinue
        } catch {}

        if ($qfe) {
            $r.OverallResult = 'Already Installed (pre-existing / logs rolled)'
            $r.ErrorCode     = '0x00000000'
            $r.Notes        += "Found KB$KBNumber in Win32_QuickFixEngineering; install predates log window; "
        }
        else {
            $r.OverallResult = 'No Activity Found'
            $r.Notes        += "No matching entries in any CCM log – update may not be targeted, logs may have rolled, or client not yet evaluated; "
        }
    }

    $r.Notes = ($r.Notes.TrimEnd('; ').Trim())
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
                  OverallResult, ErrorCode,   Notes |
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
