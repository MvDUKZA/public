#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Log Diagnostic — dumps every line containing a KB number (or known
    event keywords) from one machine's CCM logs so you can see the exact text
    and tune parsing patterns.

.USAGE
    # Dump KB-related lines from one machine to screen + file:
    .\Get-SCCMLogDiagnostic.ps1 -Machine VDIFD10208

    # Also dump ALL lines from a specific log (e.g. UpdatesHandler) to see
    # the full picture:
    .\Get-SCCMLogDiagnostic.ps1 -Machine VDIFD10208 -DumpFullLog UpdatesHandler

.OUTPUT
    .\SCCMDiag_<Machine>_<timestamp>.txt
#>

param(
    [Parameter(Mandatory)]
    [string]$Machine,

    [string]$CCMLogPath = 'C$\Windows\CCM\Logs',

    # Optional: dump every line from this log base name (no .log extension)
    [string]$DumpFullLog = '',

    [string]$OutputFile = ".\SCCMDiag_${Machine}_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
)

$logBases = @(
    'CAS',
    'ContentTransferManager',
    'DataTransferService',
    'UpdatesHandler',
    'UpdatesDeployment',
    'WUAHandler',
    'RebootCoordinator',
    'PolicyAgent'
)

$uncLogDir = "\\$Machine\$CCMLogPath"
Write-Host "Log directory: $uncLogDir" -ForegroundColor Cyan

if (-not (Test-Path $uncLogDir)) {
    Write-Error "Cannot reach $uncLogDir"
    exit 1
}

# ── Shared-read file open (same as main script) ───────────────────────────────
function Read-SharedLines([string]$Path) {
    $lines = [System.Collections.Generic.List[string]]::new()
    $fs     = [System.IO.File]::Open($Path,
                  [System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite)
    $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::Default, $true)
    try {
        while (-not $reader.EndOfStream) { $lines.Add($reader.ReadLine()) }
    } finally { $reader.Dispose(); $fs.Dispose() }
    return $lines.ToArray()
}

# ── Get ordered log files (rollover first, current last) ─────────────────────
function Get-OrderedLogs([string]$Dir, [string]$Base) {
    $rollovers = @(
        Get-ChildItem $Dir -Filter "$Base-*.log" -ErrorAction SilentlyContinue |
        ForEach-Object {
            $dt = if ($_.BaseName -match '-(\d{8})-(\d{6})') {
                try { [datetime]::ParseExact("$($Matches[1])$($Matches[2])",'yyyyMMddHHmmss',$null) }
                catch { $_.LastWriteTime }
            } else { $_.LastWriteTime }
            [PSCustomObject]@{ Path=$_.FullName; DT=$dt }
        } | Sort-Object DT | Select-Object -ExpandProperty Path
    )
    $current = Join-Path $Dir "$Base.log"
    $all = @()
    if ($rollovers) { $all += $rollovers }
    if (Test-Path $current) { $all += $current }
    return $all
}

$output = [System.Collections.Generic.List[string]]::new()
$output.Add("=== SCCM Log Diagnostic : $Machine  $(Get-Date) ===")
$output.Add("")

foreach ($base in $logBases) {
    $files = Get-OrderedLogs -Dir $uncLogDir -Base $base
    if (-not $files) { continue }

    $output.Add("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    $output.Add("LOG: $base  ($($files.Count) file(s))")
    $output.Add("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

    foreach ($f in $files) {
        $label = Split-Path $f -Leaf
        $output.Add("  ── $label ──")
        Write-Host "  Reading $label ..." -ForegroundColor DarkGray

        try { $lines = Read-SharedLines $f }
        catch { $output.Add("  ERROR: $_"); continue }

        if ($DumpFullLog -and ($base -eq $DumpFullLog)) {
            # Dump every line
            foreach ($ln in $lines) { $output.Add("    $ln") }
        } else {
            # Dump only lines that contain a KB number OR known event keywords
            $keywords = 'kb\d{6}|article|download|install|reboot|restart|content|transfer|wua|update.*status|cistate|waitfor'
            $matched = $lines | Where-Object { $_ -match $keywords }
            if ($matched) {
                foreach ($ln in $matched) { $output.Add("    $ln") }
                $output.Add("  ($($matched.Count) matching lines out of $($lines.Count) total)")
            } else {
                $output.Add("  (no matching lines in $($lines.Count) total lines)")
            }
        }
        $output.Add("")
    }
}

# Write to file and screen
$output | Set-Content $OutputFile -Encoding UTF8
Write-Host "`n✔  Diagnostic saved: $OutputFile" -ForegroundColor Green

# Also print to screen for quick review
$output | Write-Host
