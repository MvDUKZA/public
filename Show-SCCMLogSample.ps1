#Requires -Version 5.1
<#
.SYNOPSIS
    SCCM Log Screen Diagnostic

.USAGE
    .\Show-SCCMLogSample.ps1 -Machine VDIFD10208

    # To see every line in a specific log instead of just KB lines:
    .\Show-SCCMLogSample.ps1 -Machine VDIFD10208 -FullDump UpdatesHandler
#>

param(
    [Parameter(Mandatory)]
    [string]$Machine,

    [string]$CCMLogPath = 'C$\Windows\CCM\Logs',

    # Optionally dump ALL lines from one specific log base name
    [string]$FullDump = ''
)

$logBases = @(
    'CAS',
    'ContentTransferManager',
    'DataTransferService',
    'UpdatesHandler',
    'UpdatesDeployment',
    'WUAHandler',
    'RebootCoordinator'
)

$uncLogDir = "\\$Machine\$CCMLogPath"

if (-not (Test-Path $uncLogDir -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: Cannot reach $uncLogDir" -ForegroundColor Red
    exit 1
}

function Read-SharedLines([string]$Path) {
    $fs     = [System.IO.File]::Open($Path,
                  [System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::ReadWrite)
    $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::Default, $true)
    $lines  = [System.Collections.Generic.List[string]]::new()
    try   { while (-not $reader.EndOfStream) { $lines.Add($reader.ReadLine()) } }
    finally { $reader.Dispose(); $fs.Dispose() }
    return $lines.ToArray()
}

function Get-OrderedLogs([string]$Dir, [string]$Base) {
    $rollovers = @(
        Get-ChildItem $Dir -Filter "$Base-*.log" -ErrorAction SilentlyContinue |
        ForEach-Object {
            $dt = if ($_.BaseName -match '-(\d{8})-(\d{6})') {
                try   { [datetime]::ParseExact("$($Matches[1])$($Matches[2])","yyyyMMddHHmmss",$null) }
                catch { $_.LastWriteTime }
            } else { $_.LastWriteTime }
            [PSCustomObject]@{ Path = $_.FullName; DT = $dt }
        } | Sort-Object DT | Select-Object -ExpandProperty Path
    )
    $current = Join-Path $Dir "$Base.log"
    $all = @()
    if ($rollovers) { $all += $rollovers }
    if (Test-Path $current) { $all += $current }
    return $all
}

# Keywords to match — broad enough to catch any event language
$keywords = '(?i)kb\d{5,8}|article|download|install|reboot|restart|content.*avail|transfer.*complete|transfer.*success|wua|cistate|waitfor|success|failed|error'

Write-Host ""
Write-Host "══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SCCM Log Diagnostic  ►  $Machine" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════" -ForegroundColor Cyan

foreach ($base in $logBases) {
    $files = Get-OrderedLogs -Dir $uncLogDir -Base $base
    if (-not $files) {
        Write-Host "`n[$base] — no files found" -ForegroundColor DarkGray
        continue
    }

    Write-Host "`n┌─ $base ($($files.Count) file(s)) ─────────────────────────────" -ForegroundColor Yellow

    foreach ($f in $files) {
        $label = Split-Path $f -Leaf
        Write-Host "│  ── $label" -ForegroundColor DarkYellow

        try   { $lines = Read-SharedLines $f }
        catch { Write-Host "│  ERROR: $_" -ForegroundColor Red; continue }

        if ($FullDump -and ($base -eq $FullDump)) {
            # Print every line (use sparingly — can be very long)
            foreach ($ln in $lines) {
                # Strip the CMTrace metadata tail to keep lines readable
                $clean = $ln -replace '<!\[LOG\[','' -replace '\]LOG\]!>.*',''
                Write-Host "│    $clean"
            }
        } else {
            $matched = $lines | Where-Object { $_ -match $keywords }
            if ($matched) {
                foreach ($ln in $matched) {
                    # Extract just the message part (before CMTrace metadata)
                    $msg = if ($ln -match '<!\[LOG\[(.+?)\]LOG\]') { $Matches[1] } else { $ln }
                    # Extract timestamp
                    $ts  = if ($ln -match 'time="(\d{2}:\d{2}:\d{2})') { $Matches[1] } else { '??:??:??' }
                    Write-Host "│  [$ts] $msg"
                }
                Write-Host "│  ($($matched.Count) of $($lines.Count) lines shown)" -ForegroundColor DarkGray
            } else {
                Write-Host "│  (no matching lines in $($lines.Count) total)" -ForegroundColor DarkGray
            }
        }
    }
    Write-Host "└────────────────────────────────────────────────────" -ForegroundColor Yellow
}

Write-Host "`nDone. Screenshot the output above and share it." -ForegroundColor Green
