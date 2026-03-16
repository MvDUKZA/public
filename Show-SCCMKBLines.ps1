#Requires -Version 5.1
<#
.SYNOPSIS
    Prints ONLY lines containing KB article numbers from UpdatesHandler,
    WUAHandler, CAS and ContentTransferManager logs.
    Output is short enough to screenshot easily.
#>

param(
    [Parameter(Mandatory)]
    [string]$Machine,
    [string]$CCMLogPath = 'C$\Windows\CCM\Logs'
)

# Only the logs most likely to have download/install event text
$logBases = @('UpdatesHandler','WUAHandler','CAS','ContentTransferManager','UpdatesDeployment')

$uncLogDir = "\\$Machine\$CCMLogPath"
if (-not (Test-Path $uncLogDir -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: Cannot reach $uncLogDir" -ForegroundColor Red; exit 1
}

function Read-SharedLines([string]$Path) {
    $fs     = [System.IO.File]::Open($Path,[System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
    $reader = [System.IO.StreamReader]::new($fs,[System.Text.Encoding]::Default,$true)
    $lines  = [System.Collections.Generic.List[string]]::new()
    try   { while (-not $reader.EndOfStream) { $lines.Add($reader.ReadLine()) } }
    finally { $reader.Dispose(); $fs.Dispose() }
    return $lines.ToArray()
}

function Get-OrderedLogs([string]$Dir,[string]$Base) {
    $rv = @(Get-ChildItem $Dir -Filter "$Base-*.log" -ErrorAction SilentlyContinue |
        ForEach-Object {
            $dt = if ($_.BaseName -match '-(\d{8})-(\d{6})') {
                try { [datetime]::ParseExact("$($Matches[1])$($Matches[2])","yyyyMMddHHmmss",$null) }
                catch { $_.LastWriteTime }
            } else { $_.LastWriteTime }
            [PSCustomObject]@{Path=$_.FullName;DT=$dt}
        } | Sort-Object DT | Select-Object -ExpandProperty Path)
    $cur = Join-Path $Dir "$Base.log"
    $all = @(); if ($rv) {$all+=$rv}; if (Test-Path $cur) {$all+=$cur}
    return $all
}

Write-Host ""
Write-Host "══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  KB-only lines from $Machine" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════" -ForegroundColor Cyan

foreach ($base in $logBases) {
    $files = Get-OrderedLogs -Dir $uncLogDir -Base $base
    if (-not $files) { continue }

    Write-Host "`n┌─ $base ──────────────────────────────────────" -ForegroundColor Yellow

    foreach ($f in $files) {
        Write-Host "│ $(Split-Path $f -Leaf)" -ForegroundColor DarkYellow
        try { $lines = Read-SharedLines $f } catch { Write-Host "│ ERROR: $_" -ForegroundColor Red; continue }

        # ONLY lines that contain a KB number
        $kbLines = $lines | Where-Object { $_ -match '(?i)\bKB\d{5,8}\b|Article(?:ID)?[\s:=]+\d{5,8}' }

        if ($kbLines) {
            foreach ($ln in $kbLines) {
                $msg = if ($ln -match '<!\[LOG\[(.+?)\]LOG\]') { $Matches[1] } else { $ln }
                $ts  = if ($ln -match 'time="(\d{2}:\d{2}:\d{2})') { $Matches[1] } else { '??:??:??' }
                Write-Host "│  [$ts] $msg"
            }
            Write-Host "│  ($($kbLines.Count) KB lines found)" -ForegroundColor DarkGray
        } else {
            Write-Host "│  (no KB lines found in $($lines.Count) total lines)" -ForegroundColor DarkGray
        }
    }
    Write-Host "└──────────────────────────────────────────────" -ForegroundColor Yellow
}
Write-Host "`nDone — screenshot and share." -ForegroundColor Green
