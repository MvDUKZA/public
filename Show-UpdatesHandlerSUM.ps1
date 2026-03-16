#Requires -Version 5.1
<#
.SYNOPSIS
    Shows lines from UpdatesHandler.log that contain SUM GUIDs,
    so we can see the exact state-transition text SCCM logs for installs.
#>
param(
    [Parameter(Mandatory)][string]$Machine,
    [string]$CCMLogPath = 'C$\Windows\CCM\Logs'
)

$uncLogDir = "\\$Machine\$CCMLogPath"
if (-not (Test-Path $uncLogDir -ErrorAction SilentlyContinue)) { Write-Host "ERROR: $uncLogDir unreachable" -ForegroundColor Red; exit 1 }

function Read-SharedLines([string]$Path) {
    $fs = [System.IO.File]::Open($Path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
    $r  = [System.IO.StreamReader]::new($fs,[System.Text.Encoding]::Default,$true)
    $l  = [System.Collections.Generic.List[string]]::new()
    try { while (-not $r.EndOfStream) { $l.Add($r.ReadLine()) } } finally { $r.Dispose(); $fs.Dispose() }
    return $l.ToArray()
}

function Get-OrderedLogs([string]$Dir,[string]$Base) {
    $rv = @(Get-ChildItem $Dir -Filter "$Base-*.log" -EA SilentlyContinue |
        ForEach-Object {
            $dt = if ($_.BaseName -match '-(\d{8})-(\d{6})') {
                try{[datetime]::ParseExact("$($Matches[1])$($Matches[2])","yyyyMMddHHmmss",$null)}catch{$_.LastWriteTime}
            } else {$_.LastWriteTime}
            [PSCustomObject]@{P=$_.FullName;D=$dt}
        } | Sort-Object D | Select-Object -Expand P)
    $c = Join-Path $Dir "$Base.log"
    $a=@(); if($rv){$a+=$rv}; if(Test-Path $c){$a+=$c}; return $a
}

Write-Host "`n══ UpdatesHandler SUM lines — $Machine ══`n" -ForegroundColor Cyan

foreach ($f in (Get-OrderedLogs $uncLogDir 'UpdatesHandler')) {
    Write-Host "── $(Split-Path $f -Leaf)" -ForegroundColor Yellow
    try { $lines = Read-SharedLines $f } catch { Write-Host "ERROR: $_" -ForegroundColor Red; continue }

    $matched = $lines | Where-Object { $_ -match '(?i)SUM_[0-9a-f\-]{36}' }
    if ($matched) {
        foreach ($ln in $matched) {
            $msg = if ($ln -match '<!\[LOG\[(.+?)\]LOG\]') { $Matches[1] } else { $ln }
            $ts  = if ($ln -match 'time="(\d{2}:\d{2}:\d{2})') { $Matches[1] } else { '??:??:??' }
            Write-Host "  [$ts] $msg"
        }
        Write-Host "  ($($matched.Count) lines)" -ForegroundColor DarkGray
    } else {
        Write-Host "  (no SUM lines in $($lines.Count) total)" -ForegroundColor DarkGray
    }
}
Write-Host "`nDone — screenshot and share." -ForegroundColor Green
