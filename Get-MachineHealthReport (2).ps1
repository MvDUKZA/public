<#
.SYNOPSIS
    Checks a list of machines for online status, logged-on user, agent service
    state (Altiris / SCCM), free disk space, and optionally a KB patch and
    file version. Outputs results to CSV.

.DESCRIPTION
    Reads computer names from a text file (one per line), runs the checks
    against each machine, and writes a CSV report. Designed for ~30 machines,
    PowerShell 5.x compatible.

.PARAMETER ComputerListPath
    Path to a text file containing one computer name per line.

.PARAMETER OutputCsvPath
    Path where the CSV report will be written.

.PARAMETER KBNumber
    Optional. KB article number to check for (with or without 'KB' prefix,
    e.g. 'KB5034441' or '5034441').

.PARAMETER FilePath
    Optional. Local path on the remote machine to check, e.g.
    'C:\Program Files\MyApp\app.exe'. Converted to an admin share path
    automatically.

.PARAMETER ExpectedFileVersion
    Optional. File version to compare against (e.g. '10.0.1.234').
    Only meaningful when -FilePath is also supplied.

.EXAMPLE
    .\Get-MachineHealthReport.ps1 -ComputerListPath .\machines.txt

.EXAMPLE
    .\Get-MachineHealthReport.ps1 -ComputerListPath .\machines.txt `
        -KBNumber KB5034441 `
        -FilePath 'C:\Program Files\Altiris\Altiris Agent\AeXNSAgent.exe' `
        -ExpectedFileVersion '8.6.3456.0'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerListPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = ".\MachineHealthReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [string]$KBNumber,

    [Parameter(Mandatory = $false)]
    [string]$FilePath,

    [Parameter(Mandatory = $false)]
    [string]$ExpectedFileVersion
)

# Normalise KB number (strip 'KB' prefix if present, then re-add)
if ($KBNumber) {
    $KBNumber = 'KB' + ($KBNumber -replace '(?i)^kb', '')
}

# --- Validate input file ---
if (-not (Test-Path -Path $ComputerListPath)) {
    Write-Error "Computer list file not found: $ComputerListPath"
    exit 1
}

$Computers = Get-Content -Path $ComputerListPath | Where-Object { $_ -and $_.Trim() -ne '' } | ForEach-Object { $_.Trim() }

if (-not $Computers -or $Computers.Count -eq 0) {
    Write-Error "No computer names found in $ComputerListPath"
    exit 1
}

Write-Host "Processing $($Computers.Count) machine(s)..." -ForegroundColor Cyan

$Results = foreach ($Computer in $Computers) {

    Write-Host "Checking $Computer..." -ForegroundColor Yellow

    # Default record - everything starts as N/A so a partial failure still produces a complete row
    $Record = [PSCustomObject]@{
        ComputerName    = $Computer
        Online          = $false
        LoggedOnUser    = 'N/A'
        AltirisService  = 'N/A'
        SCCMService     = 'N/A'
        SystemDriveGB   = 'N/A'
        FreeSpaceGB     = 'N/A'
        FreePercent     = 'N/A'
        KBInstalled     = 'N/A'
        FileVersion     = 'N/A'
        VersionMatch    = 'N/A'
        Error           = ''
    }

    # --- 1. Ping test ---
    try {
        $Online = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction Stop
    }
    catch {
        $Online = $false
    }

    $Record.Online = $Online

    if (-not $Online) {
        $Record.Error = 'Offline / unreachable'
        $Record
        continue
    }

    # --- 2. Logged-on user (via WMI / CIM) ---
    try {
        $CompSys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop
        if ($CompSys.UserName) {
            $Record.LoggedOnUser = $CompSys.UserName
        }
        else {
            $Record.LoggedOnUser = 'None'
        }
    }
    catch {
        $Record.LoggedOnUser = 'Query failed'
        $Record.Error = "User query: $($_.Exception.Message); "
    }

    # --- 3. Service checks ---
    # Altiris DAgent (Deployment Solution) - service name varies between DS versions
    # (AClient on DS 6.x, "Altiris Object Host Service" / "Altiris Deployment Agent" on later).
    # Match by display name wildcard, then fall back to process check for dagent.exe.
    try {
        $Altiris = Get-Service -ComputerName $Computer -DisplayName 'Altiris*Deployment*','Altiris*Agent*' -ErrorAction Stop |
                   Select-Object -First 1
        if ($Altiris) {
            $Record.AltirisService = "$($Altiris.Status) ($($Altiris.Name))"
        }
        else {
            # Fall back: is the dagent.exe process actually running?
            $Proc = Get-WmiObject -Class Win32_Process -ComputerName $Computer `
                    -Filter "Name='dagent.exe'" -ErrorAction Stop
            if ($Proc) {
                $Record.AltirisService = 'Process running (no service found)'
            }
            else {
                $Record.AltirisService = 'Not installed'
            }
        }
    }
    catch {
        $Record.AltirisService = 'Not installed'
    }

    try {
        $SCCM = Get-Service -ComputerName $Computer -Name 'CcmExec' -ErrorAction Stop
        $Record.SCCMService = $SCCM.Status.ToString()
    }
    catch {
        $Record.SCCMService = 'Not installed'
    }

    # --- 4. Disk space (system drive only - usually C:) ---
    try {
        $Disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $Computer `
            -Filter "DeviceID='C:'" -ErrorAction Stop

        if ($Disk) {
            $TotalGB = [math]::Round($Disk.Size / 1GB, 2)
            $FreeGB  = [math]::Round($Disk.FreeSpace / 1GB, 2)
            $Pct     = if ($Disk.Size -gt 0) {
                [math]::Round(($Disk.FreeSpace / $Disk.Size) * 100, 1)
            } else { 0 }

            $Record.SystemDriveGB = $TotalGB
            $Record.FreeSpaceGB   = $FreeGB
            $Record.FreePercent   = $Pct
        }
    }
    catch {
        $Record.Error += "Disk query: $($_.Exception.Message)"
    }

    # --- 5. KB patch check (optional) ---
    if ($KBNumber) {
        try {
            $Hotfix = Get-HotFix -ComputerName $Computer -Id $KBNumber -ErrorAction Stop
            if ($Hotfix) {
                $InstalledOn = if ($Hotfix.InstalledOn) { $Hotfix.InstalledOn.ToString('yyyy-MM-dd') } else { 'Unknown date' }
                $Record.KBInstalled = "Installed ($InstalledOn)"
            }
        }
        catch {
            # Get-HotFix throws if the KB isn't present, so that's the "not installed" path
            if ($_.Exception.Message -match 'not found|No match') {
                $Record.KBInstalled = 'Not installed'
            }
            else {
                $Record.KBInstalled = 'Query failed'
                $Record.Error += "KB query: $($_.Exception.Message); "
            }
        }
    }

    # --- 6. File version check (optional) ---
    if ($FilePath) {
        try {
            # Convert local path (C:\foo\bar.exe) to UNC admin share (\\host\C$\foo\bar.exe)
            if ($FilePath -match '^[A-Za-z]:\\') {
                $DriveLetter = $FilePath.Substring(0, 1)
                $Remainder   = $FilePath.Substring(3)
                $RemotePath  = "\\$Computer\$DriveLetter`$\$Remainder"
            }
            else {
                $RemotePath = $FilePath
            }

            if (Test-Path -Path $RemotePath -ErrorAction Stop) {
                $FileInfo = Get-Item -Path $RemotePath -ErrorAction Stop
                $ActualVersion = $FileInfo.VersionInfo.FileVersion

                if ($ActualVersion) {
                    $Record.FileVersion = $ActualVersion

                    if ($ExpectedFileVersion) {
                        # Try a proper [version] comparison first, fall back to string match
                        try {
                            $ActualV   = [version]($ActualVersion -replace '[^\d\.].*$', '')
                            $ExpectedV = [version]($ExpectedFileVersion -replace '[^\d\.].*$', '')
                            $Record.VersionMatch = if ($ActualV -eq $ExpectedV) { 'Match' }
                                                   elseif ($ActualV -gt $ExpectedV) { 'Newer' }
                                                   else { 'Older' }
                        }
                        catch {
                            $Record.VersionMatch = if ($ActualVersion -eq $ExpectedFileVersion) { 'Match' } else { 'Mismatch' }
                        }
                    }
                }
                else {
                    $Record.FileVersion = 'No version info'
                }
            }
            else {
                $Record.FileVersion = 'File not found'
            }
        }
        catch {
            $Record.FileVersion = 'Query failed'
            $Record.Error += "File check: $($_.Exception.Message); "
        }
    }

    $Record
}

# --- Export CSV ---
try {
    $Results | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport written to: $OutputCsvPath" -ForegroundColor Green
    Write-Host "Total machines:  $($Results.Count)" -ForegroundColor Green
    Write-Host "Online:          $(($Results | Where-Object Online).Count)" -ForegroundColor Green
    Write-Host "Offline:         $(($Results | Where-Object { -not $_.Online }).Count)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to write CSV: $($_.Exception.Message)"
    exit 1
}
