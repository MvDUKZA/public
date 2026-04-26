<#
.SYNOPSIS
    Sets a registry value on a list of remote computers via PowerShell Remoting (WinRM).

.DESCRIPTION
    Reads a list of computer names from a text file (one per line) and sets the specified
    registry value on each. Creates the registry key path if it does not exist. Uses
    Invoke-Command over WinRM because Remote Registry service is disabled in the environment.
    Runs sequentially. Logs per-host results to a timestamped CSV next to the script.

.PARAMETER ComputerList
    Path to a text file containing one hostname per line. Blank lines and lines starting
    with '#' are ignored.

.PARAMETER RegPath
    Full registry path under HKLM, e.g. 'HKLM:\SOFTWARE\Contoso\Agent'
    Accepts either 'HKLM:\...' or 'HKEY_LOCAL_MACHINE\...' form.

.PARAMETER ValueName
    Name of the registry value to set.

.PARAMETER ValueData
    Data to write. For DWord/QWord pass a number; for String/ExpandString pass a string;
    for MultiString pass a string array; for Binary pass a byte array.

.PARAMETER ValueType
    Registry value type. Defaults to DWord.

.PARAMETER Credential
    Optional alternate credentials for WinRM.

.EXAMPLE
    .\Set-RemoteRegistryValue.ps1 -ComputerList .\machines.txt `
        -RegPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' `
        -ValueName 'NoAutoUpdate' -ValueData 1

.EXAMPLE
    .\Set-RemoteRegistryValue.ps1 -ComputerList .\machines.txt `
        -RegPath 'HKLM:\SOFTWARE\Contoso' -ValueName 'AgentMode' `
        -ValueData 'Production' -ValueType String
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ComputerList,

    [Parameter(Mandatory)]
    [string]$RegPath,

    [Parameter(Mandatory)]
    [string]$ValueName,

    [Parameter(Mandatory)]
    $ValueData,

    [ValidateSet('String','ExpandString','Binary','DWord','MultiString','QWord')]
    [string]$ValueType = 'DWord',

    [System.Management.Automation.PSCredential]$Credential
)

# --- Setup ---------------------------------------------------------------

# Normalise registry path to PS-Drive form (HKLM:\...) for the remote scriptblock.
if ($RegPath -match '^HKEY_LOCAL_MACHINE\\') {
    $RegPath = $RegPath -replace '^HKEY_LOCAL_MACHINE\\', 'HKLM:\'
}
if ($RegPath -notmatch '^HKLM:\\') {
    Write-Error "RegPath must be under HKLM (got '$RegPath')."
    return
}

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath   = Join-Path $scriptDir "RegistryUpdate-$timestamp.csv"

$computers = Get-Content -Path $ComputerList |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') }

if (-not $computers) {
    Write-Error "No computers found in '$ComputerList'."
    return
}

Write-Host ""
Write-Host "Target path : $RegPath" -ForegroundColor Cyan
Write-Host "Value name  : $ValueName" -ForegroundColor Cyan
Write-Host "Value data  : $ValueData" -ForegroundColor Cyan
Write-Host "Value type  : $ValueType" -ForegroundColor Cyan
Write-Host "Computers   : $($computers.Count)" -ForegroundColor Cyan
Write-Host "Log file    : $logPath" -ForegroundColor Cyan
Write-Host ""

# --- Remote scriptblock --------------------------------------------------

$remoteScript = {
    param($RegPath, $ValueName, $ValueData, $ValueType)

    $result = [pscustomobject]@{
        OldValue = $null
        NewValue = $null
        Created  = $false
        Error    = $null
    }

    try {
        if (-not (Test-Path -Path $RegPath)) {
            New-Item -Path $RegPath -Force | Out-Null
            $result.Created = $true
        } else {
            try {
                $existing = Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction Stop
                $result.OldValue = $existing.$ValueName
            } catch {
                # Value name does not yet exist under the key — leave OldValue null.
            }
        }

        New-ItemProperty -Path $RegPath -Name $ValueName -Value $ValueData `
            -PropertyType $ValueType -Force | Out-Null

        $verify = Get-ItemProperty -Path $RegPath -Name $ValueName -ErrorAction Stop
        $result.NewValue = $verify.$ValueName
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    return $result
}

# --- Main loop -----------------------------------------------------------

$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($computer in $computers) {
    $i++
    Write-Host ("[{0}/{1}] {2} ... " -f $i, $computers.Count, $computer) -NoNewline

    $row = [ordered]@{
        Timestamp    = (Get-Date -Format 's')
        ComputerName = $computer
        Status       = $null
        OldValue     = $null
        NewValue     = $null
        KeyCreated   = $null
        Error        = $null
    }

    # Connectivity pre-check (WinRM)
    try {
        $null = Test-WSMan -ComputerName $computer -ErrorAction Stop
    }
    catch {
        $row.Status = 'Offline/NoWinRM'
        $row.Error  = $_.Exception.Message
        Write-Host "OFFLINE/NoWinRM" -ForegroundColor Yellow
        $results.Add([pscustomobject]$row)
        continue
    }

    # Invoke remote change
    try {
        $invokeParams = @{
            ComputerName = $computer
            ScriptBlock  = $remoteScript
            ArgumentList = $RegPath, $ValueName, $ValueData, $ValueType
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $invokeParams.Credential = $Credential }

        $remote = Invoke-Command @invokeParams

        if ($remote.Error) {
            $row.Status     = 'Failed'
            $row.OldValue   = $remote.OldValue
            $row.KeyCreated = $remote.Created
            $row.Error      = $remote.Error
            Write-Host "FAILED ($($remote.Error))" -ForegroundColor Red
        } else {
            $row.Status     = 'Success'
            $row.OldValue   = $remote.OldValue
            $row.NewValue   = $remote.NewValue
            $row.KeyCreated = $remote.Created
            $createdTag = if ($remote.Created) { ' (key created)' } else { '' }
            Write-Host ("OK  old='{0}' -> new='{1}'{2}" -f $remote.OldValue, $remote.NewValue, $createdTag) -ForegroundColor Green
        }
    }
    catch {
        $row.Status = 'Failed'
        $row.Error  = $_.Exception.Message
        Write-Host "FAILED ($($_.Exception.Message))" -ForegroundColor Red
    }

    $results.Add([pscustomobject]$row)
}

# --- Output --------------------------------------------------------------

$results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

$summary = $results | Group-Object Status | Sort-Object Name
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
$summary | ForEach-Object { Write-Host ("  {0,-16} {1}" -f $_.Name, $_.Count) }
Write-Host ""
Write-Host "Log written to: $logPath" -ForegroundColor Cyan
