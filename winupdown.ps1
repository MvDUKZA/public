# Signed by Marinus van Deventer

<#
.SYNOPSIS
    Downloads Windows updates from the Microsoft Update Catalogue using the MSCatalogLTS module.

.DESCRIPTION
    This script installs or updates the MSCatalogLTS module if needed, searches for updates based on queries,
    filters by architecture and other criteria, downloads the files, and logs the process.

.PARAMETER SearchQueries
    Array of search strings for updates (e.g., "Windows 11 Version 24H2 Cumulative Update x64").

.PARAMETER DownloadPath
    Directory to save downloaded updates (must exist).

.PARAMETER Architecture
    Filter by architecture (all, x64, x86, arm64; default: x64).

.PARAMETER IncludePreview
    Include preview updates (default: $false).

.PARAMETER LogFile
    Path to the log file (default: C:\temp\scripts\logs\DownloadUpdates.log).

.PARAMETER StrictSearch
    Use strict (exact match) searching (default: $true).

.EXAMPLE
    .\DownloadWindowsUpdates.ps1 -SearchQueries @("Windows 11 Version 24H2 Cumulative Update x64") -DownloadPath "C:\Updates" -Architecture "x64"

.NOTES
    Requires PowerShell 7+ for best performance.
    Based on MSCatalogLTS module: https://github.com/Marco-online/MSCatalogLTS
    Reference: https://learn.microsoft.com/en-us/powershell/module/?view=powershell-7.4
    Changelog: v1.0 - Initial version.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$SearchQueries,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$DownloadPath,

    [ValidateSet("all", "x64", "x86", "arm64")]
    [string]$Architecture = "x64",

    [switch]$IncludePreview,

    [string]$LogFile = "C:\temp\scripts\logs\DownloadUpdates.log",

    [switch]$StrictSearch = $true
)

#region Initialization
# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "PowerShell 7+ recommended for optimal performance."
}

# Create log directory if needed
$logDir = Split-Path $LogFile -Parent
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

# Function to log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -FilePath $LogFile -Append
    Write-Information $Message
}
Write-Log "Script started. Parameters: SearchQueries=$($SearchQueries -join ', '), DownloadPath=$DownloadPath, Architecture=$Architecture"

# Install or update MSCatalogLTS module
try {
    if (-not (Get-Module -ListAvailable -Name MSCatalogLTS)) {
        Write-Log "Installing MSCatalogLTS module..."
        Install-Module -Name MSCatalogLTS -Scope CurrentUser -Force -ErrorAction Stop
    } else {
        Write-Log "Updating MSCatalogLTS module..."
        Update-Module -Name MSCatalogLTS -Force -ErrorAction Stop
    }
    Import-Module -Name MSCatalogLTS -ErrorAction Stop
} catch {
    Write-Log "Failed to install/update MSCatalogLTS: $_" "ERROR"
    throw
}
#endregion

#region Main Logic
try {
    $allDownloads = @()
    $totalQueries = $SearchQueries.Count
    $queryProgress = 0

    foreach ($query in $SearchQueries) {
        $queryProgress += (50 / $totalQueries)
        Write-Progress -Activity "Searching Updates" -Status "Query: $query" -PercentComplete $queryProgress

        Write-Log "Searching for: $query"
        $searchParams = @{
            Search = $query
            AllPages = $true
            Architecture = $Architecture
            IncludePreview = $IncludePreview.IsPresent
        }
        if ($StrictSearch) { $searchParams.Strict = $true }

        $updates = Get-MSCatalogUpdate @searchParams -ErrorAction Stop
        if ($updates.Count -eq 0) {
            Write-Log "No updates found for query: $query" "WARNING"
            continue
        }

        # Filter to latest (assuming sorted by LastUpdated descending)
        $latestUpdate = $updates | Sort-Object LastUpdated -Descending | Select-Object -First 1
        Write-Log "Found latest update: $($latestUpdate.Title)"

        if ($PSCmdlet.ShouldProcess($latestUpdate.Title, "Download update")) {
            Write-Log "Downloading: $($latestUpdate.Title)"
            Save-MSCatalogUpdate -Update $latestUpdate -Destination $DownloadPath -DownloadAll -ErrorAction Stop
            $allDownloads += $latestUpdate
        }
    }

    Write-Progress -Activity "Downloading Updates" -Status "Complete" -PercentComplete 100

    # Return status
    [PSCustomObject]@{
        Status = "Success"
        DownloadedUpdates = $allDownloads
        LogFile = $LogFile
        DownloadPath = $DownloadPath
    }
} catch {
    Write-Log "Script error: $_" "ERROR"
    throw
} finally {
    Write-Log "Script completed."
}
#endregion
