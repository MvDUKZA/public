# Signed by Marinus van Deventer

<#
.SYNOPSIS
    Downloads Windows updates from the Microsoft Update Catalogue using the MSCatalogLTS module.

.DESCRIPTION
    This script installs or updates the MSCatalogLTS module if needed (checking versions first), searches for updates based on queries,
    filters by architecture and other criteria, downloads the files to a monthyear subfolder, and logs the process. It creates the DownloadPath if it does not exist.

.PARAMETER SearchQueries
    Array of search strings for updates (e.g., "2025-08 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems").

.PARAMETER DownloadPath
    Base directory to save downloaded updates (subfolder with monthyear will be created; created if not exists).

.PARAMETER Architecture
    Filter by architecture (all, x64, x86, arm64; default: x64).

.PARAMETER IncludePreview
    Include preview updates (default: $false).

.PARAMETER LogFile
    Path to the log file (default: C:\temp\scripts\logs\DownloadUpdates.log).

.PARAMETER StrictSearch
    Use strict (exact match) searching (default: $true).

.EXAMPLE
    .\DownloadWindowsUpdates.ps1 -SearchQueries @("2025-08 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems") -DownloadPath "C:\Updates" -Architecture "x64"

.NOTES
    Requires PowerShell 7+ for best performance.
    Based on MSCatalogLTS module: https://github.com/Marco-online/MSCatalogLTS
    Reference: https://learn.microsoft.com/en-us/powershell/module/?view=powershell-7.4
    Changelog: v1.3 - Updated .EXAMPLE to use "2025-08 Cumulative Update for Windows 11 Version 24H2 for x64-based Systems" as the search query.
               v1.2 - Added check/create for DownloadPath if not exists; check module version before updating; updated return object.
               v1.1 - Added automatic creation of monthyear subfolder (e.g., 0825) under DownloadPath; updated return object.
               v1.0 - Initial version.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$SearchQueries,

    [Parameter(Mandatory = $true)]
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

# Create DownloadPath if not exists
if (-not (Test-Path $DownloadPath -PathType Container)) {
    Write-Log "Creating DownloadPath: $DownloadPath"
    New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
}

# Install or update MSCatalogLTS module
try {
    $installedModule = Get-Module -ListAvailable -Name MSCatalogLTS | Sort-Object Version -Descending | Select-Object -First 1
    $latestModule = Find-Module -Name MSCatalogLTS -ErrorAction Stop | Select-Object -First 1

    if (-not $installedModule) {
        Write-Log "Installing MSCatalogLTS module version $($latestModule.Version)..."
        Install-Module -Name MSCatalogLTS -Scope CurrentUser -Force -ErrorAction Stop
    } elseif ([version]$installedModule.Version -lt [version]$latestModule.Version) {
        Write-Log "Updating MSCatalogLTS from $($installedModule.Version) to $($latestModule.Version)..."
        Update-Module -Name MSCatalogLTS -Force -ErrorAction Stop
    } else {
        Write-Log "MSCatalogLTS module is up to date (version $($installedModule.Version))."
    }
    Import-Module -Name MSCatalogLTS -ErrorAction Stop
} catch {
    Write-Log "Failed to install/update MSCatalogLTS: $_" "ERROR"
    throw
}

# Create monthyear subfolder
$monthYear = Get-Date -Format "MMyy"
$fullDownloadPath = Join-Path $DownloadPath $monthYear
if (-not (Test-Path $fullDownloadPath)) {
    Write-Log "Creating subfolder: $fullDownloadPath"
    New-Item -Path $fullDownloadPath -ItemType Directory -Force | Out-Null
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
            Write-Log "Downloading: $($latestUpdate.Title) to $fullDownloadPath"
            Save-MSCatalogUpdate -Update $latestUpdate -Destination $fullDownloadPath -DownloadAll -ErrorAction Stop
            $allDownloads += $latestUpdate
        }
    }

    Write-Progress -Activity "Downloading Updates" -Status "Complete" -PercentComplete 100

    # Return status
    [PSCustomObject]@{
        Status = "Success"
        DownloadedUpdates = $allDownloads
        LogFile = $LogFile
        DownloadPath = $fullDownloadPath
    }
} catch {
    Write-Log "Script error: $_" "ERROR"
    throw
} finally {
    Write-Log "Script completed."
}
#endregion
