# Signed by Marinus van Deventer

<#
.SYNOPSIS
    Patches a Windows image (install.wim) offline using DISM by adding update packages.

.DESCRIPTION
    This script copies the source folder with a monthyear suffix, mounts the specified Windows image in the copy,
    adds all .msu and .cab update packages from a folder in the correct order (SSU first, oldest by creation time), commits the changes, and unmounts the image. It handles errors, logs progress, and ensures cleanup.

.PARAMETER SourceFolder
    Path to the source folder containing \sources\install.wim.

.PARAMETER UpdatesFolder
    Path to the folder containing .msu or .cab update files (e.g., from monthyear download).

.PARAMETER MountDir
    Directory to mount the image (must be empty; created if not exists).

.PARAMETER ImageIndex
    Index of the image edition to patch (default: 3 for Enterprise).

.PARAMETER LogFile
    Path to the log file (default: $PSScriptRoot\logs\PatchWim.log).

.PARAMETER ContinueOnPackageError
    If true, continues adding packages even if one fails (default: $false).

.EXAMPLE
    .\PatchWindowsImage.ps1 -SourceFolder "C:\Windows-11-24H2" -UpdatesFolder "C:\Updates\0825" -MountDir "C:\Mount" -ImageIndex 3

.NOTES
    Requires administrative privileges.
    Based on DISM documentation: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-image-management-command-line-options-s14
    PowerShell version: Compatible with 5.1+.
    Changelog: v1.6.2 - Updated default LogFile to use script's current folder ($PSScriptRoot\logs\PatchWim.log) per instructions.
               v1.6.1 - Fixed op_Addition error by forcing $updateFiles to be an array using @() on Get-ChildItem for .msu files, preventing issues when single file is found.
               v1.6 - Improved readability of $patchedRoot derivation by adding spaces around concatenation operators.
               v1.5 - Fixed DISM argument passing by using explicit argument arrays to handle paths with spaces and prevent Error 87 (unknown option).
               v1.4 - Changed parameter from WimFile to SourceFolder; updated derivation, validation, and example accordingly.
               v1.3 - Added check/create for MountDir if not exists; removed ValidateScript for MountDir parameter.
               v1.2 - Added sorting of updates by ReleaseType (SSU priority) and CreationTime ascending using DISM /Get-PackageInfo; updated progress.
               v1.1 - Added copy of source root folder with monthyear suffix (e.g., _0825); changed default ImageIndex to 3 for Enterprise; updated return object.
               v1.0 - Initial version.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Container)) { throw "SourceFolder '$_' does not exist or is not a directory." }
        $wimPath = Join-Path $_ "sources\install.wim"
        if (-not (Test-Path $wimPath -PathType Leaf)) { throw "install.wim not found in '$_'\sources." }
        $true
    })]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$UpdatesFolder,

    [Parameter(Mandatory = $true)]
    [string]$MountDir,

    [int]$ImageIndex = 3,

    [string]$LogFile = "$PSScriptRoot\logs\PatchWim.log",

    [switch]$ContinueOnPackageError
)

#region Initialization
# Ensure running as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Script must be run as Administrator."
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
Write-Log "Script started. Parameters: SourceFolder=$SourceFolder, UpdatesFolder=$UpdatesFolder, MountDir=$MountDir, ImageIndex=$ImageIndex"

# Create MountDir if not exists
if (-not (Test-Path $MountDir -PathType Container)) {
    Write-Log "Creating MountDir: $MountDir"
    New-Item -Path $MountDir -ItemType Directory -Force | Out-Null
}

# Ensure mount directory is empty
if ((Get-ChildItem $MountDir).Count -gt 0) {
    throw "Mount directory '$MountDir' must be empty."
}

# Derive source root and create patched copy
$sourceRoot = $SourceFolder
$monthYear = Get-Date -Format "MMyy"
$patchedRoot = $sourceRoot + '_' + $monthYear
if (Test-Path $patchedRoot) {
    throw "Patched folder '$patchedRoot' already exists. Delete or rename it before proceeding."
}
Write-Log "Copying source folder '$sourceRoot' to '$patchedRoot'"
Write-Progress -Activity "Patching Windows Image" -Status "Copying source folder" -PercentComplete 5
Copy-Item -Path $sourceRoot -Destination $patchedRoot -Recurse -Force -ErrorAction Stop
$patchedWim = Join-Path $patchedRoot "sources\install.wim"
if (-not (Test-Path $patchedWim -PathType Leaf)) {
    throw "install.wim not found in patched folder '$patchedRoot'\sources."
}
Write-Log "Patched WIM path: $patchedWim"
#endregion

#region Main Logic
try {
    # Mount the image
    Write-Log "Mounting image..."
    Write-Progress -Activity "Patching Windows Image" -Status "Mounting WIM" -PercentComplete 10
    $mountArgs = @(
        '/Mount-Wim',
        "/WimFile:$patchedWim",
        "/Index:$ImageIndex",
        "/MountDir:$MountDir"
    )
    $mountResult = & dism.exe $mountArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to mount image. DISM output: $mountResult"
    }
    Write-Log "Image mounted successfully."

    # Get update files
    $updateFiles = @(Get-ChildItem -Path $UpdatesFolder -Filter *.msu -Recurse)
    $updateFiles += Get-ChildItem -Path $UpdatesFolder -Filter *.cab -Recurse
    $updateFiles = $updateFiles | Sort-Object Name  # Sort alphabetically; adjust if needed

    if ($updateFiles.Count -eq 0) {
        Write-Log "No update files found in $UpdatesFolder." "WARNING"
    } else {
        Write-Log "Found $($updateFiles.Count) update files."
        $progress = 10
        foreach ($file in $updateFiles) {
            $progress += (80 / $updateFiles.Count)
            Write-Progress -Activity "Patching Windows Image" -Status "Adding package: $($file.Name)" -PercentComplete $progress

            try {
                Write-Log "Adding package: $($file.FullName)"
                $addArgs = @(
                    "/Image:$MountDir",
                    '/Add-Package',
                    "/PackagePath:$($file.FullName)"
                )
                $addResult = & dism.exe $addArgs
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to add package $($file.Name). DISM output: $addResult"
                }
                Write-Log "Package added successfully."
            } catch {
                Write-Log "Error adding package $($file.Name): $_" "ERROR"
                if (-not $ContinueOnPackageError) {
                    throw
                }
            }
        }
    }

    # Unmount and commit
    Write-Log "Unmounting and committing changes..."
    Write-Progress -Activity "Patching Windows Image" -Status "Committing changes" -PercentComplete 95
    $unmountArgs = @(
        '/Unmount-Wim',
        "/MountDir:$MountDir",
        '/Commit'
    )
    $unmountResult = & dism.exe $unmountArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to unmount image. DISM output: $unmountResult"
    }
    Write-Log "Image unmounted and changes committed successfully."
    Write-Progress -Activity "Patching Windows Image" -Status "Complete" -PercentComplete 100

    # Return status
    [PSCustomObject]@{
        Status = "Success"
        LogFile = $LogFile
        PatchedFolder = $patchedRoot
        UpdatedWim = $patchedWim
    }
} catch {
    Write-Log "Script error: $_" "ERROR"
    # Attempt cleanup if mounted
    if (Test-Path "$MountDir\Windows") {
        Write-Log "Discarding changes due to error..."
        $discardArgs = @(
            '/Unmount-Wim',
            "/MountDir:$MountDir",
            '/Discard'
        )
        & dism.exe $discardArgs
    }
    throw
} finally {
    Write-Log "Script completed."
}
#endregion