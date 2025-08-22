# Signed by Marinus van Deventer

<#
.SYNOPSIS
    Patches a Windows image (install.wim) offline using DISM by adding update packages.

.DESCRIPTION
    This script copies the source image folder with a monthyear suffix, mounts the specified Windows image in the copy,
    adds all .msu and .cab update packages from a folder, commits the changes, and unmounts the image. It handles errors, logs progress, and ensures cleanup.

.PARAMETER WimFile
    Path to the install.wim file (assumed structure: <root>\sources\install.wim).

.PARAMETER UpdatesFolder
    Path to the folder containing .msu or .cab update files (e.g., from monthyear download).

.PARAMETER MountDir
    Directory to mount the image (must be empty).

.PARAMETER ImageIndex
    Index of the image edition to patch (default: 3 for Enterprise).

.PARAMETER LogFile
    Path to the log file (default: C:\temp\scripts\logs\PatchWim.log).

.PARAMETER ContinueOnPackageError
    If true, continues adding packages even if one fails (default: $false).

.EXAMPLE
    .\PatchWindowsImage.ps1 -WimFile "C:\Windows-11-24H2\sources\install.wim" -UpdatesFolder "C:\Updates\0825" -MountDir "C:\Mount" -ImageIndex 3

.NOTES
    Requires administrative privileges.
    Based on DISM documentation: https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-image-management-command-line-options-s14
    PowerShell version: Compatible with 5.1+.
    Changelog: v1.1 - Added copy of source root folder with monthyear suffix (e.g., _0825); changed default ImageIndex to 3 for Enterprise; updated return object.
               v1.0 - Initial version.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$WimFile,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$UpdatesFolder,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$MountDir,

    [int]$ImageIndex = 3,

    [string]$LogFile = "C:\temp\scripts\logs\PatchWim.log",

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
Write-Log "Script started. Parameters: WimFile=$WimFile, UpdatesFolder=$UpdatesFolder, MountDir=$MountDir, ImageIndex=$ImageIndex"

# Ensure mount directory is empty
if ((Get-ChildItem $MountDir).Count -gt 0) {
    throw "Mount directory '$MountDir' must be empty."
}

# Derive source root and create patched copy
$sourcesDir = Split-Path $WimFile -Parent
$sourceRoot = Split-Path $sourcesDir -Parent
$monthYear = Get-Date -Format "MMyy"
$patchedRoot = "$sourceRoot_$monthYear"
if (Test-Path $patchedRoot) {
    throw "Patched folder '$patchedRoot' already exists. Delete or rename it before proceeding."
}
Write-Log "Copying source folder '$sourceRoot' to '$patchedRoot'"
Write-Progress -Activity "Patching Windows Image" -Status "Copying source folder" -PercentComplete 5
Copy-Item -Path $sourceRoot -Destination $patchedRoot -Recurse -Force -ErrorAction Stop
$patchedWim = Join-Path (Join-Path $patchedRoot "sources") (Split-Path $WimFile -Leaf)
Write-Log "Patched WIM path: $patchedWim"
#endregion

#region Main Logic
try {
    # Mount the image
    Write-Log "Mounting image..."
    Write-Progress -Activity "Patching Windows Image" -Status "Mounting WIM" -PercentComplete 10
    $mountArgs = "/Mount-Wim /WimFile:`"$patchedWim`" /Index:$ImageIndex /MountDir:`"$MountDir`""
    $mountResult = & dism.exe $mountArgs.Split(' ')
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to mount image. DISM output: $mountResult"
    }
    Write-Log "Image mounted successfully."

    # Get update files
    $updateFiles = Get-ChildItem -Path $UpdatesFolder -Filter *.msu -Recurse
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
                $addArgs = "/Image:`"$MountDir`" /Add-Package /PackagePath:`"$($file.FullName)`""
                $addResult = & dism.exe $addArgs.Split(' ')
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
    $unmountArgs = "/Unmount-Wim /MountDir:`"$MountDir`" /Commit"
    $unmountResult = & dism.exe $unmountArgs.Split(' ')
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
        & dism.exe /Unmount-Wim /MountDir:"$MountDir" /Discard
    }
    throw
} finally {
    Write-Log "Script completed."
}
#endregion
