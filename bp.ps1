<#
.SYNOPSIS
    Creates the Packer project folder structure and dummy files based on the specified root path.

.DESCRIPTION
    This script builds the defined folder structure for a Packer project, including OS-specific configs for Windows and Linux, and creates dummy placeholder files. It uses Git for versioning but does not initialise the repo. Folders are created using New-Item, and dummy files are populated with basic content using Set-Content. Logs are written to C:\temp\scripts\logs.

.PARAMETER RootPath
    The root path where the structure will be created (e.g., C:\packer). Must be a valid directory path.

.EXAMPLE
    .\Build-PackerStructure.ps1 -RootPath "C:\packer"

.NOTES
    Author: Marinus van Deventer
    Version: 1.1
    Date: August 01, 2025
    Changes: Added floppy/ and isos/ folders with common subfolders for each OS family.
    Reference: New-Item[](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/new-item?view=powershell-7.4)
    Reference: Set-Content[](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/set-content?view=powershell-7.4)
    Reference: Test-Path[](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-path?view=powershell-7.4)
    Working Directory: C:\temp\scripts
    Logs: C:\temp\scripts\logs\Build-PackerStructure.log
    The script validates inputs and handles errors with try/catch.
    To sign the script: Set-AuthenticodeSignature -FilePath .\Build-PackerStructure.ps1 -Certificate (Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert)
#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -IsValid })]
    [string]$RootPath
)

#region Initial Setup
$WorkingDir = "C:\temp\scripts"
$LogDir = "$WorkingDir\logs"
$LogFile = "$LogDir\Build-PackerStructure.log"

# Create working and log directories if not exist
if (-not (Test-Path $WorkingDir -PathType Container)) {
    New-Item -Path $WorkingDir -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path $LogDir -PathType Container)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

# Function for logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Information $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry
}

Write-Log "Script execution started with RootPath: $RootPath"

# Validate RootPath
if (-not (Test-Path $RootPath -PathType Container)) {
    try {
        New-Item -Path $RootPath -ItemType Directory -Force | Out-Null
        Write-Log "Created RootPath: $RootPath"
    } catch {
        Write-Log "Error creating RootPath: $($_.Exception.Message)" "ERROR"
        throw "Failed to create RootPath."
    }
}
#endregion

#region Create Folder Structure and Dummy Files
try {
    # bin\
    $BinPath = "$RootPath\bin"
    New-Item -Path $BinPath -ItemType Directory -Force | Out-Null
    Set-Content -Path "$BinPath\packer.exe" -Value "# Dummy packer.exe placeholder"
    Set-Content -Path "$BinPath\sdelete.exe" -Value "# Dummy sdelete.exe placeholder"
    Set-Content -Path "$BinPath\shred" -Value "# Dummy shred placeholder"
    Write-Log "Created bin/ with dummy files"

    # builds\
    $BuildsPath = "$RootPath\builds"
    New-Item -Path $BuildsPath -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\artefacts" -ItemType Directory -Force | Out-Null
    # builds/artefacts/windows\
    $ArtefactsWindows = "$BuildsPath\artefacts\windows"
    New-Item -Path $ArtefactsWindows -ItemType Directory -Force | Out-Null
    # builds/artefacts/linux\
    $ArtefactsLinux = "$BuildsPath\artefacts\linux"
    New-Item -Path $ArtefactsLinux -ItemType Directory -Force | Out-Null
    # builds/artefacts/macos\
    $ArtefactsMacos = "$BuildsPath\artefacts\macos"
    New-Item -Path $ArtefactsMacos -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\logs" -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\manifests" -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\reports" -ItemType Directory -Force | Out-Null
    Write-Log "Created builds/ subfolders with OS artefacts"

    # configs\
    $ConfigsPath = "$RootPath\configs"
    New-Item -Path $ConfigsPath -ItemType Directory -Force | Out-Null
    # configs/windows\
    $WindowsConfigs = "$ConfigsPath\windows"
    New-Item -Path $WindowsConfigs -ItemType Directory -Force | Out-Null
    # configs/windows/client\
    $WindowsClient = "$WindowsConfigs\client"
    New-Item -Path $WindowsClient -ItemType Directory -Force | Out-Null
    # configs/windows/client/vsphere\
    $WindowsClientVsphere = "$WindowsClient\vsphere"
    New-Item -Path $WindowsClientVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WindowsClientVsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Windows client vSphere -->"
    # configs/windows/client/azure\
    $WindowsClientAzure = "$WindowsClient\azure"
    New-Item -Path $WindowsClientAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WindowsClientAzure\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Windows client Azure -->"
    # configs/windows/server\
    $WindowsServer = "$WindowsConfigs\server"
    New-Item -Path $WindowsServer -ItemType Directory -Force | Out-Null
    # configs/windows/server/vsphere\
    $WindowsServerVsphere = "$WindowsServer\vsphere"
    New-Item -Path $WindowsServerVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WindowsServerVsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Windows server vSphere -->"
    # configs/windows/server/azure\
    $WindowsServerAzure = "$WindowsServer\azure"
    New-Item -Path $WindowsServerAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WindowsServerAzure\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Windows server Azure -->"
    # configs/windows/common\
    $WindowsCommon = "$WindowsConfigs\common"
    New-Item -Path $WindowsCommon -ItemType Directory -Force | Out-Null
    New-Item -Path "$WindowsCommon\autounattend_fragments" -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WindowsCommon\autounattend_fragments\dummy_fragment.xml" -Value "<!-- Dummy fragment for Windows common -->"
    # configs/linux\
    $LinuxConfigs = "$ConfigsPath\linux"
    New-Item -Path $LinuxConfigs -ItemType Directory -Force | Out-Null
    # configs/linux/ubuntu\
    $UbuntuConfigs = "$LinuxConfigs\ubuntu"
    New-Item -Path $UbuntuConfigs -ItemType Directory -Force | Out-Null
    # configs/linux/ubuntu/vsphere\
    $UbuntuVsphere = "$UbuntuConfigs\vsphere"
    New-Item -Path $UbuntuVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$UbuntuVsphere\kickstart.cfg" -Value "# Dummy kickstart.cfg for Ubuntu vSphere"
    # configs/linux/ubuntu/azure\
    $UbuntuAzure = "$UbuntuConfigs\azure"
    New-Item -Path $UbuntuAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$UbuntuAzure\kickstart.cfg" -Value "# Dummy kickstart.cfg for Ubuntu Azure"
    # configs/linux/rhel\
    $RhelConfigs = "$LinuxConfigs\rhel"
    New-Item -Path $RhelConfigs -ItemType Directory -Force | Out-Null
    # configs/linux/rhel/vsphere\
    $RhelVsphere = "$RhelConfigs\vsphere"
    New-Item -Path $RhelVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$RhelVsphere\kickstart.cfg" -Value "# Dummy kickstart.cfg for RHEL vSphere"
    # configs/linux/rhel/azure\
    $RhelAzure = "$RhelConfigs\azure"
    New-Item -Path $RhelAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$RhelAzure\kickstart.cfg" -Value "# Dummy kickstart.cfg for RHEL Azure"
    # configs/linux/common\
    $LinuxCommon = "$LinuxConfigs\common"
    New-Item -Path $LinuxCommon -ItemType Directory -Force | Out-Null
    New-Item -Path "$LinuxCommon\cloud-init" -ItemType Directory -Force | Out-Null
    Set-Content -Path "$LinuxCommon\cloud-init\dummy_cloudinit.yaml" -Value "# Dummy cloud-init for Linux common"
    # configs/common\
    $CommonConfigs = "$ConfigsPath\common"
    New-Item -Path $CommonConfigs -ItemType Directory -Force | Out-Null
    Write-Log "Created configs/ structure with dummy files"

    # templates
