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
    Version: 1.2
    Date: August 01, 2025
    Changes: Added floppy/ and isos/ with common subfolders for each OS family, full structure creation.
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
    # builds/artefacts\
    $ArtefactsPath = "$BuildsPath\artefacts"
    New-Item -Path $ArtefactsPath -ItemType Directory -Force | Out-Null
    # builds/artefacts/windows\
    $ArtefactsWindows = "$ArtefactsPath\windows"
    New-Item -Path $ArtefactsWindows -ItemType Directory -Force | Out-Null
    # builds/artefacts/windows/win11\
    $ArtefactsWin11 = "$ArtefactsWindows\win11"
    New-Item -Path $ArtefactsWin11 -ItemType Directory -Force | Out-Null
    # builds/artefacts/windows/winsvr2022\
    $ArtefactsWinSvr2022 = "$ArtefactsWindows\winsvr2022"
    New-Item -Path $ArtefactsWinSvr2022 -ItemType Directory -Force | Out-Null
    # builds/artefacts/linux\
    $ArtefactsLinux = "$ArtefactsPath\linux"
    New-Item -Path $ArtefactsLinux -ItemType Directory -Force | Out-Null
    # builds/artefacts/linux/ubuntu\
    $ArtefactsUbuntu = "$ArtefactsLinux\ubuntu"
    New-Item -Path $ArtefactsUbuntu -ItemType Directory -Force | Out-Null
    # builds/artefacts/linux/rhel\
    $ArtefactsRhel = "$ArtefactsLinux\rhel"
    New-Item -Path $ArtefactsRhel -ItemType Directory -Force | Out-Null
    # builds/artefacts/macos\
    $ArtefactsMacos = "$ArtefactsPath\macos"
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
    # configs/windows/win11\
    $Win11Configs = "$WindowsConfigs\win11"
    New-Item -Path $Win11Configs -ItemType Directory -Force | Out-Null
    # configs/windows/win11/vsphere\
    $Win11Vsphere = "$Win11Configs\vsphere"
    New-Item -Path $Win11Vsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$Win11Vsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Win11 vSphere -->"
    Set-Content -Path "$Win11Vsphere\variables.pkrvars.hcl" -Value "# Dummy variables for Win11 vSphere"
    # configs/windows/win11/azure\
    $Win11Azure = "$Win11Configs\azure"
    New-Item -Path $Win11Azure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$Win11Azure\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Win11 Azure -->"
    Set-Content -Path "$Win11Azure\variables.pkrvars.hcl" -Value "# Dummy variables for Win11 Azure"
    # configs/windows/winsvr2022\
    $WinSvr2022Configs = "$WindowsConfigs\winsvr2022"
    New-Item -Path $WinSvr2022Configs -ItemType Directory -Force | Out-Null
    # configs/windows/winsvr2022/vsphere\
    $WinSvr2022Vsphere = "$WinSvr2022Configs\vsphere"
    New-Item -Path $WinSvr2022Vsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WinSvr2022Vsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for WinSvr2022 vSphere -->"
    Set-Content -Path "$WinSvr2022Vsphere\variables.pkrvars.hcl" -Value "# Dummy variables for WinSvr2022 vSphere"
    # configs/windows/winsvr2022/azure\
    $WinSvr2022Azure = "$WinSvr2022Configs\azure"
    New-Item -Path $WinSvr2022Azure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$WinSvr2022Azure\autounattend.xml" -Value "<!-- Dummy autounattend.xml for WinSvr2022 Azure -->"
    Set-Content -Path "$WinSvr2022Azure\variables.pkrvars.hcl" -Value "# Dummy variables for WinSvr2022 Azure"
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

    # templates\
    $TemplatesPath = "$RootPath\templates"
    New-Item -Path $TemplatesPath -ItemType Directory -Force | Out-Null
    # templates/windows\
    $TemplatesWindows = "$TemplatesPath\windows"
    New-Item -Path $TemplatesWindows -ItemType Directory -Force | Out-Null
    Set-Content -Path "$TemplatesWindows\vsphere-base.pkr.hcl" -Value "# Dummy vSphere base template for Windows"
    Set-Content -Path "$TemplatesWindows\azure-base.pkr.hcl" -Value "# Dummy Azure base template for Windows"
    # templates/linux\
    $TemplatesLinux = "$TemplatesPath\linux"
    New-Item -Path $TemplatesLinux -ItemType Directory -Force | Out-Null
    Set-Content -Path "$TemplatesLinux\vsphere-base.pkr.hcl" -Value "# Dummy vSphere base template for Linux"
    Set-Content -Path "$TemplatesLinux\azure-base.pkr.hcl" -Value "# Dummy Azure base template for Linux"
    # templates/components\
    $TemplatesComponents = "$TemplatesPath\components"
    New-Item -Path $TemplatesComponents -ItemType Directory -Force | Out-Null
    Set-Content -Path "$TemplatesComponents\winrm-setup.pkr.hcl" -Value "# Dummy WinRM setup component"
    Set-Content -Path "$TemplatesComponents\ssh-setup.pkr.hcl" -Value "# Dummy SSH setup component"
    Set-Content -Path "$TemplatesComponents\cloud-init.pkr.hcl" -Value "# Dummy cloud-init component"
    Write-Log "Created templates/ with dummy files"

    # scripts\
    $ScriptsPath = "$RootPath\scripts"
    New-Item -Path $ScriptsPath -ItemType Directory -Force | Out-Null
    # scripts/windows\
    $ScriptsWindows = "$ScriptsPath\windows"
    New-Item -Path $ScriptsWindows -ItemType Directory -Force | Out-Null
    # scripts/windows/client\
    $ScriptsWindowsClient = "$ScriptsWindows\client"
    New-Item -Path $ScriptsWindowsClient -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsWindowsClient\install-horizon.ps1" -Value "# Dummy install-horizon.ps1 for Windows client"
    Set-Content -Path "$ScriptsWindowsClient\vdi-optimize.ps1" -Value "# Dummy vdi-optimize.ps1 for Windows client"
    # scripts/windows/server\
    $ScriptsWindowsServer = "$ScriptsWindows\server"
    New-Item -Path $ScriptsWindowsServer -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsWindowsServer\install-roles.ps1" -Value "# Dummy install-roles.ps1 for Windows server"
    Set-Content -Path "$ScriptsWindowsServer\configure-domain.ps1" -Value "# Dummy configure-domain.ps1 for Windows server"
    # scripts/windows/common\
    $ScriptsWindowsCommon = "$ScriptsWindows\common"
    New-Item -Path $ScriptsWindowsCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsWindowsCommon\security.ps1" -Value "# Dummy security.ps1 for Windows common"
    Set-Content -Path "$ScriptsWindowsCommon\updates.ps1" -Value "# Dummy updates.ps1 for Windows common"
    # scripts/linux\
    $ScriptsLinux = "$ScriptsPath\linux"
    New-Item -Path $ScriptsLinux -ItemType Directory -Force | Out-Null
    # scripts/linux/ubuntu\
    $ScriptsLinuxUbuntu = "$ScriptsLinux\ubuntu"
    New-Item -Path $ScriptsLinuxUbuntu -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLinuxUbuntu\install-docker.sh" -Value "# Dummy install-docker.sh for Ubuntu"
    Set-Content -Path "$ScriptsLinuxUbuntu\configure-networking.sh" -Value "# Dummy configure-networking.sh for Ubuntu"
    # scripts/linux/rhel\
    $ScriptsLinuxRhel = "$ScriptsLinux\rhel"
    New-Item -Path $ScriptsLinuxRhel -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLinuxRhel\install-packages.sh" -Value "# Dummy install-packages.sh for RHEL"
    Set-Content -Path "$ScriptsLinuxRhel\security-harden.sh" -Value "# Dummy security-harden.sh for RHEL"
    # scripts/linux/common\
    $ScriptsLinuxCommon = "$ScriptsLinux\common"
    New-Item -Path $ScriptsLinuxCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLinuxCommon\cloud-init.sh" -Value "# Dummy cloud-init.sh for Linux common"
    Set-Content -Path "$ScriptsLinuxCommon\kernel-tune.sh" -Value "# Dummy kernel-tune.sh for Linux common"
    # scripts/lib\
    $ScriptsLib = "$ScriptsPath\lib"
    New-Item -Path $ScriptsLib -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLib\package-manager.psm1" -Value "# Dummy package-manager.psm1 module"
    Set-Content -Path "$ScriptsLib\os-detection.psm1" -Value "# Dummy os-detection.psm1 module"
    Write-Log "Created scripts/ structure with dummy files"

    # variables\
    $VariablesPath = "$RootPath\variables"
    New-Item -Path $VariablesPath -ItemType Directory -Force | Out-Null
    Set-Content -Path "$VariablesPath\global.pkrvars.hcl" -Value "# Dummy global variables"
    # variables/platforms\
    $VariablesPlatforms = "$VariablesPath\platforms"
    New-Item -Path $VariablesPlatforms -ItemType Directory -Force | Out-Null
    Set-Content -Path "$VariablesPlatforms\vsphere.pkrvars.hcl" -Value "# Dummy vSphere platform variables"
    Set-Content -Path "$VariablesPlatforms\azure.pkrvars.hcl" -Value "# Dummy Azure platform variables"
    # variables/os\
    $VariablesOS = "$VariablesPath\os"
    New-Item -Path $VariablesOS -ItemType Directory -Force | Out-Null
    Set-Content -Path "$VariablesOS\win11.pkrvars.hcl" -Value "# Dummy Win11 OS variables"
    Set-Content -Path "$VariablesOS\winsvr2022.pkrvars.hcl" -Value "# Dummy WinSvr2022 OS variables"
    Set-Content -Path "$VariablesOS\ubuntu.pkrvars.hcl" -Value "# Dummy Ubuntu OS variables"
    Set-Content -Path "$VariablesOS\rhel.pkrvars.hcl" -Value "# Dummy RHEL OS variables"
    Write-Log "Created variables/ structure with dummy files"

    # .gitignore, .env.example, README.md, Build-Template.ps1, Build-Template.sh
    Set-Content -Path "$RootPath\.gitignore" -Value "# Dummy .gitignore content"
    Set-Content -Path "$RootPath\.env.example" -Value "# Dummy .env.example content"
    Set-Content -Path "$RootPath\README.md" -Value "# Dummy README.md content"
    Set-Content -Path "$RootPath\Build-Template.ps1" -Value "# Dummy Build-Template.ps1"
    Set-Content -Path "$RootPath\Build-Template.sh" -Value "# Dummy Build-Template.sh"

    # tests\
    $TestsPath = "$RootPath\tests"
    New-Item -Path $TestsPath -ItemType Directory -Force | Out-Null
    # tests/windows\
    $TestsWindows = "$TestsPath\windows"
    New-Item -Path $TestsWindows -ItemType Directory -Force | Out-Null
    Set-Content -Path "$TestsWindows\script-tests.ps1" -Value "# Dummy Pester tests for Windows"
    # tests/linux\
    $TestsLinux = "$TestsPath\linux"
    New-Item -Path $TestsLinux -ItemType Directory -Force | Out-Null
    Set-Content -Path "$TestsLinux\script-tests.bats" -Value "# Dummy BATS tests for Linux"
    Write-Log "Created root files and tests/ structure with dummy files"

    # isos\
    $IsosPath = "$RootPath\isos"
    New-Item -Path $IsosPath -ItemType Directory -Force | Out-Null
    # isos/windows\
    $IsosWindows = "$IsosPath\windows"
    New-Item -Path $IsosWindows -ItemType Directory -Force | Out-Null
    # isos/windows/win11\
    $IsosWin11 = "$IsosWindows\win11"
    New-Item -Path $IsosWin11 -ItemType Directory -Force | Out-Null
    # isos/windows/win11/vsphere\
    $IsosWin11Vsphere = "$IsosWin11\vsphere"
    New-Item -Path $IsosWin11Vsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosWin11Vsphere\README.txt" -Value "# Dummy README for Win11 vSphere ISO reference"
    # isos/windows/win11/azure\
    $IsosWin11Azure = "$IsosWin11\azure"
    New-Item -Path $IsosWin11Azure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosWin11Azure\README.txt" -Value "# Dummy README for Win11 Azure ISO reference"
    # isos/windows/winsvr2022\
    $IsosWinSvr2022 = "$IsosWindows\winsvr2022"
    New-Item -Path $IsosWinSvr2022 -ItemType Directory -Force | Out-Null
    # isos/windows/winsvr2022/vsphere\
    $IsosWinSvr2022Vsphere = "$IsosWinSvr2022\vsphere"
    New-Item -Path $IsosWinSvr2022Vsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosWinSvr2022Vsphere\README.txt" -Value "# Dummy README for WinSvr2022 vSphere ISO reference"
    # isos/windows/winsvr2022/azure\
    $IsosWinSvr2022Azure = "$IsosWinSvr2022\azure"
    New-Item -Path $IsosWinSvr2022Azure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosWinSvr2022Azure\README.txt" -Value "# Dummy README for WinSvr2022 Azure ISO reference"
    # isos/windows/common\
    $IsosWindowsCommon = "$IsosWindows\common"
    New-Item -Path $IsosWindowsCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosWindowsCommon\README.txt" -Value "# Dummy README for Windows common ISO reference"
    # isos/linux\
    $IsosLinux = "$IsosPath\linux"
    New-Item -Path $IsosLinux -ItemType Directory -Force | Out-Null
    # isos/linux/ubuntu\
    $IsosUbuntu = "$IsosLinux\ubuntu"
    New-Item -Path $IsosUbuntu -ItemType Directory -Force | Out-Null
    # isos/linux/ubuntu/vsphere\
    $IsosUbuntuVsphere = "$IsosUbuntu\vsphere"
    New-Item -Path $IsosUbuntuVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosUbuntuVsphere\README.txt" -Value "# Dummy README for Ubuntu vSphere ISO reference"
    # isos/linux/ubuntu/azure\
    $IsosUbuntuAzure = "$IsosUbuntu\azure"
    New-Item -Path $IsosUbuntuAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosUbuntuAzure\README.txt" -Value "# Dummy README for Ubuntu Azure ISO reference"
    # isos/linux/rhel\
    $IsosRhel = "$IsosLinux\rhel"
    New-Item -Path $IsosRhel -ItemType Directory -Force | Out-Null
    # isos/linux/rhel/vsphere\
    $IsosRhelVsphere = "$IsosRhel\vsphere"
    New-Item -Path $IsosRhelVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosRhelVsphere\README.txt" -Value "# Dummy README for RHEL vSphere ISO reference"
    # isos/linux/rhel/azure\
    $IsosRhelAzure = "$IsosRhel\azure"
    New-Item -Path $IsosRhelAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosRhelAzure\README.txt" -Value "# Dummy README for RHEL Azure ISO reference"
    # isos/linux/common\
    $IsosLinuxCommon = "$IsosLinux\common"
    New-Item -Path $IsosLinuxCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosLinuxCommon\README.txt" -Value "# Dummy README for Linux common ISO reference"
    # isos/macos\
    $IsosMacos = "$IsosPath\macos"
    New-Item -Path $IsosMacos -ItemType Directory -Force | Out-Null
    # isos/macos/common\
    $IsosMacosCommon = "$IsosMacos\common"
    New-Item -Path $IsosMacosCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosMacosCommon\README.txt" -Value "# Dummy README for MacOS common ISO reference"
    # isos/macos/vsphere\
    $IsosMacosVsphere = "$IsosMacos\vsphere"
    New-Item -Path $IsosMacosVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$IsosMacosVsphere\README.txt" -Value "# Dummy README for MacOS vSphere ISO reference"
    Write-Log "Created isos/ structure with dummy README files"

    # floppy\
    $FloppyPath = "$RootPath\floppy"
    New-Item -Path $FloppyPath -ItemType Directory -Force | Out-Null
    # floppy/windows\
    $FloppyWindows = "$FloppyPath\windows"
    New-Item -Path $FloppyWindows -ItemType Directory -Force | Out-Null
    # floppy/windows/win11\
    $FloppyWin11 = "$FloppyWindows\win11"
    New-Item -Path $FloppyWin11 -ItemType Directory -Force | Out-Null
    # floppy/windows/win11/vsphere\
    $FloppyWin11Vsphere = "$FloppyWin11\vsphere"
    New-Item -Path $FloppyWin11Vsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyWin11Vsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Win11 vSphere floppy -->"
    # floppy/windows/win11/azure\
    $FloppyWin11Azure = "$FloppyWin11\azure"
    New-Item -Path $FloppyWin11Azure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyWin11Azure\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Win11 Azure floppy -->"
    # floppy/windows/winsvr2022\
    $FloppyWinSvr2022 = "$FloppyWindows\winsvr2022"
    New-Item -Path $FloppyWinSvr2022 -ItemType Directory -Force | Out-Null
    # floppy/windows/winsvr2022/vsphere\
    $FloppyWinSvr2022Vsphere = "$FloppyWinSvr2022\vsphere"
    New-Item -Path $FloppyWinSvr2022Vsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyWinSvr2022Vsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for WinSvr2022 vSphere floppy -->"
    # floppy/windows/winsvr2022/azure\
    $FloppyWinSvr2022Azure = "$FloppyWinSvr2022\azure"
    New-Item -Path $FloppyWinSvr2022Azure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyWinSvr2022Azure\autounattend.xml" -Value "<!-- Dummy autounattend.xml for WinSvr2022 Azure floppy -->"
    # floppy/windows/common\
    $FloppyWindowsCommon = "$FloppyWindows\common"
    New-Item -Path $FloppyWindowsCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyWindowsCommon\autounattend.xml" -Value "<!-- Dummy autounattend.xml for Windows common floppy -->"
    # floppy/linux\
    $FloppyLinux = "$FloppyPath\linux"
    New-Item -Path $FloppyLinux -ItemType Directory -Force | Out-Null
    # floppy/linux/ubuntu\
    $FloppyUbuntu = "$FloppyLinux\ubuntu"
    New-Item -Path $FloppyUbuntu -ItemType Directory -Force | Out-Null
    # floppy/linux/ubuntu/vsphere\
    $FloppyUbuntuVsphere = "$FloppyUbuntu\vsphere"
    New-Item -Path $FloppyUbuntuVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyUbuntuVsphere\ks.cfg" -Value "# Dummy ks.cfg for Ubuntu vSphere floppy"
    # floppy/linux/ubuntu/azure\
    $FloppyUbuntuAzure = "$FloppyUbuntu\azure"
    New-Item -Path $FloppyUbuntuAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyUbuntuAzure\ks.cfg" -Value "# Dummy ks.cfg for Ubuntu Azure floppy"
    # floppy/linux/rhel\
    $FloppyRhel = "$FloppyLinux\rhel"
    New-Item -Path $FloppyRhel -ItemType Directory -Force | Out-Null
    # floppy/linux/rhel/vsphere\
    $FloppyRhelVsphere = "$FloppyRhel\vsphere"
    New-Item -Path $FloppyRhelVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyRhelVsphere\ks.cfg" -Value "# Dummy ks.cfg for RHEL vSphere floppy"
    # floppy/linux/rhel/azure\
    $FloppyRhelAzure = "$FloppyRhel\azure"
    New-Item -Path $FloppyRhelAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyRhelAzure\ks.cfg" -Value "# Dummy ks.cfg for RHEL Azure floppy"
    # floppy/linux/common\
    $FloppyLinuxCommon = "$FloppyLinux\common"
    New-Item -Path $FloppyLinuxCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyLinuxCommon\ks.cfg" -Value "# Dummy ks.cfg for Linux common floppy"
    # floppy/macos\
    $FloppyMacos = "$FloppyPath\macos"
    New-Item -Path $FloppyMacos -ItemType Directory -Force | Out-Null
    # floppy/macos/common\
    $FloppyMacosCommon = "$FloppyMacos\common"
    New-Item -Path $FloppyMacosCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyMacosCommon\autounattend.xml" -Value "<!-- Dummy autounattend.xml for MacOS common floppy -->"
    # floppy/macos/vsphere\
    $FloppyMacosVsphere = "$FloppyMacos\vsphere"
    New-Item -Path $FloppyMacosVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$FloppyMacosVsphere\autounattend.xml" -Value "<!-- Dummy autounattend.xml for MacOS vSphere floppy -->"
    Write-Log "Created floppy/ structure with dummy files"

    Write-Log "Folder structure and dummy files created successfully"
} catch {
    Write-Log "Error during structure creation: $($_.Exception.Message)" "ERROR"
    throw "Script execution failed."
} finally {
    Write-Log "Script execution completed"
}
#endregion
