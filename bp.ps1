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
    Version: 1.0
    Date: August 01, 2025
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
    Write-Log "Created bin/ with dummy files"

    # builds\
    $BuildsPath = "$RootPath\builds"
    New-Item -Path $BuildsPath -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\artefacts" -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\logs" -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\manifests" -ItemType Directory -Force | Out-Null
    New-Item -Path "$BuildsPath\reports" -ItemType Directory -Force | Out-Null
    Write-Log "Created builds/ subfolders"

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
    Set-Content -Path "$WindowsCommon\autounattend_fragments\dummy_fragment.xml" -Value "<!-- Dummy fragment for Windows -->"

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
    Set-Content -Path "$UbuntuVsphere\variables.pkrvars.hcl" -Value "# Dummy variables for Ubuntu vSphere"
    # configs/linux/ubuntu/azure\
    $UbuntuAzure = "$UbuntuConfigs\azure"
    New-Item -Path $UbuntuAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$UbuntuAzure\kickstart.cfg" -Value "# Dummy kickstart.cfg for Ubuntu Azure"
    Set-Content -Path "$UbuntuAzure\variables.pkrvars.hcl" -Value "# Dummy variables for Ubuntu Azure"
    # configs/linux/rhel\
    $RhelConfigs = "$LinuxConfigs\rhel"
    New-Item -Path $RhelConfigs -ItemType Directory -Force | Out-Null
    # configs/linux/rhel/vsphere\
    $RhelVsphere = "$RhelConfigs\vsphere"
    New-Item -Path $RhelVsphere -ItemType Directory -Force | Out-Null
    Set-Content -Path "$RhelVsphere\kickstart.cfg" -Value "# Dummy kickstart.cfg for RHEL vSphere"
    Set-Content -Path "$RhelVsphere\variables.pkrvars.hcl" -Value "# Dummy variables for RHEL vSphere"
    # configs/linux/rhel/azure\
    $RhelAzure = "$RhelConfigs\azure"
    New-Item -Path $RhelAzure -ItemType Directory -Force | Out-Null
    Set-Content -Path "$RhelAzure\kickstart.cfg" -Value "# Dummy kickstart.cfg for RHEL Azure"
    Set-Content -Path "$RhelAzure\variables.pkrvars.hcl" -Value "# Dummy variables for RHEL Azure"
    # configs/linux/common\
    $LinuxCommon = "$LinuxConfigs\common"
    New-Item -Path $LinuxCommon -ItemType Directory -Force | Out-Null
    New-Item -Path "$LinuxCommon\cloud-init" -ItemType Directory -Force | Out-Null
    Set-Content -Path "$LinuxCommon\cloud-init\dummy_cloudinit.yaml" -Value "# Dummy cloud-init for Linux"

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
    # scripts/windows/server\
    $ScriptsWindowsServer = "$ScriptsWindows\server"
    New-Item -Path $ScriptsWindowsServer -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsWindowsServer\install-rds.ps1" -Value "# Dummy install-rds.ps1 for Windows server"
    # scripts/windows/common\
    $ScriptsWindowsCommon = "$ScriptsWindows\common"
    New-Item -Path $ScriptsWindowsCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsWindowsCommon\optimizations.ps1" -Value "# Dummy optimizations.ps1 for Windows common"
    # scripts/linux\
    $ScriptsLinux = "$ScriptsPath\linux"
    New-Item -Path $ScriptsLinux -ItemType Directory -Force | Out-Null
    # scripts/linux/ubuntu\
    $ScriptsLinuxUbuntu = "$ScriptsLinux\ubuntu"
    New-Item -Path $ScriptsLinuxUbuntu -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLinuxUbuntu\install-packages.sh" -Value "# Dummy install-packages.sh for Ubuntu"
    # scripts/linux/rhel\
    $ScriptsLinuxRhel = "$ScriptsLinux\rhel"
    New-Item -Path $ScriptsLinuxRhel -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLinuxRhel\install-packages.sh" -Value "# Dummy install-packages.sh for RHEL"
    # scripts/linux/common\
    $ScriptsLinuxCommon = "$ScriptsLinux\common"
    New-Item -Path $ScriptsLinuxCommon -ItemType Directory -Force | Out-Null
    Set-Content -Path "$ScriptsLinuxCommon\optimizations.sh" -Value "# Dummy optimizations.sh for Linux common"
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

    Write-Log "Folder structure and dummy files created successfully"
} catch {
    Write-Log "Error during structure creation: $($_.Exception.Message)" "ERROR"
    throw "Script execution failed."
} finally {
    Write-Log "Script execution completed"
}
#endregion
