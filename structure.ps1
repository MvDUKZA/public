PSPackageBuilder/
│
├─ README.md
├─ .gitignore
│
├─ config/
│   ├─ TeamMap.json
│   └─ VendorMap.json
│
├─ src/
│   └─ PSPackageBuilder/
│       ├─ PSPackageBuilder.psd1
│       ├─ PSPackageBuilder.psm1
│       │
│       ├─ Public/
│       │   ├─ Invoke-PSPackageBuild.ps1        # Script 1 entry
│       │   └─ Invoke-PSPBWrapperBuild.ps1      # Script 2 entry (wrappers)
│       │
│       ├─ Private/
│       │   ├─ Defaults.ps1                     # ONLY place for hardcoded paths
│       │   ├─ Resolve-Setting.ps1
│       │   ├─ Read-Config.ps1
│       │   ├─ Get-UserTeam.ps1
│       │   ├─ Get-CandidateFiles.ps1
│       │   ├─ Select-Mode.ps1
│       │   ├─ Select-Payload.ps1
│       │   ├─ Select-ExistingPackage.ps1
│       │   ├─ Get-PackageMetadata.ps1
│       │   ├─ Get-MsiMetadata.ps1
│       │   ├─ Get-AppxMetadata.ps1
│       │   ├─ Get-ExeMetadata.ps1
│       │   ├─ Normalize-Vendor.ps1
│       │   ├─ Sanitize-Token.ps1
│       │   ├─ Normalize-Version.ps1
│       │   ├─ New-FolderName.ps1
│       │   ├─ Find-NextRevision.ps1
│       │   ├─ Test-Duplicate.ps1
│       │   ├─ Show-EditableNameDialog.ps1
│       │   ├─ Write-PackageManifest.ps1
│       │   ├─ Write-PackageReadme.ps1
│       │   ├─ New-PackageMaterialized.ps1
│       │   │
│       │   ├─ Select-PackageFolder.ps1         # Script 2 helpers
│       │   ├─ Select-PayloadFromPackage.ps1
│       │   ├─ Get-MsiProductCode.ps1
│       │   ├─ Find-MstTransforms.ps1
│       │   ├─ Build-InstallLine.ps1
│       │   ├─ Update-InvokeAppDeployToolkit.ps1
│       │   ├─ Write-InstallWrapper.ps1
│       │   ├─ Write-UninstallWrapper.ps1
│       │   ├─ Write-DetectScript.ps1
│       │   └─ Update-PackageJson.ps1
│
└─ scripts/
    ├─ Script1.ps1   # Thin runner → Invoke-PSPackageBuild
    └─ Script2.ps1   # Thin runner → Invoke-PSPBWrapperBuild
