# MECM Windows In-Place Upgrade Remediation — Task Sequence Build Guide

**Purpose:** Repair corrupt Windows component stores by re-laying Windows 24H2 from a
pre-patched WIM on a network share. Preserves all apps, data, and settings.  
**Audience:** Senior MECM/SCCM engineer — assumes console access and CM site drive access.  
**Last updated:** June 2026

---

## Contents

1. [Prerequisites](#1-prerequisites)
2. [Monthly maintenance — the only thing that changes](#2-monthly-maintenance)
3. [Task Sequence structure overview](#3-task-sequence-structure-overview)
4. [Step-by-step console build instructions](#4-step-by-step-console-build-instructions)
5. [Export / Import PowerShell commands](#5-export--import-powershell-commands)
6. [Monthly variable update script](#6-monthly-variable-update-script)
7. [Troubleshooting reference](#7-troubleshooting-reference)

---

## 1. Prerequisites

| Item | Detail |
|---|---|
| Share path | `\\computer\share` |
| Source folder naming | `Microsoft_Windows_Enterprise_24H2_MMYY` |
| WIM image index | `3` (Windows Enterprise) |
| MECM version | Current Branch (2207 or later recommended) |
| PowerShell for scripts | **5.1 only** — CM module does not support PS7 |
| Free space requirement | 20 GB on system drive |

The share must be readable by the MECM Computer Account (or Network Access Account if configured).

---

## 2. Monthly maintenance

> **This is the only thing you change each month.**

Update the `IPU_Month` Task Sequence Variable in the TS.  
Format: `MMYY` — e.g. `0626` for June 2026, `0726` for July 2026.

You can do this manually in the console, or run the script in [Section 6](#6-monthly-variable-update-script).

---

## 3. Task Sequence structure overview

```
[TS] Windows IPU Remediation - 24H2
│
├── GROUP: Set Variables
│     ├── Set TS Variable — IPU_Month
│     ├── Set TS Variable — IPU_SourcePath
│     └── Set TS Variable — OSDSetupAdditionalUpgradeOptions
│
├── GROUP: Pre-Flight Checks
│     ├── Check Readiness          (native step — free space)
│     └── Run Command Line         (verify share path accessible)
│
├── GROUP: Upgrade Operating System
│     └── Upgrade Operating System (native step — handles all reboots)
│
├── GROUP: Post-Processing
│     ├── Run Command Line         (remove Windows.old)
│     ├── Run Command Line         (remove $Windows.~BT)
│     ├── Run Command Line         (remove $Windows.~WS)
│     └── Run Command Line         (DISM component store cleanup)
│
├── GROUP: Rollback
│     └── Run Command Line         (log rollback event)
│
└── GROUP: Run Actions on Failure
      └── Run Command Line         (collect Panther logs)
```

---

## 4. Step-by-step console build instructions

### Create the Task Sequence

1. Open the MECM console
2. Navigate to: `Software Library > Operating Systems > Task Sequences`
3. Right-click > **Create Task Sequence**
4. Select: **Upgrade an operating system from an upgrade package**
5. Name: `Windows IPU Remediation - 24H2`
6. Upgrade package: browse to `\\computer\share\Microsoft_Windows_Enterprise_24H2_0626`
   *(you will replace the hardcoded path with a variable in the step later)*
7. Edition index: `3`
8. Complete the wizard — click **Next** through remaining pages, then **Close**

> The wizard creates the basic skeleton. You will now edit it to add all groups and steps.

---

### Edit the Task Sequence

Right-click the new TS > **Edit**

---

### GROUP: Set Variables

The wizard will have created some default steps. **Delete all default steps first**, then
build the groups below from scratch using the **Add** menu in the TS editor.

**Add group:** Click **Add > New Group**  
Name: `Set Variables`  
Move this group to the **top** of the TS.

---

#### Step 1 — Set TS Variable: IPU_Month

- Click the `Set Variables` group to select it
- Click **Add > General > Set Task Sequence Variable**
- **Name:** `Set IPU_Month`
- **Task Sequence Variable:** `IPU_Month`
- **Value:** `0626`

> Change this value each month. Everything else references this variable.

---

#### Step 2 — Set TS Variable: IPU_SourcePath

- Click **Add > General > Set Task Sequence Variable**
- **Name:** `Set IPU_SourcePath`
- **Task Sequence Variable:** `IPU_SourcePath`
- **Value:**
```
\\computer\share\Microsoft_Windows_Enterprise_24H2_%IPU_Month%
```

---

#### Step 3 — Set TS Variable: OSDSetupAdditionalUpgradeOptions

- Click **Add > General > Set Task Sequence Variable**
- **Name:** `Set Setup Additional Options`
- **Task Sequence Variable:** `OSDSetupAdditionalUpgradeOptions`
- **Value:**
```
/resizerecoverypartition disable /showoobe none /dynamicupdate disable /telemetry disable /diagnosticprompt enable
```

> This variable is natively read by the **Upgrade Operating System** step and appended
> directly to the setup.exe command line. No need to touch the Upgrade step itself.

---

### GROUP: Pre-Flight Checks

**Add group:** Click **Add > New Group**  
Name: `Pre-Flight Checks`

---

#### Step 4 — Check Readiness (native step)

- Click **Add > General > Check Readiness**
- **Name:** `Check Free Disk Space`
- Tick: **Ensure minimum free disk space (MB)**
- Value: `20480`
- Leave all other checks **unticked**
- **Options tab:** Continue on error: **No** (fail TS if not enough space)

> This is a 100% native step — no script required.

---

#### Step 5 — Verify source path accessible

- Click **Add > General > Run Command Line**
- **Name:** `Verify Source Path Accessible`
- **Command line:**
```cmd
cmd /c if not exist "%IPU_SourcePath%\setup.exe" exit 1
```
- **Options tab:** Continue on error: **No**

---

### GROUP: Upgrade Operating System

**Add group:** Click **Add > New Group**  
Name: `Upgrade Operating System`

---

#### Step 6 — Upgrade Operating System (native step — handles all reboots)

- Click **Add > Images > Upgrade Operating System**
- **Name:** `Upgrade Operating System`
- **Upgrade package:** Browse and select the package you imported in the wizard
  *(or if using a direct path, set the path to `%IPU_SourcePath%`)*
- **Edition index:** `3`
- **Product key:** leave blank (preserves existing activation)
- **Provide the following driver content to Windows Setup during upgrade:** leave blank
- **Time-out (minutes):** `180`
- **Options tab:** Continue on error: **No**

> **Reboot handling:** This native step uses SetupComplete.cmd and TSMBootstrap.exe to
> survive all setup reboots and resume the TS automatically. Do NOT add a Restart
> Computer step between this group and Post-Processing.

---

### GROUP: Post-Processing

**Add group:** Click **Add > New Group**  
Name: `Post-Processing`

> These steps run after the TS resumes post-upgrade reboot.

---

#### Step 7 — Remove Windows.old

- Click **Add > General > Run Command Line**
- **Name:** `Remove Windows.old`
- **Command line:**
```cmd
cmd /c rd /s /q "%SystemDrive%\Windows.old"
```
- **Options tab:** Continue on error: **Yes**

---

#### Step 8 — Remove $Windows.~BT

- Click **Add > General > Run Command Line**
- **Name:** `Remove $Windows.~BT`
- **Command line:**
```cmd
cmd /c rd /s /q "%SystemDrive%\$Windows.~BT"
```
- **Options tab:** Continue on error: **Yes**

---

#### Step 9 — Remove $Windows.~WS

- Click **Add > General > Run Command Line**
- **Name:** `Remove $Windows.~WS`
- **Command line:**
```cmd
cmd /c rd /s /q "%SystemDrive%\$Windows.~WS"
```
- **Options tab:** Continue on error: **Yes**

---

#### Step 10 — DISM Component Store Cleanup

- Click **Add > General > Run Command Line**
- **Name:** `DISM Component Store Cleanup`
- **Command line:**
```cmd
dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase
```
- **Options tab:** Continue on error: **Yes**

> This reclaims significant space from superseded components. Can take 5-15 minutes.
> Run last so it operates on the clean post-upgrade state.

---

### GROUP: Rollback

**Add group:** Click **Add > New Group**  
Name: `Rollback`

> This group is automatically triggered by MECM when setup.exe initiates a rollback.
> SetupRollback.cmd kicks the TS back in and execution jumps here.

---

#### Step 11 — Log Rollback Event

- Click **Add > General > Run Command Line**
- **Name:** `Log Rollback Event`
- **Command line:**
```cmd
cmd /c echo %DATE% %TIME% IPU_ROLLBACK >> "%SystemDrive%\Windows\Logs\IPU_Rollback.log"
```
- **Options tab:** Continue on error: **Yes**

---

### GROUP: Run Actions on Failure

**Add group:** Click **Add > New Group**  
Name: `Run Actions on Failure`

> Steps here run if any step in the TS fails (not just rollback).

---

#### Step 12 — Collect Panther Logs

- Click **Add > General > Run Command Line**
- **Name:** `Collect Panther Logs on Failure`
- **Command line** *(update the log share path for your environment)*:
```cmd
cmd /c xcopy /y "%SystemDrive%\$Windows.~BT\Sources\Panther\setupact.log" "\\logserver\IPULogs\%_SMSTSMachineName%_setupact.log*"
```
- **Options tab:** Continue on error: **Yes**

---

### Final step — Save the TS

Click **OK** in the TS editor to save.

---

## 5. Export / Import PowerShell commands

> Run from PowerShell 5.1 only. Must be run from the CM site drive.

### Load the CM module and connect to site

```powershell
# Replace XYZ with your site code
$SiteCode   = "XYZ"
$SiteServer = "yourcmserver.domain.local"

$CMFolder = Split-Path $env:SMS_ADMIN_UI_PATH
Import-Module "$CMFolder\ConfigurationManager.psd1"
Set-Location "$SiteCode`:"
```

### Export the TS to a zip (for GitHub / backup)

```powershell
Export-CMTaskSequence `
    -Name "Windows IPU Remediation - 24H2" `
    -ExportFilePath "\\fileserver\MECM_Source\TS\IPU_Remediation.zip" `
    -WithDependence $false
```

### Import the TS from zip (new site or rebuild)

```powershell
Import-CMTaskSequence `
    -Path "\\fileserver\MECM_Source\TS\IPU_Remediation.zip" `
    -ImportActionType DirectImport
```

### Extract raw sequence XML (for source control / diffing)

```powershell
(Get-CMTaskSequence | Where-Object { $_.Name -eq "Windows IPU Remediation - 24H2" }).Sequence |
    Out-File "C:\Temp\IPU_TS_Sequence.xml" -Encoding UTF8
```

---

## 6. Monthly variable update script

> Run once a month before the new image folder goes live on the share.  
> PS 5.1 only. Adjust `$SiteCode` and `$SiteServer` for your environment.

```powershell
#Requires -Version 5.1
<#
.SYNOPSIS
    Updates the IPU_Month Task Sequence Variable to the current month.
    Run once a month when the new patched WIM folder is live on the share.

.NOTES
    Must be run on a machine with the MECM console installed.
    Run as an account with MECM full administrator or TS edit rights.
#>

param(
    [string] $SiteCode    = "XYZ",                          # <- update
    [string] $SiteServer  = "yourcmserver.domain.local",   # <- update
    [string] $TSName      = "Windows IPU Remediation - 24H2"
)

# Load CM module
$CMFolder = Split-Path $env:SMS_ADMIN_UI_PATH
Import-Module "$CMFolder\ConfigurationManager.psd1" -ErrorAction Stop
Set-Location "$SiteCode`:"

# Build new month stamp: MMYY format
$NewMonth = Get-Date -Format "MMyy"

Write-Host "Updating '$TSName' — setting IPU_Month to: $NewMonth"

# Get the TS
$TS = Get-CMTaskSequence -Name $TSName -ErrorAction Stop

if (-not $TS) {
    Write-Error "Task Sequence '$TSName' not found. Check the name and site code."
    exit 1
}

# Find and update the Set Task Sequence Variable step named "Set IPU_Month"
$Step = Get-CMTaskSequenceStep -TaskSequenceName $TSName |
        Where-Object { $_.Name -eq "Set IPU_Month" }

if (-not $Step) {
    Write-Error "Step 'Set IPU_Month' not found in the TS. Check step name matches exactly."
    exit 1
}

Set-CMTSStepSetVariable `
    -TaskSequenceName $TSName `
    -StepName         "Set IPU_Month" `
    -VariableName     "IPU_Month" `
    -VariableValue    $NewMonth

Write-Host "Done. IPU_Month is now: $NewMonth"
Write-Host "Source path will resolve to: \\computer\share\Microsoft_Windows_Enterprise_24H2_$NewMonth"
```

---

## 7. Troubleshooting reference

### Key log files

| Log | Location | What it tells you |
|---|---|---|
| smsts.log | `%SystemDrive%\Windows\CCM\Logs\` (during TS) | Full TS execution log — start here |
| setupact.log | `%SystemDrive%\$Windows.~BT\Sources\Panther\` | Windows Setup activity during upgrade |
| setuperr.log | `%SystemDrive%\$Windows.~BT\Sources\Panther\` | Setup errors only |
| setupact.log (post) | `%SystemDrive%\Windows\Panther\` | Post-upgrade setup log |

### Common setup.exe exit codes

| Exit code | Meaning |
|---|---|
| `0x00000000` | Success |
| `0xC1900101` | Driver compatibility failure — check Panther logs for which driver |
| `0xC1900200` | System not eligible for upgrade (hardware check failed) |
| `0xC1900208` | Compatibility check failed — incompatible app or driver |
| `0x800F0923` | Driver or app incompatibility detected |
| `0xC1900210` | No issues found during compat scan (not an error — internal) |

### TS did not resume after reboot

If the TS stops after the upgrade reboot and does not continue into Post-Processing:

1. Confirm you are using the native **Upgrade Operating System** step, not Run Command Line
2. Check `smsts.log` for `TSMBootstrap` entries around the reboot
3. Check if the CM client is stuck in provisioning mode:
```powershell
# Run on the affected machine as SYSTEM / admin
Invoke-WmiMethod -Namespace "root\CCM" `
                 -Class "SMS_Client" `
                 -Name "SetClientProvisioningMode" `
                 -ArgumentList $false
```

### Verify the TS variable resolved correctly

Add a temporary **Run Command Line** step at the top of Pre-Flight Checks during testing:

```cmd
cmd /c echo Source=%IPU_SourcePath% >> "%SystemDrive%\Temp\IPU_Debug.txt"
```

Remove before production deployment.

### DISM cleanup fails or hangs

DISM `/ResetBase` can take 15+ minutes on a heavily patched machine. If it times out:

```powershell
# Run manually post-upgrade if the TS step failed
dism.exe /Online /Cleanup-Image /StartComponentCleanup
# Then separately:
dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase
```
