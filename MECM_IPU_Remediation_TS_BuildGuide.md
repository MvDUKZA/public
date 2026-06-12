# MECM Windows In-Place Upgrade Remediation — Task Sequence Build Guide

**Purpose:** Repair corrupt Windows component stores by re-laying Windows 24H2 from a
pre-patched WIM on a network share. Preserves all apps, data, and settings.  
**Audience:** Senior MECM/SCCM engineer — assumes console access and CM site drive access.  
**Last updated:** June 2026

---

## Contents

1. [Prerequisites and architecture decision](#1-prerequisites-and-architecture-decision)
2. [Monthly maintenance — the only thing that changes](#2-monthly-maintenance)
3. [Task Sequence structure overview](#3-task-sequence-structure-overview)
4. [Step-by-step console build instructions](#4-step-by-step-console-build-instructions)
5. [Export / Import PowerShell commands](#5-export--import-powershell-commands)
6. [Monthly variable update script](#6-monthly-variable-update-script)
7. [Troubleshooting reference](#7-troubleshooting-reference)

---

## 1. Prerequisites and architecture decision

| Item | Detail |
|---|---|
| Share path | `\\computer\share` |
| Source folder naming | `Microsoft_Windows_Enterprise_24H2_MMYY` |
| WIM image index | `3` (Windows Enterprise) |
| MECM version | Current Branch 2207 or later |
| PowerShell for admin scripts | **5.1 only** — CM module does not support PS7 |
| Free space requirement | 20 GB on system drive |

### Why "Source path" mode, not "Upgrade Package"

The **Upgrade Operating System** TS step has two modes:

- **Upgrade package** — requires a package object imported into MECM, distributed to DPs,
  and referenced by Package ID. You cannot point it at a raw UNC path or use a variable
  for the path. With a monthly-rotating source folder this would mean re-importing and
  re-distributing a new package every month.

- **Source path** — accepts a direct UNC path to the folder containing setup.exe. MECM
  runs setup directly from the share at execution time with no DP involvement. Accepts
  TS variable references like `%IPU_SourcePath%`.

**For this use case (remediation tool, patched WIM already on a share) — Source path is
the correct choice.** You get the full native step reboot handling with none of the
package management overhead. The only monthly change is updating one TS variable.

The share must be readable by the MECM Network Access Account (or Computer Account if
NAA is not configured).

---

## 2. Monthly maintenance

> **This is the only thing you change each month — one variable value.**

Update `IPU_Month` in the TS Set Variables group.  
Format: `MMYY` — e.g. `0626` = June 2026, `0726` = July 2026, `0826` = August 2026.

Do this manually in the console, or run the script in [Section 6](#6-monthly-variable-update-script).

---

## 3. Task Sequence structure overview

```
[TS] Windows IPU Remediation - 24H2
│
├── GROUP: Set Variables
│     ├── Set TS Variable — IPU_Month           <-- ONLY THIS CHANGES MONTHLY
│     ├── Set TS Variable — IPU_SourcePath      (references IPU_Month)
│     └── Set TS Variable — OSDSetupAdditionalUpgradeOptions
│
├── GROUP: Pre-Flight Checks
│     ├── Check Readiness     (native step — free space check, no script needed)
│     └── Run Command Line    (verify share path is accessible before proceeding)
│
├── GROUP: Upgrade Operating System
│     └── Upgrade Operating System  (native step, Source path mode)
│                                   (handles ALL reboots via SetupComplete.cmd)
│                                   (TS resumes automatically after upgrade reboots)
│
├── GROUP: Post-Processing          (runs after TS resumes post-upgrade)
│     ├── Run Command Line    (remove Windows.old)
│     ├── Run Command Line    (remove $Windows.~BT)
│     ├── Run Command Line    (remove $Windows.~WS)
│     └── Run Command Line    (DISM component store cleanup)
│
├── GROUP: Rollback                 (auto-triggered by SetupRollback.cmd on failure)
│     └── Run Command Line    (log rollback event)
│
└── GROUP: Run Actions on Failure
      └── Run Command Line    (collect Panther logs)
```

---

## 4. Step-by-step console build instructions

### Create the Task Sequence shell

1. Open the MECM console
2. Navigate to: `Software Library > Operating Systems > Task Sequences`
3. Right-click > **Create Task Sequence**
4. Select: **Create a custom task sequence**
   *(Do NOT use "Upgrade an operating system" wizard — it forces an Upgrade Package object
   which can't use a dynamic UNC path. Build the steps manually instead.)*
5. Name: `Windows IPU Remediation - 24H2`
6. Boot image: **none required** — this TS runs from the full OS, not WinPE
7. Click **Next** > **Next** > **Close**

Right-click the new TS > **Edit** to open the TS editor for all steps below.

---

### GROUP: Set Variables

**Add group:** `Add > New Group`  
Name: `Set Variables`

---

#### Step 1 — IPU_Month

- `Add > General > Set Task Sequence Variable`
- **Name:** `Set IPU_Month`
- **Task Sequence Variable:** `IPU_Month`
- **Value:** `0626`

> **This is the only value you change each month.**

---

#### Step 2 — IPU_SourcePath

- `Add > General > Set Task Sequence Variable`
- **Name:** `Set IPU_SourcePath`
- **Task Sequence Variable:** `IPU_SourcePath`
- **Value:**
```
\\computer\share\Microsoft_Windows_Enterprise_24H2_%IPU_Month%
```

---

#### Step 3 — OSDSetupAdditionalUpgradeOptions

- `Add > General > Set Task Sequence Variable`
- **Name:** `Set Setup Additional Options`
- **Task Sequence Variable:** `OSDSetupAdditionalUpgradeOptions`
- **Value:**
```
/resizerecoverypartition disable /showoobe none /dynamicupdate disable /telemetry disable /diagnosticprompt enable
```

> This variable is natively read by the Upgrade Operating System step and appended
> to the setup.exe command line automatically. No need to type switches into the step itself.

---

### GROUP: Pre-Flight Checks

**Add group:** `Add > New Group`  
Name: `Pre-Flight Checks`

---

#### Step 4 — Check Readiness (native — no script needed)

- `Add > General > Check Readiness`
- **Name:** `Check Free Disk Space`
- Tick: **Ensure minimum free disk space (MB)**
- Value: `20480`
- Leave all other checks unticked
- **Options tab > Continue on error:** No

---

#### Step 5 — Verify source path accessible

- `Add > General > Run Command Line`
- **Name:** `Verify Source Path Accessible`
- **Command line:**
```cmd
cmd /c if not exist "%IPU_SourcePath%\setup.exe" exit 1
```
- **Options tab > Continue on error:** No

---

### GROUP: Upgrade Operating System

**Add group:** `Add > New Group`  
Name: `Upgrade Operating System`

---

#### Step 6 — Upgrade Operating System (native step)

- `Add > Images > Upgrade Operating System`
- **Name:** `Upgrade Operating System`
- Select: **Source path** radio button  
  *(NOT "Upgrade package" — that requires a Package ID, not a UNC path)*
- **Source path value:**
```
%IPU_SourcePath%
```
- **Edition index:** `3`
- **Product key:** leave blank (preserves existing licence)
- **Time-out (minutes):** `180`
- **Dynamically update:** Disable (do not select — matches our variable setting)
- **Options tab > Continue on error:** No

> **How reboots work here:** The native Upgrade Operating System step injects
> SetupComplete.cmd before launching setup.exe. When setup finishes and reboots,
> SetupComplete.cmd runs on the next boot, restores TSMBootstrap.exe via a registry
> value, and the TS engine resumes automatically into Post-Processing.
> You do NOT need a Restart Computer step between Upgrade and Post-Processing.

---

### GROUP: Post-Processing

**Add group:** `Add > New Group`  
Name: `Post-Processing`

---

#### Step 7 — Remove Windows.old

- `Add > General > Run Command Line`
- **Name:** `Remove Windows.old`
- **Command line:**
```cmd
cmd /c rd /s /q "%SystemDrive%\Windows.old"
```
- **Options tab > Continue on error:** Yes

---

#### Step 8 — Remove $Windows.~BT

- `Add > General > Run Command Line`
- **Name:** `Remove $Windows.~BT`
- **Command line:**
```cmd
cmd /c rd /s /q "%SystemDrive%\$Windows.~BT"
```
- **Options tab > Continue on error:** Yes

---

#### Step 9 — Remove $Windows.~WS

- `Add > General > Run Command Line`
- **Name:** `Remove $Windows.~WS`
- **Command line:**
```cmd
cmd /c rd /s /q "%SystemDrive%\$Windows.~WS"
```
- **Options tab > Continue on error:** Yes

---

#### Step 10 — DISM Component Store Cleanup

- `Add > General > Run Command Line`
- **Name:** `DISM Component Store Cleanup`
- **Command line:**
```cmd
dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase
```
- **Options tab > Continue on error:** Yes

> Runs last so it operates on the clean post-upgrade state.
> Takes 5-15 minutes on heavily patched machines — this is normal.

---

### GROUP: Rollback

**Add group:** `Add > New Group`  
Name: `Rollback`

> Automatically triggered by MECM when setup.exe initiates a rollback.
> SetupRollback.cmd fires TSMBootstrap and execution resumes here instead of
> Post-Processing.

---

#### Step 11 — Log Rollback Event

- `Add > General > Run Command Line`
- **Name:** `Log Rollback Event`
- **Command line:**
```cmd
cmd /c echo %DATE% %TIME% IPU_ROLLBACK on %_SMSTSMachineName% >> "%SystemDrive%\Windows\Logs\IPU_Rollback.log"
```
- **Options tab > Continue on error:** Yes

---

### GROUP: Run Actions on Failure

**Add group:** `Add > New Group`  
Name: `Run Actions on Failure`

---

#### Step 12 — Collect Panther Logs on Failure

- `Add > General > Run Command Line`
- **Name:** `Collect Panther Logs on Failure`
- **Command line** *(update \\logserver\IPULogs to your actual log share)*:
```cmd
cmd /c xcopy /y "%SystemDrive%\$Windows.~BT\Sources\Panther\setupact.log" "\\logserver\IPULogs\%_SMSTSMachineName%_setupact.log*"
```
- **Options tab > Continue on error:** Yes

---

### Save and close

Click **OK** in the TS editor to save all steps.

---

## 5. Export / Import PowerShell commands

> Run from PowerShell 5.1 only. Must be run from the CM site drive.
> CM module does not load in PS7.

### Load CM module and connect

```powershell
# Update SiteCode and SiteServer for your environment
$SiteCode   = "XYZ"
$SiteServer = "yourcmserver.domain.local"

$CMFolder = Split-Path $env:SMS_ADMIN_UI_PATH
Import-Module "$CMFolder\ConfigurationManager.psd1"
Set-Location "$SiteCode`:"
```

### Export the TS to zip (run once after building — store in source control)

```powershell
Export-CMTaskSequence `
    -Name           "Windows IPU Remediation - 24H2" `
    -ExportFilePath "\\fileserver\MECM_Source\TS\IPU_Remediation.zip" `
    -WithDependence $false
```

### Import the TS from zip (new site or rebuild scenario)

```powershell
Import-CMTaskSequence `
    -Path            "\\fileserver\MECM_Source\TS\IPU_Remediation.zip" `
    -ImportActionType DirectImport
```

### Extract raw sequence XML (for diffing / source control)

```powershell
(Get-CMTaskSequence | Where-Object { $_.Name -eq "Windows IPU Remediation - 24H2" }).Sequence |
    Out-File "C:\Temp\IPU_TS_Sequence.xml" -Encoding UTF8
```

---

## 6. Monthly variable update script

> Run once a month when the new image folder is live on the share.
> Only updates the IPU_Month variable — nothing else in the TS changes.

```powershell
#Requires -Version 5.1
<#
.SYNOPSIS
    Updates IPU_Month TS variable to the current month stamp (MMYY).
    Run monthly when the new patched WIM folder is live on the share.

.NOTES
    Requires MECM console installed on the machine running this script.
    Run as an account with TS edit rights in MECM.
    PowerShell 5.1 only — CM module does not support PS7.
#>

param(
    [string] $SiteCode   = "XYZ",                         # <- update
    [string] $SiteServer = "yourcmserver.domain.local",   # <- update
    [string] $TSName     = "Windows IPU Remediation - 24H2",
    [string] $StepName   = "Set IPU_Month"
)

# Load CM module
$CMFolder = Split-Path $env:SMS_ADMIN_UI_PATH
Import-Module "$CMFolder\ConfigurationManager.psd1" -ErrorAction Stop
Set-Location "$SiteCode`:"

# Calculate new month stamp: MMYY
$NewMonth = Get-Date -Format "MMyy"

Write-Host "Updating '$TSName'"
Write-Host "Setting IPU_Month to: $NewMonth"
Write-Host "Source will resolve to: \\computer\share\Microsoft_Windows_Enterprise_24H2_$NewMonth"

# Verify TS exists
$TS = Get-CMTaskSequence -Name $TSName -Fast -ErrorAction Stop
if (-not $TS) {
    Write-Error "Task Sequence '$TSName' not found."
    exit 1
}

# Verify step exists
$Step = Get-CMTaskSequenceStep -TaskSequenceName $TSName |
        Where-Object { $_.Name -eq $StepName }

if (-not $Step) {
    Write-Error "Step '$StepName' not found in TS. Check step name matches exactly."
    exit 1
}

# Update the variable value
Set-CMTSStepSetVariable `
    -TaskSequenceName $TSName `
    -StepName         $StepName `
    -VariableName     "IPU_Month" `
    -VariableValue    $NewMonth

Write-Host "Done. IPU_Month updated to: $NewMonth"
```

---

## 7. Troubleshooting reference

### Key log files

| Log | Location on client | What it tells you |
|---|---|---|
| smsts.log | `%WinDir%\CCM\Logs\` | Full TS execution — start here |
| setupact.log | `%SystemDrive%\$Windows.~BT\Sources\Panther\` | Setup activity during upgrade |
| setuperr.log | `%SystemDrive%\$Windows.~BT\Sources\Panther\` | Setup errors only |
| setupact.log | `%WinDir%\Panther\` | Post-upgrade setup log |
| IPU_Rollback.log | `%WinDir%\Logs\` | Written by Step 11 on rollback |

### Common setup.exe exit codes

| Exit code | Meaning |
|---|---|
| `0x00000000` | Success |
| `0xC1900101` | Driver compatibility failure — check Panther log for which driver |
| `0xC1900200` | Hardware not eligible (unusual for same-version repair) |
| `0xC1900208` | Compatibility check failed — incompatible app or driver |
| `0x800F0923` | Driver or app incompatibility detected |
| `0xC1900210` | No issues found during compat scan — internal code, not an error |

### TS did not resume after upgrade reboot

If the TS stops and does not continue into Post-Processing:

1. Confirm the Upgrade OS step is using **Source path** mode, not Run Command Line
2. Check smsts.log for `TSMBootstrap` entries around the reboot timestamp
3. Check if the CM client is stuck in provisioning mode — run on the affected machine:

```powershell
# Run as admin/SYSTEM on the affected machine
Invoke-WmiMethod -Namespace "root\CCM" `
                 -Class       "SMS_Client" `
                 -Name        "SetClientProvisioningMode" `
                 -ArgumentList $false
```

### Verify the source path variable resolved correctly

Add a temporary step at the top of Pre-Flight Checks during initial testing, remove before production:

```cmd
cmd /c echo Source=%IPU_SourcePath% >> "%SystemDrive%\Temp\IPU_Debug.txt"
```

### DISM cleanup takes too long or fails

DISM `/ResetBase` can take 15+ minutes on heavily patched machines. If the step times out
or fails, run manually after the upgrade:

```cmd
dism.exe /Online /Cleanup-Image /StartComponentCleanup
dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase
```
