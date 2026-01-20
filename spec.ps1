# PSPackageBuilder – Updated Specifications (Renamed Scripts)

This document consolidates and updates the requirements for the PSPackageBuilder project, including the final script names and responsibilities.

---

## Project Name

**PSPackageBuilder**

## Purpose

Provide a PowerShell-based tooling workflow to:

1. Create and materialise standards-compliant application package folders (Script 1)
2. Generate PSAppDeployToolkit v4 “light wrapper” scripts and inject install/uninstall blocks (Script 2)

The repository contains **code only** (GitHub). Application payloads/binaries are **never** stored in GitHub.

---

# Script 1 Specification

## Script 1 Name

**scripts/New-PSPackage.ps1**

## Script 1 Public Entry Function

**Invoke-NewPSPackage**

## Script 1 Purpose

Create a correctly named package folder in the package repository, populated with:

* PSAppDeployToolkit template (PSADT)
* Source payload files
* `package.json` manifest
* `README.md` documentation

Script 1 supports **New Package** and **Clone Existing Package** modes.

## Script 1 In Scope

* New package creation
* Clone existing package
* Payload selection from a source drop
* Metadata extraction (MSI/EXE/MSIX-APPX/scripts)
* Vendor normalisation
* Version normalisation
* Revision selection
* Folder name generation and validation
* Duplicate/collision prevention
* Final editable folder-name confirmation
* Materialisation (create folder, copy PSADT, copy payload, write manifest + README)

## Script 1 Out of Scope

* Creating install/uninstall/detect logic (Script 2)
* SCCM / Intune integration
* Modifying PSADT logic beyond copying the template
* Source control actions (committing binaries)

---

## Script 1 Defaults and Overrides

### Defaults Location (Hard Requirement)

All hardcoded paths and defaults MUST exist ONLY in:

* `src/PSPackageBuilder/Private/Defaults.ps1`

### Required Defaults

* `RepoRoot` (package repository share/folder root)
* `SourceRoot` (source drop for payload selection)
* `PsadtTemplateRoot` (PSADT template location)
* Supported extensions:

  * `.msi`, `.exe`, `.msix`, `.appx`, `.msixbundle`, `.appxbundle`, `.ps1`, `.cmd`, `.bat`
* `DefaultTarget = 'X'` (do not infer target from file type)
* Filenames: `package.json`, `README.md`
* Behaviour flags: collision/duplicate policies, version fallback, etc.

### Override Precedence (Hard Requirement)

1. Explicit function parameter
2. Pipeline context object property
3. Defaults.ps1

---

## Script 1 Team Resolution (Hard Requirement)

* Determine current user and resolve AD group membership
* Map AD group → 3-letter Team code using `config/TeamMap.json`
* If user cannot be mapped to a Team code, Script 1 MUST fail fast with a clear error
* Team code is immutable in New mode; in Clone mode inherit from existing package by default

---

## Script 1 Target Rule (Hard Requirement)

* Target MUST default to `X`
* Target MUST NOT be inferred from file type
* Clone mode should inherit the Target from existing package unless overridden

---

## Script 1 Mode Selection (Hard Requirement)

Prompt user to select one mode:

1. New Package
2. Clone Existing Package

---

## Script 1 Payload Selection

* Payload selection root is `SourceRoot`
* Multiple files may be selected
* Use `Out-GridView` if available; provide CLI fallback otherwise
* Only supported extensions are selectable

---

## Script 1 Clone Mode Behaviour

* User selects an existing package folder under `RepoRoot`
* Script MUST read existing metadata from `package.json` (do NOT parse folder name)
* Inherit tokens from cloned package unless overridden (at minimum: Team/Target, and typically Vendor/Product/Type)

---

## Script 1 Metadata Extraction (Best-Effort)

### MSI

* ProductName
* ProductVersion
* Manufacturer

### MSIX/APPX and bundles

* Read `AppxManifest.xml` from the package container
* Extract:

  * Identity.Version
  * Properties.DisplayName (product)
  * Identity.Publisher / PublisherDisplayName (vendor mapping)

### EXE

* FileVersionInfo where available (CompanyName, ProductName/FileDescription, ProductVersion/FileVersion)

### Scripts

* Filename-based metadata (optionally future header parsing)

Each extractor returns a standard metadata object including:

* VendorRaw, ProductRaw, VersionRaw
* Vendor, Product, Version (normalised)
* Confidence (High/Medium/Low)
* Source (which fields were used)

---

## Script 1 Vendor Normalisation (Hard Requirement)

* Use `config/VendorMap.json`
* Case-insensitive matching and trimming
* Canonicalise vendor strings (e.g. Microsoft Corporation / Microsoft Inc → Microsoft)

---

## Script 1 Version + Revision Rules (Authoritative)

### New Package

* Version = extracted or fallback
* Revision = `01`

### Clone Package

* If new payload Version != existing Version:

  * Version = new payload Version
  * Revision = `01`
* If new payload Version == existing Version:

  * Version unchanged
  * Revision increments (`02`, `03`, ...)

Revisions are always 2 digits, zero-padded.

---

## Script 1 Duplication and Collision Rules

* Identical Vendor + Product + Version + identical payload hash (SHA256) MUST be blocked
* Same Vendor + Product + Version with different payload hash MAY create a new revision
* Collision handling must follow configured policy (e.g., auto-increment revision)

---

## Script 1 Final Editable Confirmation (Hard Requirement)

* At the end, display the proposed folder name in an editable textbox dialog (WinForms)
* User can edit; script validates:

  * naming compliance
  * illegal characters
  * uniqueness
* OK continues; Cancel exits without changes

---

## Script 1 Materialisation (Hard Requirement)

After confirmation:

1. Create `<RepoRoot>\<FinalFolderName>`
2. Copy PSADT template content from `PsadtTemplateRoot` into the package folder
3. Copy selected payload(s) into:

   * `Files\Source\`
4. Hash payloads (SHA256)
5. Write `package.json`
6. Write `README.md`

### Output Structure (Guaranteed)

<PackageFolder>
README.md
package.json
(PSADT template files)
Files
Source <payload files>

---

# Script 2 Specification

## Script 2 Name

**scripts/New-PSPackageWrappers.ps1**

## Script 2 Public Entry Function

**Invoke-NewPSPackageWrappers**

## Script 2 Purpose

Generate PSAppDeployToolkit v4 wrapper scripts and inject installation/uninstallation blocks for an existing package created by Script 1.

Script 2 creates/overwrites exactly three files in the selected package folder:

* `install.ps1`
* `uninstall.ps1`
* `detect.ps1`

Script 2 also inserts/updates auto-generated sections in `Invoke-AppDeployToolkit.ps1` between clearly delimited markers.

Reference: PSADT v4 documentation for installing applications and process execution:
[https://psappdeploytoolkit.com/docs/usage/installing%20applications](https://psappdeploytoolkit.com/docs/usage/installing%20applications)

---

## Script 2 In Scope

* Selecting an existing package folder from the repository
* Reading package metadata (`package.json`)
* Selecting a payload within the package (`Files\Source\`)
* Building best-effort silent install logic based on payload type
* Detecting MST transforms for MSI and applying them
* Displaying the final install line in an editable textbox
* Writing wrapper scripts (install/uninstall/detect)
* Updating/injecting blocks into `Invoke-AppDeployToolkit.ps1`
* Updating `package.json` with wrapper metadata

## Script 2 Out of Scope

* Rebuilding package folder naming (Script 1 job)
* SCCM/Intune packaging/export workflows beyond producing detect.ps1
* Deep installer introspection (e.g. unpacking EXE types beyond best-effort)

---

## Script 2 Preconditions

Selected package folder MUST contain:

* PSADT v4 entrypoint (`Invoke-AppDeployToolkit.exe` or `Invoke-AppDeployToolkit.ps1`)
* `Files\` folder
* `package.json`

If any are missing, Script 2 MUST fail fast with a clear error.

---

## Script 2 User Flow (Hard Requirement)

1. User selects a package folder (Out-GridView preferred; CLI fallback)
2. Script reads `package.json`
3. User selects payload under `Files\Source\` (Out-GridView preferred; CLI fallback)
4. Script determines payload type and builds:

   * Install logic
   * Uninstall logic (best effort)
   * Detect logic (best effort)
5. If MSI: detect `.mst` files in same folder (or under Files\Source):

   * If exactly one MST found: auto-select it
   * If multiple: allow selection (multi-select)
6. Show computed “install line” in editable textbox; user can modify
7. On OK: write/overwrite wrapper files and inject PSADT blocks
8. Update `package.json` with wrapper metadata

Cancel at editable textbox exits without writing files.

---

## Script 2 Payload Types and Logic

### MSI

Install:

* Use PSADT v4 `Start-ADTMsiProcess`
* Include `-Transforms` when MST selected
* MSI path should be referenced relative to PSADT files directory (`$adtSession.DirFiles` preferred)

Uninstall:

* Prefer ProductCode uninstall if obtainable:

  * `Start-ADTMsiProcess -Action Uninstall -ProductCode '{GUID}'`
* If ProductCode not available, provide a best-effort fallback (e.g. uninstall by name/filter) and clearly mark as TODO

Detect:

* Prefer MSI ProductCode detection via uninstall registry keys
* Optional validation: DisplayName contains AppName and DisplayVersion equals package.json Version

### EXE

Install:

* Use PSADT v4 `Start-ADTProcess -FilePath <exe> -ArgumentList <silent args>`
* Provide best-effort silent args suggestions but never assume perfect
* Require final user review/edit in textbox

Uninstall:

* Best-effort: registry-based detection of uninstall string and execute it, or use AppName lookup logic; if ambiguous, insert TODO placeholder and require user edit later

Detect:

* Registry-based detection on DisplayName (contains AppName)
* Optional version match to package.json Version

### Script payload (.ps1 / .cmd / .bat)

Install:

* .ps1: `Start-ADTProcess -FilePath 'powershell.exe' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File "<script>"'`
* .cmd/.bat: `Start-ADTProcess -FilePath 'cmd.exe' -ArgumentList '/c "<script>"'`

Uninstall:

* Default placeholder (TODO) unless user provides explicit uninstall command

Detect:

* Default to a clear TODO and return “not detected” unless a marker mechanism is configured

---

## Script 2 Editable Install Line (Hard Requirement)

* Must show the computed install statement as a single editable line in a WinForms textbox dialog
* User can edit it
* On OK: validate not empty and can be written into a PS script file
* On Cancel: exit without changes

---

## Script 2 PSADT v4 Injection (Hard Requirement)

Script 2 must insert/update auto-generated blocks inside `Invoke-AppDeployToolkit.ps1`:

Markers:

* Install block:

  * `### PSPackageBuilder:BEGIN AUTO-INSTALL ###`
  * `### PSPackageBuilder:END AUTO-INSTALL ###`
* Uninstall block:

  * `### PSPackageBuilder:BEGIN AUTO-UNINSTALL ###`
  * `### PSPackageBuilder:END AUTO-UNINSTALL ###`

Within these blocks, insert the final install/uninstall statements using PSADT v4 functions (`Start-ADTProcess`, `Start-ADTMsiProcess`).

---

## Script 2 Wrapper Files (Hard Requirement)

Script 2 creates/overwrites exactly these three files in the package root:

### install.ps1

* Calls PSADT v4 entrypoint with:

  * `-DeploymentType Install`
  * `-DeployMode Silent` (default; allow override)
* Should not contain installer logic; installer logic lives in injected block in `Invoke-AppDeployToolkit.ps1`

### uninstall.ps1

* Calls PSADT v4 entrypoint with:

  * `-DeploymentType Uninstall`
  * `-DeployMode Silent` (default; allow override)

### detect.ps1

* Standalone detection script (usable by SCCM/Intune):

  * Exit 0 = detected
  * Exit 1 = not detected
* Reads `package.json` for AppName/AppVersion and any MSI ProductCode captured
* Implements detection best-effort for MSI/EXE; optional MSIX/AppX support

---

## Script 2 package.json Update (Hard Requirement)

Script 2 must update `package.json` with wrapper metadata, including:

* SelectedPayloadPath (relative to package root)
* PayloadType (MSI/EXE/SCRIPT)
* InstallLine (final edited line)
* UninstallLine (generated or placeholder)
* Detect method and keys used
* MSI ProductCode (if obtained)
* Timestamp and author

---

# Repository Structure (Updated Names)

scripts/

* `New-PSPackage.ps1`
* `New-PSPackageWrappers.ps1`

Module public functions:

* `Invoke-NewPSPackage`
* `Invoke-NewPSPackageWrappers`

---

# GitHub Policy

* Repository contains code only
* Payload/binaries are not stored in GitHub
* `.gitignore` must exclude common installer extensions and staging/cache folders
