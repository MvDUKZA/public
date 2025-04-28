# Function to write consistent log messages with timestamps
function Write-Log {
    param([string]$Message)
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
}

# Function to parse the .audit file content
function Parse-AuditFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    # Regex pattern to extract each <custom_item> block
    $pattern = "<custom_item>(.*?)</custom_item>"
    $customItemMatches = [regex]::Matches($Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

    if ($customItemMatches.Count -eq 0) {
        Write-Warning "No <custom_item> entries found in audit file."
        return
    }

    # Initialize global storage
    $global:dataDict = @{}
    $global:allItems = @()
    $global:noIndexItems = @()

    # Loop through each custom_item found
    foreach ($match in $customItemMatches) {
        $item = $match.Groups[1].Value

        # Extract 'type' field
        $typeMatch = [regex]::Match($item, 'type\s*:\s*(.+?)\n')
        if (!$typeMatch.Success) { continue }
        $type = $typeMatch.Groups[1].Value.Trim()

        # Skip AUDIT_POWERSHELL items
        if ($type -eq "AUDIT_POWERSHELL") { continue }

        # Extract 'description' field
        $descriptionMatch = [regex]::Match($item, 'description\s*:\s*(.+?)\n')
        if (!$descriptionMatch.Success) { continue }
        $description = $descriptionMatch.Groups[1].Value.Trim()
        $description = $description.Trim('"')

        # Extract CIS Index (start of description)
        $index = $null
        $cisMatch = [regex]::Match($description, '^\s*(\d+(\.\d+)+)')
        if ($cisMatch.Success) {
            $index = $cisMatch.Groups[1].Value.Trim()
        }

        # Extract optional fields
        $solutionMatch = [regex]::Match($item, 'solution\s*:\s*(.+?)\n\s*reference')
        $solution = if ($solutionMatch.Success) { $solutionMatch.Groups[1].Value.Trim().Replace("`n", ' ') } else { "" }

        $valueDataMatch = [regex]::Match($item, 'value_data\s*:\s*(.+?)\n')
        $valueData = if ($valueDataMatch.Success) { $valueDataMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        $valueTypeMatch = [regex]::Match($item, 'value_type\s*:\s*(.+?)\n')
        $valueType = if ($valueTypeMatch.Success) { 
            switch ($valueTypeMatch.Groups[1].Value.Trim()) {
                "POLICY_DWORD"      { "DWord" }
                "POLICY_STRING"     { "String" }
                "POLICY_MULTI_TEXT" { "MultiString" }
                Default              { $valueTypeMatch.Groups[1].Value.Trim() }
            }
        } else { "" }

        $regKeyMatch = [regex]::Match($item, 'reg_key\s*:\s*(.+?)\n')
        $regKey = if ($regKeyMatch.Success) { $regKeyMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        $regItemMatch = [regex]::Match($item, 'reg_item\s*:\s*(.+?)\n')
        $regItem = if ($regItemMatch.Success) { $regItemMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        $regOptionMatch = [regex]::Match($item, 'reg_option\s*:\s*(.+?)\n')
        $regOption = if ($regOptionMatch.Success) { $regOptionMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        $keyItemMatch = [regex]::Match($item, 'key_item\s*:\s*(.+?)\n')
        $keyItem = if ($keyItemMatch.Success) { $keyItemMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        $auditPolicySubcategoryMatch = [regex]::Match($item, 'audit_policy_subcategory\s*:\s*(.+?)\n')
        $auditPolicySubcategory = if ($auditPolicySubcategoryMatch.Success) { $auditPolicySubcategoryMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        $rightTypeMatch = [regex]::Match($item, 'right_type\s*:\s*(.+?)\n')
        $rightType = if ($rightTypeMatch.Success) { $rightTypeMatch.Groups[1].Value.Trim().Replace('"', '') } else { "" }

        # If key_item exists, overwrite regItem
        if ($keyItem) { $regItem = $keyItem }

        # Build parsed object
        $parsedItem = [PSCustomObject]@{
            Checklist = 1
            Type = $type
            Index = $index
            Description = $description
            Solution = $solution
            "Reg Key" = $regKey
            "Reg Item" = $regItem
            "Reg Option" = $regOption
            "Audit Policy Subcategory" = $auditPolicySubcategory
            "Right Type" = $rightType
            "Value Data" = $valueData
            "Value Type" = $valueType
        }

        # Add to master list
        $global:allItems += $parsedItem

        # If no Index, add separately
        if (-not $index) {
            $global:noIndexItems += $parsedItem
        }

        # Add to type group
        if (-not $global:dataDict.ContainsKey($type)) {
            $global:dataDict[$type] = @()
        }
        $global:dataDict[$type] += $parsedItem
    }

    Write-Log "Parsed $($customItemMatches.Count) <custom_item> entries."
}

# Helper for natural CIS Index sort
function NaturalSort {
    param([string]$Index)
    if (-not $Index) { return @(9999) }
    return ($Index -split '\.') | ForEach-Object { [int]$_ }
}

# Export everything
function Export-DataToCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputFolder
    )

    if (-not $global:allItems) {
        Write-Log "No data found to export."
        return
    }

    if (-not (Test-Path $OutputFolder)) {
        try {
            New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
            Write-Log "Created output folder: $OutputFolder"
        } catch {
            Write-Log "Failed to create output folder: $OutputFolder. Error: $_"
            return
        }
    }

    # Export All settings
    try {
        $allOutputFile = Join-Path -Path $OutputFolder -ChildPath "All_Settings.csv"
        $global:allItems | Sort-Object @{Expression={NaturalSort $_.Index}; Ascending=$true} | Export-Csv -Path $allOutputFile -NoTypeInformation -Encoding UTF8
        Write-Log "Exported ALL settings to $allOutputFile"
    } catch {
        Write-Log "Failed to export all settings. Error: $_"
    }

    # Export per Type
    foreach ($type in $global:dataDict.Keys) {
        $safeType = ($type -replace '[\\/:*?"<>|]', '_')
        $outputFile = Join-Path -Path $OutputFolder -ChildPath "$safeType.csv"

        try {
            $global:dataDict[$type] | Sort-Object @{Expression={NaturalSort $_.Index}; Ascending=$true} | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
            Write-Log "Exported type '$type' to $outputFile"
        } catch {
            Write-Log "Failed to export type '$type'. Error: $_"
        }
    }

    # Export NoIndex items
    if ($global:noIndexItems.Count -gt 0) {
        try {
            $noIndexFile = Join-Path -Path $OutputFolder -ChildPath "NoIndexItems.csv"
            $global:noIndexItems | Export-Csv -Path $noIndexFile -NoTypeInformation -Encoding UTF8
            Write-Log "Exported entries without CIS Index to $noIndexFile"
        } catch {
            Write-Log "Failed to export NoIndexItems. Error: $_"
        }
    }
}

# Main driver function
function Main {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AuditFilePath,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputFolder
    )

    if (-not (Test-Path $AuditFilePath)) {
        Write-Log "Audit file not found: $AuditFilePath"
        return
    }

    try {
        $content = Get-Content -Path $AuditFilePath -Raw -ErrorAction Stop
    } catch {
        Write-Log "Failed to read audit file. Error: $_"
        return
    }

    Parse-AuditFile -Content $content
    Export-DataToCsv -OutputFolder $OutputFolder
}

# Example usage
$inputAuditFile = "C:\temp\CIS_Microsoft_Windows_11_Enterprise_v4.0.0_L1.audit"
$outputCsvFolder = "C:\temp"
Main -AuditFilePath $inputAuditFile -OutputFolder $outputCsvFolder
