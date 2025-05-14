<#
.SYNOPSIS
    Converts GPResult XML to readable CSV.
.DESCRIPTION
    Parses a GPResult XML file produced by `gpresult /x` and exports the group policy settings to a CSV file.
.PARAMETER XmlPath
    Path to the GPResult XML file.
.PARAMETER CsvPath
    Path to the output CSV file. Defaults to the same directory and name as the XML file with a .csv extension.
.EXAMPLE
    .\Convert-GPResultXmlToCsv.ps1 -XmlPath "C:\Reports\gpresult.xml"
#>
param (
    [Parameter(Mandatory = $true)]
    [string]$XmlPath,

    [Parameter(Mandatory = $false)]
    [string]$CsvPath = ""
)

function Write-Log {
    param (
        [Parameter(Mandatory = $true)][ValidateSet("INFO","WARN","ERROR")][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Output "$timestamp [$Level] $Message"
}

# Verify XML file exists
if (-Not (Test-Path -Path $XmlPath -PathType Leaf)) {
    Write-Log -Level "ERROR" -Message "XML file not found: $XmlPath"
    exit 1
}

# Determine CSV path
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($XmlPath)
    $directory = [System.IO.Path]::GetDirectoryName($XmlPath)
    $CsvPath   = Join-Path -Path $directory -ChildPath "$baseName.csv"
}

# Load XML and strip namespaces/prefixes robustly
try {
    $rawXml = Get-Content -Path $XmlPath -Raw -ErrorAction Stop
    Write-Log -Level "INFO" -Message "Loaded raw XML"

    # Remove all xmlns declarations (xmlns and xmlns:prefix)
    $cleanXml = $rawXml -replace 'xmlns(:\w+)?="[^"]+"',''
    # Remove namespace prefixes in element names e.g., <rsop:Node> -> <Node>
    $cleanXml = [regex]::Replace($cleanXml, '(</?)(\w+:)', '$1', 'IgnoreCase')
    # Remove namespace prefixes in attribute names e.g., xsi:type -> type
    $cleanXml = [regex]::Replace($cleanXml, '(\s)(\w+:)(?=\w+=)', '$1', 'IgnoreCase')

    [xml]$gpoReport = $cleanXml
    Write-Log -Level "INFO" -Message "Stripped namespaces/prefixes and parsed XML"
}
catch {
    Write-Log -Level "ERROR" -Message "Failed to load or parse XML: $_"
    exit 1
}

$results = @()

# Iterate Computer and User scopes
foreach ($scope in 'Computer','User') {
    $parent = $gpoReport.RSOP.$scope
    if ($null -eq $parent) {
        Write-Log -Level "WARN" -Message "Scope not found: $scope"
        continue
    }

    $extensions = $parent.ExtensionData.Extension
    if ($null -eq $extensions -or $extensions.Count -eq 0) {
        Write-Log -Level "WARN" -Message "No extensions under scope: $scope"
        continue
    }

    foreach ($ext in $extensions) {
        $extName = $ext.Name
        Write-Log -Level "INFO" -Message "Processing extension: $extName ($scope)"

        # Find any element with Name and State/Value attributes
        $policyNodes = $ext.SelectNodes(".//*[@Name and (@State or @Value)]")
        foreach ($node in $policyNodes) {
            $results += [PSCustomObject]@{
                Scope     = $scope
                Extension = $extName
                Setting   = $node.GetAttribute('Name')
                State     = $node.GetAttribute('State')
                Value     = $node.GetAttribute('Value')
            }
        }
    }
}

# Export to CSV
try {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Log -Level "INFO" -Message "Exported CSV: $CsvPath"
}
catch {
    Write-Log -Level "ERROR" -Message "CSV export failed: $_"
    exit 1
}
