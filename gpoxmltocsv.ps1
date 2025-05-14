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

# Check if XML file exists
if (-Not (Test-Path -Path $XmlPath -PathType Leaf)) {
    Write-Log -Level "ERROR" -Message "XML file not found: $XmlPath"
    exit 1
}

# Set default CsvPath if not provided
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($XmlPath)
    $directory = [System.IO.Path]::GetDirectoryName($XmlPath)
    $CsvPath = Join-Path -Path $directory -ChildPath "$baseName.csv"
}

# Load XML
try {
    [xml]$gpoReport = Get-Content -Path $XmlPath -Raw -ErrorAction Stop
    Write-Log -Level "INFO" -Message "Loaded XML file: $XmlPath"
}
catch {
    Write-Log -Level "ERROR" -Message "Failed to load XML file: $_"
    exit 1
}

# Setup XML namespace manager
$nsManager = New-Object System.Xml.XmlNamespaceManager($gpoReport.NameTable)
$nsManager.AddNamespace("rsop", "http://www.microsoft.com/GroupPolicy/Rsop")

$results = @()

# Process scopes
foreach ($scope in @("Computer", "User")) {
    # Select the scope node using XPath with namespace
    $parent = $gpoReport.SelectSingleNode("/rsop:RSOP/rsop:$scope", $nsManager)
    if ($null -eq $parent) {
        Write-Log -Level "WARN" -Message "Scope not found: $scope"
        continue
    }
    
    # Select ExtensionData/Extension nodes with namespace
    $extensions = $parent.SelectNodes("rsop:ExtensionData/rsop:Extension", $nsManager)
    if ($null -eq $extensions -or $extensions.Count -eq 0) {
        Write-Log -Level "WARN" -Message "No extensions found for scope: $scope"
        continue
    }
    
    foreach ($ext in $extensions) {
        $extName = $ext.GetAttribute("Name")
        Write-Log -Level "INFO" -Message "Processing extension: $extName in scope: $scope"
        
        # Select all elements with Name and State/Value attributes within the extension, using namespace
        $policyNodes = $ext.SelectNodes(".//rsop:*[@Name and (@State or @Value)]", $nsManager)
        foreach ($node in $policyNodes) {
            $name = $node.GetAttribute("Name")
            $state = $node.GetAttribute("State")
            $value = $node.GetAttribute("Value")
            $results += [PSCustomObject]@{
                Scope = $scope
                Extension = $extName
                Setting = $name
                State = $state
                Value = $value
            }
        }
    }
}

# Export to CSV
try {
    $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Log -Level "INFO" -Message "Exported results to CSV: $CsvPath"
}
catch {
    Write-Log -Level "ERROR" -Message "Failed to export CSV: $_"
    exit 1
}
