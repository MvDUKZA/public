# Extract-UniqueControlIDs.ps1
param(
    [Parameter(Mandatory=$true)]
    [string]$XmlFilePath
)

# Load the XML file
$xmlContent = [xml](Get-Content -Path $XmlFilePath)

# Select all CONTROL nodes and extract their ID values
$uniqueIDs = $xmlContent.SelectNodes('//CONTROL/ID') | 
    ForEach-Object { $_.InnerText } |
    Sort-Object -Unique

# Output the results
$uniqueIDs