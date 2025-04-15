# Import the GroupPolicy module if not already loaded
Import-Module GroupPolicy

# Define key variables
$csvPath = "C:\Path\To\Your\file.csv"           # Path to your CSV file
$gpoName = "Win11 VDI Policy"                   # Name of the target GPO
$collectionName = "CIS-Hardening_01"            # Group Policy Preferences Registry Collection Name
$context = "Computer"                           # Use "Computer" for machine policy

# Load registry settings from CSV
$registryItems = Import-Csv -Path $csvPath

# Try to get the GPO; throw an error if it doesn't exist
$gpo = Get-GPO -Name $gpoName -ErrorAction Stop

# Loop through each row in the CSV
foreach ($item in $registryItems) {
    # Extract root hive and subkey
    $rootKey, $subKey = $item.Key -split "\\", 2

    # Validate the root hive
    $hiveMap = @{
        "HKLM" = "HKLM"
        "HKEY_LOCAL_MACHINE" = "HKLM"
        "HKCU" = "HKCU"
        "HKEY_CURRENT_USER" = "HKCU"
    }

    $hive = $hiveMap[$rootKey]
    if (-not $hive) {
        Write-Warning "Unsupported registry hive in key: $($item.Key)"
        continue
    }

    # Display progress
    Write-Host "Adding registry item: $($item.Key)\$($item.Name) - Value: $($item.Value) - Type: $($item.Type)"

    # Set the registry preference item in the specified GPO and collection
    Set-GPPrefRegistryValue -Name $gpoName `
        -Context $context `
        -Key "$hive\$subKey" `
        -ValueName $item.Name `
        -Type $item.Type `
        -Value $item.Value `
        -Action Update `
        -Collection $collectionName
}

Write-Host ""
Write-Host "Registry preferences successfully applied to GPO '$gpoName' under COMPUTER CONFIGURATION in collection '$collectionName'."
