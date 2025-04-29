# Final version: reads simplified CSV, builds Registry.xml for Group Policy Preferences

$CsvPath = "C:\Temp\RegistrySettings.csv"
$OutputXmlPath = "C:\Temp\Registry.xml"

# Load CSV
$RegistrySettings = Import-Csv -Path $CsvPath

# Create XML document
$xml = New-Object System.Xml.XmlDocument
$declaration = $xml.CreateXmlDeclaration("1.0", "UTF-8", $null)
$xml.AppendChild($declaration) | Out-Null

# Create root element
$root = $xml.CreateElement("Registry")
$xml.AppendChild($root)

foreach ($setting in $RegistrySettings) {
    # Split Hive and KeyPath
    if ($setting.Key -match "^([^\\]+)\\(.+)$") {
        $cls = $matches[1]
        $path = $matches[2]
    } else {
        Write-Warning "Invalid key format: $($setting.Key)"
        continue
    }

    # Construct the unique display name: ControlID-Name
    $nameAttr = "$($setting.ControlID)-$($setting.Name)"

    # Create Registry item
    $regItem = $xml.CreateElement("Registry")
    $regItem.SetAttribute("cls", $cls)
    $regItem.SetAttribute("name", $nameAttr)
    $regItem.SetAttribute("path", $path)
    $regItem.SetAttribute("status", "Enabled")
    $regItem.SetAttribute("image", "2")
    $regItem.SetAttribute("changed", (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
    $regItem.SetAttribute("uid", [guid]::NewGuid().ToString())
    $regItem.SetAttribute("removePolicy", "1")  # Important: auto-delete if policy not applied
    $regItem.SetAttribute("key", $path)
    $regItem.SetAttribute("valueName", $setting.Name)
    $regItem.SetAttribute("action", "U")         # "U" = Update (replaces if exists)
    $regItem.SetAttribute("type", $setting.Type)
    $regItem.SetAttribute("value", $setting.Value)

    $root.AppendChild($regItem) | Out-Null
}

# Save output
$xml.Save($OutputXmlPath)

Write-Host "`n✔ Registry.xml created at: $OutputXmlPath" -ForegroundColor Green
