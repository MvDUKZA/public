# Omnissa VM Optimization Extractor Script

# CONFIGURATION VARIABLES
$jsonPath = "c:\temp\work\Windows 10, 11 and Server 2019, 2022 2025-05-27-161320.json"
$xmlPath = "c:\temp\work\Windows 10, 11 and Server 2019, 2022.xml"
$outputDir = "C:\temp\work"
$logFile = "$outputDir\ExtractionLog.txt"

# Setup Logging
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir }
if (Test-Path $logFile) { Remove-Item $logFile }
Start-Transcript -Path $logFile -Append

Write-Host "Loading JSON and XML files..."
$jsonData = Get-Content $jsonPath | ConvertFrom-Json
[xml]$xmlData = Get-Content $xmlPath

# Extract Steps with IsSelected:true from JSON
$selectedSteps = $jsonData.TemplateItemList | Where-Object { $_.IsSelected -eq $true } | ForEach-Object { $_.Step.Trim().ToLower() }
Write-Host "Found $($selectedSteps.Count) selected steps in JSON."

# Prepare output files
$outputFiles = @{
    Registry      = "$outputDir\Registry.csv"
    Service       = "$outputDir\Service.csv"
    SchTasks      = "$outputDir\SchTasks.csv"
    ShellExecute  = "$outputDir\ShellExecute.csv"
    InternalClass = "$outputDir\InternalClass.csv"
}
$outputData = @{
    Registry      = @()
    Service       = @()
    SchTasks      = @()
    ShellExecute  = @()
    InternalClass = @()
}
$unmatchedSteps = @()

# Recursive Function to Get All Steps
function Get-AllSteps($xmlGroup) {
    $allSteps = @()
    if ($xmlGroup.step) { $allSteps += $xmlGroup.step }
    if ($xmlGroup.group) { foreach ($subGroup in $xmlGroup.group) { $allSteps += Get-AllSteps $subGroup } }
    return $allSteps
}

Write-Host "Searching XML for matching steps..."
$allXmlSteps = Get-AllSteps $xmlData.sequence

foreach ($stepName in $selectedSteps) {
    $matchedXmlStep = $allXmlSteps | Where-Object { $_.name.Trim().ToLower() -eq $stepName }
    if ($matchedXmlStep) {
        foreach ($step in $matchedXmlStep) {
            foreach ($action in $step.action) {
                $type = $action.type
                $controlId = $step.nodeId
                switch ($type) {
                    "Registry" {
                        $row = [PSCustomObject]@{
                            ControlID = $controlId
                            Key       = $action.params.keyName
                            Name      = $action.params.valueName
                            Value     = $action.params.data
                            Type      = $action.params.type
                        }
                        $outputData.Registry += $row
                    }
                    "Service" {
                        $row = [PSCustomObject]@{
                            ControlID   = $controlId
                            ServiceName = $action.params.serviceName
                            StartMode   = $action.params.startMode
                        }
                        $outputData.Service += $row
                    }
                    "SchTasks" {
                        $row = [PSCustomObject]@{
                            ControlID = $controlId
                            TaskName  = $action.params.taskName
                            Status    = $action.params.status
                        }
                        $outputData.SchTasks += $row
                    }
                    "ShellExecute" {
                        $row = [PSCustomObject]@{
                            ControlID = $controlId
                            Command   = $action.command
                        }
                        $outputData.ShellExecute += $row
                    }
                    "InternalClass" {
                        $row = [PSCustomObject]@{
                            ControlID = $controlId
                            Command   = $action.command
                            Params    = ($action.params | Get-Member -MemberType NoteProperty | ForEach-Object { "$($_.Name)=$($action.params.$($_.Name))" }) -join "; "
                        }
                        $outputData.InternalClass += $row
                    }
                }
            }
        }
    } else {
        Write-Warning "No match in XML for JSON step: $stepName"
        $unmatchedSteps += $stepName
    }
}

# Export to CSV
Write-Host "Writing CSV files..."
foreach ($type in $outputData.Keys) {
    if ($outputData[$type].Count -gt 0) {
        $outputData[$type] | Export-Csv -Path $outputFiles[$type] -NoTypeInformation -Force
        Write-Host "$type exported to $($outputFiles[$type])"
    } else {
        Write-Host "$type has no data to export."
    }
}

# Log unmatched steps
if ($unmatchedSteps.Count -gt 0) {
    $unmatchedStepsFile = "$outputDir\UnmatchedSteps.txt"
    $unmatchedSteps | Out-File -FilePath $unmatchedStepsFile
    Write-Warning "Some JSON steps did not match XML. See $unmatchedStepsFile for details."
}

Stop-Transcript
Write-Host "Extraction complete. Log saved to $logFile"