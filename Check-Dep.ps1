# Function to check if an executable is DEP-compatible using dumpbin
function Test-DepCompatibility {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ExePath
    )

    # Check if the file exists
    if (-not (Test-Path $ExePath)) {
        Write-Warning "File not found: $ExePath"
        return $null
    }

    # Path to dumpbin.exe (adjust based on your Visual Studio installation)
    $dumpbinPath = "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Tools\MSVC\14.29.30133\bin\Hostx64\x64\dumpbin.exe"
    
    if (-not (Test-Path $dumpbinPath)) {
        Write-Error "dumpbin.exe not found at $dumpbinPath. Please ensure Visual Studio is installed and update the path."
        return $null
    }

    try {
        # Run dumpbin to get headers
        $output = & $dumpbinPath /headers $ExePath 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Failed to analyze $ExePath with dumpbin."
            return $null
        }

        # Check for NXCOMPAT flag in the output
        $isNxCompat = $output -match "NX compatible"

        return [PSCustomObject]@{
            ExePath     = $ExePath
            DEPCompatible = $isNxCompat
        }
    }
    catch {
        Write-Warning "Error analyzing $ExePath : $_"
        return $null
    }
}

# Function to check DEP compatibility for all running processes
function Test-RunningProcesses {
    Write-Host "Checking DEP compatibility for all running processes..." -ForegroundColor Cyan

    # Get all running processes and filter for .exe files
    $processes = Get-Process | Where-Object { $_.Path -and $_.Path.EndsWith(".exe") } | Select-Object -Unique Path

    $results = @()
    foreach ($process in $processes) {
        $result = Test-DepCompatibility -ExePath $process.Path
        if ($result) {
            $results += $result
        }
    }

    return $results
}

# Function to check DEP compatibility for specific EXEs by file path
function Test-SpecificExes {
    Write-Host "Please select one or more .exe files to check for DEP compatibility..." -ForegroundColor Cyan

    # Use OpenFileDialog to select multiple .exe files
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    $openFileDialog.Filter = "Executable Files (*.exe)|*.exe"
    $openFileDialog.Title = "Select Executable Files"

    $results = @()
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        foreach ($file in $openFileDialog.FileNames) {
            $result = Test-DepCompatibility -ExePath $file
            if ($result) {
                $results += $result
            }
        }
    }
    else {
        Write-Host "No files selected." -ForegroundColor Yellow
    }

    return $results
}

# Main script
Write-Host "DEP Compatibility Checker" -ForegroundColor Green
Write-Host "1. Check all running processes"
Write-Host "2. Check specific .exe files"
$choice = Read-Host "Enter your choice (1 or 2)"

$results = @()
if ($choice -eq "1") {
    $results = Test-RunningProcesses
}
elseif ($choice -eq "2") {
    $results = Test-SpecificExes
}
else {
    Write-Host "Invalid choice. Please run the script again and select 1 or 2." -ForegroundColor Red
    exit
}

# Display results
Write-Host "`nResults:" -ForegroundColor Cyan
$results | Format-Table -Property ExePath, @{Label="DEP Compatible"; Expression={$_.DEPCompatible}} -AutoSize

# Optionally save results to a CSV file
$save = Read-Host "Would you like to save the results to a CSV file? (y/n)"
if ($save -eq 'y' -or $save -eq 'Y') {
    $outputPath = Join-Path $env:USERPROFILE "Desktop\DEP_Compatibility_Report.csv"
    $results | Export-Csv -Path $outputPath -NoTypeInformation
    Write-Host "Results saved to $outputPath" -ForegroundColor Green
}