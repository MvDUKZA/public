<#
.SYNOPSIS
Exports Qualys Compliance Policy reports with secure authentication and error handling.

.DESCRIPTION
Automates export of Qualys Compliance Policies in XML, CSV, or PDF format with:
- Secure credential management
- Interactive policy selection
- Asynchronous report handling
- Comprehensive error checking

.PARAMETER SavePath
Output path for the report (including filename)

.PARAMETER ExportFormat
Report format (xml, csv, pdf)

.PARAMETER PolicyId
[Optional] Specific policy ID to skip interactive selection

.EXAMPLE
PS> .\Export-QualysComplianceReport.ps1 -SavePath "C:\Audit\Q4_Report.pdf" -ExportFormat pdf
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SavePath,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("xml","csv","pdf")]
    [string]$ExportFormat,
    
    [string]$PolicyId
)

# === INITIALIZATION ===
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"

# === SECURE AUTHENTICATION ===
$credential = Get-Credential -Message "Enter Qualys API credentials"
$auth = [Convert]::ToBase64String(
    [Text.Encoding]::ASCII.GetBytes(
        "$($credential.UserName):$($credential.GetNetworkCredential().Password)"
    )
)
$headers = @{
    Authorization = "Basic $auth"
    Accept        = "application/xml"
}

# === API BASE URL ===
$baseUrl = "https://qualysapi.qualys.com"

# === FUNCTIONS ===
function Invoke-QualysApi {
    param(
        [string]$Uri,
        [hashtable]$Body
    )
    
    try {
        $response = Invoke-RestMethod -Uri $Uri `
            -Method Post `
            -Headers $headers `
            -Body $Body `
            -ContentType "application/x-www-form-urlencoded"
        
        # Check for API errors
        if ($response.RESPONSE.ERROR) {
            throw "API Error $($response.RESPONSE.ERROR.NUMBER): $($response.RESPONSE.ERROR.TEXT)"
        }
        
        return $response
    }
    catch {
        Write-Error "API call failed: $_"
        exit 1
    }
}

# === MAIN SCRIPT ===
try {
    # Ensure output directory exists
    $outputDir = Split-Path $SavePath -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    # === STEP 1: List Compliance Policies ===
    Write-Verbose "Fetching policy list..."
    $listResponse = Invoke-QualysApi -Uri "$baseUrl/api/2.0/fo/compliance/policy/" -Body @{ action = "list" }
    
    $policies = $listResponse.POLICY_LIST_OUTPUT.POLICY_LIST.POLICY
    
    if (-not $policies) {
        Write-Error "No compliance policies found"
        exit 1
    }

    # === STEP 2: Policy Selection ===
    if (-not $PolicyId) {
        Write-Host "`nAvailable Policies:`n"
        $PolicyId = $policies | 
            Select-Object @{Name="PolicyID"; Expression={$_.ID}}, Title |
            Out-GridView -Title "Select a Compliance Policy" -OutputMode Single |
            Select-Object -ExpandProperty PolicyID
        
        if (-not $PolicyId) {
            Write-Error "No policy selected"
            exit 1
        }
    }

    # === STEP 3: Initiate Export ===
    Write-Host "`nExporting Policy ID $PolicyId ($ExportFormat)..."
    $exportParams = @{
        action        = "fetch"
        policy_id     = $PolicyId
        output_format = $ExportFormat
    }
    
    $exportResponse = Invoke-QualysApi -Uri "$baseUrl/api/2.0/fo/compliance/policy/" -Body $exportParams

    # === STEP 4: Handle Report Output ===
    if ($ExportFormat -in "csv","pdf") {
        # Handle async report generation
        $downloadId = $exportResponse.COMPLIANCE_POLICY_OUTPUT.RESPONSE.ITEM.ID
        
        Write-Host "Report generation started (ID: $downloadId). Checking status..."
        do {
            $statusResponse = Invoke-QualysApi -Uri "$baseUrl/api/2.0/fo/report/" -Body @{
                action = "status"
                id     = $downloadId
            }
            
            $status = $statusResponse.REPORT_OUTPUT.RESPONSE.ITEM.STATUS.STATE
            Write-Verbose "Current status: $status"
            Start-Sleep -Seconds 10
            
        } while ($status -ne "finished")

        # Download completed report
        $downloadUrl = $statusResponse.REPORT_OUTPUT.RESPONSE.ITEM.DOWNLOAD_URL
        Write-Host "Downloading report from: $downloadUrl"
        Invoke-WebRequest -Uri $downloadUrl -OutFile $SavePath -Headers $headers
    }
    else {
        # XML is returned inline
        $exportResponse | Out-File -FilePath $SavePath
    }

    Write-Host "`nSuccess! Report saved to: $SavePath"
    if ((Get-Item $SavePath).Length -eq 0) {
        Write-Warning "Output file is empty - verify policy content and API permissions"
    }
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}