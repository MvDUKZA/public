<#
.SYNOPSIS
    Retrieves details of failed Qualys Policy Compliance postures from a CSV file and provides autofix functions for specific controls.

.DESCRIPTION
    This script imports a CSV file containing failed Qualys posture IDs, fetches detailed evidence for failures using the Qualys PCRS API, and outputs the reasons for failure. It includes autofix functions for certain controls, such as updating Microsoft Defender virus definitions. Autofix requires remote access to the affected hosts.

    The script follows best practices for error handling, input validation, logging, and security. It uses the Qualys PCRS API for fetching posture information with evidence.

    Note: For PCRS API, ensure the QualysBaseUrl is set to your subscription's Gateway URL (e.g., https://gateway.qg1.apps.qualys.eu for EU Platform 1). Refer to https://www.qualys.com/platform-identification/ for the correct URL. Do not use the standard API URL like https://qualysapi.qualys.eu, as /auth may not be available there.

    Reference: Qualys API Documentation (VM/PC User Guide, updated July 7, 2025) - https://www.qualys.com/docs/qualys-api-vmpc-user-guide.pdf

.PARAMETER CsvPath
    The path to the CSV file containing failed posture data.

.PARAMETER QualysBaseUrl
    The base URL for the Qualys API (e.g., https://gateway.qg1.apps.qualys.eu).

.PARAMETER QualysCredential
    PSCredential object for Qualys API authentication.

.PARAMETER RemoteCredential
    PSCredential object for remote host access (required for autofix).

.PARAMETER SubscriptionId
    Optional subscription ID for multi-subscription environments.

.PARAMETER Fix
    Switch to enable autofix for supported controls.

.PARAMETER LogPath
    Path for log file. Defaults to C:\temp\scripts\logs\QualysFix.log.

.EXAMPLE
    $qualysCred = Get-Credential -Message "Enter Qualys credentials"
    $remoteCred = Get-Credential -Message "Enter remote admin credentials"
    .\Get-QualysFailureDetailsAndFix.ps1 -CsvPath "C:\path\to\failed_postures.csv" -QualysBaseUrl "https://gateway.qg1.apps.qualys.eu" -QualysCredential $qualysCred -RemoteCredential $remoteCred -Fix

.NOTES
    Author: Marinus van Deventer
    Version: 1.1
    Date: 21 July 2025
    Dependencies: Requires PowerShell 7+ for optimal performance. Uses Invoke-RestMethod for API calls.
    Changelog:
    - 1.0: Initial version with PCRS API integration and Defender autofix.
    - 1.1: Updated to PCRS v3.0 with adjusted endpoint and request body to include policyId and optional subscriptionId. Added parameter for SubscriptionId. Updated authentication to use Invoke-RestMethod. Added note on using Gateway URL. Grouped requests by policy for efficient API calls.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ if (Test-Path $_ -PathType Leaf) { $true } else { throw "CSV file not found: $_" } })]
    [string]$CsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [uri]$QualysBaseUrl,

    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$QualysCredential,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$RemoteCredential,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [switch]$Fix,

    [string]$LogPath = "C:\temp\scripts\logs\QualysFix.log"
)

#region Initialization
# Ensure working directory exists
$workingDir = "C:\temp\scripts"
if (-not (Test-Path $workingDir -PathType Container)) {
    New-Item -Path $workingDir -ItemType Directory -ErrorAction Stop | Out-Null
}

# Ensure logs directory exists
$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir -PathType Container)) {
    New-Item -Path $logDir -ItemType Directory -ErrorAction Stop | Out-Null
}

# Set error action preference
$ErrorActionPreference = 'Stop'

# Function to log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
    Write-Verbose $Message
}
#endregion

#region Helper Functions
# Function to get Qualys JWT token
function Get-QualysToken {
    param (
        [uri]$BaseUrl,
        [PSCredential]$Credential
    )
    try {
        $authUri = "$BaseUrl/auth"
        $body = @{
            username = $Credential.UserName
            password = $Credential.GetNetworkCredential().Password
            token    = 'true'
        }
        $response = Invoke-RestMethod -Uri $authUri -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
        Write-Log "Successfully obtained Qualys token."
        return $response.Trim()
    } catch {
        Write-Log "Error obtaining Qualys token: $_" "ERROR"
        throw
    }
}

# Function to fetch posture info with evidence
function Get-PostureDetails {
    param (
        [uri]$BaseUrl,
        [string]$Token,
        [array]$PolicyHostMappings
    )
    try {
        $postureUri = "$BaseUrl/pcrs/3.0/posture/postureInfo?evidenceRequired=1&compressionRequired=0"
        $body = $PolicyHostMappings | ConvertTo-Json -Depth 3
        $headers = @{
            Authorization = "Bearer $Token"
            'Content-Type' = 'application/json'
            'Accept' = 'application/json'
        }
        $response = Invoke-RestMethod -Uri $postureUri -Method Post -Headers $headers -Body $body
        Write-Log "Fetched posture details for $($PolicyHostMappings.Count) policy mappings."
        return $response
    } catch {
        Write-Log "Error fetching posture details: $_" "ERROR"
        throw
    }
}
#endregion

#region Fix Functions
# Fix function for Control ID 2781: Update Microsoft Defender definitions
function Fix-DefenderDefinitions {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [string]$ComputerName,
        [PSCredential]$Credential
    )
    try {
        if ($PSCmdlet.ShouldProcess($ComputerName, "Update Microsoft Defender definitions")) {
            Invoke-Command -ComputerName $ComputerName -Credential $Credential -ScriptBlock {
                # Reference: https://learn.microsoft.com/en-us/powershell/module/defender/update-mpsignature?view=windowsserver2022-ps
                Update-MpSignature -ErrorAction Stop
            }
            Write-Log "Updated Defender definitions on $ComputerName."
        }
    } catch {
        Write-Log "Error updating Defender on $ComputerName: $_" "ERROR"
        throw
    }
}

# Placeholder for other fixes, e.g., Control ID 2816: Authorized processes
# function Fix-AuthorizedProcesses {
#     param (
#         [string]$ComputerName,
#         [PSCredential]$Credential,
#         [string]$Evidence  # Use evidence to determine unauthorized processes
#     )
#     # Implement logic to stop unauthorized processes if safe
#     # Caution: This can be risky; manual review recommended
# }
#endregion

#region Main Logic
try {
    Write-Log "Script started. Processing CSV: $CsvPath"

    # Import CSV
    $failedPostures = Import-Csv -Path $CsvPath
    if ($failedPostures.Count -eq 0) {
        Write-Log "No failed postures found in CSV." "WARN"
        return
    }

    # Group by Policy Id and collect unique hosts per policy
    $policyGroups = $failedPostures | Group-Object -Property 'Policy Id'
    $policyHostMappings = @()
    foreach ($group in $policyGroups) {
        $policyId = $group.Name
        $hostIds = $group.Group | Select-Object -ExpandProperty 'Host Id' -Unique
        $mapping = @{
            policyId = $policyId
            hostIds = $hostIds
        }
        if ($SubscriptionId) {
            $mapping.subscriptionId = $SubscriptionId
        }
        $policyHostMappings += $mapping
    }

    # Get Qualys token
    $token = Get-QualysToken -BaseUrl $QualysBaseUrl -Credential $QualysCredential

    # Fetch posture details
    $postureData = Get-PostureDetails -BaseUrl $QualysBaseUrl -Token $token -PolicyHostMappings $policyHostMappings

    # Process each failed posture
    foreach ($posture in $failedPostures) {
        $hostId = $posture.'Host Id'
        $controlId = $posture.'Control Id'
        $postureId = $posture.'Posture Id'
        $dnsHostname = $posture.'DNS Hostname'

        # Find matching data in posture response (assuming structure: postureData contains array of host objects with controls)
        # Note: Adapt based on actual JSON structure from API; placeholder logic here
        $hostData = $postureData | Where-Object { $_.hostId -eq $hostId }
        $controlData = $hostData.controls | Where-Object { $_.controlId -eq $controlId -and $_.postureId -eq $postureId }
        $evidence = $controlData.evidence  # Or extendedEvidence/causeOfFailure

        Write-Output "Posture ID: $postureId | Host: $dnsHostname | Control ID: $controlId"
        Write-Output "Failure Details/Evidence: $evidence"

        if ($Fix) {
            if (-not $RemoteCredential) {
                throw "RemoteCredential required for autofix."
            }
            switch ($controlId) {
                '2781' {
                    Fix-DefenderDefinitions -ComputerName $dnsHostname -Credential $RemoteCredential
                }
                '2816' {
                    Write-Log "Autofix for Control 2816 not implemented (manual review required)." "WARN"
                }
                default {
                    Write-Log "No autofix available for Control ID $controlId." "INFO"
                }
            }
        }
    }

    Write-Log "Script completed successfully."
} catch {
    Write-Log "Script error: $_" "ERROR"
    throw
} finally {
    # Clean up if needed
}
#endregion
