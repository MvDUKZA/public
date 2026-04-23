<#
.SYNOPSIS
    Triggers a Qualys Cloud Agent Vulnerability Management "Scan on Demand" on one or more
    remote Windows machines by setting the registry key that the agent monitors.

.DESCRIPTION
    Qualys Cloud Agent for Windows watches HKLM\SOFTWARE\Qualys\QualysAgent\ScanOnDemand
    in real time. Writing ScanOnDemand = 1 under the Vulnerability subkey forces an
    on-demand VM manifest collection, which is the supported mechanism for a post-patch
    rescan without waiting for the normal assessment interval.

    Reference:
    https://docs.qualys.com/en/ca/install-guide/windows/configuration/configure_scan_on_demand.htm

    Because the Remote Registry service is disabled in this environment, the script uses
    PowerShell Remoting (WinRM) to execute the registry write locally on each target.

    Registry values written (all REG_DWORD):
        HKLM\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability\ScanOnDemand = 1
        HKLM\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability\CpuLimit     = <CpuLimit>
        (Optional) ...\ScanOnStartup = 1   if -ScanOnStartup is specified

    Agent behaviour:
        ScanOnDemand data value:
            1 = execute now (what we set)
            2 = scan in progress (agent sets this)
            0 = scan complete   (agent sets this)

.PARAMETER ComputerName
    One or more target computers. Accepts pipeline input.

.PARAMETER InputFile
    Path to a text file containing one computer name per line. Blank lines and lines
    starting with '#' are ignored.

.PARAMETER CpuLimit
    CPU throttle percentage for the on-demand scan only (2-100). Default 100 = no throttle,
    which is the recommended value inside a patch/change window for fastest turnaround.

.PARAMETER ScanOnStartup
    Also set ScanOnStartup = 1 so the agent rescans after the next service start / reboot.
    Useful when the patch job requires a reboot to finish remediation.

.PARAMETER Credential
    Alternate credentials for PS Remoting. Defaults to the current user.

.PARAMETER ThrottleLimit
    Max concurrent remote sessions when multiple computers are supplied. Default 32.

.PARAMETER LogPath
    Optional CSV path to write per-host results. Directory is created if missing.

.EXAMPLE
    .\Invoke-QualysScanOnDemand.ps1 -ComputerName PC001

.EXAMPLE
    .\Invoke-QualysScanOnDemand.ps1 -ComputerName PC001,PC002,PC003 -Verbose

.EXAMPLE
    .\Invoke-QualysScanOnDemand.ps1 -InputFile .\patched-hosts.txt -ScanOnStartup `
        -LogPath .\QualysScan-Results.csv

.EXAMPLE
    Get-ADComputer -Filter "Name -like 'VDI-*'" | Select-Object -ExpandProperty Name |
        .\Invoke-QualysScanOnDemand.ps1 -CpuLimit 100

.NOTES
    Requires:
        - PowerShell Remoting enabled on targets (WinRM).
        - Caller must have local administrator rights on targets (HKLM write).
        - Qualys Cloud Agent installed and running on targets.
        - VM module activated for the agent (agent will not scan an unactivated manifest).

    Agent versions earlier than 4.8 may not have the Vulnerability subkey pre-created;
    the script creates any missing keys.
#>

[CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess)]
param(
    [Parameter(ParameterSetName = 'ByName', Mandatory, Position = 0,
               ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CN', 'Name', 'Computer')]
    [string[]] $ComputerName,

    [Parameter(ParameterSetName = 'ByFile', Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $InputFile,

    [ValidateRange(2, 100)]
    [int] $CpuLimit = 100,

    [switch] $ScanOnStartup,

    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [ValidateRange(1, 256)]
    [int] $ThrottleLimit = 32,

    [string] $LogPath
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # Scriptblock executed on each target. Self-contained - no outer variables referenced.
    $remoteScript = {
        param(
            [int] $CpuLimit,
            [bool] $AlsoScanOnStartup
        )

        $result = [ordered]@{
            ComputerName    = $env:COMPUTERNAME
            AgentInstalled  = $false
            AgentService    = $null
            AgentVersion    = $null
            PreviousState   = $null
            KeyCreated      = $false
            ScanOnDemandSet = $false
            CpuLimitSet     = $null
            ScanOnStartup   = $null
            Success         = $false
            Message         = $null
        }

        try {
            # 1. Sanity-check that the Qualys agent is present. No point writing the key
            #    on a host that has no agent to read it.
            $svc = Get-Service -Name 'QualysAgent' -ErrorAction SilentlyContinue
            if (-not $svc) {
                $result.Message = 'Qualys Cloud Agent service (QualysAgent) not found.'
                return [pscustomobject]$result
            }
            $result.AgentInstalled = $true
            $result.AgentService   = $svc.Status.ToString()

            $agentRoot = 'HKLM:\SOFTWARE\Qualys\QualysAgent'
            if (Test-Path $agentRoot) {
                $ver = (Get-ItemProperty -Path $agentRoot -ErrorAction SilentlyContinue).ProductVersion
                if ($ver) { $result.AgentVersion = $ver }
            }

            # 2. Ensure the full key path exists. On agent < 4.8 only root keys are
            #    auto-created; the Vulnerability subkey may be missing.
            $sodPath = 'HKLM:\SOFTWARE\Qualys\QualysAgent\ScanOnDemand'
            $vulnPath = Join-Path $sodPath 'Vulnerability'

            foreach ($p in @($sodPath, $vulnPath)) {
                if (-not (Test-Path $p)) {
                    New-Item -Path $p -Force | Out-Null
                    $result.KeyCreated = $true
                }
            }

            # 3. Record the previous ScanOnDemand value so we can see whether a scan is
            #    already mid-flight (agent sets it to 2 while running, 0 when done).
            $prev = Get-ItemProperty -Path $vulnPath -Name 'ScanOnDemand' -ErrorAction SilentlyContinue
            if ($prev) { $result.PreviousState = $prev.ScanOnDemand }

            if ($result.PreviousState -eq 2) {
                # Agent is already scanning. Writing 1 now is a no-op and misleading.
                $result.Message = 'A Qualys scan is already in progress (ScanOnDemand = 2). No change made.'
                $result.Success = $true
                return [pscustomobject]$result
            }

            # 4. Write the values the agent watches. All must be REG_DWORD.
            New-ItemProperty -Path $vulnPath -Name 'CpuLimit' -PropertyType DWord `
                -Value $CpuLimit -Force | Out-Null
            $result.CpuLimitSet = $CpuLimit

            New-ItemProperty -Path $vulnPath -Name 'ScanOnDemand' -PropertyType DWord `
                -Value 1 -Force | Out-Null
            $result.ScanOnDemandSet = $true

            if ($AlsoScanOnStartup) {
                New-ItemProperty -Path $vulnPath -Name 'ScanOnStartup' -PropertyType DWord `
                    -Value 1 -Force | Out-Null
                $result.ScanOnStartup = 1
            }

            $result.Success = $true
            $result.Message = 'ScanOnDemand trigger written. Agent will pick it up in real time.'
        }
        catch {
            $result.Success = $false
            $result.Message = $_.Exception.Message
        }

        [pscustomobject]$result
    }

    # Resolve computer list from either parameter set.
    $targets = [System.Collections.Generic.List[string]]::new()
    if ($PSCmdlet.ParameterSetName -eq 'ByFile') {
        Get-Content -LiteralPath $InputFile |
            ForEach-Object { $_.Trim() } |
            Where-Object   { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object { $targets.Add($_) }

        if ($targets.Count -eq 0) {
            throw "Input file '$InputFile' contained no usable computer names."
        }
    }

    $allResults = [System.Collections.Generic.List[object]]::new()
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'ByName') {
        foreach ($c in $ComputerName) { $targets.Add($c) }
    }
}

end {
    # De-duplicate while preserving order.
    $targets = $targets |
        Where-Object { $_ } |
        Select-Object -Unique

    if (-not $targets) {
        Write-Warning 'No target computers supplied.'
        return
    }

    Write-Verbose ("Targeting {0} computer(s). CpuLimit={1}. ScanOnStartup={2}." -f `
        @($targets).Count, $CpuLimit, [bool]$ScanOnStartup)

    $shouldMsg = "Set HKLM\SOFTWARE\Qualys\QualysAgent\ScanOnDemand\Vulnerability\ScanOnDemand = 1"
    if (-not $PSCmdlet.ShouldProcess(($targets -join ', '), $shouldMsg)) { return }

    $icmParams = @{
        ComputerName   = $targets
        ScriptBlock    = $remoteScript
        ArgumentList   = @($CpuLimit, [bool]$ScanOnStartup)
        ThrottleLimit  = $ThrottleLimit
        ErrorAction    = 'Continue'  # keep going on per-host failures
        ErrorVariable  = 'remoteErrors'
    }
    if ($Credential -and $Credential.UserName) {
        $icmParams['Credential'] = $Credential
    }

    $results = Invoke-Command @icmParams

    # Fold any remoting-level failures (unreachable host, WinRM refused, auth, etc.) into
    # the same result shape so the output is uniform.
    foreach ($err in $remoteErrors) {
        $badHost = $err.TargetObject
        if (-not $badHost -and $err.OriginInfo) { $badHost = $err.OriginInfo.PSComputerName }
        if (-not $badHost) { $badHost = '<unknown>' }

        $allResults.Add([pscustomobject]@{
            ComputerName    = $badHost
            AgentInstalled  = $null
            AgentService    = $null
            AgentVersion    = $null
            PreviousState   = $null
            KeyCreated      = $false
            ScanOnDemandSet = $false
            CpuLimitSet     = $null
            ScanOnStartup   = $null
            Success         = $false
            Message         = "Remote connection failed: $($err.Exception.Message)"
        })
    }

    foreach ($r in $results) {
        # Strip PSRemoting-added properties so the object is clean on screen / in CSV.
        $allResults.Add(
            $r | Select-Object ComputerName, AgentInstalled, AgentService, AgentVersion,
                               PreviousState, KeyCreated, ScanOnDemandSet, CpuLimitSet,
                               ScanOnStartup, Success, Message
        )
    }

    # Emit to pipeline.
    $allResults

    # Summary to the host so an operator can see the result at a glance.
    $ok   = @($allResults | Where-Object Success).Count
    $fail = @($allResults).Count - $ok
    Write-Host ""
    Write-Host ("Qualys ScanOnDemand trigger complete: {0} succeeded, {1} failed." -f $ok, $fail) `
        -ForegroundColor (if ($fail) { 'Yellow' } else { 'Green' })

    if ($LogPath) {
        $dir = Split-Path -Parent $LogPath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $allResults | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results written to: $LogPath" -ForegroundColor Cyan
    }
}
