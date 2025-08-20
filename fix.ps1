<#
.SYNOPSIS
    Fixer for CID 27981: Force Microsoft Defender signature update.

.DESCRIPTION
    Attempts to update Microsoft Defender signatures on the local system.
    Uses Update-MpSignature when available; otherwise falls back to MpCmdRun.exe.
    Starts the WinDefend service if required. Retries a few times. Returns a
    PSCustomObject with Outcome and Details for the orchestrator to capture.

.OUTPUTS
    [pscustomobject] with properties:
        Outcome : Succeeded | Failed | Not Applicable
        Details : Text describing what happened

.NOTES
    Working dir: C:\temp\scripts
    Logs: C:\temp\scripts\logs\Fix_27981_<yyyyMMdd_HHmm>.log
    Signed by Marinus van Deventer
#>

[CmdletBinding()]
param()

begin {
    $ErrorActionPreference = 'Stop'

    #region Paths and Logging
    $workingDir = 'C:\temp\scripts'
    $logsDir    = Join-Path $workingDir 'logs'
    foreach ($d in @($workingDir,$logsDir)) {
        if (-not (Test-Path $d -PathType Container)) {
            New-Item -Path $d -ItemType Directory -Force | Out-Null
        }
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
    $logPath   = Join-Path $logsDir "Fix_27981_$timestamp.log"

    function Write-Log {
        param([string]$Message, [ValidateSet('INFO','WARNING','ERROR')][string]$Level='INFO')
        $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        Add-Content -Path $logPath -Value $line -Encoding UTF8
        Write-Verbose $line
    }
    #endregion

    #region Helpers
    function Get-MpCmdRunPath {
        # Prefer the latest Platform subfolder; fall back to the base folder
        $base = Join-Path $env:ProgramFiles 'Windows Defender'
        $plat = Join-Path $base 'Platform'
        $candidates = @()

        if (Test-Path $plat) {
            $latest = Get-ChildItem -Path $plat -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
            if ($latest) {
                $candidates += (Join-Path $latest.FullName 'MpCmdRun.exe')
            }
        }
        $candidates += (Join-Path $base 'MpCmdRun.exe')

        foreach ($p in $candidates) {
            if (Test-Path $p) { return $p }
        }
        return $null
    }

    function Get-DefenderStatus {
        # Returns a hashtable with useful status facts; tolerates missing Defender module
        $status = @{
            HasModule    = $false
            ServiceName  = 'WinDefend'
            ServiceState = $null
            AVSigVer     = $null
            AVSigAge     = $null
            AVLastUpdate = $null
            EngineVer    = $null
            ProductVer   = $null
        }

        $svc = Get-Service -Name 'WinDefend' -ErrorAction SilentlyContinue
        if ($svc) { $status.ServiceState = $svc.Status.ToString() }

        $mp = Get-Command -Name Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($mp) {
            $status.HasModule = $true
            try {
                $s = Get-MpComputerStatus -ErrorAction Stop
                $status.AVSigVer     = $s.AntivirusSignatureVersion
                $status.AVSigAge     = $s.AntivirusSignatureAge
                $status.AVLastUpdate = $s.AntivirusSignatureLastUpdated
                $status.EngineVer    = $s.AMEngineVersion
                $status.ProductVer   = $s.AMProductVersion
            } catch {
                # Ignore telemetry failures
            }
        }
        return $status
    }

    function Ensure-WinDefendRunning {
        $svc = Get-Service -Name 'WinDefend' -ErrorAction SilentlyContinue
        if (-not $svc) {
            Write-Log 'WinDefend service not found. Microsoft Defender may not be installed.' 'WARNING'
            return $false
        }
        if ($svc.Status -ne 'Running') {
            try {
                Write-Log 'Starting WinDefend service...'
                Start-Service -Name 'WinDefend' -ErrorAction Stop
                $svc.WaitForStatus('Running','00:00:20')
            } catch {
                Write-Log ("Failed to start WinDefend: $($_.Exception.Message)") 'ERROR'
                return $false
            }
        }
        return $true
    }

    function Try-UpdateWithCmdlet {
        param([int]$RetryCount = 3, [int]$RetryDelaySeconds = 20)

        $cmd = Get-Command -Name Update-MpSignature -ErrorAction SilentlyContinue
        if (-not $cmd) {
            Write-Log 'Update-MpSignature not available in this session.' 'WARNING'
            return $false, 'Update-MpSignature not available'
        }

        $sources = @('MicrosoftUpdateServer','MMPC','InternalDefinitionUpdateServer')
        for ($i=1; $i -le $RetryCount; $i++) {
            foreach ($src in $sources) {
                try {
                    Write-Log ("Attempt $i using Update-MpSignature -UpdateSource $src")
                    Update-MpSignature -UpdateSource $src -ErrorAction Stop
                    return $true, ("Update-MpSignature succeeded via $src")
                } catch {
                    Write-Log ("Update-MpSignature via $src failed: $($_.Exception.Message)") 'WARNING'
                }
            }
            if ($i -lt $RetryCount) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
        return $false, 'Update-MpSignature attempts exhausted'
    }

    function Try-UpdateWithMpCmdRun {
        param([int]$RetryCount = 3, [int]$RetryDelaySeconds = 20)

        $mpPath = Get-MpCmdRunPath
        if (-not $mpPath) {
            Write-Log 'MpCmdRun.exe not found.' 'ERROR'
            return $false, 'MpCmdRun.exe not found'
        }

        for ($i=1; $i -le $RetryCount; $i++) {
            try {
                Write-Log ("Attempt $i using MpCmdRun.exe -SignatureUpdate at $($mpPath)")
                $p = Start-Process -FilePath $mpPath -ArgumentList '-SignatureUpdate' -PassThru -Wait -WindowStyle Hidden
                if ($p.ExitCode -eq 0) {
                    return $true, 'MpCmdRun.exe -SignatureUpdate returned 0'
                } else {
                    Write-Log ("MpCmdRun.exe exit code $($p.ExitCode)") 'WARNING'
                }
            } catch {
                Write-Log ("MpCmdRun.exe failed: $($_.Exception.Message)") 'WARNING'
            }
            if ($i -lt $RetryCount) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
        return $false, 'MpCmdRun.exe attempts exhausted'
    }

    function Format-StatusDetails {
        param($Before,$After,$MethodNote)
        $parts = @()
        if ($MethodNote) { $parts += $MethodNote }
        if ($Before -and $After) {
            $parts += ("AVVersion {0} -> {1}" -f ($Before.AVSigVer ?? '<unknown>'), ($After.AVSigVer ?? '<unknown>'))
            if ($After.AVLastUpdate) {
                $parts += ("LastUpdate {0:yyyy-MM-dd HH:mm}" -f $After.AVLastUpdate)
            }
        }
        if ($After.EngineVer)  { $parts += ("Engine {0}"  -f $After.EngineVer) }
        if ($After.ProductVer) { $parts += ("Product {0}" -f $After.ProductVer) }
        return ($parts -join '; ')
    }
    #endregion
}

process {
    try {
        # Detect presence
        $before = Get-DefenderStatus
        if (-not (Ensure-WinDefendRunning)) {
            return [pscustomobject]@{ Outcome='Not Applicable'; Details='WinDefend service missing or cannot start' }
        }

        # Decide path: cmdlet if available, else MpCmdRun
        $usedMethod = $null
        $ok = $false; $note = $null

        $cmdPresent = [bool](Get-Command -Name Update-MpSignature -ErrorAction SilentlyContinue)
        if ($cmdPresent) {
            $res = Try-UpdateWithCmdlet -RetryCount 3 -RetryDelaySeconds 20
            $ok  = $res[0]; $note = $res[1]; $usedMethod = 'Update-MpSignature'
        }

        if (-not $ok) {
            $res = Try-UpdateWithMpCmdRun -RetryCount 3 -RetryDelaySeconds 20
            $ok  = $res[0]; $note = $res[1]; if ($ok) { $usedMethod = 'MpCmdRun.exe' }
        }

        # Gather after state and produce result
        $after = Get-DefenderStatus
        $details = Format-StatusDetails -Before $before -After $after -MethodNote $note

        if ($ok) {
            return [pscustomobject]@{ Outcome='Succeeded'; Details=$details }
        } else {
            return [pscustomobject]@{ Outcome='Failed'; Details=$details }
        }
    } catch {
        Write-Log ("Unhandled error: $($_.Exception.Message)") 'ERROR'
        return [pscustomobject]@{ Outcome='Failed'; Details=$_.Exception.Message }
    }
}

end {
    # no-op
}

# Signed by Marinus van Deventer