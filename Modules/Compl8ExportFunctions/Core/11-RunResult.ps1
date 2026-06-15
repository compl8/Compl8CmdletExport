#region Run Result / Exit-Code Contract

# ─────────────────────────────────────────────────────────────────────────────
#  Exit-code map (B1 contract).
#  Consumers obtain codes via Get-ExportExitCode -Status <name>.
#  Code 1 is NOT listed here — it remains the unhandled-exception sentinel
#  already used by the top-level catch in MainExecution.ps1.
# ─────────────────────────────────────────────────────────────────────────────
$script:RunResultExitCodes = [ordered]@{
    Completed   = 0   # All sections OK, no remaining tasks
    Partial     = 2   # Some tasks incomplete / PartialFailure sections
    AuthFailed  = 3   # Authentication / token error (B3 will wire this)
    ConfigError = 4   # Configuration validation failure (B3 will wire this)
    Locked      = 5   # Export dir locked / already in use (future B4)
}

# RunSummary.json schema version. Bump when the emitted shape changes in a way
# consumers must branch on. Surfaced as the file's "schemaVersion" field.
$script:RunResultSchemaVersion = 1

function Get-ExportExitCode {
    <#
    .SYNOPSIS
        Returns the process exit code integer for a named export status.

    .DESCRIPTION
        Centralises the status → exit-code mapping so callers never hard-code
        integers.  All status names used by Write-RunSummary are valid inputs.

        Exit codes:
          0  Completed   — clean run, all sections succeeded
          1  (not in map) — reserved for unhandled / fatal exceptions (top-level catch)
          2  Partial     — one or more sections had partial failures or remaining tasks
          3  AuthFailed  — authentication / token error  (wired by B3)
          4  ConfigError — configuration validation failure (wired by B3)
          5  Locked      — export directory already in use (wired by B4)

    .PARAMETER Status
        One of: Completed, Partial, Failed, AuthFailed, ConfigError, Locked.
        'Failed' maps to 1 (reserved for the top-level catch; returned for
        completeness so callers can query it without hard-coding 1).

    .OUTPUTS
        [int] exit code.

    .EXAMPLE
        $code = Get-ExportExitCode -Status 'Partial'
        exit $code
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Completed', 'Partial', 'Failed', 'AuthFailed', 'ConfigError', 'Locked')]
        [string]$Status
    )

    if ($Status -eq 'Failed') { return 1 }
    return [int]$script:RunResultExitCodes[$Status]
}

function Write-RunSummary {
    <#
    .SYNOPSIS
        Writes a machine-readable RunSummary.json to the export run directory.

    .DESCRIPTION
        Captures the final state of an export run in a structured JSON file that
        can be consumed by schedulers, monitoring pipelines, and future unattended
        callers (B2).  The file is written atomically (temp + File.Move) so a
        reader never sees a partial file.

        Schema (schemaVersion = $script:RunResultSchemaVersion, currently 1):
          {
            "schemaVersion": 1,
            "startedUtc":    "2026-01-25T12:00:00.000Z",
            "endedUtc":      "2026-01-25T13:05:00.000Z",
            "mode":          "ContentExplorer",
            "exitCode":      0,
            "status":        "Completed",        // Completed | Partial | Failed | AuthFailed | ConfigError | Locked
            "sections":      [{ "name", "status", "recordCount", "errorCount" }],
            "remainingTasks": 0,
            "errors":        [{ "timestamp", "message" }],  // capped to last 20
            "droppedErrors": 0
          }

    .PARAMETER ExportDir
        Path to the export run directory.  Must exist (or be creatable) before
        this function is called.  The file is written as
        <ExportDir>\RunSummary.json.

    .PARAMETER Result
        Hashtable describing the outcome.  Expected keys:

          Mode            [string]    Export mode label (e.g. "ContentExplorer")
          Status          [string]    One of: Completed | Partial | Failed | AuthFailed | ConfigError | Locked
          StartedUtc      [datetime]  UTC start time (defaults to now if absent)
          Sections        [array]     Optional — array of @{ Name; Status; RecordCount; ErrorCount }
          RemainingTasks  [int]       Count of non-completed tasks (0 = all done)
          Errors          [array]     Optional — raw error entries from ExportStats

    .EXAMPLE
        Write-RunSummary -ExportDir $script:ExportRunDirectory -Result @{
            Mode           = $exportMode
            Status         = 'Completed'
            StartedUtc     = $script:SessionStartTime
            RemainingTasks = 0
            Errors         = @($stats.Errors)
        }

    .NOTES
        Failures to write RunSummary.json are non-fatal: a warning is logged and
        execution continues so the exit code is still set correctly even if the
        file could not be written.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir,

        [Parameter(Mandatory)]
        [hashtable]$Result
    )

    # ── Guard: ensure the target directory exists ──────────────────────────
    if (-not (Test-Path $ExportDir)) {
        try {
            New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null
        }
        catch {
            Write-ExportLog -Message ("Write-RunSummary: cannot create ExportDir '{0}': {1}" -f $ExportDir, $_.Exception.Message) -Level Warning
            return
        }
    }

    $summaryPath = Join-Path $ExportDir "RunSummary.json"

    try {
        # ── Times ────────────────────────────────────────────────────────────
        $endedUtc   = [datetime]::UtcNow
        $startedUtc = if ($Result.ContainsKey('StartedUtc') -and $Result.StartedUtc -is [datetime]) {
            [datetime]$Result.StartedUtc
        } else {
            $endedUtc
        }

        # ── Status & exit code ────────────────────────────────────────────────
        $statusValue = if ($Result.ContainsKey('Status') -and $Result.Status) {
            [string]$Result.Status
        } else {
            'Completed'
        }
        $exitCodeValue = Get-ExportExitCode -Status $statusValue

        # ── Mode ──────────────────────────────────────────────────────────────
        $modeValue = if ($Result.ContainsKey('Mode') -and $Result.Mode) { [string]$Result.Mode } else { '' }

        # ── Sections ──────────────────────────────────────────────────────────
        $rawSections = if ($Result.ContainsKey('Sections') -and $Result.Sections) {
            @($Result.Sections)
        } else {
            @()
        }
        $sectionsArray = @(
            $rawSections | ForEach-Object {
                [ordered]@{
                    name        = [string]($_.Name       ?? $_.name       ?? '')
                    status      = [string]($_.Status     ?? $_.status     ?? 'Unknown')
                    recordCount = [int]   ($_.RecordCount ?? $_.recordCount ?? 0)
                    errorCount  = [int]   ($_.ErrorCount  ?? $_.errorCount  ?? 0)
                }
            }
        )

        # ── Remaining tasks ───────────────────────────────────────────────────
        $remainingTasksValue = if ($Result.ContainsKey('RemainingTasks') -and $null -ne $Result.RemainingTasks) {
            [int]$Result.RemainingTasks
        } else {
            0
        }

        # ── Errors (capped to last 20) ────────────────────────────────────────
        $maxErrors  = 20
        $rawErrors  = if ($Result.ContainsKey('Errors') -and $Result.Errors) { @($Result.Errors) } else { @() }
        $dropped    = [Math]::Max(0, $rawErrors.Count - $maxErrors)
        $cappedErrors = @(
            ($rawErrors | Select-Object -Last $maxErrors) | ForEach-Object {
                $ts  = if ($_ -is [hashtable]) { $_.Timestamp } elseif ($_ -is [System.Collections.IDictionary]) { $_['Timestamp'] } else { $_.Timestamp }
                $msg = if ($_ -is [hashtable]) { $_.Message   } elseif ($_ -is [System.Collections.IDictionary]) { $_['Message']   } else { $_.Message   }
                [ordered]@{
                    timestamp = [string]($ts  ?? '')
                    message   = [string]($msg ?? '')
                }
            }
        )

        # ── Build summary object ──────────────────────────────────────────────
        $summary = [ordered]@{
            schemaVersion  = $script:RunResultSchemaVersion
            startedUtc     = $startedUtc.ToUniversalTime().ToString('o')
            endedUtc       = $endedUtc.ToUniversalTime().ToString('o')
            mode           = $modeValue
            exitCode       = $exitCodeValue
            status         = $statusValue
            sections       = $sectionsArray
            remainingTasks = $remainingTasksValue
            errors         = $cappedErrors
            droppedErrors  = $dropped
        }

        # ── Atomic write ──────────────────────────────────────────────────────
        $tmpPath = "${summaryPath}.tmp.${PID}"
        $json    = $summary | ConvertTo-Json -Depth 10
        Set-Content -Path $tmpPath -Value $json -Encoding UTF8
        [System.IO.File]::Move($tmpPath, $summaryPath, $true)

        Write-ExportLog -Message ("RunSummary written: status={0}, exitCode={1}, remainingTasks={2}" -f $statusValue, $exitCodeValue, $remainingTasksValue) -Level Info
    }
    catch {
        # Non-fatal: warn but do not throw — the exit code still gets set by the caller.
        Write-ExportLog -Message ("Write-RunSummary failed (non-fatal): {0}" -f $_.Exception.Message) -Level Warning
        # Clean up orphaned temp file if it exists
        $tmpPath = "${summaryPath}.tmp.${PID}"
        if (Test-Path $tmpPath) { Remove-Item -Path $tmpPath -Force -ErrorAction SilentlyContinue }
    }
}

#endregion
