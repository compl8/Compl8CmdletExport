#region Content Explorer Detail Export Parameter Builder

# Shared builder for the splat passed to Export-ContentExplorerWithProgress by the four
# single-terminal Content Explorer detail loops:
#   - Invoke-ContentExplorerExport       (fresh "Phase 7")
#   - Invoke-ContentExplorerResume       (single-terminal resume branch)
#   - Invoke-ContentExplorerRetry        (retry-bucket re-export)
#   - Invoke-ContentExplorerFromTasksCsv (single-terminal tasks-CSV branch)
#
# Each of those loops built a near-identical $exportParams hashtable inline: the same
# fixed keys (Task, ProgressLogPath, AdaptivePageSize=$true, TelemetryDatabasePath,
# OutputDirectory), the same SiteUrl/UserPrincipalName location-filter selection from the
# task's LocationType/Location, and a per-site-resolved PageSize. Only two things differed
# between sites and both are passed IN so behavior is preserved exactly:
#   - PageSize: each site computes its own value (per-task vs $cePageSize, and the fresh
#     loop's Error-status floor of max(500, $cePageSize)). The caller resolves it and
#     passes the result; this builder does not encode any page-size rule.
#   - Telemetry: the resume and tasks-CSV loops pass a New-ContentExplorerTelemetry object;
#     the fresh and retry loops pass none. To stay byte-equivalent to the loops that never
#     added the key, the Telemetry key is added ONLY when a non-null object is supplied.
#
# IMPORTANT (behavior preservation): this builder is for the single-terminal loops only.
# It deliberately does NOT add -CleanPriorPages — no single-terminal loop ever set it; the
# only caller of CleanPriorPages is the worker's auth-recovery retry path, which keeps its
# own inline param construction. The per-site bookkeeping each loop performs AFTER the
# export (status write-back, CSV flush cadence, run-tracker mutation, skip predicates) is
# intentionally left inline at each call site and is not the concern of this builder.

function Build-CEDetailExportParams {
    <#
    .SYNOPSIS
        Builds the hashtable to splat into Export-ContentExplorerWithProgress for a
        single-terminal Content Explorer detail task.
    .DESCRIPTION
        Returns a hashtable with the fixed keys shared by all four single-terminal detail
        loops plus the SiteUrl/UserPrincipalName location filter derived from the task's
        LocationType/Location. The page size and (optional) telemetry object are supplied
        by the caller so each loop's exact behavior is preserved.

        Location filter (identical to the inline logic it replaces):
          - LocationType -eq "SiteUrl" and a non-empty Location => adds SiteUrl
          - LocationType -eq "UPN"     and a non-empty Location => adds UserPrincipalName
          - otherwise (WorkloadFallback / no location)          => no location filter

        Telemetry: the Telemetry key is added only when -Telemetry is a non-null object, so
        callers that previously omitted it produce a splat with no Telemetry key (matching
        the fresh and retry loops), while callers that passed one produce a splat with the
        Telemetry key (matching the resume and tasks-CSV loops).
    .PARAMETER Task
        The detail task object. Must expose LocationType and Location for the location
        filter; passed through unchanged as the -Task argument.
    .PARAMETER PageSize
        The page size resolved by the caller (per-task value, $cePageSize, or the fresh
        loop's max(500, $cePageSize) Error floor). Passed through as -PageSize.
    .PARAMETER ProgressLogPath
        Path to the tailable progress log. Passed through as -ProgressLogPath.
    .PARAMETER TelemetryDatabasePath
        Path to the JSONL telemetry database. Passed through as -TelemetryDatabasePath.
    .PARAMETER OutputDirectory
        Classifier output directory (Get-CEClassifierDir result). Passed through as
        -OutputDirectory.
    .PARAMETER Telemetry
        Optional telemetry object from New-ContentExplorerTelemetry. When non-null, a
        Telemetry key is added to the returned hashtable; when null/omitted, no Telemetry
        key is added.
    .OUTPUTS
        Hashtable ready to splat into Export-ContentExplorerWithProgress.
    #>
    param(
        # Duck-typed detail task: PSCustomObject (from Read-TaskCsv) or hashtable;
        # reads .LocationType, .Location (and is forwarded whole as the Task key).
        [Parameter(Mandatory)]
        $Task,

        [Parameter(Mandatory)]
        [int]$PageSize,

        [string]$ProgressLogPath,

        [string]$TelemetryDatabasePath,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        $Telemetry = $null
    )

    $exportParams = @{
        Task                  = $Task
        PageSize              = $PageSize
        ProgressLogPath       = $ProgressLogPath
        AdaptivePageSize      = $true
        TelemetryDatabasePath = $TelemetryDatabasePath
        OutputDirectory       = $OutputDirectory
    }

    # Telemetry key added only when supplied — keeps the splat byte-equivalent to the
    # fresh/retry loops (which never added the key) and the resume/tasks-CSV loops (which did).
    if ($null -ne $Telemetry) {
        $exportParams["Telemetry"] = $Telemetry
    }

    # Location filter — identical selection to the inline blocks this replaces.
    if ($Task.LocationType -eq "SiteUrl" -and $Task.Location) {
        $exportParams["SiteUrl"] = $Task.Location
    }
    elseif ($Task.LocationType -eq "UPN" -and $Task.Location) {
        $exportParams["UserPrincipalName"] = $Task.Location
    }

    return $exportParams
}

#endregion
