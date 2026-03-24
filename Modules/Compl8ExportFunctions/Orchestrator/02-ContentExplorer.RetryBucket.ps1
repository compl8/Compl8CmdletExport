#region Content Explorer - Retry Bucket

function Get-RetryBucketTasks {
    <#
    .SYNOPSIS
        Identifies completed detail tasks where actual exported count differs from expected by more than a threshold.
    .DESCRIPTION
        Iterates completed detail tasks from a DetailTasks.csv and returns those where the discrepancy
        between OriginalExpectedCount and the actual exported count (stored in ExpectedCount after overwrite)
        exceeds the specified threshold. Skips tasks with errors, missing data, or very small counts.
    .PARAMETER DetailTasks
        Array of task objects from Read-TaskCsv (DetailTasks.csv).
    .PARAMETER Threshold
        Fractional threshold for discrepancy detection. Default 0.02 (2%).
    .PARAMETER MinCount
        Minimum OriginalExpectedCount to consider. Tasks below this are excluded
        because percentage-based thresholds are meaningless for tiny counts. Default 10.
    .OUTPUTS
        Array of PSCustomObjects with: TagType, TagName, Workload, OriginalExpectedCount, ActualCount, DiscrepancyPct, PageSize
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$DetailTasks,

        [double]$Threshold = 0.02,

        [int]$MinCount = 10
    )

    $retryTasks = @()

    foreach ($task in $DetailTasks) {
        # Only check completed tasks
        if ($task.Status -ne "Completed") { continue }

        # Skip tasks with error messages (aggregate errors, etc.)
        if ($task.ErrorMessage -and $task.ErrorMessage.Trim() -ne "") { continue }

        # Get original expected count - skip if column missing (backward compat) or zero
        $originalExpected = $task.OriginalExpectedCount -as [int]
        if (-not $originalExpected -or $originalExpected -eq 0) { continue }

        # Skip small tasks where percentage threshold is meaningless
        if ($originalExpected -lt $MinCount) { continue }

        # ActualCount is stored in ExpectedCount after the orchestrator overwrites it
        $actualCount = $task.ExpectedCount -as [int]
        if ($null -eq $actualCount) { $actualCount = 0 }

        # Compute discrepancy
        $discrepancy = [Math]::Abs($actualCount - $originalExpected) / $originalExpected

        if ($discrepancy -gt $Threshold) {
            $discrepancyPct = [Math]::Round(($actualCount - $originalExpected) / $originalExpected * 100, 1)
            $retryTasks += [PSCustomObject]@{
                TagType               = $task.TagType
                TagName               = $task.TagName
                Workload              = $task.Workload
                OriginalExpectedCount = $originalExpected
                ActualCount           = $actualCount
                DiscrepancyPct        = $discrepancyPct
                PageSize              = ($task.PageSize -as [int])
            }
        }
    }

    return $retryTasks
}

#endregion

