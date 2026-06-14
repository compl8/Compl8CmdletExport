#region Content Explorer - Work Plan & Aggregate Queries

function New-ContentExplorerWorkPlan {
    <#
    .SYNOPSIS
        Runs aggregate queries to build a Content Explorer export work plan.
    .DESCRIPTION
        For each TagName+Workload combination, calls Export-ContentExplorerData
        with -Aggregate to get location-level counts. Results are written to
        the aggregate CSV for progress tracking and potential reuse.
        Includes retry logic for transient errors.
    .PARAMETER TagType
        The classifier type (e.g. SensitiveInformationType, Sensitivity).
    .PARAMETER TagNames
        Array of tag names to query.
    .PARAMETER Workloads
        Array of workloads to query.
    .PARAMETER AggregateCsvPath
        Path to the aggregate CSV file for writing results.
    .PARAMETER ExportRunDirectory
        Directory for the current export run.
    .OUTPUTS
        Hashtable with:
          Tasks              - Array of task objects with TagType, TagName, Workload,
                               ExpectedCount, ExportedCount, Locations, Status
          TotalExpectedRecords - Sum of all expected record counts
          HasErrors          - Boolean indicating if any queries failed
          ErrorTasks         - Array of "TagName|Workload" strings that failed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TagType,

        [Parameter(Mandatory)]
        [string[]]$TagNames,

        [Parameter(Mandatory)]
        [string[]]$Workloads,

        [string]$AggregateCsvPath,

        [string]$ExportRunDirectory
    )

    $workPlan = @{
        Tasks                = @()
        TotalExpectedRecords = 0
        HasErrors            = $false
        ErrorTasks           = @()
    }

    foreach ($tagName in $TagNames) {
        foreach ($workload in $Workloads) {
            Write-ExportLog -Message ("    Aggregate: " + $tagName + " / " + $workload) -Level Info

            $allAggregates = @()
            $aggError = $null
            $aggSuccess = $false

            try {
                # Shared aggregate pagination loop (BackoffHelper strategy:
                # Invoke-RetryWithBackoff -MaxRetries 3, Context "Aggregate: <name>/<workload>").
                # Cookie null/empty + same-cookie guards live inside the function.
                $allAggregates = @(Invoke-CEAggregatePaging `
                    -TagType $TagType -TagName $tagName -Workload $workload `
                    -PageSize 5000 -RetryMode 'BackoffHelper' `
                    -BackoffContext ("Aggregate: " + $tagName + "/" + $workload))

                $aggSuccess = $true
            }
            catch {
                $aggError = $_.Exception.Message
                Write-ExportLog -Message ("      AGGREGATE FAILED: " + $aggError) -Level Error
                $workPlan.HasErrors = $true
                $workPlan.ErrorTasks += ($tagName + "|" + $workload)
            }

            # Build task from aggregate results
            $locations = [System.Collections.ArrayList]::new()
            $totalCount = 0

            if ($aggSuccess) {
                foreach ($agg in $allAggregates) {
                    [void]$locations.Add(@{
                        Name          = $agg.Name
                        ExpectedCount = [int]$agg.Count
                        ExportedCount = 0
                    })
                    $totalCount += [int]$agg.Count
                }
                $displayCount = $totalCount.ToString('N0')
                $locationCount = $locations.Count
                Write-ExportLog -Message ("      -> " + $displayCount + " items in " + $locationCount + " locations") -Level Success
            }

            # Write to aggregate CSV (atomic append)
            if ($AggregateCsvPath) {
                $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

                if ($aggSuccess) {
                    $csvLines = [System.Collections.ArrayList]::new()
                    foreach ($agg in $allAggregates) {
                        $locationName = $agg.Name -replace '"', '""'
                        if ($locationName -match '[,"]') {
                            $locationName = '"' + $locationName + '"'
                        }
                        $line = $timestamp + "," + $TagType + "," + (ConvertTo-CsvField $tagName) + "," + $workload + "," + $locationName + "," + $agg.Count
                        [void]$csvLines.Add($line)
                    }

                    if ($csvLines.Count -gt 0) {
                        $csvContent = $csvLines -join [Environment]::NewLine
                        try {
                            [System.IO.File]::AppendAllText($AggregateCsvPath, $csvContent + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
                        }
                        catch {
                            Write-ExportLog -Message ("      Failed to append to aggregate CSV: " + $_.Exception.Message) -Level Warning
                        }
                    }
                }
                else {
                    # Write error row
                    $escapedError = $aggError -replace '"', '""'
                    $errorLine = $timestamp + "," + $TagType + "," + (ConvertTo-CsvField $tagName) + "," + $workload + ',ERROR,0,"' + $escapedError + '"'
                    try {
                        [System.IO.File]::AppendAllText($AggregateCsvPath, $errorLine + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
                    }
                    catch {
                        Write-ExportLog -Message ("      Failed to write error to aggregate CSV: " + $_.Exception.Message) -Level Warning
                    }
                }
            }

            # Add task to work plan
            $task = @{
                TagType       = $TagType
                TagName       = $tagName
                Workload      = $workload
                ExpectedCount = $totalCount
                ExportedCount = 0
                Locations     = @($locations)
                Status        = if ($aggSuccess) { "Pending" } else { "Error" }
                PageMetrics   = @()
                ResponseTimes = @()
            }

            if ($aggError) {
                $task.AggregateError = $aggError
            }

            $workPlan.Tasks += $task
            $workPlan.TotalExpectedRecords += $totalCount
        }
    }

    return $workPlan
}

#endregion

