# ContentExplorer Retry — re-export tasks with >2% count discrepancy
function Invoke-ContentExplorerRetry {
    <#
    .SYNOPSIS
        Re-exports Content Explorer tasks that had >2% discrepancy between expected and actual counts.
    .DESCRIPTION
        Reads RetryTasks.csv from a previous export, re-runs detail export for those specific tasks
        (skipping aggregation), then re-evaluates retry buckets.
    .PARAMETER ExportDir
        Path to the export run directory containing RetryTasks.csv.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir
    )

    # Validate RetryTasks.csv exists
    $retryTasksCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "RetryTasks.csv"
    if (-not (Test-Path $retryTasksCsvPath)) {
        Write-Host "  ERROR: No RetryTasks.csv found in $ExportDir" -ForegroundColor Red
        return
    }

    # Load retry tasks
    $retryTasks = @(Import-Csv -Path $retryTasksCsvPath -Encoding UTF8)
    if ($retryTasks.Count -eq 0) {
        Write-Host "  No retry tasks found in RetryTasks.csv" -ForegroundColor Yellow
        return
    }

    # Display retry summary
    $retryLines = @(
        ("Export: {0}" -f (Split-Path $ExportDir -Leaf)),
        ("Tasks to retry: {0}" -f $retryTasks.Count)
    )
    foreach ($rt in $retryTasks) {
        $sign = if (($rt.DiscrepancyPct -as [double]) -ge 0) { "+" } else { "" }
        $retryLines += ("  {0} / {1}: {2}{3}%%" -f $rt.TagName, $rt.Workload, $sign, $rt.DiscrepancyPct)
    }
    Write-Host ""
    Write-Banner -Title 'CONTENT EXPLORER - RETRY MODE' -Lines $retryLines -Color 'Cyan'

    # Confirm
    Write-Host ""
    if (-not $script:Unattended) {
        $confirm = Read-Host "  Retry these tasks? [Y/N]"
        if ([string]::IsNullOrEmpty($confirm) -or $confirm.Trim().ToUpper() -ne "Y") {
            Write-Host "  Retry cancelled." -ForegroundColor Yellow
            return
        }
    } else {
        Write-ExportLog -Message "Unattended: proceeding with retry without confirmation (prompt D skipped)." -Level Info
    }

    # Initialize logging
    Initialize-ExportLog -LogDirectory (Get-LogsDir $ExportDir)
    $script:ExportRunDirectory = $ExportDir
    $script:SharedExportDirectory = $ExportDir
    $script:ErrorLogPath = Join-Path (Get-LogsDir $ExportDir) "ExportProject-Errors.log"

    Write-ExportLog -Message ("Retrying {0} discrepant tasks" -f $retryTasks.Count) -Level Info

    # Load Content Explorer page size (manifest overrides config file for consistency)
    $configPath = Join-Path $scriptRoot "ConfigFiles" "ContentExplorerClassifiers.json"
    $resolved = Resolve-CEPageSize -ExportRunDirectory $ExportDir -ConfigPath $configPath -FallbackPageSize $PageSize
    $cePageSize = $resolved.PageSize

    $progressLogPath = Join-Path (Get-LogsDir $ExportDir) "ContentExplorer-Progress.log"
    $aggregateCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "ContentExplorer-Aggregates.csv"
    $trackerPath = Join-Path (Get-CoordinationDir $ExportDir) "RunTracker.json"
    $tracker = Get-ContentExplorerRunTracker -TrackerPath $trackerPath
    $telemetryDbPath = Join-Path $scriptRoot "TelemetryDB" "content-explorer-telemetry.jsonl"

    # Build work plan tasks from retry CSV (reuse aggregate data, skip aggregation)
    $script:AllWorkPlanTasks = @()
    $workloads = @($script:CEDefaultWorkloads)

    foreach ($rt in $retryTasks) {
        $rtLocation = if ($rt.Location) { $rt.Location } else { "" }
        $rtLocationType = if ($rt.LocationType) { $rt.LocationType } else { "WorkloadFallback" }

        $script:AllWorkPlanTasks += @{
            TagType               = $rt.TagType
            TagName               = $rt.TagName
            Workload              = $rt.Workload
            Location              = $rtLocation
            LocationType          = $rtLocationType
            ExpectedCount         = ($rt.OriginalExpectedCount -as [int])
            OriginalExpectedCount = ($rt.OriginalExpectedCount -as [int])
            PageSize              = if ($rt.PageSize) { ($rt.PageSize -as [int]) } else { $cePageSize }
            ExportedCount         = 0
            Status                = "Pending"
        }
    }

    Write-ExportLog -Message ("  Built {0} retry tasks" -f $script:AllWorkPlanTasks.Count) -Level Info

    # Set phase to Detail
    Write-ExportPhase -ExportDir $ExportDir -Phase "Detail"

    # Run detail export for retry tasks (single-terminal)
    $completedTaskCounts = @{}

    foreach ($task in @($script:AllWorkPlanTasks)) {
        $taskLocation = if ($task.Location) { $task.Location } else { "" }
        $taskKey = "{0}|{1}|{2}|{3}" -f $task.TagType, $task.TagName, $task.Workload, $taskLocation

        if (-not $task.ExpectedCount -or $task.ExpectedCount -eq 0) {
            Write-ExportLog -Message ("    Skipping {0} / {1} - no data" -f $task.TagName, $task.Workload) -Level Info
            continue
        }

        $locationLabel = if ($taskLocation) { " @ $taskLocation" } else { "" }
        Write-ExportLog -Message ("    Retrying: {0} / {1}{2} (expected {3})" -f $task.TagName, $task.Workload, $locationLabel, $task.ExpectedCount) -Level Info

        # Export task with progress tracking — output to Data/ContentExplorer/TagType/TagName/
        $classifierDir = Get-CEClassifierDir $ExportDir $task.TagType $task.TagName

        $taskPageSize = if ($task.PageSize -and ($task.PageSize -as [int]) -gt 0) { ($task.PageSize -as [int]) } else { $cePageSize }
        # Single-terminal detail splat (retry): per-task page size, no telemetry object.
        $exportParams = Build-CEDetailExportParams -Task $task -PageSize $taskPageSize `
            -ProgressLogPath $progressLogPath -TelemetryDatabasePath $telemetryDbPath `
            -OutputDirectory $classifierDir

        try {
            Export-ContentExplorerWithProgress @exportParams | Out-Null
            $exportedCount = if ($task.ExportedCount) { $task.ExportedCount } else { 0 }
            $completedTaskCounts[$taskKey] = $exportedCount

            if ($exportedCount -gt 0) {
                Write-ExportLog -Message ("    Completed: {0} / {1}{2} - {3} records" -f $task.TagName, $task.Workload, $locationLabel, $exportedCount) -Level Success
            }
            else {
                Write-ExportLog -Message ("    Completed: {0} / {1}{2} - 0 records" -f $task.TagName, $task.Workload, $locationLabel) -Level Info
            }
        }
        catch {
            Write-ExportLog -Message ("    FAILED: {0} / {1}{2} - {3}" -f $task.TagName, $task.Workload, $locationLabel, $_.Exception.Message) -Level Error
            if ($script:ErrorLogPath) {
                Write-ExportErrorLog -ErrorLogPath $script:ErrorLogPath -Context "Retry Detail Export" -TaskKey $taskKey -ErrorRecord $_
            }
        }
    }

    # Re-evaluate retry bucket after re-export
    $detailCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
    $remainingRetryTasks = @()
    if (Test-Path $detailCsvPath) {
        # Update the DetailTasks.csv with new actual counts for retried tasks
        $detailTasks = Read-TaskCsv -Path $detailCsvPath
        foreach ($rt in $retryTasks) {
            $rtLocation = if ($rt.Location) { $rt.Location } else { "" }
            $taskKey = "{0}|{1}|{2}|{3}" -f $rt.TagType, $rt.TagName, $rt.Workload, $rtLocation
            if ($completedTaskCounts.ContainsKey($taskKey)) {
                $matchTask = $detailTasks | Where-Object {
                    $_.TagType -eq $rt.TagType -and $_.TagName -eq $rt.TagName -and $_.Workload -eq $rt.Workload -and $(if ($_.Location) { $_.Location } else { "" }) -eq $rtLocation
                } | Select-Object -First 1
                if ($matchTask) {
                    $matchTask.ExpectedCount = $completedTaskCounts[$taskKey]
                }
            }
        }
        Write-TaskCsv -Path $detailCsvPath -Tasks $detailTasks
        $remainingRetryTasks = @(Get-RetryBucketTasks -DetailTasks $detailTasks)
    }

    # Update RetryTasks.csv
    if ($remainingRetryTasks.Count -gt 0) {
        Write-RetryTasksCsv -Path $retryTasksCsvPath -RetryTasks $remainingRetryTasks
    }
    else {
        # All tasks now pass - remove retry CSV
        Remove-Item -Path $retryTasksCsvPath -Force -ErrorAction SilentlyContinue
    }

    # Display final retry summary
    Show-RetryBucketSummary -RetryTasks $remainingRetryTasks -ExportDir $ExportDir

    # Write remaining (non-completed) tasks for follow-on runs
    $remainingCount = Write-RemainingTasksCsv -ExportDir $ExportDir
    if ($remainingCount -gt 0) {
        Write-ExportLog -Message ("  Remaining tasks: {0} (see RemainingTasks.csv)" -f $remainingCount) -Level Warning
        Write-ExportLog -Message ("  To re-run: .\Export-Compl8Configuration.ps1 -CETasksCsv ""{0}""" -f (Join-Path (Get-CoordinationDir $ExportDir) "RemainingTasks.csv")) -Level Info
    }

    # Write manifest and set final phase
    Write-CEManifest -ExportDir $ExportDir
    Write-ExportPhase -ExportDir $ExportDir -Phase "Completed"
    Write-ExportLog -Message "Retry complete" -Level Success
}
