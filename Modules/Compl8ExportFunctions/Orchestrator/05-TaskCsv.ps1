#region Task CSV I/O

function Write-AETaskCsv {
    <#
    .SYNOPSIS
        Writes an Activity Explorer day task CSV file (AEDayTasks.csv) atomically.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][array]$Tasks
    )
    $tmpPath = $Path + ".tmp.$PID"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("Day,StartTime,EndTime,AssignedPID,Status,PageCount,RecordCount,ErrorMessage")
    foreach ($task in $Tasks) {
        $escapedErr = if ($task.ErrorMessage) { ($task.ErrorMessage -replace '"','""') } else { "" }
        $line = '{0},{1},{2},{3},{4},{5},{6},"{7}"' -f $task.Day, $task.StartTime, $task.EndTime, ($task.AssignedPID -as [int]), $task.Status, ($task.PageCount -as [int]), ($task.RecordCount -as [int]), $escapedErr
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::Move($tmpPath, $Path, $true)
}

function Read-AETaskCsv {
    <#
    .SYNOPSIS
        Reads an Activity Explorer day task CSV file and returns task objects.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        return @(Import-Csv -Path $Path -Encoding UTF8)
    }
    catch {
        return @()
    }
}

function Write-TaskCsv {
    <#
    .SYNOPSIS
        Writes a task CSV file (AggregateTasks.csv or DetailTasks.csv) atomically.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][array]$Tasks
    )
    $tmpPath = $Path + ".tmp.$PID"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("TagType,TagName,Workload,Location,LocationType,ExpectedCount,PageSize,AssignedPID,Status,ErrorMessage,OriginalExpectedCount")
    foreach ($task in $Tasks) {
        $escapedTag = $task.TagName -replace '"','""'
        $escapedLoc = if ($task.Location) { ($task.Location -replace '"','""') } else { "" }
        $locType = if ($task.LocationType) { $task.LocationType } else { "" }
        $escapedErr = if ($task.ErrorMessage) { ($task.ErrorMessage -replace '"','""') } else { "" }
        $origExpected = if ($task.OriginalExpectedCount) { $task.OriginalExpectedCount -as [int] } else { $task.ExpectedCount -as [int] }
        $line = '{0},"{1}",{2},"{3}",{4},{5},{6},{7},{8},"{9}",{10}' -f $task.TagType, $escapedTag, $task.Workload, $escapedLoc, $locType, ($task.ExpectedCount -as [int]), ($task.PageSize -as [int]), ($task.AssignedPID -as [int]), $task.Status, $escapedErr, $origExpected
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::Move($tmpPath, $Path, $true)
}

function Read-TaskCsv {
    <#
    .SYNOPSIS
        Reads a task CSV file and returns task objects.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return @() }
    try {
        return @(Import-Csv -Path $Path -Encoding UTF8)
    }
    catch {
        return @()
    }
}

function Update-DetailTaskPageSizes {
    <#
    .SYNOPSIS
        Recalculates PageSize for pending location-level tasks in an existing DetailTasks.csv.
    .DESCRIPTION
        Updates PageSize for tasks that haven't started yet, using single-location sizing:
          <100 items -> 100, <1000 -> 500, <10000 -> 2000, else -> 5000.
        WorkloadFallback tasks are left unchanged. Already completed/in-progress tasks are untouched.
    .PARAMETER Path
        Path to the DetailTasks.csv file.
    .PARAMETER WhatIf
        If set, shows what would change without writing.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$WhatIf
    )

    if (-not (Test-Path $Path)) {
        Write-Host "DetailTasks.csv not found: $Path" -ForegroundColor Red
        return
    }

    $tasks = @(Import-Csv -Path $Path -Encoding UTF8)
    $updated = 0
    $skipped = 0

    foreach ($task in $tasks) {
        # Only update pending tasks with a location (not WorkloadFallback, not started)
        if ($task.Status -ne "Pending") { $skipped++; continue }
        if (-not $task.Location -or $task.LocationType -eq "WorkloadFallback") { $skipped++; continue }

        $expected = [int]$task.ExpectedCount
        $oldPageSize = [int]$task.PageSize

        # Single-location page sizing: no multi-folder overhead
        $newPageSize = if ($expected -lt 100) { 100 }
                       elseif ($expected -lt 1000) { 500 }
                       elseif ($expected -lt 10000) { 2000 }
                       else { 5000 }

        if ($newPageSize -ne $oldPageSize) {
            if ($WhatIf) {
                Write-Host ("  {0}/{1} [{2}]: {3} -> {4} (expected: {5:N0})" -f $task.TagName, $task.Workload, $task.Location.Substring(0, [Math]::Min(40, $task.Location.Length)), $oldPageSize, $newPageSize, $expected)
            }
            $task.PageSize = $newPageSize
            $updated++
        }
    }

    if ($WhatIf) {
        Write-Host "`n$updated tasks would be updated, $skipped skipped (completed/in-progress/fallback)" -ForegroundColor Cyan
    } else {
        Write-TaskCsv -Path $Path -Tasks $tasks
        Write-Host "$updated tasks updated, $skipped skipped" -ForegroundColor Green
    }
}

function Write-RetryTasksCsv {
    <#
    .SYNOPSIS
        Writes a RetryTasks.csv file from retry bucket task objects.
    .PARAMETER Path
        Output file path.
    .PARAMETER RetryTasks
        Array of retry task objects from Get-RetryBucketTasks.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][array]$RetryTasks
    )
    $tmpPath = $Path + ".tmp.$PID"
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("TagType,TagName,Workload,OriginalExpectedCount,ActualCount,DiscrepancyPct,PageSize")
    foreach ($task in $RetryTasks) {
        $escapedTag = ($task.TagName -replace '"','""')
        $line = '{0},"{1}",{2},{3},{4},{5},{6}' -f $task.TagType, $escapedTag, $task.Workload, ($task.OriginalExpectedCount -as [int]), ($task.ActualCount -as [int]), $task.DiscrepancyPct, ($task.PageSize -as [int])
        [void]$sb.AppendLine($line)
    }
    [System.IO.File]::WriteAllText($tmpPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    [System.IO.File]::Move($tmpPath, $Path, $true)
}

function Show-RetryBucketSummary {
    <#
    .SYNOPSIS
        Displays the retry bucket summary after combination phase.
    .PARAMETER RetryTasks
        Array of retry task objects from Get-RetryBucketTasks.
    .PARAMETER ExportDir
        Path to the export directory (for displaying the retry command).
    #>
    param(
        [array]$RetryTasks,
        [string]$ExportDir
    )
    Write-ExportLog -Message "`n--- Retry Bucket ---" -Level Info
    if ($RetryTasks -and $RetryTasks.Count -gt 0) {
        Write-ExportLog -Message ("  Tasks with >2%% discrepancy: {0}" -f $RetryTasks.Count) -Level Warning
        foreach ($rt in $RetryTasks) {
            $sign = if ($rt.DiscrepancyPct -ge 0) { "+" } else { "" }
            Write-ExportLog -Message ("    {0} / {1}: expected {2}, got {3} ({4}{5}%%)" -f $rt.TagName, $rt.Workload, $rt.OriginalExpectedCount.ToString('N0'), $rt.ActualCount.ToString('N0'), $sign, $rt.DiscrepancyPct) -Level Warning
        }
        Write-ExportLog -Message "  Retry file: RetryTasks.csv" -Level Info
        Write-ExportLog -Message ("  To retry: .\Export-Compl8Configuration.ps1 -CERetryDir ""{0}""" -f $ExportDir) -Level Info
    }
    else {
        Write-ExportLog -Message "  All tasks within 2%% tolerance" -Level Success
    }
}

function Write-RemainingTasksCsv {
    <#
    .SYNOPSIS
        Writes non-completed detail tasks to RemainingTasks.csv for follow-on runs.
    .PARAMETER ExportDir
        Path to the export run directory.
    .OUTPUTS
        Int - count of remaining tasks written (0 if all completed).
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir
    )
    $detailCsvPath = Join-Path (Get-CoordinationDir $ExportDir) "DetailTasks.csv"
    if (-not (Test-Path $detailCsvPath)) { return 0 }

    $allTasks = Read-TaskCsv -Path $detailCsvPath
    $remaining = @($allTasks | Where-Object { $_.Status -ne "Completed" })
    if ($remaining.Count -eq 0) { return 0 }

    $remainingPath = Join-Path (Get-CoordinationDir $ExportDir) "RemainingTasks.csv"
    Write-TaskCsv -Path $remainingPath -Tasks $remaining
    return $remaining.Count
}

#endregion

