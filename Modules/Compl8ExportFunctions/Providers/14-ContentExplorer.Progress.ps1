#region Content Explorer - Progress Tracking & Display

function Get-ContentExplorerAggregateProgress {
    <#
    .SYNOPSIS
        Calculates aggregate progress for Content Explorer export.
    .DESCRIPTION
        Reads the aggregate CSV to determine total expected records and tasks,
        then merges with completed task counts to calculate progress.
    .PARAMETER AggregateCsvPath
        Path to the ContentExplorer-Aggregates.csv file.
    .PARAMETER CompletedTasks
        Hashtable of completed task keys mapped to their exported record counts.
    .OUTPUTS
        Hashtable with:
          TotalExpected  - Total expected records from all tasks
          TotalExported  - Total records exported so far
          CompletedTasks - Number of completed tasks
          TotalTasks     - Total number of tasks
          Tasks          - Hashtable of task details keyed by "TagType|TagName|Workload"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$AggregateCsvPath,

        [hashtable]$CompletedTasks = @{}
    )

    $progress = @{
        TotalExpected  = 0
        TotalExported  = 0
        CompletedTasks = 0
        TotalTasks     = 0
        Tasks          = @{}
    }

    if (-not (Test-Path $AggregateCsvPath)) {
        return $progress
    }

    try {
        $csvData = Import-Csv -Path $AggregateCsvPath -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-ExportLog -Message ("Failed to read aggregate CSV for progress: " + $_.Exception.Message) -Level Warning
        return $progress
    }

    # Aggregate by task key
    $taskTotals = @{}
    foreach ($row in $csvData) {
        if ($row.Location -eq "ERROR") { continue }

        $taskKey = "{0}|{1}|{2}" -f $row.TagType, $row.TagName, $row.Workload

        if (-not $taskTotals.ContainsKey($taskKey)) {
            $taskTotals[$taskKey] = @{
                TagType       = $row.TagType
                TagName       = $row.TagName
                Workload      = $row.Workload
                ExpectedCount = 0
                ExportedCount = 0
                IsCompleted   = $false
            }
        }

        $count = $row.Count -as [int]
        if ($count) { $taskTotals[$taskKey].ExpectedCount += $count }
    }

    # Merge with completed task data
    foreach ($taskKey in $taskTotals.Keys) {
        $taskInfo = $taskTotals[$taskKey]
        $progress.TotalExpected += $taskInfo.ExpectedCount
        $progress.TotalTasks++

        if ($CompletedTasks.ContainsKey($taskKey)) {
            $exported = $CompletedTasks[$taskKey] -as [int]
            $taskInfo.ExportedCount = $exported
            $taskInfo.IsCompleted = $true
            $progress.TotalExported += $exported
            $progress.CompletedTasks++
        }
    }

    $progress.Tasks = $taskTotals
    return $progress
}

function Write-ContentExplorerProgress {
    <#
    .SYNOPSIS
        Displays formatted Content Explorer export progress.
    .DESCRIPTION
        Shows overall progress and highlights the current task being processed.
    .PARAMETER Progress
        Progress hashtable from Get-ContentExplorerAggregateProgress.
    .PARAMETER CurrentTaskKey
        The "TagType|TagName|Workload" key of the task currently being exported.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Progress,

        [string]$CurrentTaskKey
    )

    $totalTasks = $Progress.TotalTasks
    $completedTasks = $Progress.CompletedTasks
    $totalExpected = $Progress.TotalExpected
    $totalExported = $Progress.TotalExported

    $taskPct = if ($totalTasks -gt 0) { [Math]::Round(($completedTasks / $totalTasks) * 100, 0) } else { 0 }
    $recordPct = if ($totalExpected -gt 0) { [Math]::Round(($totalExported / $totalExpected) * 100, 1) } else { 0 }

    Write-ExportLog -Message ("  Progress: Tasks " + $completedTasks + "/" + $totalTasks + " (" + $taskPct + "%) | Records " + $totalExported.ToString('N0') + "/" + $totalExpected.ToString('N0') + " (" + $recordPct + "%)") -Level Info

    # Show current task info
    if ($CurrentTaskKey -and $Progress.Tasks.ContainsKey($CurrentTaskKey)) {
        $current = $Progress.Tasks[$CurrentTaskKey]
        Write-ExportLog -Message ("  Current: " + $current.TagName + " / " + $current.Workload + " (expected: " + $current.ExpectedCount.ToString('N0') + ")") -Level Info
    }
}

#endregion

