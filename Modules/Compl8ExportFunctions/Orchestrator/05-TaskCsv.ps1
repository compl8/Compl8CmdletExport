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

function Get-ContentExplorerLocationType {
    <#
    .SYNOPSIS
        Returns the Content Explorer location filter type for a workload.
    #>
    param(
        [Parameter(Mandatory)][string]$Workload
    )

    switch ($Workload) {
        'Exchange'   { return 'UPN' }
        'Teams'      { return 'UPN' }
        'SharePoint' { return 'SiteUrl' }
        'OneDrive'   { return 'SiteUrl' }
        default      { return 'WorkloadFallback' }
    }
}

function Get-ContentExplorerDetailPageSize {
    <#
    .SYNOPSIS
        Selects the page size used for a generated Content Explorer detail task.
    #>
    param(
        [int]$ExpectedCount,
        [int]$DefaultPageSize = 1000,
        [switch]$LocationScoped,
        [switch]$AggregateError
    )

    if ($LocationScoped) {
        if ($ExpectedCount -ge 10000) { return 5000 }
        if ($ExpectedCount -ge 4000) { return 2000 }
        if ($ExpectedCount -ge 500) { return 1000 }
        return 500
    }

    $floor = if ($AggregateError) { 500 } else { 100 }
    $pageSize = [Math]::Max($floor, $DefaultPageSize)
    if ($ExpectedCount -gt 0) {
        $maxPageSize = [Math]::Max($floor, 2 * $ExpectedCount)
        $pageSize = [Math]::Max($floor, [Math]::Min($pageSize, $maxPageSize))
    }
    return $pageSize
}

function New-ContentExplorerDetailTasks {
    <#
    .SYNOPSIS
        Expands aggregate/work-plan rows into executable Content Explorer detail tasks.
    .DESCRIPTION
        Generates location-level tasks when location data is available. For workloads in
        WorkloadFallbackWorkloads, or aggregate-error rows, generates one workload-level
        WorkloadFallback task and leaves Location empty so workers do not pass SiteUrl or
        UserPrincipalName filters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$WorkPlanTasks,
        [int]$DefaultPageSize = 1000,
        [string[]]$WorkloadFallbackWorkloads = @(),
        [switch]$Sort
    )

    $detailTasks = @()
    $fallbackLookup = @{}
    foreach ($workload in @($WorkloadFallbackWorkloads)) {
        if (-not [string]::IsNullOrWhiteSpace($workload)) {
            $fallbackLookup[$workload] = $true
        }
    }

    foreach ($task in @($WorkPlanTasks)) {
        if (-not $task) { continue }

        $tagType = $task.TagType
        $tagName = $task.TagName
        $workload = $task.Workload
        if ([string]::IsNullOrWhiteSpace($tagType) -or [string]::IsNullOrWhiteSpace($tagName) -or [string]::IsNullOrWhiteSpace($workload)) {
            continue
        }

        $locations = if ($task.Locations) { @($task.Locations) } else { @() }
        $expectedCount = 0
        if ($null -ne $task.TotalCount) {
            $expectedCount = $task.TotalCount -as [int]
        }
        elseif ($null -ne $task.ExpectedCount) {
            $expectedCount = $task.ExpectedCount -as [int]
        }
        if ($null -eq $expectedCount) { $expectedCount = 0 }

        $hasError = $false
        if ($task.HasError -eq $true -or $task.Status -eq 'Error' -or $task.AggregateError) {
            $hasError = $true
        }

        $locationType = Get-ContentExplorerLocationType -Workload $workload
        $useWorkloadFallback = $hasError -or $fallbackLookup.ContainsKey($workload) -or ($locationType -eq 'WorkloadFallback')

        if ($locations.Count -gt 0 -and -not $useWorkloadFallback) {
            foreach ($loc in $locations) {
                $locationName = $loc.Name
                $locExpected = $loc.ExpectedCount -as [int]
                if ($null -eq $locExpected -or $locExpected -le 0) { continue }

                $detailTasks += @{
                    Phase                 = 'Detail'
                    TagType               = $tagType
                    TagName               = $tagName
                    Workload              = $workload
                    Location              = $locationName
                    LocationType          = $locationType
                    ExpectedCount         = $locExpected
                    OriginalExpectedCount = $locExpected
                    PageSize              = Get-ContentExplorerDetailPageSize -ExpectedCount $locExpected -DefaultPageSize $DefaultPageSize -LocationScoped
                    Status                = 'Pending'
                    AssignedPID           = 0
                    ErrorMessage          = ''
                }
            }
            continue
        }

        if ($expectedCount -le 0 -and -not $hasError) { continue }

        $detailTasks += @{
            Phase                 = 'Detail'
            TagType               = $tagType
            TagName               = $tagName
            Workload              = $workload
            Location              = ''
            LocationType          = 'WorkloadFallback'
            ExpectedCount         = $expectedCount
            OriginalExpectedCount = $expectedCount
            PageSize              = Get-ContentExplorerDetailPageSize -ExpectedCount $expectedCount -DefaultPageSize $DefaultPageSize -AggregateError:$hasError
            Status                = 'Pending'
            AssignedPID           = 0
            ErrorMessage          = if ($hasError) { 'Aggregate failed' } else { '' }
        }
    }

    if ($Sort -and $detailTasks.Count -gt 1) {
        $detailTasks = @($detailTasks | Sort-Object { [int]$_.ExpectedCount } -Descending)
    }

    return @($detailTasks)
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
    [void]$sb.AppendLine("TagType,TagName,Workload,Location,LocationType,OriginalExpectedCount,ActualCount,DiscrepancyPct,PageSize")
    foreach ($task in $RetryTasks) {
        $escapedTag = ($task.TagName -replace '"','""')
        $location = if ($task.Location) { $task.Location } else { "" }
        $escapedLocation = ($location -replace '"','""')
        $locationType = if ($task.LocationType) { $task.LocationType } else { "WorkloadFallback" }
        $line = '{0},"{1}",{2},"{3}",{4},{5},{6},{7},{8}' -f $task.TagType, $escapedTag, $task.Workload, $escapedLocation, $locationType, ($task.OriginalExpectedCount -as [int]), ($task.ActualCount -as [int]), $task.DiscrepancyPct, ($task.PageSize -as [int])
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

