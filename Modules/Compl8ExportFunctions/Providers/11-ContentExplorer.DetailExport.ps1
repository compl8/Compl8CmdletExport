#region Content Explorer - Detail Export with Pagination

function Export-ContentExplorerWithProgress {
    <#
    .SYNOPSIS
        Main detail export with pagination, progress tracking, and retries.
    .DESCRIPTION
        Calls Export-ContentExplorerData in a pagination loop with:
        - Retry logic for transient and non-transient errors
        - PageCookie anomaly detection (new cookie but no records)
        - Progress logging to a tailable file
        - Adaptive page size selection
        - Partial error tracking in the Task object
    .PARAMETER Task
        Hashtable with TagType, TagName, Workload, ExpectedCount, Locations.
        Modified in place: ExportedCount, Status, TotalPages, TotalTimeMs, PartialErrors.
    .PARAMETER PageSize
        Base page size for queries. May be overridden by adaptive sizing.
    .PARAMETER ProgressLogPath
        Path to progress log file (tailable).
    .PARAMETER Telemetry
        Telemetry tracking object from New-ContentExplorerTelemetry.
    .PARAMETER TelemetryDatabasePath
        Path to the JSONL telemetry database file.
    .PARAMETER AdaptivePageSize
        If set, calls Get-AdaptivePageSize to select optimal page size.
    .PARAMETER OutputDirectory
        Directory where per-page JSON files are written. Each page creates a separate file
        named {Workload}-{NNN}.json (or {Workload}-{LocationHash}-{NNN}.json for location-filtered tasks).
        Each file contains: {PageNumber, ExportTimestamp, TagType, TagName, Workload, RecordCount, Records}.
    .OUTPUTS
        Exported record count (int).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Task,

        [int]$PageSize = 1000,

        [string]$ProgressLogPath,

        $Telemetry,

        [string]$TelemetryDatabasePath,

        [switch]$AdaptivePageSize,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [string]$SiteUrl,

        [string]$UserPrincipalName
    )

    # Early parameter validation
    if ([string]::IsNullOrWhiteSpace($Task.TagName)) {
        throw "TagName cannot be null or empty"
    }
    if ([string]::IsNullOrWhiteSpace($Task.TagType)) {
        throw "TagType cannot be null or empty"
    }
    if ([string]::IsNullOrWhiteSpace($Task.Workload)) {
        throw "Workload cannot be null or empty"
    }

    $tagType = $Task.TagType
    $tagName = $Task.TagName
    $workload = $Task.Workload
    $expectedCount = if ($Task.ExpectedCount) { $Task.ExpectedCount -as [int] } else { 0 }

    # Select page size
    $effectivePageSize = $PageSize
    if ($AdaptivePageSize -and $Task.Locations -and $Task.Locations.Count -gt 0) {
        try {
            $effectivePageSize = Get-AdaptivePageSize -Task $Task -Workload $workload -TelemetryDatabasePath $TelemetryDatabasePath
            Write-ExportLog -Message ("      Adaptive page size: " + $effectivePageSize + " (workload: " + $workload + ")") -Level Info
        }
        catch {
            Write-ExportLog -Message ("      Adaptive page size failed, using default " + $PageSize + ": " + $_.Exception.Message) -Level Warning
            $effectivePageSize = $PageSize
        }
    }

    # Clamp page size: floor 100, ceiling 2x expected count
    if ($expectedCount -gt 0) {
        $maxPageSize = [Math]::Max(100, 2 * $expectedCount)
        $unclamped = $effectivePageSize
        $effectivePageSize = [Math]::Max(100, [Math]::Min($effectivePageSize, $maxPageSize))
        if ($effectivePageSize -ne $unclamped) {
            Write-ExportLog -Message ("      Page size clamped: {0} -> {1} (expected: {2})" -f $unclamped, $effectivePageSize, $expectedCount) -Level Info
        }
    } else {
        $effectivePageSize = [Math]::Max(100, $effectivePageSize)
    }

    # Initialize task tracking
    $Task.Status = "InProgress"
    $Task.ExportedCount = 0
    $Task.TotalPages = 0
    $Task.TotalTimeMs = 0
    if (-not $Task.ContainsKey('PartialErrors')) { $Task.PartialErrors = @() }

    # Build per-page filename prefix
    $locationSuffix = ""
    if ($SiteUrl) { $locationSuffix = "-" + ([Math]::Abs($SiteUrl.GetHashCode())).ToString("X8") }
    elseif ($UserPrincipalName) { $locationSuffix = "-" + ([Math]::Abs($UserPrincipalName.GetHashCode())).ToString("X8") }
    $pageFilePrefix = "{0}{1}" -f $workload, $locationSuffix

    # Ensure output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }

    $pageCookie = $null
    $previousCookie = $null
    $pageNumber = 0
    $maxRetries = 3
    $transientDelaySec = 60
    $nonTransientDelaySec = 5
    $finalAttemptDelaySec = 120
    $emptyPageRetried = $false
    $isRetryingSamePage = $false
    $startTime = Get-Date

    # Log export start
    $locationSuffix = if ($SiteUrl) { " [Site:$SiteUrl]" } elseif ($UserPrincipalName) { " [User:$UserPrincipalName]" } else { "" }
    $logEntry = "[{0}] START {1}/{2}/{3}{4} Expected:{5} PageSize:{6}" -f
        (Get-Date).ToString("HH:mm:ss"), $tagType, $tagName, $workload, $locationSuffix, $expectedCount, $effectivePageSize
    Write-ProgressEntry -LogPath $ProgressLogPath -Message $logEntry

    # Populate telemetry location fields
    if ($Telemetry) {
        if ($SiteUrl) {
            $Telemetry.Location = $SiteUrl
            $Telemetry.LocationType = "SiteUrl"
        } elseif ($UserPrincipalName) {
            $Telemetry.Location = $UserPrincipalName
            $Telemetry.LocationType = "UPN"
        } else {
            $Telemetry.LocationType = "WorkloadFallback"
        }
    }

    # Collect per-page metrics for telemetry
    $pageMetrics = [System.Collections.ArrayList]::new()

    try {
        do {
            if ($isRetryingSamePage) {
                $isRetryingSamePage = $false  # Reset flag; keep same page number for this retry
            }
            else {
                $pageNumber++
            }
            $pageStartTime = Get-Date

            # Build export parameters
            $exportParams = @{
                TagType     = $tagType
                TagName     = $tagName
                Workload    = $workload
                PageSize    = $effectivePageSize
                ErrorAction = 'Stop'
            }
            if ($pageCookie) { $exportParams['PageCookie'] = $pageCookie }
            if ($SiteUrl) { $exportParams['SiteUrl'] = $SiteUrl }
            if ($UserPrincipalName) { $exportParams['UserPrincipalName'] = $UserPrincipalName }

            $pageSuccess = $false
            $retryCount = 0
            $result = $null

            # Retry loop for current page
            while (-not $pageSuccess -and $retryCount -le $maxRetries) {
                try {
                    $result = Export-ContentExplorerData @exportParams
                    $pageSuccess = $true
                }
                catch {
                    # Connection lost - cmdlet not available; no point retrying
                    if ($_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                        Write-ExportLog -Message ("      Page " + $pageNumber + " FATAL: S&C cmdlet not available - connection lost") -Level Error
                        $Task.PartialErrors += @{
                            Page         = $pageNumber
                            RetryCount   = 0
                            ErrorMessage = $_.Exception.Message
                            IsTransient  = $false
                            Timestamp    = (Get-Date).ToString("o")
                            PageCookie   = $pageCookie
                            Location     = if ($SiteUrl) { $SiteUrl } elseif ($UserPrincipalName) { $UserPrincipalName } else { "" }
                        }
                        throw
                    }

                    $retryCount++
                    $errMsg = $_.Exception.Message
                    $errorInfo = Get-HttpErrorExplanation -ErrorMessage $errMsg -ErrorRecord $_
                    $isTransient = $errorInfo.IsTransient

                    # Track partial error
                    $partialError = @{
                        Page         = $pageNumber
                        RetryCount   = $retryCount
                        ErrorMessage = $errMsg
                        IsTransient  = $isTransient
                        Timestamp    = (Get-Date).ToString("o")
                        PageCookie   = $pageCookie
                        Location     = if ($SiteUrl) { $SiteUrl } elseif ($UserPrincipalName) { $UserPrincipalName } else { "" }
                    }
                    $Task.PartialErrors += $partialError

                    # Auth error - throw immediately for caller to handle (before exhaustion check)
                    if ($errorInfo.Category -eq "AuthError") {
                        Write-ExportLog -Message ("      Page " + $pageNumber + " AUTH ERROR - throwing for caller") -Level Error
                        throw
                    }

                    if ($retryCount -gt $maxRetries) {
                        # All retries exhausted
                        Write-ExportLog -Message ("      Page " + $pageNumber + " FAILED after " + $maxRetries + " retries: " + $errMsg) -Level Error
                        $logMsg = "[{0}] FAIL Page {1} after {2} retries: {3}" -f (Get-Date).ToString("HH:mm:ss"), $pageNumber, $maxRetries, $errMsg
                        Write-ProgressEntry -LogPath $ProgressLogPath -Message $logMsg
                        break
                    }

                    $delay = Get-RetryDelay -AttemptNumber $retryCount -MaxRetries $maxRetries `
                        -IsTransient $isTransient -TransientDelaySeconds $transientDelaySec `
                        -NonTransientDelaySeconds $nonTransientDelaySec -FinalAttemptDelaySeconds $finalAttemptDelaySec

                    $levelPrefix = if ($isTransient) { "Transient error" } else { "Error" }
                    Write-ExportLog -Message ("      " + $levelPrefix + " page " + $pageNumber + " (attempt " + $retryCount + "/" + $maxRetries + ") - retrying in " + $delay + "s") -Level Warning
                    Start-Sleep -Seconds $delay
                }
            }

            # Page failed after all retries
            if (-not $pageSuccess) {
                $Task.Status = "PartialFailure"
                break
            }

            # Process page results
            if ($null -eq $result -or $result.Count -eq 0) {
                break
            }

            $metadata = $result[0]
            $recordsInPage = $metadata.RecordsReturned -as [int]
            if ($null -eq $recordsInPage) { $recordsInPage = 0 }

            # On first page, correct expected count from API metadata if available
            if ($pageNumber -eq 1 -and $null -ne $metadata.TotalCount) {
                $apiTotalCount = $metadata.TotalCount -as [int]
                if ($apiTotalCount -and $apiTotalCount -gt 0 -and $apiTotalCount -ne $expectedCount) {
                    Write-ExportLog -Message ("      Expected count corrected: {0} -> {1} (from API metadata)" -f $expectedCount, $apiTotalCount) -Level Info
                    $expectedCount = $apiTotalCount
                    $Task.ExpectedCount = $apiTotalCount
                }
            }

            if ($recordsInPage -gt 0) {
                $pageRecords = $result[1..$recordsInPage]

                # Add export metadata to each record
                foreach ($record in $pageRecords) {
                    if ($record -is [PSCustomObject]) {
                        $record | Add-Member -NotePropertyName '_ExportTagType' -NotePropertyValue $tagType -Force
                        $record | Add-Member -NotePropertyName '_ExportTagName' -NotePropertyValue $tagName -Force
                    }
                }

                # Write per-page file
                $pageFileName = "{0}-{1:D3}.json" -f $pageFilePrefix, $pageNumber
                $pageFilePath = Join-Path $OutputDirectory $pageFileName

                $pageData = @{
                    PageNumber      = $pageNumber
                    ExportTimestamp  = (Get-Date).ToString("o")
                    TagType         = $tagType
                    TagName         = $tagName
                    Workload        = $workload
                    RecordCount     = $recordsInPage
                    Records         = @($pageRecords)
                }

                try {
                    $serializablePage = ConvertTo-SerializableObject -InputObject $pageData
                    $pageJson = $serializablePage | ConvertTo-Json -Depth 20
                    Set-Content -Path $pageFilePath -Value $pageJson -Encoding UTF8
                }
                catch {
                    Write-ExportLog -Message ("      Page {0}: Failed to save page file: {1}" -f $pageNumber, $_.Exception.Message) -Level Error
                }

                $Task.ExportedCount += $recordsInPage
                $emptyPageRetried = $false
            }

            # Page timing
            $pageElapsed = ((Get-Date) - $pageStartTime).TotalMilliseconds
            $Task.TotalPages = $pageNumber

            # Record per-page metric
            [void]$pageMetrics.Add(@{
                PageNumber   = $pageNumber
                PageTimeMs   = [int]$pageElapsed
                RecordCount  = $recordsInPage
                RetryCount   = $retryCount
                Timestamp    = (Get-Date).ToString("o")
            })

            # Log progress
            $totalExported = $Task.ExportedCount
            $pctStr = if ($expectedCount -gt 0) { [Math]::Round(($totalExported / $expectedCount) * 100, 1).ToString() + "%" } else { "N/A" }
            $logMsg = "[{0}] Page {1}: +{2} records (Total: {3}/{4} = {5}) [{6}ms]" -f
                (Get-Date).ToString("HH:mm:ss"), $pageNumber, $recordsInPage, $totalExported, $expectedCount, $pctStr, [int]$pageElapsed
            Write-ProgressEntry -LogPath $ProgressLogPath -Message $logMsg

            # Check for more pages
            $morePagesAvailable = ($metadata.MorePagesAvailable -eq $true -or $metadata.MorePagesAvailable -eq "True")
            if (-not $morePagesAvailable) {
                break
            }

            # PageCookie tracking
            $newCookie = $metadata.PageCookie

            if ($recordsInPage -eq 0 -and $newCookie -and $newCookie -ne $pageCookie) {
                # New cookie received but no records - anomaly
                if (-not $emptyPageRetried) {
                    # Retry with previous cookie once (same page number, not a new page)
                    Write-ExportLog -Message ("      Page " + $pageNumber + ": New cookie but 0 records - retrying previous cookie (30s wait)") -Level Warning
                    Start-Sleep -Seconds 30
                    $emptyPageRetried = $true
                    $isRetryingSamePage = $true  # Signal loop to skip page increment on next iteration
                    # Keep using current pageCookie (previous cookie)
                    continue
                }
                else {
                    # Already retried, continue with new cookie
                    Write-ExportLog -Message ("      Page " + $pageNumber + ": Still 0 records after retry - continuing with new cookie") -Level Warning
                    $emptyPageRetried = $false
                }
            }

            $previousCookie = $pageCookie
            $pageCookie = $newCookie

        } while ($true)
    }
    catch {
        Write-ExportLog -Message ("      Export exception: " + $_.Exception.Message) -Level Error
        $Task.Status = "Failed"
    }

    # Finalize task
    $totalElapsed = ((Get-Date) - $startTime).TotalMilliseconds
    $Task.TotalTimeMs = [int]$totalElapsed

    if ($Task.Status -eq "InProgress") {
        $Task.Status = "Completed"
    }

    # Log completion
    $totalSec = [Math]::Round($totalElapsed / 1000, 1)
    $logMsg = "[{0}] END {1}/{2}/{3} Records:{4} Pages:{5} Time:{6}s Status:{7}" -f
        (Get-Date).ToString("HH:mm:ss"), $tagType, $tagName, $workload,
        $Task.ExportedCount, $Task.TotalPages, $totalSec, $Task.Status
    Write-ProgressEntry -LogPath $ProgressLogPath -Message $logMsg

    # Save telemetry
    if ($Telemetry) {
        $Telemetry.RecordCount = $Task.ExportedCount
        $Telemetry.PageCount = $Task.TotalPages
        $Telemetry.TotalTimeMs = $Task.TotalTimeMs
        $Telemetry.PageSize = $effectivePageSize
        $Telemetry.Status = $Task.Status
        $Telemetry.CompletedTime = (Get-Date).ToString("o")
        $Telemetry.PageMetrics = @($pageMetrics)

        if ($TelemetryDatabasePath) {
            Save-ContentExplorerTelemetry -Telemetry $Telemetry -DatabasePath $TelemetryDatabasePath
        }
    }

    # Write _task.json summary
    if ($Task.ExportedCount -gt 0) {
        $taskSummary = @{
            TagType        = $tagType
            TagName        = $tagName
            Workload       = $workload
            ExpectedCount  = $expectedCount
            ActualCount    = $Task.ExportedCount
            Pages          = $Task.TotalPages
            Status         = $Task.Status
            TotalTimeMs    = $Task.TotalTimeMs
            ExportDate     = (Get-Date).ToString("o")
            PageFilePrefix = $pageFilePrefix
        }
        if ($Task.PartialErrors.Count -gt 0) {
            $taskSummary.PartialErrors = @($Task.PartialErrors)
        }
        try {
            $taskJsonPath = Join-Path $OutputDirectory ("_task-{0}.json" -f $pageFilePrefix)
            $taskSummary | ConvertTo-Json -Depth 10 | Set-Content -Path $taskJsonPath -Encoding UTF8
        }
        catch {
            Write-ExportLog -Message ("      Failed to write _task summary: " + $_.Exception.Message) -Level Warning
        }
    }

    return $Task.ExportedCount
}

