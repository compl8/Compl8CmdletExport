#region Activity Explorer Export Functions

function Export-ActivityExplorerWithProgress {
    <#
    .SYNOPSIS
        Resilient per-page Activity Explorer export with progress tracking and resume capability.
    .DESCRIPTION
        Exports Activity Explorer data page by page, saving each page immediately to disk.
        Supports resume from last successful page after failures or interruptions.
        Includes retry logic for transient errors and automatic end-time adjustment for
        "future dates" API errors.

    .PARAMETER StartTime
        Start of the date range to export (UTC).
    .PARAMETER EndTime
        End of the date range to export (UTC).
    .PARAMETER PageSize
        Records per page (1-5000). Default: 5000
    .PARAMETER Filters
        Hashtable of filters (Activity, Workload, etc.).
    .PARAMETER OutputDirectory
        Path to the ActivityExplorer output subfolder.
    .PARAMETER Tracker
        Run tracker hashtable for state persistence.
    .PARAMETER TrackerPath
        File path where the run tracker is saved.
    .PARAMETER ProgressLogPath
        File path for the tailable progress log.
    .PARAMETER Resume
        Switch to enable resume from last successful page.

    .OUTPUTS
        Hashtable with TotalRecords, PageCount, and optionally ResumedFrom.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$StartTime,

        [Parameter(Mandatory)]
        [datetime]$EndTime,

        [ValidateRange(1, 5000)]
        [int]$PageSize = 5000,

        [hashtable]$Filters,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [Parameter(Mandatory)]
        [hashtable]$Tracker,

        [Parameter(Mandatory)]
        [string]$TrackerPath,

        [string]$ProgressLogPath,

        [switch]$Resume
    )

    # Early parameter validation
    if ($StartTime -ge $EndTime) {
        throw "StartTime ($($StartTime.ToString('o'))) must be before EndTime ($($EndTime.ToString('o')))"
    }

    # Ensure output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }

    # Format dates for the cmdlet (MM/dd/yyyy HH:mm:ss format)
    $startStr = $StartTime.ToString("MM/dd/yyyy HH:mm:ss")
    $currentEndTime = $EndTime
    $endStr = $currentEndTime.ToString("MM/dd/yyyy HH:mm:ss")

    # Build base export parameters
    $exportParams = @{
        StartTime    = $startStr
        EndTime      = $endStr
        OutputFormat = "Json"
        PageSize     = $PageSize
        ErrorAction  = 'Stop'
    }

    # Add filters (up to 5 supported by the API)
    if ($Filters -and $Filters.Count -gt 0) {
        if ($Filters.Count -gt 5) {
            Write-Warning "Activity Explorer API supports max 5 filters per category. $($Filters.Count) items enabled — only the first 5 will be used. Disable some items in the config file."
        }
        $filterIndex = 1
        foreach ($filterName in $Filters.Keys) {
            if ($filterIndex -le 5) {
                $filterValues = @($filterName) + @($Filters[$filterName])
                $exportParams["Filter$filterIndex"] = $filterValues
                $filterIndex++
            }
        }
    }

    # Initialize tracking variables
    $pageNumber = 0
    $totalRecords = 0
    $resumedFrom = $null
    $previousWaterMark = $null

    # Handle resume
    if ($Resume -and $Tracker.LastWaterMark) {
        $pageNumber = $Tracker.CompletedPages
        $totalRecords = $Tracker.TotalRecords
        $exportParams['PageCookie'] = $Tracker.LastWaterMark
        $previousWaterMark = $Tracker.LastWaterMark

        $resumedFrom = @{
            PageNumber  = $pageNumber
            RecordCount = $totalRecords
        }

        $msg = "  RESUMING from page {0} ({1} records already exported)" -f $pageNumber, $totalRecords
        Write-ExportLog -Message $msg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
    }
    else {
        $msg = "  Starting Activity Explorer export: {0} to {1}" -f $startStr, $endStr
        Write-ExportLog -Message $msg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
    }

    # Initial query with future-dates retry
    $result = $null
    $futureDateRetries = 0
    $maxFutureDateRetries = 3

    while ($null -eq $result -and $futureDateRetries -le $maxFutureDateRetries) {
        try {
            $result = Export-ActivityExplorerData @exportParams
        }
        catch {
            $errorMsg = Get-PageErrorMessage -ErrorRecord $_
            $isFutureDate = $errorMsg -match "future" -and $errorMsg -match "date"

            if ($isFutureDate -and $futureDateRetries -lt $maxFutureDateRetries) {
                $futureDateRetries++
                $currentEndTime = $currentEndTime.AddHours(-1)
                $endStr = $currentEndTime.ToString("MM/dd/yyyy HH:mm:ss")
                $exportParams['EndTime'] = $endStr

                $msg = "  Future date error - reducing end time by 1 hour (attempt {0}/{1}): {2}" -f $futureDateRetries, $maxFutureDateRetries, $endStr
                Write-ExportLog -Message $msg -Level Warning
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg

                Add-PartialError -Tracker $Tracker -PageNumber 0 -ErrorMessage $errorMsg -ErrorType "FutureDateRetry"
            }
            else {
                # Check for auth error
                $initialErrorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMsg -ErrorRecord $_
                if ($initialErrorInfo.Category -eq "AuthError") {
                    Write-ExportLog -Message "  AUTH ERROR - throwing for caller to handle" -Level Error
                    throw
                }

                # Non-recoverable error on initial query
                $msg = "  Initial query failed: {0}" -f $errorMsg
                Write-ExportLog -Message $msg -Level Error
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                Add-PartialError -Tracker $Tracker -PageNumber 0 -ErrorMessage $errorMsg -ErrorType "InitialQueryFailed"

                $Tracker['Status'] = "Failed"
                $Tracker['PartialFailure'] = $true
                Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

                return @{
                    TotalRecords = $totalRecords
                    PageCount    = $pageNumber
                    ResumedFrom  = $resumedFrom
                }
            }
        }
    }

    # Check if we got any data
    if ($null -eq $result) {
        Write-ExportLog -Message "  No results returned after future-date retries" -Level Warning
        $Tracker['Status'] = "Completed"
        Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath
        return @{
            TotalRecords = 0
            PageCount    = 0
            ResumedFrom  = $resumedFrom
        }
    }

    # Log total result count
    $totalAvailable = $result.TotalResultCount
    if ($totalAvailable) {
        $msg = "  Total activities available: {0}" -f $totalAvailable
        Write-ExportLog -Message $msg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
        $Tracker['TotalAvailable'] = $totalAvailable -as [long]
        Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath
    }

    # Page processing loop
    $exportStartTime = Get-Date
    $done = $false

    while (-not $done) {
        $pageNumber++
        $pageStartTime = Get-Date

        # Parse records from this page
        $pageRecords = @()
        $hasContent = Test-PageHasContent -Result $result

        if ($hasContent) {
            try {
                $parsed = $result.ResultData | ConvertFrom-Json
                $pageRecords = @($parsed)
            }
            catch {
                $msg = "  Page {0}: Failed to parse ResultData: {1}" -f $pageNumber, $_.Exception.Message
                Write-ExportLog -Message $msg -Level Warning
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
            }
        }

        $recordCount = $pageRecords.Count

        if ($recordCount -gt 0) {
            # Calculate record time range
            $recordTimeRange = @{
                Earliest = $null
                Latest   = $null
            }

            try {
                $timestamps = @($pageRecords | Where-Object { $_.Happened } | ForEach-Object {
                    if ($_.Happened -is [datetime]) { $_.Happened } else { [DateTime]::Parse($_.Happened, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind) }
                })
                if ($timestamps.Count -gt 0) {
                    $recordTimeRange.Earliest = ($timestamps | Measure-Object -Minimum).Minimum.ToString("o")
                    $recordTimeRange.Latest = ($timestamps | Measure-Object -Maximum).Maximum.ToString("o")
                }
            }
            catch {
                # Timestamp parsing is optional for progress display - non-critical
                Write-Verbose "Timestamp parsing failed for page $pageNumber : $($_.Exception.Message)"
            }

            # Save page file immediately (JSON array by default, JSONL when COMPL8_JSONL_OUTPUT=1)
            $useJsonl = $env:COMPL8_JSONL_OUTPUT -eq "1"
            $pageExt = if ($useJsonl) { "jsonl" } else { "json" }
            $pageFileName = "Page-{0:D3}.{1}" -f $pageNumber, $pageExt
            $pageFilePath = Join-Path $OutputDirectory $pageFileName

            try {
                if ($useJsonl) {
                    $sb = [System.Text.StringBuilder]::new()
                    foreach ($rec in $pageRecords) {
                        $serRec = ConvertTo-SerializableObject -InputObject $rec
                        [void]$sb.AppendLine(($serRec | ConvertTo-Json -Depth 20 -Compress))
                    }
                    [System.IO.File]::WriteAllText($pageFilePath, $sb.ToString(), [System.Text.Encoding]::UTF8)
                }
                else {
                    $pageData = @{
                        PageNumber      = $pageNumber
                        ExportTimestamp = (Get-Date).ToString("o")
                        RecordTimeRange = $recordTimeRange
                        RecordCount     = $recordCount
                        WaterMark       = $result.WaterMark
                        Records         = $pageRecords
                    }
                    $serializablePage = ConvertTo-SerializableObject -InputObject $pageData
                    $pageJson = $serializablePage | ConvertTo-Json -Depth 20
                    Set-Content -Path $pageFilePath -Value $pageJson -Encoding UTF8
                }
            }
            catch {
                $msg = "  Page {0}: Failed to save page file: {1}" -f $pageNumber, $_.Exception.Message
                Write-ExportLog -Message $msg -Level Error
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
            }

            $totalRecords += $recordCount
        }

        # Update tracker
        $pageElapsed = (Get-Date) - $pageStartTime
        $Tracker['CompletedPages'] = $pageNumber
        $Tracker['TotalRecords'] = $totalRecords
        $Tracker['LastWaterMark'] = $result.WaterMark
        $Tracker['LastPageTime'] = $pageElapsed.TotalSeconds

        Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

        # Progress logging
        $totalElapsed = (Get-Date) - $exportStartTime
        $pagesPerMin = if ($totalElapsed.TotalMinutes -gt 0) {
            [math]::Round($pageNumber / $totalElapsed.TotalMinutes, 1)
        }
        else { 0 }

        $pctMsg = ""
        if ($totalAvailable -and $totalAvailable -gt 0) {
            $pct = [Math]::Round(($totalRecords / $totalAvailable) * 100, 1)
            $pctMsg = ", {0}% of {1:N0}" -f $pct, $totalAvailable
        }
        $progressMsg = "  Page {0}: {1} records ({2:N0} total{3}, {4} pages/min)" -f $pageNumber, $recordCount, $totalRecords, $pctMsg, $pagesPerMin
        Write-ExportLog -Message $progressMsg -Level Info
        Write-ProgressEntry -Path $ProgressLogPath -Message $progressMsg

        # Check if this is the last page
        if ($result.LastPage -eq $true) {
            $done = $true
            continue
        }

        # Get next page with retry logic
        $newWaterMark = $result.WaterMark
        $previousWaterMark = if ($exportParams.ContainsKey('PageCookie')) { $exportParams['PageCookie'] } else { $null }
        $exportParams['PageCookie'] = $newWaterMark

        $nextResult = $null
        $retryCount = 0
        $maxRetries = 3
        $sameCookieRetries = 0

        while ($null -eq $nextResult) {
            try {
                $nextResult = Export-ActivityExplorerData @exportParams
            }
            catch {
                $errorMsg = Get-PageErrorMessage -ErrorRecord $_
                $errorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMsg -ErrorRecord $_

                # Auth error - throw immediately for caller to handle
                if ($errorInfo.Category -eq "AuthError") {
                    Write-ExportLog -Message "  AUTH ERROR on page retrieval - throwing for caller to handle" -Level Error
                    throw
                }

                # Future dates error on subsequent pages
                $isFutureDate = $errorMsg -match "future" -and $errorMsg -match "date"
                if ($isFutureDate) {
                    $futureDateRetries++
                    if ($futureDateRetries -le $maxFutureDateRetries) {
                        $currentEndTime = $currentEndTime.AddHours(-1)
                        $endStr = $currentEndTime.ToString("MM/dd/yyyy HH:mm:ss")
                        $exportParams['EndTime'] = $endStr

                        $msg = "  Future date error on page {0} - reducing end time: {1}" -f ($pageNumber + 1), $endStr
                        Write-ExportLog -Message $msg -Level Warning
                        Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                        Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) -ErrorMessage $errorMsg -ErrorType "FutureDateRetry"
                        continue
                    }
                }

                $retryCount++
                Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) -ErrorMessage $errorMsg -ErrorType "PageRetrievalError"

                if ($retryCount -gt $maxRetries) {
                    # All retries exhausted - save what we have
                    $msg = "  Page {0}: All {1} retries exhausted. Saving progress." -f ($pageNumber + 1), $maxRetries
                    Write-ExportLog -Message $msg -Level Error
                    Write-ProgressEntry -Path $ProgressLogPath -Message $msg

                    $Tracker['PartialFailure'] = $true
                    $Tracker['Status'] = "PartialFailure"
                    Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

                    return @{
                        TotalRecords = $totalRecords
                        PageCount    = $pageNumber
                        ResumedFrom  = $resumedFrom
                    }
                }

                $delay = Get-RetryDelay -AttemptNumber $retryCount -MaxRetries $maxRetries `
                    -IsTransient $true -ScaleByAttempt
                $msg = "  Page {0}: Retry {1}/{2} - waiting {3}s" -f ($pageNumber + 1), $retryCount, $maxRetries, $delay
                Write-ExportLog -Message $msg -Level Warning
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                Start-Sleep -Seconds $delay
            }
        }

        # Analyze the new result
        $nextHasContent = Test-PageHasContent -Result $nextResult
        $nextWaterMark = $nextResult.WaterMark
        $cookieChanged = ($nextWaterMark -ne $newWaterMark)

        if ($cookieChanged -and $nextHasContent) {
            # Normal success: new cookie with content - continue
            $result = $nextResult
        }
        elseif ($cookieChanged -and -not $nextHasContent) {
            # New cookie but no content: retry previous cookie once
            $msg = "  Page {0}: New PageCookie but no content - retrying previous cookie (30s wait)" -f ($pageNumber + 1)
            Write-ExportLog -Message $msg -Level Warning
            Write-ProgressEntry -Path $ProgressLogPath -Message $msg

            Start-Sleep -Seconds 30

            # Retry with the previous (working) cookie
            $retryParams = $exportParams.Clone()
            $retryParams['PageCookie'] = $newWaterMark

            try {
                $retryResult = Export-ActivityExplorerData @retryParams
                $retryHasContent = Test-PageHasContent -Result $retryResult

                if ($retryHasContent) {
                    # Previous cookie now has content
                    $result = $retryResult
                }
                else {
                    # Still no content - continue with the new cookie regardless
                    $msg = "  Page {0}: Retry also returned no content - continuing with new cookie" -f ($pageNumber + 1)
                    Write-ExportLog -Message $msg -Level Warning
                    Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                    $result = $nextResult
                }
            }
            catch {
                # Retry failed - continue with new cookie
                $msg = "  Page {0}: Previous cookie retry failed - continuing with new cookie" -f ($pageNumber + 1)
                Write-ExportLog -Message $msg -Level Warning
                $result = $nextResult
            }
        }
        elseif (-not $cookieChanged) {
            # Same cookie returned - retry with delays
            $sameCookieRetries++

            if ($sameCookieRetries -gt $maxRetries) {
                # Same cookie exhausted all retries
                $msg = "  Page {0}: Same PageCookie returned {1} times - saving progress" -f ($pageNumber + 1), $sameCookieRetries
                Write-ExportLog -Message $msg -Level Error
                Write-ProgressEntry -Path $ProgressLogPath -Message $msg

                Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) `
                    -ErrorMessage "Same PageCookie returned after $sameCookieRetries retries" `
                    -ErrorType "SameCookieRetryExhausted"

                if ($sameCookieRetries -eq $maxRetries + 1) {
                    # One final attempt after 120s
                    $msg = "  Page {0}: Final attempt after 120s wait" -f ($pageNumber + 1)
                    Write-ExportLog -Message $msg -Level Warning
                    Write-ProgressEntry -Path $ProgressLogPath -Message $msg
                    Start-Sleep -Seconds 120

                    try {
                        $finalResult = Export-ActivityExplorerData @exportParams
                        $finalCookie = $finalResult.WaterMark
                        if ($finalCookie -ne $newWaterMark -or (Test-PageHasContent -Result $finalResult)) {
                            $result = $finalResult
                            $sameCookieRetries = 0
                            continue
                        }
                    }
                    catch {
                        # Final attempt also failed - will fall through to PartialFailure handling below
                        Write-ExportLog -Message ("  Page {0}: Final attempt also failed: {1}" -f ($pageNumber + 1), $_.Exception.Message) -Level Warning
                    }
                }

                $Tracker['PartialFailure'] = $true
                $Tracker['Status'] = "PartialFailure"
                Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

                return @{
                    TotalRecords = $totalRecords
                    PageCount    = $pageNumber
                    ResumedFrom  = $resumedFrom
                }
            }

            # Retry with 60s delay
            $delay = 60
            $msg = "  Page {0}: Same PageCookie (retry {1}/{2}) - waiting {3}s" -f ($pageNumber + 1), $sameCookieRetries, $maxRetries, $delay
            Write-ExportLog -Message $msg -Level Warning
            Write-ProgressEntry -Path $ProgressLogPath -Message $msg
            Start-Sleep -Seconds $delay

            try {
                $retryResult = Export-ActivityExplorerData @exportParams
                $result = $retryResult
                # Reset same-cookie counter if cookie changed
                if ($retryResult.WaterMark -ne $newWaterMark) {
                    $sameCookieRetries = 0
                }
            }
            catch {
                $errorMsg = Get-PageErrorMessage -ErrorRecord $_
                Add-PartialError -Tracker $Tracker -PageNumber ($pageNumber + 1) -ErrorMessage $errorMsg -ErrorType "SameCookieRetryError"

                # Auth error - throw
                $sameCookieErrorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMsg -ErrorRecord $_
                if ($sameCookieErrorInfo.Category -eq "AuthError") {
                    throw
                }

                # Continue retry loop
                $result = $nextResult
            }
        }
        else {
            # Fallback: use whatever we got
            $result = $nextResult
        }
    }

    # Export completed successfully
    $totalElapsed = (Get-Date) - $exportStartTime
    $Tracker['Status'] = "Completed"
    $Tracker['EndTime'] = (Get-Date).ToString("o")
    $Tracker['Duration'] = $totalElapsed.ToString()
    Save-ActivityExplorerRunTracker -Tracker $Tracker -TrackerPath $TrackerPath

    $completionMsg = "  Export completed: {0} records in {1} pages ({2:N1} minutes)" -f $totalRecords, $pageNumber, $totalElapsed.TotalMinutes
    Write-ExportLog -Message $completionMsg -Level Success
    Write-ProgressEntry -Path $ProgressLogPath -Message $completionMsg

    return @{
        TotalRecords = $totalRecords
        PageCount    = $pageNumber
        ResumedFrom  = $resumedFrom
    }
}

function Merge-ActivityExplorerPages {
    <#
    .SYNOPSIS
        Merges and deduplicates Activity Explorer page files from multi-terminal worker directories.
    .DESCRIPTION
        Scans Data/ActivityExplorer/YYYY-MM-DD/Page-*.json files across all day directories,
        deduplicates records by RecordIdentity, and writes a combined output file.
        Uses streaming JSON write for large datasets (>50k records) to minimize memory.
    .PARAMETER ExportDirectory
        Root export directory containing Data/ActivityExplorer/ subdirectories.
    .PARAMETER OutputPath
        Path for the combined output file. Defaults to ExportDirectory/Data/ActivityExplorer/ActivityExplorer-Combined.json.
    .PARAMETER StreamingOutput
        Force streaming output mode regardless of record count.
    .OUTPUTS
        Hashtable with TotalRecords, UniqueRecords, DuplicatesRemoved, PagesProcessed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExportDirectory,

        [string]$OutputPath,

        [switch]$StreamingOutput
    )

    if (-not $OutputPath) {
        $OutputPath = Join-Path (Get-AEDataDir $ExportDirectory) "ActivityExplorer-Combined.json"
    }

    # Find all page files in Data/ActivityExplorer/ day directories
    $aeDataDir = Get-AEDataDir $ExportDirectory
    $pageFiles = @(Get-ChildItem -Path $aeDataDir -Recurse -Filter "Page-*.json" -ErrorAction SilentlyContinue |
        Sort-Object FullName)

    if ($pageFiles.Count -eq 0) {
        Write-Verbose "No AE page files found in worker directories"
        return @{
            TotalRecords      = 0
            UniqueRecords     = 0
            DuplicatesRemoved = 0
            PagesProcessed    = 0
        }
    }

    Write-Verbose ("Found {0} page file(s) across worker directories" -f $pageFiles.Count)

    # First pass: collect all records and dedup by RecordIdentity
    $seenIdentities = @{}
    $allRecords = [System.Collections.ArrayList]::new()
    $totalRecords = 0
    $pagesProcessed = 0

    foreach ($pageFile in $pageFiles) {
        try {
            $pageData = Get-Content -Raw -Path $pageFile.FullName -ErrorAction Stop | ConvertFrom-Json
            if ($null -eq $pageData) {
                Write-Warning ("Page file {0} parsed as null, skipping" -f $pageFile.Name)
                continue
            }
            if ($pageData.Records) {
                foreach ($record in $pageData.Records) {
                    $totalRecords++
                    $identity = $record.RecordIdentity
                    if ($identity -and $seenIdentities.ContainsKey($identity)) {
                        continue  # Duplicate
                    }
                    if ($identity) {
                        $seenIdentities[$identity] = $true
                    }
                    [void]$allRecords.Add($record)
                }
            }
            $pagesProcessed++
        }
        catch {
            Write-Warning ("Failed to read page file {0}: {1}" -f $pageFile.Name, $_.Exception.Message)
        }
    }

    $uniqueRecords = $allRecords.Count
    $duplicatesRemoved = $totalRecords - $uniqueRecords

    Write-Verbose ("Total: {0}, Unique: {1}, Duplicates removed: {2}" -f $totalRecords, $uniqueRecords, $duplicatesRemoved)

    # Write output
    $useStreaming = $StreamingOutput -or ($uniqueRecords -gt 50000)

    if ($useStreaming) {
        Write-Verbose "Writing combined file (streaming mode)..."
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.FileStream]::new($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None, 65536)
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.Encoding]::UTF8, 65536)

            $writer.WriteLine("[")
            $isFirst = $true
            $written = 0

            foreach ($record in $allRecords) {
                if (-not $isFirst) {
                    $writer.WriteLine(",")
                }
                $json = $record | ConvertTo-Json -Depth 20 -Compress
                $writer.Write("  ")
                $writer.Write($json)
                $isFirst = $false
                $written++

                if ($written % 10000 -eq 0) {
                    Write-Verbose "  Written $written records..."
                }
            }

            $writer.WriteLine("")
            $writer.WriteLine("]")
        }
        finally {
            if ($writer) { $writer.Dispose() }
            if ($stream) { $stream.Dispose() }
        }
    }
    else {
        Write-Verbose "Writing combined file (standard mode)..."
        if ($uniqueRecords -gt 0) {
            $json = $allRecords.ToArray() | ConvertTo-Json -Depth 20
            [System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.Encoding]::UTF8)
        }
        else {
            [System.IO.File]::WriteAllText($OutputPath, "[]", [System.Text.Encoding]::UTF8)
        }
    }

    return @{
        TotalRecords      = $totalRecords
        UniqueRecords     = $uniqueRecords
        DuplicatesRemoved = $duplicatesRemoved
        PagesProcessed    = $pagesProcessed
    }
}

function Find-UnknownActivityTypes {
    <#
    .SYNOPSIS
        Detects activity types or workloads in exported data that aren't in the config.
    .DESCRIPTION
        Compares unique Activity and Workload values found in exported records against
        the known lists from the configuration file. Logs any unknown values as warnings
        so the config can be updated for future exports.
    .PARAMETER Records
        Array of Activity Explorer records to analyze.
    .PARAMETER KnownActivities
        Array of activity type names from the configuration file.
    .PARAMETER KnownWorkloads
        Array of workload names from the configuration file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Records,

        [string[]]$KnownActivities,

        [string[]]$KnownWorkloads
    )

    if ($Records.Count -eq 0) {
        return
    }

    # Extract unique activity types from records
    $recordActivities = @($Records | Where-Object { $_.Activity } | ForEach-Object { $_.Activity } | Select-Object -Unique)
    $recordWorkloads = @($Records | Where-Object { $_.Workload } | ForEach-Object { $_.Workload } | Select-Object -Unique)

    $foundUnknown = $false

    # Check for unknown activities
    if ($KnownActivities -and $KnownActivities.Count -gt 0 -and $recordActivities.Count -gt 0) {
        $unknownActivities = @($recordActivities | Where-Object { $_ -notin $KnownActivities })

        if ($unknownActivities.Count -gt 0) {
            $foundUnknown = $true
            $activityList = $unknownActivities -join ", "
            $msg = "  UNKNOWN ACTIVITY TYPES found in export ({0}): {1}" -f $unknownActivities.Count, $activityList
            Write-ExportLog -Message $msg -Level Warning
            Write-ExportLog -Message "  Consider adding these to ConfigFiles\ActivityExplorerSelector.json" -Level Warning
        }
    }

    # Check for unknown workloads
    if ($KnownWorkloads -and $KnownWorkloads.Count -gt 0 -and $recordWorkloads.Count -gt 0) {
        $unknownWorkloads = @($recordWorkloads | Where-Object { $_ -notin $KnownWorkloads })

        if ($unknownWorkloads.Count -gt 0) {
            $foundUnknown = $true
            $workloadList = $unknownWorkloads -join ", "
            $msg = "  UNKNOWN WORKLOADS found in export ({0}): {1}" -f $unknownWorkloads.Count, $workloadList
            Write-ExportLog -Message $msg -Level Warning
            Write-ExportLog -Message "  Consider adding these to ConfigFiles\ActivityExplorerSelector.json" -Level Warning
        }
    }

    if (-not $foundUnknown) {
        Write-ExportLog -Message "  All activity types and workloads match configuration" -Level Info
    }
}

function Get-ActivityExplorerFilters {
    <#
    .SYNOPSIS
        Loads Activity Explorer filter configuration from ActivityExplorerSelector.json.
    .DESCRIPTION
        Reads the config file, extracts enabled activities and workloads, and returns
        a hashtable suitable for passing to Export-ActivityExplorerWithProgress.
        Only adds a filter category if some (but not all) items are disabled.
        Returns $null if no filters are needed (all items enabled or no config found).
    .PARAMETER ConfigPath
        Path to the ActivityExplorerSelector.json config file.
    .PARAMETER ConfigObject
        Pre-loaded config object (e.g., from ExportSettings.json manifest).
        When provided, skips file read and uses this object directly.
    .PARAMETER LogDetails
        When set, logs details about which filters are active using Write-ExportLog.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ByPath')]
    param(
        [Parameter(ParameterSetName = 'ByPath')]
        [string]$ConfigPath,

        [Parameter(Mandatory, ParameterSetName = 'ByObject')]
        [PSCustomObject]$ConfigObject,

        [switch]$LogDetails
    )

    if ($PSCmdlet.ParameterSetName -eq 'ByObject') {
        $config = $ConfigObject
    }
    elseif ($ConfigPath) {
        $config = Read-JsonConfig -Path $ConfigPath
    }
    else {
        $config = $null
    }

    if (-not $config) {
        if ($LogDetails) {
            Write-ExportLog -Message "  No config file found, exporting all activities" -Level Info
        }
        return $null
    }

    $filters = @{}

    # Only add activity filter if some (but not all) activities are disabled
    if ($config.Activities) {
        $enabledActivities = Get-EnabledItems -Config $config.Activities
        $totalActivities = @($config.Activities.PSObject.Properties | Where-Object { $_.Name -notlike "_*" }).Count
        if ($enabledActivities.Count -gt 0 -and $enabledActivities.Count -lt $totalActivities) {
            $filters['Activity'] = $enabledActivities
            if ($LogDetails) {
                Write-ExportLog -Message "  Filtering to $($enabledActivities.Count) of $totalActivities activities" -Level Info
            }
        }
        elseif ($LogDetails) {
            Write-ExportLog -Message "  All $totalActivities activities enabled (no activity filter)" -Level Info
        }
    }

    # Only add workload filter if some (but not all) workloads are disabled
    if ($config.Workloads) {
        $enabledWorkloads = Get-EnabledItems -Config $config.Workloads
        $totalWorkloads = @($config.Workloads.PSObject.Properties | Where-Object { $_.Name -notlike "_*" }).Count
        if ($enabledWorkloads.Count -gt 0 -and $enabledWorkloads.Count -lt $totalWorkloads) {
            $filters['Workload'] = $enabledWorkloads
            if ($LogDetails) {
                Write-ExportLog -Message "  Filtering to $($enabledWorkloads.Count) of $totalWorkloads workloads" -Level Info
            }
        }
        elseif ($LogDetails) {
            Write-ExportLog -Message "  All $totalWorkloads workloads enabled (no workload filter)" -Level Info
        }
    }

    if ($filters.Count -eq 0) {
        if ($LogDetails) {
            Write-ExportLog -Message "  No filters applied (exporting all activities)" -Level Info
        }
        return $null
    }

    return $filters
}

#endregion

