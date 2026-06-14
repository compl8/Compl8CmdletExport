#region Content Explorer - Aggregate Pagination (shared loop)

function Invoke-CEAggregatePaging {
    <#
    .SYNOPSIS
        Shared Content Explorer aggregate-discovery pagination loop.
    .DESCRIPTION
        Performs the paginated aggregate query for a single TagType/TagName/Workload
        combination and returns the accumulated aggregate records. This is the loop
        body shared by the three call sites (worker, resume, fresh work-plan); the
        differences between those sites are parameterized so each remains
        byte-equivalent to its prior inline implementation:

          - Core paging: Export-ContentExplorerData with -Aggregate -PageSize 5000
            -ErrorAction Stop (+PageCookie), record accumulation, cookie guards
            (null/empty -> throw, same-cookie-stuck -> throw), and termination
            (empty/null page -> break; no-more-pages -> break) are IDENTICAL across
            all RetryModes.

          - RetryMode selects the per-page failure handling:
              * 'WorkerReconnect' — the worker's inline retry (maxAggRetries = 3):
                  - CommandNotFoundException / AuthError -> Disconnect, reconnect via
                    Connect-Compl8Compliance @AuthParams; on success retry the SAME
                    page in place (does NOT consume an attempt); on failure throw the
                    captured page error.
                  - Other errors -> write to the shared + per-worker error logs (when
                    -WriteWorkerErrorLog), then for transient errors wait 60*attempt
                    (or 120s on the final attempt) and retry; otherwise throw.
              * 'None' — a single Export-ContentExplorerData call, no per-page retry.
                Any error propagates to the caller (the resume site's behavior).
              * 'BackoffHelper' — wrap the page call in Invoke-RetryWithBackoff
                -MaxRetries 3 (the fresh work-plan site's behavior). CommandNotFound
                and AuthError throw immediately; transient errors back off 60*attempt
                / 120s.

        Returns an ARRAY of aggregate record objects (each with .Name and .Count),
        in API order, accumulated across all pages. Each caller consumes the array
        exactly as before (worker/resume did `+=`; fresh did ArrayList.Add — all
        three iterate the returned enumerable identically).

    .PARAMETER TagType
        Classifier type (e.g. SensitiveInformationType, Sensitivity).
    .PARAMETER TagName
        The tag name to query.
    .PARAMETER Workload
        The workload to query (Exchange, SharePoint, OneDrive, Teams).
    .PARAMETER PageSize
        Page size for the aggregate query. Default 5000 (the value all three sites use).
    .PARAMETER RetryMode
        'WorkerReconnect' | 'None' | 'BackoffHelper'. Selects the per-page retry
        strategy as described above.
    .PARAMETER AuthParams
        Hashtable of auth parameters for the WorkerReconnect reconnect-in-place path.
        Required (effectively) for 'WorkerReconnect'; the worker passes
        $script:AuthParams from the App/orchestrator scope. Module-scope $script:
        variables are NOT visible here, so this MUST be passed explicitly.
    .PARAMETER WriteWorkerErrorLog
        When set (WorkerReconnect only), per-retry errors are written to both
        ErrorLogPath and WorkerErrorLogPath via Write-ExportErrorLog, matching the
        worker's prior behavior.
    .PARAMETER ErrorLogPath
        Shared error log path (passed from the App scope $script:ErrorLogPath).
    .PARAMETER WorkerErrorLogPath
        Per-worker error log path (passed from the App scope $script:WorkerErrorLogPath).
    .PARAMETER TaskKey
        "TagType|TagName|Workload" key used in the worker error-log entries.
    .PARAMETER BackoffContext
        Context string passed to Invoke-RetryWithBackoff (BackoffHelper only). The
        fresh site uses "Aggregate: <tagName>/<workload>"; pass it verbatim so the
        log message is unchanged.
    .OUTPUTS
        [object[]] — accumulated aggregate records (may be empty).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TagType,

        [Parameter(Mandatory)]
        [string]$TagName,

        [Parameter(Mandatory)]
        [string]$Workload,

        [int]$PageSize = 5000,

        [Parameter(Mandatory)]
        [ValidateSet('WorkerReconnect', 'None', 'BackoffHelper')]
        [string]$RetryMode,

        [hashtable]$AuthParams,

        [switch]$WriteWorkerErrorLog,

        [string]$ErrorLogPath,

        [string]$WorkerErrorLogPath,

        [string]$TaskKey,

        [string]$BackoffContext = "operation"
    )

    # WorkerReconnect tuning (identical to the worker's inline constants).
    $maxAggRetries = 3
    $aggFinalAttemptDelay = 120

    $allAggregates = @()
    $pageCookie = $null
    $aggPageNum = 0

    do {
        $aggParams = @{
            TagType     = $TagType
            TagName     = $TagName
            Workload    = $Workload
            PageSize    = $PageSize
            Aggregate   = $true
            ErrorAction = 'Stop'
        }
        if ($pageCookie) { $aggParams['PageCookie'] = $pageCookie }

        # ── Per-page query with the selected retry strategy ──
        $aggResult = $null

        switch ($RetryMode) {
            'None' {
                # Single call, no per-page retry (resume site). Errors propagate.
                $aggResult = Export-ContentExplorerData @aggParams
            }

            'BackoffHelper' {
                # Fresh work-plan site: Invoke-RetryWithBackoff with MaxRetries=3.
                $aggResult = Invoke-RetryWithBackoff -ScriptBlock {
                    Export-ContentExplorerData @aggParams
                } -MaxRetries $maxAggRetries -Context $BackoffContext
            }

            'WorkerReconnect' {
                # Worker site: inline reconnect-in-place + transient backoff.
                $pageSuccess = $false
                $pageRetry = 0

                while (-not $pageSuccess -and $pageRetry -le $maxAggRetries) {
                    try {
                        $aggResult = Export-ContentExplorerData @aggParams
                        $pageSuccess = $true
                    }
                    catch {
                        $lastPageError = $_
                        $pageRetry++
                        $errorInfo = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_
                        $statusStr = if ($errorInfo.StatusCode) { "HTTP $($errorInfo.StatusCode)" } else { $errorInfo.Category }

                        # Connection lost - cmdlet not available (session dropped or never established)
                        if ($_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                            Write-ExportLog -Message "    CONNECTION LOST: S&C cmdlet not available - attempting reconnection..." -Level Warning
                            try {
                                Disconnect-Compl8Compliance
                                if ($AuthParams -and $AuthParams.Count -gt 0) {
                                    $reAuthResult = Connect-Compl8Compliance @AuthParams
                                    if ($reAuthResult) {
                                        Write-ExportLog -Message "    Reconnection successful - retrying" -Level Success
                                        $pageRetry--
                                        continue
                                    }
                                }
                            }
                            catch {
                                Write-ExportLog -Message ("    Reconnection failed: {0}" -f $_.Exception.Message) -Level Error
                            }
                            # Cannot recover - throw to outer catch (bad session tracking will exit the worker)
                            throw $lastPageError
                        }

                        # Auth recovery
                        if ($errorInfo.Category -eq "AuthError") {
                            Write-ExportLog -Message "    AUTH EXPIRED during aggregate - attempting recovery..." -Level Warning
                            try {
                                Disconnect-Compl8Compliance
                                if ($AuthParams -and $AuthParams.Count -gt 0) {
                                    $reAuthResult = Connect-Compl8Compliance @AuthParams
                                    if ($reAuthResult) {
                                        Write-ExportLog -Message "    Re-authentication successful - retrying" -Level Success
                                        $pageRetry--
                                        continue
                                    }
                                }
                            }
                            catch {
                                Write-ExportLog -Message ("    Re-authentication failed: {0}" -f $_.Exception.Message) -Level Error
                            }
                            throw $lastPageError
                        }

                        # Log the error
                        if ($WriteWorkerErrorLog) {
                            if ($ErrorLogPath) {
                                Write-ExportErrorLog -ErrorLogPath $ErrorLogPath -Context "Worker Aggregate (Page Retry)" -TaskKey $TaskKey -ErrorRecord $_ -AdditionalData @{ RetryCount = $pageRetry; MaxRetries = $maxAggRetries; Page = $aggPageNum }
                            }
                            if ($WorkerErrorLogPath) {
                                Write-ExportErrorLog -ErrorLogPath $WorkerErrorLogPath -Context "Worker Aggregate (Page Retry)" -TaskKey $TaskKey -ErrorRecord $_ -AdditionalData @{ RetryCount = $pageRetry; MaxRetries = $maxAggRetries; Page = $aggPageNum }
                            }
                        }

                        if ($errorInfo.IsTransient -and $pageRetry -le $maxAggRetries) {
                            if ($pageRetry -eq $maxAggRetries) {
                                $msg = "    Aggregate TRANSIENT ERROR [{0}] (attempt {1}/{2}) - final attempt in {3}s" -f $statusStr, $pageRetry, $maxAggRetries, $aggFinalAttemptDelay
                                Write-ExportLog -Message $msg -Level Warning
                                Start-Sleep -Seconds $aggFinalAttemptDelay
                            }
                            else {
                                $retryDelay = 60 * $pageRetry
                                $msg = "    Aggregate TRANSIENT ERROR [{0}] (attempt {1}/{2}) - waiting {3}s" -f $statusStr, $pageRetry, $maxAggRetries, $retryDelay
                                Write-ExportLog -Message $msg -Level Warning
                                Start-Sleep -Seconds $retryDelay
                            }
                        }
                        else {
                            $msg = "    Aggregate FAILED [{0}] after {1} attempts" -f $statusStr, $pageRetry
                            Write-ExportLog -Message $msg -Level Error
                            throw
                        }
                    }
                }
            }
        }

        # ── Page accumulation + cookie advance with guards (shared) ──
        $aggPageNum++

        if ($null -eq $aggResult -or $aggResult.Count -eq 0) { break }

        $metadata = $aggResult[0]
        if ($metadata.RecordsReturned -gt 0) {
            $allAggregates += $aggResult[1..$metadata.RecordsReturned]
        }

        if ($metadata.MorePagesAvailable -eq $true -or $metadata.MorePagesAvailable -eq "True") {
            $newAggCookie = $metadata.PageCookie
            if ([string]::IsNullOrEmpty($newAggCookie)) {
                throw "MorePagesAvailable=true but PageCookie is null/empty - cannot advance aggregate cursor"
            }
            if ($newAggCookie -eq $pageCookie) {
                throw "API returned same PageCookie as previous aggregate page - cursor stuck"
            }
            $pageCookie = $newAggCookie
        }
        else { break }
    } while ($true)

    # Return the FLAT accumulated record collection. Every call site wraps the
    # result with @(...), which normalizes 0/1/N correctly: empty -> count 0;
    # one record -> count 1; N records -> count N. Do NOT use the comma operator
    # or an [object[]] re-wrap here — that would nest the whole record set inside
    # a 1-element outer array, and @(...) does not unwrap a nested array element,
    # so callers would always see .Count == 1 with $agg bound to the entire array.
    return $allAggregates
}

#endregion
