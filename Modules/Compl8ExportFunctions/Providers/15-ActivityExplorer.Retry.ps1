#region Activity Explorer Retry Helpers

function Get-PageErrorMessage {
    <#
    .SYNOPSIS
        Extracts a clean error message from an error record.
    .PARAMETER ErrorRecord
        The PowerShell error record to extract the message from.
    .OUTPUTS
        String containing the cleaned error message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    if ($null -eq $ErrorRecord) {
        return "Unknown error (null error record)"
    }

    # Try the exception message first
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Message) {
            return $ErrorRecord.Exception.Message
        }
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            return $ErrorRecord.ErrorDetails.Message
        }
        return $ErrorRecord.ToString()
    }

    # Handle plain exceptions
    if ($ErrorRecord -is [System.Exception]) {
        return $ErrorRecord.Message
    }

    # Fallback: convert to string
    return "$ErrorRecord"
}

function Test-PageHasContent {
    <#
    .SYNOPSIS
        Checks if an Activity Explorer API result page has actual records.
    .PARAMETER Result
        The API response object from Export-ActivityExplorerData.
    .OUTPUTS
        Boolean indicating whether the page contains records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Result
    )

    if ($null -eq $Result) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($Result.ResultData)) {
        return $false
    }

    try {
        $records = $Result.ResultData | ConvertFrom-Json
        if ($null -eq $records) {
            return $false
        }
        return (@($records).Count -gt 0)
    }
    catch {
        Write-Verbose "Test-PageHasContent: Failed to parse ResultData: $($_.Exception.Message)"
        return $false
    }
}

function Add-PartialError {
    <#
    .SYNOPSIS
        Adds an error entry to the tracker's PartialErrors array.
    .PARAMETER Tracker
        The run tracker hashtable to update.
    .PARAMETER PageNumber
        The page number where the error occurred.
    .PARAMETER ErrorMessage
        The error message to record.
    .PARAMETER ErrorType
        Classification of the error (e.g., "SameCookieRetryExhausted", "ApiError", "FutureDateError").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Tracker,

        [Parameter(Mandatory)]
        [int]$PageNumber,

        [Parameter(Mandatory)]
        [string]$ErrorMessage,

        [Parameter(Mandatory)]
        [string]$ErrorType
    )

    if (-not $Tracker.ContainsKey('PartialErrors') -or $null -eq $Tracker['PartialErrors']) {
        $Tracker['PartialErrors'] = @()
    }

    $errorEntry = @{
        Timestamp    = (Get-Date).ToString("o")
        PageNumber   = $PageNumber
        ErrorType    = $ErrorType
        ErrorMessage = $ErrorMessage
    }

    $Tracker['PartialErrors'] = @($Tracker['PartialErrors']) + @($errorEntry)
}

function Get-RetryDelay {
    <#
    .SYNOPSIS
        Calculates retry delay based on error classification and attempt number.
    .DESCRIPTION
        Centralizes the retry delay calculation pattern used across Content Explorer
        and Activity Explorer pagination. Delay varies by transient vs non-transient
        errors, with a longer final-attempt delay.
    .PARAMETER AttemptNumber
        Current retry attempt (1-based).
    .PARAMETER MaxRetries
        Maximum number of retries allowed.
    .PARAMETER IsTransient
        Whether the error is transient (server/network). Transient gets longer delays.
    .PARAMETER TransientDelaySeconds
        Base delay for transient errors. Default: 60
    .PARAMETER NonTransientDelaySeconds
        Delay for non-transient errors. Default: 5
    .PARAMETER FinalAttemptDelaySeconds
        Delay before the final retry attempt. Default: 120
    .PARAMETER ScaleByAttempt
        If set, multiplies transient delay by attempt number. Default: $false (flat delay).
    .OUTPUTS
        Integer - number of seconds to wait before retrying.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$AttemptNumber,

        [int]$MaxRetries = 3,

        [bool]$IsTransient = $true,

        [int]$TransientDelaySeconds = 60,

        [int]$NonTransientDelaySeconds = 5,

        [int]$FinalAttemptDelaySeconds = 120,

        [switch]$ScaleByAttempt
    )

    # Final attempt: always use the longer delay regardless of error type
    if ($AttemptNumber -ge $MaxRetries) {
        return $FinalAttemptDelaySeconds
    }

    if ($IsTransient) {
        if ($ScaleByAttempt) {
            return ($TransientDelaySeconds * $AttemptNumber)
        }
        return $TransientDelaySeconds
    }

    return $NonTransientDelaySeconds
}

function Invoke-RetryWithBackoff {
    <#
    .SYNOPSIS
        Generic retry wrapper with backoff delays.
    .DESCRIPTION
        Executes a script block with retry logic. Delay pattern varies based on whether
        the error is transient (longer delays) or non-transient (shorter delays).
        On the final attempt, uses a longer delay to allow recovery.
    .PARAMETER ScriptBlock
        The script block to execute with retry.
    .PARAMETER MaxRetries
        Maximum number of retry attempts. Default: 3
    .PARAMETER InitialDelaySeconds
        Base delay for transient errors (multiplied by attempt number). Default: 60
    .PARAMETER FinalAttemptDelaySeconds
        Delay before the final retry attempt. Default: 120
    .PARAMETER Context
        Description of the operation for logging purposes.
    .OUTPUTS
        Returns the script block's result on success. Throws on final failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [int]$MaxRetries = 3,

        [int]$InitialDelaySeconds = 60,

        [int]$FinalAttemptDelaySeconds = 120,

        [string]$Context = "operation"
    )

    $attempt = 0
    $lastError = $null

    while ($attempt -le $MaxRetries) {
        try {
            $result = & $ScriptBlock
            return $result
        }
        catch {
            $lastError = $_
            $attempt++
            $errorMessage = Get-PageErrorMessage -ErrorRecord $_

            # Connection lost - cmdlet not available; no point retrying
            if ($_.Exception -is [System.Management.Automation.CommandNotFoundException]) {
                Write-ExportLog -Message ("  FATAL during {0}: S&C cmdlet not available - connection lost" -f $Context) -Level Error
                throw
            }

            # Check if this is an auth error - do not retry, throw immediately
            $errorInfo = Get-HttpErrorExplanation -ErrorMessage $errorMessage -ErrorRecord $_
            if ($errorInfo.Category -eq "AuthError") {
                Write-ExportLog -Message ("  AUTH ERROR during {0} - not retrying" -f $Context) -Level Error
                throw
            }

            if ($attempt -gt $MaxRetries) {
                # All retries exhausted
                $msg = "  {0}: All {1} retries exhausted. Last error: {2}" -f $Context, $MaxRetries, $errorMessage
                Write-ExportLog -Message $msg -Level Error
                throw
            }

            $isTransient = $errorInfo.IsTransient
            $statusStr = if ($errorInfo.StatusCode) { "HTTP {0}" -f $errorInfo.StatusCode } else { $errorInfo.Category }

            $delay = Get-RetryDelay -AttemptNumber $attempt -MaxRetries $MaxRetries `
                -IsTransient $isTransient -TransientDelaySeconds $InitialDelaySeconds `
                -FinalAttemptDelaySeconds $FinalAttemptDelaySeconds -ScaleByAttempt

            $levelPrefix = if ($isTransient -and $attempt -lt $MaxRetries) { "TRANSIENT " } else { "" }
            $msg = "  {0}: {1}{2} (attempt {3}/{4}) - waiting {5}s" -f $Context, $levelPrefix, $statusStr, $attempt, $MaxRetries, $delay
            Write-ExportLog -Message $msg -Level Warning
            Start-Sleep -Seconds $delay
        }
    }

    # Should not reach here, but safety net
    throw $lastError
}

#endregion

