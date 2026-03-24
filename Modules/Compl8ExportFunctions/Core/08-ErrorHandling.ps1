#region Error Handling Functions

function Write-ExportErrorLog {
    <#
    .SYNOPSIS
        Writes detailed error information to an error log file using atomic NTFS append.

    .DESCRIPTION
        Appends a structured error entry to the specified log file. Each entry includes
        a timestamp, operation context, full exception details, HTTP error classification,
        inner exception chain, and optional additional context data.

        Uses Add-Content for atomic NTFS append, making it safe for multiple terminals
        to write to the same log file concurrently.

    .PARAMETER ErrorLogPath
        Full path to the error log file. Created if it does not exist.

    .PARAMETER Context
        Description of the operation that failed (e.g., "Worker Aggregate", "Detail Export").

    .PARAMETER TaskKey
        Identifier for the task that failed (e.g., "SensitiveInformationType|Credit Card|Exchange").

    .PARAMETER ErrorRecord
        The PowerShell ErrorRecord object from the catch block.

    .PARAMETER AdditionalData
        Optional hashtable of extra context to include in the log entry
        (e.g., retry count, page number, page cookie).

    .EXAMPLE
        Write-ExportErrorLog -ErrorLogPath $logPath -Context "Worker Aggregate" `
            -TaskKey "SIT|Credit Card|Exchange" -ErrorRecord $_

    .EXAMPLE
        Write-ExportErrorLog -ErrorLogPath $logPath -Context "Detail Export" `
            -TaskKey "Sensitivity|Confidential|SharePoint" -ErrorRecord $_ `
            -AdditionalData @{ RetryCount = 3; PageNumber = 12 }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ErrorLogPath,

        [string]$Context = "Unknown",

        [string]$TaskKey,

        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [hashtable]$AdditionalData
    )

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $separator = "=" * 80

        # Build the error entry using a StringBuilder for efficiency
        $sb = [System.Text.StringBuilder]::new(2048)
        [void]$sb.AppendLine($separator)
        [void]$sb.AppendLine("TIMESTAMP:  $timestamp")
        [void]$sb.AppendLine("CONTEXT:    $Context")

        if ($TaskKey) {
            [void]$sb.AppendLine("TASK:       $TaskKey")
        }

        [void]$sb.AppendLine("PID:        $PID")

        # Exception details
        $exception = $ErrorRecord.Exception
        if ($exception) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("EXCEPTION:")
            [void]$sb.AppendLine("  Message:  $($exception.Message)")
            [void]$sb.AppendLine("  Type:     $($exception.GetType().FullName)")
        }

        # Stack trace
        if ($ErrorRecord.ScriptStackTrace) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("STACK TRACE:")
            [void]$sb.AppendLine($ErrorRecord.ScriptStackTrace)
        }

        # HTTP error classification
        $httpInfo = Get-HttpErrorExplanation -ErrorMessage $exception.Message -ErrorRecord $ErrorRecord
        if ($httpInfo) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("HTTP ERROR ANALYSIS:")
            if ($httpInfo.StatusCode) {
                [void]$sb.AppendLine("  Status Code: $($httpInfo.StatusCode)")
            }
            [void]$sb.AppendLine("  Category:    $($httpInfo.Category)")
            [void]$sb.AppendLine("  IsTransient: $($httpInfo.IsTransient)")
            [void]$sb.AppendLine("  Explanation: $($httpInfo.Explanation)")
        }

        # Inner exception chain (up to 3 levels)
        $innerException = $exception.InnerException
        $innerLevel = 1
        while ($innerException -and $innerLevel -le 3) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("INNER EXCEPTION (Level $innerLevel):")
            [void]$sb.AppendLine("  Message:  $($innerException.Message)")
            [void]$sb.AppendLine("  Type:     $($innerException.GetType().FullName)")
            $innerException = $innerException.InnerException
            $innerLevel++
        }

        # Additional context data
        if ($AdditionalData -and $AdditionalData.Count -gt 0) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("ADDITIONAL DATA:")
            foreach ($key in $AdditionalData.Keys) {
                $value = $AdditionalData[$key]
                # Truncate long values to keep log readable
                $valueStr = if ($null -eq $value) { "(null)" } else { "$value" }
                if ($valueStr.Length -gt 500) {
                    $valueStr = $valueStr.Substring(0, 500) + "... (truncated)"
                }
                [void]$sb.AppendLine("  ${key}: $valueStr")
            }
        }

        [void]$sb.AppendLine("")

        # Ensure parent directory exists
        $logDir = Split-Path $ErrorLogPath -Parent
        if ($logDir -and -not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Force -Path $logDir | Out-Null
        }

        # Atomic NTFS append - safe for concurrent writes from multiple terminals
        Add-Content -Path $ErrorLogPath -Value $sb.ToString() -Encoding UTF8
    }
    catch {
        # If error logging itself fails, write to console as last resort
        Write-Warning "Failed to write to error log '$ErrorLogPath': $($_.Exception.Message)"
    }
}

function Format-ErrorDetail {
    <#
    .SYNOPSIS
        Formats a PowerShell ErrorRecord into a human-readable string.

    .DESCRIPTION
        Extracts the exception message, exception type, and script stack trace from
        an ErrorRecord and formats them into a single readable string suitable for
        log messages or console output.

    .PARAMETER ErrorRecord
        The PowerShell ErrorRecord object to format.

    .OUTPUTS
        String containing the formatted error details.

    .EXAMPLE
        try { ... } catch { $detail = Format-ErrorDetail -ErrorRecord $_; Write-Host $detail }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $parts = [System.Collections.ArrayList]::new()

    # Exception message
    if ($ErrorRecord.Exception) {
        [void]$parts.Add("Message: $($ErrorRecord.Exception.Message)")
        [void]$parts.Add("Type: $($ErrorRecord.Exception.GetType().FullName)")
    }
    else {
        [void]$parts.Add("Message: $($ErrorRecord.ToString())")
    }

    # Stack trace
    if ($ErrorRecord.ScriptStackTrace) {
        [void]$parts.Add("Stack Trace:")
        [void]$parts.Add($ErrorRecord.ScriptStackTrace)
    }

    return $parts -join [Environment]::NewLine
}

function Get-HttpErrorExplanation {
    <#
    .SYNOPSIS
        Classifies HTTP errors from Purview API responses.

    .DESCRIPTION
        Analyzes an error message and/or ErrorRecord to determine the HTTP status code,
        error category, whether the error is transient (retryable), and a human-readable
        explanation. This is used to drive retry logic and error reporting.

        Categories:
        - AuthError:   HTTP 401 (token expired or invalid credentials)
        - Throttle:    HTTP 429 (rate limiting / too many requests)
        - ServerError: HTTP 500, 502, 503, 504 (transient server-side failures)
        - ClientError: HTTP 400, 403, 404 (non-retryable client errors)
        - Network:     WebException, timeout, connection reset (transient)
        - Unknown:     Unrecognized error pattern

    .PARAMETER ErrorMessage
        The exception message string to analyze for HTTP status codes.

    .PARAMETER ErrorRecord
        The full PowerShell ErrorRecord for additional context extraction.

    .OUTPUTS
        Hashtable with keys: StatusCode (int or $null), Category (string),
        IsTransient (bool), Explanation (string).

    .EXAMPLE
        $info = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_
        if ($info.IsTransient) { Start-Sleep -Seconds 60; continue }
    #>
    [CmdletBinding()]
    param(
        [string]$ErrorMessage,

        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $result = @{
        StatusCode  = $null
        Category    = "Unknown"
        IsTransient = $false
        Explanation = "Unrecognized error"
    }

    # Combine error message sources for pattern matching
    $fullMessage = $ErrorMessage
    if ($ErrorRecord -and $ErrorRecord.Exception) {
        $fullMessage = "$ErrorMessage $($ErrorRecord.Exception.Message)"
        # Check inner exception messages too
        $inner = $ErrorRecord.Exception.InnerException
        if ($inner) {
            $fullMessage = "$fullMessage $($inner.Message)"
        }
    }

    if ([string]::IsNullOrWhiteSpace($fullMessage)) {
        return $result
    }

    # Try to extract HTTP status code from the message
    # Common patterns: "(401)", "Status code: 401", "HttpStatusCode: 401", "Response status code does not indicate success: 401"
    $statusCode = $null
    if ($fullMessage -match '\b(4\d{2}|5\d{2})\b') {
        $statusCode = [int]$Matches[1]
    }

    # Classify by HTTP status code first
    if ($statusCode) {
        $result.StatusCode = $statusCode

        switch ($statusCode) {
            401 {
                $result.Category = "AuthError"
                $result.IsTransient = $false
                $result.Explanation = "Authentication expired or invalid. Token refresh or re-authentication required."
            }
            403 {
                $result.Category = "ClientError"
                $result.IsTransient = $false
                $result.Explanation = "Access denied. Insufficient permissions for the requested operation."
            }
            404 {
                $result.Category = "ClientError"
                $result.IsTransient = $false
                $result.Explanation = "Resource not found. The requested item or endpoint does not exist."
            }
            429 {
                $result.Category = "Throttle"
                $result.IsTransient = $true
                $result.Explanation = "Request throttled. Too many requests - wait before retrying."
            }
            400 {
                $result.Category = "ClientError"
                $result.IsTransient = $false
                $result.Explanation = "Bad request. Check parameters and date ranges."
            }
            500 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Internal server error. Transient failure on the server side."
            }
            502 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Bad gateway. Upstream server returned an invalid response."
            }
            503 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Service unavailable. Server is temporarily overloaded or under maintenance."
            }
            504 {
                $result.Category = "ServerError"
                $result.IsTransient = $true
                $result.Explanation = "Gateway timeout. Server did not respond in time."
            }
            default {
                if ($statusCode -ge 500) {
                    $result.Category = "ServerError"
                    $result.IsTransient = $true
                    $result.Explanation = "Server error (HTTP $statusCode). Transient failure."
                }
                elseif ($statusCode -ge 400) {
                    $result.Category = "ClientError"
                    $result.IsTransient = $false
                    $result.Explanation = "Client error (HTTP $statusCode). Review request parameters."
                }
            }
        }

        return $result
    }

    # No status code found - classify by message patterns
    $messageLower = $fullMessage.ToLower()

    # Network / connectivity errors
    if ($messageLower -match 'webexception|web exception|socketexception|socket exception') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "Network error (WebException). Check connectivity and retry."
        return $result
    }

    if ($messageLower -match 'timeout|timed\s*out|operation\s*was\s*canceled') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "Request timed out. The server did not respond within the allowed time."
        return $result
    }

    if ($messageLower -match 'connection\s*(was\s*)?reset|connection\s*(was\s*)?closed|connection\s*(was\s*)?refused') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "Connection was reset or refused. Network disruption detected."
        return $result
    }

    if ($messageLower -match 'ssl|tls|certificate|secure\s*channel') {
        $result.Category = "Network"
        $result.IsTransient = $true
        $result.Explanation = "SSL/TLS error. Secure channel could not be established."
        return $result
    }

    # Authentication patterns without HTTP status codes
    if ($messageLower -match 'unauthorized|authentication\s*failed|token\s*expired|access\s*token') {
        $result.Category = "AuthError"
        $result.IsTransient = $false
        $result.Explanation = "Authentication error detected in message. Re-authentication may be required."
        return $result
    }

    # Throttling patterns without HTTP status codes
    if ($messageLower -match 'throttl|rate\s*limit|too\s*many\s*requests') {
        $result.Category = "Throttle"
        $result.IsTransient = $true
        $result.Explanation = "Throttling detected in message. Wait before retrying."
        return $result
    }

    # Generic transient server-side patterns
    if ($messageLower -match 'service\s*unavailable|server\s*(is\s*)?busy|temporarily\s*unavailable|internal\s*server\s*error') {
        $result.Category = "ServerError"
        $result.IsTransient = $true
        $result.Explanation = "Server-side transient error. Retry after delay."
        return $result
    }

    # Role definition not ready (EXO session initialization race condition)
    if ($messageLower -match 'not present in the role definition') {
        $result.Category = "ServerError"
        $result.IsTransient = $true
        $result.Explanation = "Role definition not yet provisioned. Backend session still initializing - retry after delay."
        return $result
    }

    # AggregateException from EXO REST module (wraps internal connection/HTTP failures)
    if ($messageLower -match 'aggregateexception|one or more errors occurred') {
        $result.Category = "ServerError"
        $result.IsTransient = $true
        $result.Explanation = "AggregateException from EXO module. Likely transient server-side or connection failure."
        return $result
    }

    # Fall through - unknown error
    $result.Explanation = "Unrecognized error: $($ErrorMessage.Substring(0, [Math]::Min($ErrorMessage.Length, 200)))"
    return $result
}

#endregion

