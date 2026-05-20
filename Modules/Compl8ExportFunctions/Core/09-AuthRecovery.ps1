#region Auth Recovery

function Invoke-WithAuthRecovery {
    <#
    .SYNOPSIS
        Wraps a scriptblock with automatic authentication token expiry recovery.

    .DESCRIPTION
        Executes the provided scriptblock and monitors for HTTP 401 (authentication expired)
        errors. On detecting a 401:

        - Certificate auth: Silently disconnects and reconnects using the stored auth
          parameters. No user interaction required.
        - Interactive auth (non-worker): Prompts the user before opening a browser login
          window. User can press Enter to re-authenticate or Q to abort.
        - Worker mode: Throws the error immediately to let the worker handle it (workers
          cannot prompt for interactive auth; they save progress and exit).

        On successful re-authentication, the scriptblock is retried once.

    .PARAMETER ScriptBlock
        The scriptblock to execute with auth recovery protection.

    .PARAMETER AuthParams
        Hashtable of authentication parameters (same format as Connect-Compl8Compliance).
        If it contains AppId and CertificateThumbprint, certificate auth is assumed.

    .PARAMETER Context
        Description of the operation for logging purposes.

    .PARAMETER IsWorkerMode
        When $true, indicates this is a spawned worker terminal that cannot prompt
        for interactive authentication. Auth failures will throw immediately.

    .OUTPUTS
        The return value of the scriptblock on success.

    .EXAMPLE
        $result = Invoke-WithAuthRecovery -ScriptBlock {
            Export-ContentExplorerData -TagType "SensitiveInformationType" -TagName "Credit Card" -PageSize 1000
        } -AuthParams $script:AuthParams -Context "Content Explorer export"

    .EXAMPLE
        $data = Invoke-WithAuthRecovery -ScriptBlock { Get-DlpCompliancePolicy } `
            -AuthParams $authParams -Context "DLP Policy retrieval"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [hashtable]$AuthParams,

        [string]$Context = "API call",

        [bool]$IsWorkerMode = $false
    )

    try {
        # Execute the scriptblock
        $result = & $ScriptBlock
        return $result
    }
    catch {
        # Check if this is an auth error
        $errorInfo = Get-HttpErrorExplanation -ErrorMessage $_.Exception.Message -ErrorRecord $_

        if ($errorInfo.Category -ne "AuthError") {
            # Not an auth error - rethrow as-is
            throw
        }

        Write-ExportLog -Message "AUTH EXPIRED during $Context - attempting recovery..." -Level Warning

        # Worker mode cannot do interactive auth - throw to let worker save progress and exit
        if ($IsWorkerMode) {
            $isCertAuth = $AuthParams -and $AuthParams.ContainsKey('AppId') -and
                         ($AuthParams.ContainsKey('CertificateThumbprint') -or $AuthParams.ContainsKey('Certificate'))

            if (-not $isCertAuth) {
                Write-ExportLog -Message "Worker cannot re-authenticate interactively. Throwing to save progress." -Level Error
                throw
            }
        }

        # Determine auth mode
        $isCertAuth = $AuthParams -and $AuthParams.ContainsKey('AppId') -and
                     ($AuthParams.ContainsKey('CertificateThumbprint') -or $AuthParams.ContainsKey('Certificate'))

        if ($isCertAuth) {
            # Certificate auth - reconnect silently
            Write-ExportLog -Message "  Reconnecting with certificate authentication..." -Level Info
            try {
                Disconnect-Compl8Compliance
                $reconnected = Connect-Compl8Compliance @AuthParams
                if (-not $reconnected) {
                    Write-ExportLog -Message "  Certificate re-authentication returned false" -Level Error
                    throw $_
                }
                Write-ExportLog -Message "  Re-authentication successful - retrying $Context" -Level Success
            }
            catch {
                Write-ExportLog -Message "  Certificate re-authentication failed: $($_.Exception.Message)" -Level Error
                throw
            }
        }
        else {
            # Interactive auth - prompt user (non-worker only)
            Write-ExportLog -Message "  Authentication token has expired." -Level Warning
            Write-Host ""
            Write-Host "Press ENTER to open browser for re-authentication, or Q to abort: " -ForegroundColor Yellow -NoNewline
            $userInput = Read-Host

            if ($userInput -and $userInput.Trim().ToUpper() -eq 'Q') {
                Write-ExportLog -Message "  User chose to abort re-authentication" -Level Error
                throw $_
            }

            try {
                Disconnect-Compl8Compliance
                $connectParams = if ($AuthParams -and $AuthParams.Count -gt 0) { $AuthParams } else { @{} }
                $reconnected = Connect-Compl8Compliance @connectParams
                if (-not $reconnected) {
                    Write-ExportLog -Message "  Interactive re-authentication returned false" -Level Error
                    throw $_
                }
                Write-ExportLog -Message "  Re-authentication successful - retrying $Context" -Level Success
            }
            catch {
                Write-ExportLog -Message "  Interactive re-authentication failed: $($_.Exception.Message)" -Level Error
                throw
            }
        }

        # Retry the scriptblock once after successful re-authentication
        try {
            $retryResult = & $ScriptBlock
            return $retryResult
        }
        catch {
            Write-ExportLog -Message "  Retry after re-authentication failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
}

function Invoke-WorkerReconnect {
    <#
    .SYNOPSIS
        Attempts to silently re-authenticate from a spawned worker after an auth
        or lost-session error. Workers cannot prompt interactively.

    .DESCRIPTION
        Only certificate auth can be recovered silently. If AuthParams describe a
        certificate connection, the worker disconnects and reconnects. For
        interactive auth (or missing params), recovery is impossible from a worker
        and the function returns $false so the caller can exit and let the
        orchestrator reclaim the in-flight task.

    .PARAMETER AuthParams
        Hashtable of authentication parameters (same format as Connect-Compl8Compliance).

    .OUTPUTS
        [bool] $true if reconnected, $false if the worker should exit.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$AuthParams
    )

    if (-not $AuthParams -or $AuthParams.Count -eq 0) {
        Write-ExportLog -Message "    Worker has no stored auth params - cannot re-authenticate" -Level Error
        return $false
    }

    $isCertAuth = $AuthParams.ContainsKey('AppId') -and
                  ($AuthParams.ContainsKey('CertificateThumbprint') -or $AuthParams.ContainsKey('Certificate'))

    if (-not $isCertAuth) {
        Write-ExportLog -Message "    Worker uses interactive auth - cannot re-authenticate silently" -Level Error
        return $false
    }

    try {
        Disconnect-Compl8Compliance -LogOnly
        $reconnected = Connect-Compl8Compliance @AuthParams -LogOnly
        if ($reconnected) {
            Write-ExportLog -Message "    Worker re-authentication successful" -Level Success
            return $true
        }
        Write-ExportLog -Message "    Worker re-authentication returned false" -Level Error
        return $false
    }
    catch {
        Write-ExportLog -Message ("    Worker re-authentication failed: {0}" -f $_.Exception.Message) -Level Error
        return $false
    }
}

#endregion

