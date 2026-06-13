#region Content Explorer Dispatch Callback Helpers

# Shared building blocks for the Content Explorer detail dispatch callbacks used by
# Invoke-DispatchLoop. Previously the resume (Invoke-ContentExplorerResume) and
# task-CSV (Invoke-ContentExplorerFromTasksCsv) orchestration functions each carried
# verbatim copies of the detail done/error signal scan, the OnMatch task lookup, and
# the OnDispatch payload builder. These functions hold the single implementation so
# the inline scriptblocks become thin wrappers.
#
# IMPORTANT (behavior preservation): the fresh multi-terminal orchestrator
# (Invoke-ContentExplorerExport) does NOT use Read-CEDetailSignals / Find-CEDetailTaskMatch.
# Its scan interleaves aggregate-CSV claiming and aggregate-error handling with the detail
# scan, tags each completion/error with TaskType, scans only the Completions dir (not worker
# dirs), and removes (rather than quarantines) a null detail payload. Those differences are
# intentional, so the fresh function keeps its own inline scan/match. Only the resume and
# task-CSV copies — which were byte-identical to each other — are consolidated here.

function Read-CEDetailSignals {
    <#
    .SYNOPSIS
        Scans Content Explorer detail completion/error signal files and returns the
        completed/error task lists for the dispatch loop (resume & task-CSV paths).
    .DESCRIPTION
        Scans the central Completions directory and (for backward compatibility) each
        worker coordination directory for detail-done-*.txt / done-detail-*.txt and
        error-detail-*.txt signal files. Each file is verified via
        ConvertFrom-SignedEnvelopeJson. Successfully parsed signals are appended to the
        completed/error lists and the file is removed. A signal that parses to $null
        (empty/unsigned/forged) or throws is QUARANTINED to <file>.invalid (falling back
        to deletion if the rename fails) so it is not re-scanned and re-warned forever.

        This is the exact logic that Invoke-ContentExplorerResume and
        Invoke-ContentExplorerFromTasksCsv shared verbatim. Returns a hashtable shaped
        for Invoke-DispatchLoop: @{ CompletedTasks = @(...); ErrorTasks = @(...) }.
    .PARAMETER ExportDir
        The export run directory. Used to locate the Completions dir and signing key.
    .PARAMETER WorkerDirs
        Worker coordination directories to additionally scan (backward compatibility).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ExportDir,

        [string[]]$WorkerDirs = @()
    )

    $completed = @()
    $errors = @()
    $completionsDir = Get-CompletionsDir $ExportDir
    $signingKey = Get-ExportRunSigningKey -ExportDir $ExportDir

    # Scan central Completions/ directory
    $doneSignalFiles = @(Get-ChildItem -Path $completionsDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
    $errorSignalFiles = @(Get-ChildItem -Path $completionsDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)

    # Also scan worker dirs (backward compat)
    foreach ($wDir in $WorkerDirs) {
        $doneSignalFiles += @(Get-ChildItem -Path $wDir -Filter "detail-done-*.txt" -File -ErrorAction SilentlyContinue)
        $doneSignalFiles += @(Get-ChildItem -Path $wDir -Filter "done-detail-*.txt" -File -ErrorAction SilentlyContinue)
        $errorSignalFiles += @(Get-ChildItem -Path $wDir -Filter "error-detail-*.txt" -File -ErrorAction SilentlyContinue)
    }

    foreach ($doneFile in $doneSignalFiles) {
        try {
            $doneContent = [System.IO.File]::ReadAllText($doneFile.FullName)
            $doneData = ConvertFrom-SignedEnvelopeJson -Json $doneContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail completion file {0}" -f $doneFile.Name)
            if ($null -eq $doneData) {
                Write-ExportLog -Message ("  Warning: Empty/null detail done file {0} - quarantined" -f $doneFile.Name) -Level Warning -LogOnly
                try { [System.IO.File]::Move($doneFile.FullName, ($doneFile.FullName + '.invalid'), $true) }
                catch { try { Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
                continue
            }
            $doneLocation = if ($doneData.Location) { $doneData.Location } else { "" }
            $completed += @{
                TagType     = $doneData.TagType
                TagName     = $doneData.TagName
                Workload    = $doneData.Workload
                Location    = $doneLocation
                RecordCount = $doneData.RecordCount
                Message     = "{0}/{1}{2} -> {3} records" -f $doneData.TagName, $doneData.Workload, $(if ($doneLocation) { "/$doneLocation" } else { "" }), $doneData.RecordCount
            }
            Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-ExportLog -Message ("  Warning: Could not parse detail done file {0} - quarantined" -f $doneFile.Name) -Level Warning -LogOnly
            try { [System.IO.File]::Move($doneFile.FullName, ($doneFile.FullName + '.invalid'), $true) }
            catch { try { Remove-Item -Path $doneFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
        }
    }

    foreach ($errFile in $errorSignalFiles) {
        try {
            $errContent = [System.IO.File]::ReadAllText($errFile.FullName)
            $errData = ConvertFrom-SignedEnvelopeJson -Json $errContent -SigningKey $signingKey -RequireSignature:([bool]$signingKey) -Context ("detail error file {0}" -f $errFile.Name)
            if ($null -eq $errData) {
                Write-ExportLog -Message ("  Warning: Empty/null detail error file {0} - quarantined" -f $errFile.Name) -Level Warning -LogOnly
                try { [System.IO.File]::Move($errFile.FullName, ($errFile.FullName + '.invalid'), $true) }
                catch { try { Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
                continue
            }
            $errLocation = if ($errData.Location) { $errData.Location } else { "" }
            $errors += @{
                TagType      = $errData.TagType
                TagName      = $errData.TagName
                Workload     = $errData.Workload
                Location     = $errLocation
                ErrorMessage = $errData.Error
                Message      = "{0}/{1}{2}: {3}" -f $errData.TagName, $errData.Workload, $(if ($errLocation) { "/$errLocation" } else { "" }), $errData.Error
            }
            Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-ExportLog -Message ("  Warning: Could not parse error file {0} - quarantined" -f $errFile.Name) -Level Warning -LogOnly
            try { [System.IO.File]::Move($errFile.FullName, ($errFile.FullName + '.invalid'), $true) }
            catch { try { Remove-Item -Path $errFile.FullName -Force -ErrorAction SilentlyContinue } catch {} }
        }
    }

    return @{ CompletedTasks = $completed; ErrorTasks = $errors }
}

function Find-CEDetailTaskMatch {
    <#
    .SYNOPSIS
        Finds the in-progress Content Explorer detail task matching a completion/error
        signal (resume & task-CSV OnMatch logic).
    .DESCRIPTION
        Matches on TagType/TagName/Workload/Location with Status = "InProgress". If no
        location-specific match is found and the signal carries no location, falls back
        to matching by TagType/TagName/Workload only. This is the exact logic the resume
        and task-CSV orchestration functions shared verbatim ($ceDetailOnMatch).

        Returns the matching task object, or $null.
    .PARAMETER Data
        The completion/error signal hashtable (carries TagType/TagName/Workload/Location).
    .PARAMETER Tasks
        The task collection to search.
    #>
    param(
        [Parameter(Mandatory)]
        $Data,

        [Parameter(Mandatory)]
        $Tasks
    )

    $loc = if ($Data.Location) { $Data.Location } else { "" }
    $match = $Tasks | Where-Object {
        $_.TagType -eq $Data.TagType -and
        $_.TagName -eq $Data.TagName -and
        $_.Workload -eq $Data.Workload -and
        $(if ($_.Location) { $_.Location } else { "" }) -eq $loc -and
        $_.Status -eq "InProgress"
    } | Select-Object -First 1
    if (-not $match -and -not $loc) {
        $match = $Tasks | Where-Object {
            $_.TagType -eq $Data.TagType -and
            $_.TagName -eq $Data.TagName -and
            $_.Workload -eq $Data.Workload -and
            $_.Status -eq "InProgress"
        } | Select-Object -First 1
    }
    return $match
}

function New-CEDetailDispatchPayload {
    <#
    .SYNOPSIS
        Builds the file-drop task payload for a Content Explorer detail task.
    .DESCRIPTION
        Produces the 8-field Detail task hashtable sent to a worker via Send-WorkerTask.
        This construction was identical across the resume and task-CSV OnDispatch
        callbacks and the Detail branch of the fresh orchestrator's OnDispatch callback.
        Callers remain responsible for sending it and (where applicable) recording
        Context.DispatchTimes — that surrounding logic differs between call sites and is
        deliberately NOT folded in here.
    .PARAMETER NextTask
        The task object selected for dispatch.
    #>
    param(
        [Parameter(Mandatory)]
        $NextTask
    )

    return @{
        Phase         = "Detail"
        TagType       = $NextTask.TagType
        TagName       = $NextTask.TagName
        Workload      = $NextTask.Workload
        Location      = if ($NextTask.Location) { $NextTask.Location } else { "" }
        LocationType  = if ($NextTask.LocationType) { $NextTask.LocationType } else { "" }
        ExpectedCount = ($NextTask.ExpectedCount -as [int])
        PageSize      = ($NextTask.PageSize -as [int])
    }
}

#endregion
