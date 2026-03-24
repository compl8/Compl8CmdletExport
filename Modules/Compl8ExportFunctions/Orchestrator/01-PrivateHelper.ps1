#region Private Helper

function Write-ProgressEntry {
    <#
    .SYNOPSIS
        Writes a timestamped entry to the progress log file.
    .DESCRIPTION
        Appends a line to the progress log with a timestamp prefix.
        Silently skips if no path is provided or the write fails.
    .PARAMETER Path
        Path to the progress log file. Also accepts -LogPath.
    .PARAMETER Message
        The message to write.
    #>
    [CmdletBinding()]
    param(
        [Alias('LogPath')]
        [string]$Path,
        [string]$Message
    )

    if ([string]::IsNullOrEmpty($Path)) {
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $entry = "[{0}] {1}" -f $timestamp, $Message
        [System.IO.File]::AppendAllText($Path, $entry + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    catch {
        # Progress log write failures are non-critical - export continues without progress logging
        Write-Verbose "Progress log write failed (non-critical): $($_.Exception.Message)"
    }
}

#endregion

