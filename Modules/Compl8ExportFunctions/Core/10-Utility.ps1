#region Utility Functions

function Format-TimeSpan {
    <#
    .SYNOPSIS
        Formats a duration in seconds into a human-readable time string.

    .DESCRIPTION
        Converts a number of seconds into a concise display string using hours,
        minutes, and seconds components. Omits zero-value leading components
        for brevity (e.g., does not show "0h" for sub-hour durations).

    .PARAMETER Seconds
        The duration in seconds to format. Accepts fractional values.

    .OUTPUTS
        String in format "Xh Ym Zs", "Ym Zs", or "Zs" depending on magnitude.

    .EXAMPLE
        Format-TimeSpan -Seconds 3725.4
        # Returns "1h 2m 5s"

    .EXAMPLE
        Format-TimeSpan -Seconds 330
        # Returns "5m 30s"

    .EXAMPLE
        Format-TimeSpan -Seconds 45.7
        # Returns "45s"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Seconds
    )

    # Handle negative or zero values
    if ($Seconds -le 0) {
        return "0s"
    }

    $totalSeconds = [int][Math]::Floor($Seconds)
    $hours = [int][Math]::Floor($totalSeconds / 3600)
    $minutes = [int][Math]::Floor(($totalSeconds % 3600) / 60)
    $secs = $totalSeconds % 60

    if ($hours -gt 0) {
        return "${hours}h ${minutes}m ${secs}s"
    }
    elseif ($minutes -gt 0) {
        return "${minutes}m ${secs}s"
    }
    else {
        return "${secs}s"
    }
}

#endregion

