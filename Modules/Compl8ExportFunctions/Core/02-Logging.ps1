#region Logging Functions

function Initialize-ExportLog {
    <#
    .SYNOPSIS
        Initializes a log file for the export session.
    .PARAMETER LogDirectory
        Directory where log files will be created.
    .PARAMETER Prefix
        Prefix for the log file name.
    .OUTPUTS
        String path to the created log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogDirectory,

        [string]$Prefix = "Compl8Export"
    )

    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $script:LogFile = Join-Path $LogDirectory "$Prefix-$timestamp.log"
    $script:SessionStartTime = Get-Date

    $header = @"
================================================================================
Compl8 Cmdlet Export
Started: $($script:SessionStartTime.ToString("yyyy-MM-dd HH:mm:ss"))
PowerShell Version: $($PSVersionTable.PSVersion)
================================================================================

"@
    Set-Content -Path $script:LogFile -Value $header
    return $script:LogFile
}

function Write-ExportLog {
    <#
    .SYNOPSIS
        Writes a log message to console and log file.
    .PARAMETER Message
        The message to log.
    .PARAMETER Level
        Log level: Info, Warning, Error, Success.
    .PARAMETER NoNewline
        Don't add newline (for status updates).
    .PARAMETER LogOnly
        Write to log file only (skip console output). Useful when a dashboard is managing console display.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info",

        [switch]$NoNewline,

        [switch]$LogOnly
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Console output (unless LogOnly)
    if (-not $LogOnly) {
        $color = switch ($Level) {
            "Info"    { "White" }
            "Warning" { "Yellow" }
            "Error"   { "Red" }
            "Success" { "Green" }
        }

        if ($NoNewline) {
            Write-Host $Message -ForegroundColor $color -NoNewline
        }
        else {
            Write-Host $Message -ForegroundColor $color
        }
    }

    # Write to log file
    if ($script:LogFile -and (Test-Path $script:LogFile -ErrorAction SilentlyContinue)) {
        Add-Content -Path $script:LogFile -Value $logEntry
    }

    # Track errors and warnings (capped to most-recent N entries; oldest dropped when exceeded)
    $cap = if ($script:ExportStatsMaxEntries) { $script:ExportStatsMaxEntries } else { 500 }
    if ($Level -eq "Error") {
        [void]$script:ExportStats.Errors.Add(@{ Timestamp = $timestamp; Message = $Message })
        while ($script:ExportStats.Errors.Count -gt $cap) { $script:ExportStats.Errors.RemoveAt(0) }
    }
    elseif ($Level -eq "Warning") {
        [void]$script:ExportStats.Warnings.Add(@{ Timestamp = $timestamp; Message = $Message })
        while ($script:ExportStats.Warnings.Count -gt $cap) { $script:ExportStats.Warnings.RemoveAt(0) }
    }
}

function Get-ExportStatistics {
    <#
    .SYNOPSIS
        Returns statistics from the current export session.
    #>
    [CmdletBinding()]
    param()

    return @{
        StartTime = $script:SessionStartTime
        EndTime = Get-Date
        Duration = if ($script:SessionStartTime) { (Get-Date) - $script:SessionStartTime } else { $null }
        ItemsExported = $script:ExportStats.ItemsExported.Clone()
        ErrorCount = $script:ExportStats.Errors.Count
        WarningCount = $script:ExportStats.Warnings.Count
        Errors = $script:ExportStats.Errors.ToArray()
        Warnings = $script:ExportStats.Warnings.ToArray()
        LogFile = $script:LogFile
    }
}

function Add-ExportCount {
    <#
    .SYNOPSIS
        Tracks the count of exported items by category.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Category,

        [Parameter(Mandatory)]
        [int]$Count
    )

    if ($script:ExportStats.ItemsExported.ContainsKey($Category)) {
        $script:ExportStats.ItemsExported[$Category] += $Count
    }
    else {
        $script:ExportStats.ItemsExported[$Category] = $Count
    }
}

#endregion

