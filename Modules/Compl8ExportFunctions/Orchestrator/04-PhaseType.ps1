#region Phase/Type I/O

function Write-ExportPhase {
    <#
    .SYNOPSIS
        Writes the current export phase to ExportPhase.txt (atomic write).
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][ValidateSet('Aggregate','Detail','Completed','AEExport','AECompleted')][string]$Phase
    )
    $coordDir = Get-CoordinationDir $ExportDir
    if (-not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $phasePath = Join-Path $coordDir "ExportPhase.txt"
    $tmpPath = $phasePath + ".tmp.$PID"
    [System.IO.File]::WriteAllText($tmpPath, $Phase)
    [System.IO.File]::Move($tmpPath, $phasePath, $true)
}

function Read-ExportPhase {
    <#
    .SYNOPSIS
        Reads the current export phase from ExportPhase.txt.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    $phasePath = Join-Path (Get-CoordinationDir $ExportDir) "ExportPhase.txt"
    if (Test-Path $phasePath) {
        return ([System.IO.File]::ReadAllText($phasePath)).Trim()
    }
    return $null
}

function Write-ExportType {
    <#
    .SYNOPSIS
        Writes the export type marker to ExportType.txt (atomic write).
    .DESCRIPTION
        Used to identify whether an export directory contains a Content Explorer
        or Activity Explorer multi-terminal export. Workers read this at startup
        to determine which worker function to invoke.
    #>
    param(
        [Parameter(Mandatory)][string]$ExportDir,
        [Parameter(Mandatory)][ValidateSet('ContentExplorer','ActivityExplorer')][string]$Type
    )
    $coordDir = Get-CoordinationDir $ExportDir
    if (-not (Test-Path $coordDir)) { New-Item -ItemType Directory -Force -Path $coordDir | Out-Null }
    $typePath = Join-Path $coordDir "ExportType.txt"
    $tmpPath = $typePath + ".tmp.$PID"
    [System.IO.File]::WriteAllText($tmpPath, $Type)
    [System.IO.File]::Move($tmpPath, $typePath, $true)
}

function Read-ExportType {
    <#
    .SYNOPSIS
        Reads the export type from ExportType.txt.
    .OUTPUTS
        'ContentExplorer', 'ActivityExplorer', or $null if file doesn't exist.
    #>
    param([Parameter(Mandatory)][string]$ExportDir)
    $typePath = Join-Path (Get-CoordinationDir $ExportDir) "ExportType.txt"
    if (Test-Path $typePath) {
        return ([System.IO.File]::ReadAllText($typePath)).Trim()
    }
    return $null
}

#endregion

