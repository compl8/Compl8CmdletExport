#region Configuration Functions

function Read-JsonConfig {
    <#
    .SYNOPSIS
        Reads and parses a JSON configuration file.
    .PARAMETER Path
        Path to the JSON file.
    .PARAMETER Required
        Throw error if file doesn't exist.
    .OUTPUTS
        PSCustomObject from parsed JSON.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$Required
    )

    if (-not (Test-Path $Path)) {
        if ($Required) {
            throw "Required configuration file not found: $Path"
        }
        Write-ExportLog -Message "Config file not found: $Path (using defaults)" -Level Warning
        return $null
    }

    try {
        $content = Get-Content -Raw -Path $Path -ErrorAction Stop
        $config = ConvertFrom-Json -InputObject $content -ErrorAction Stop
        return $config
    }
    catch {
        Write-ExportLog -Message "Failed to parse config file: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-EnabledItems {
    <#
    .SYNOPSIS
        Returns items from a config object where value is "True".
    .PARAMETER Config
        PSCustomObject with True/False string values.
    .OUTPUTS
        Array of enabled property names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config
    )

    $enabled = @()
    foreach ($prop in $Config.PSObject.Properties) {
        if ($prop.Value -eq "True") {
            $enabled += $prop.Name
        }
    }
    return $enabled
}

function Get-ContentExplorerSettings {
    <#
    .SYNOPSIS
        Normalizes Content Explorer settings from saved manifests and config files.
    .DESCRIPTION
        Saved settings win, then the current config file, then context-specific defaults.
        Supports both the current Settings.{BatchSize,PageSize,Workloads} shape and
        the legacy _BatchSize/_PageSize/_Workloads shape for backward compatibility.
    #>
    [CmdletBinding()]
    param(
        [PSCustomObject]$ConfigObject,
        $SavedSettings,
        [int]$DefaultBatchSize = 10,
        [string[]]$DefaultWorkloads = @("SharePoint", "OneDrive"),
        [int]$DefaultPageSize = 100
    )

    $configSettings = if ($ConfigObject -and $ConfigObject.Settings) { $ConfigObject.Settings } else { $null }

    $batchSize = $null
    if ($SavedSettings -and $SavedSettings.BatchSize) {
        $batchSize = $SavedSettings.BatchSize -as [int]
    }
    elseif ($configSettings -and $configSettings.BatchSize) {
        $batchSize = $configSettings.BatchSize -as [int]
    }
    elseif ($ConfigObject -and $ConfigObject._BatchSize) {
        $batchSize = $ConfigObject._BatchSize -as [int]
    }
    if (-not $batchSize -or $batchSize -lt 1) {
        $batchSize = $DefaultBatchSize
    }

    $workloads = @()
    if ($SavedSettings -and $SavedSettings.Workloads) {
        $workloads = @($SavedSettings.Workloads)
    }
    elseif ($configSettings -and $configSettings.Workloads) {
        $workloads = @($configSettings.Workloads)
    }
    elseif ($ConfigObject -and $ConfigObject._Workloads) {
        $workloads = @($ConfigObject._Workloads)
    }
    if ($workloads.Count -eq 0) {
        $workloads = @($DefaultWorkloads)
    }

    $pageSize = $null
    if ($SavedSettings -and $SavedSettings.PageSize) {
        $pageSize = $SavedSettings.PageSize -as [int]
    }
    elseif ($configSettings -and $configSettings.PageSize) {
        $pageSize = $configSettings.PageSize -as [int]
    }
    elseif ($ConfigObject -and $ConfigObject._PageSize) {
        $pageSize = $ConfigObject._PageSize -as [int]
    }
    if (-not $pageSize -or $pageSize -lt 1) {
        $pageSize = $DefaultPageSize
    }

    return [PSCustomObject]@{
        BatchSize = $batchSize
        Workloads = @($workloads)
        PageSize  = $pageSize
    }
}

#endregion

