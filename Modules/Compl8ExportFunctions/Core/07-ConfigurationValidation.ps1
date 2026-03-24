#region Configuration Validation

function Test-ExportConfiguration {
    <#
    .SYNOPSIS
        Validates that required configuration files exist for the chosen export mode.

    .DESCRIPTION
        Checks for the presence of JSON configuration files needed by the selected
        export mode before establishing a connection to Security & Compliance PowerShell.
        This prevents wasted authentication time when config files are missing.

        For FullExport mode, Content Explorer and Activity Explorer config files are
        checked unless the corresponding -NoContent or -NoActivity switches are set.

    .PARAMETER ExportMode
        The export mode to validate. Valid values: Full, DLP, Labels, ContentExplorer,
        ActivityExplorer, eDiscovery, RBAC.

    .PARAMETER ScriptRoot
        Root directory of the export script, used to locate ConfigFiles subfolder.

    .PARAMETER NoContent
        When set, skips validation of Content Explorer config files (used with FullExport).

    .PARAMETER NoActivity
        When set, skips validation of Activity Explorer config files (used with FullExport).

    .OUTPUTS
        Boolean. Returns $true if all required config files exist, $false otherwise.

    .EXAMPLE
        Test-ExportConfiguration -ExportMode "ContentExplorer" -ScriptRoot $PSScriptRoot

    .EXAMPLE
        Test-ExportConfiguration -ExportMode "Full" -ScriptRoot $PSScriptRoot -NoActivity
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Full", "DLP", "Labels", "ContentExplorer", "ActivityExplorer", "eDiscovery", "RBAC")]
        [string]$ExportMode,

        [Parameter(Mandatory)]
        [string]$ScriptRoot,

        [switch]$NoContent,

        [switch]$NoActivity
    )

    $configDir = Join-Path $ScriptRoot "ConfigFiles"
    $allValid = $true

    # Build the list of required config files based on export mode
    $requiredConfigs = @()

    switch ($ExportMode) {
        "ContentExplorer" {
            $requiredConfigs += @{
                Name = "ContentExplorerClassifiers.json"
                Description = "Content Explorer classifier configuration"
            }
        }
        "ActivityExplorer" {
            $requiredConfigs += @{
                Name = "ActivityExplorerSelector.json"
                Description = "Activity Explorer activity/workload filter"
            }
        }
        "Full" {
            if (-not $NoContent) {
                $requiredConfigs += @{
                    Name = "ContentExplorerClassifiers.json"
                    Description = "Content Explorer classifier configuration"
                }
            }
            if (-not $NoActivity) {
                $requiredConfigs += @{
                    Name = "ActivityExplorerSelector.json"
                    Description = "Activity Explorer activity/workload filter"
                }
            }
        }
        # DLP, Labels, eDiscovery, RBAC modes do not require additional config files
    }

    # Validate each required config file
    foreach ($config in $requiredConfigs) {
        $configPath = Join-Path $configDir $config.Name
        if (Test-Path $configPath) {
            Write-ExportLog -Message "  Config OK: $($config.Name)" -Level Success
        }
        else {
            Write-ExportLog -Message "  Config MISSING: $($config.Name) ($($config.Description))" -Level Error
            Write-ExportLog -Message "    Expected at: $configPath" -Level Info
            $allValid = $false
        }
    }

    if ($allValid) {
        Write-ExportLog -Message "Configuration validation passed" -Level Success
    }
    else {
        Write-ExportLog -Message "Configuration validation failed - missing required files" -Level Error
    }

    return $allValid
}

#endregion

