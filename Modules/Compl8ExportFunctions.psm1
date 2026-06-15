#Requires -Version 7.4
<#
.SYNOPSIS
    PowerShell module for Compl8 Cmdlet Export functions.

.DESCRIPTION
    This module provides functions for connecting to Security & Compliance PowerShell,
    exporting compliance configuration data, and handling pagination for large datasets.

    Based on current Microsoft Learn documentation:
    - https://learn.microsoft.com/powershell/module/exchange/connect-ippssession
    - https://learn.microsoft.com/powershell/module/exchange/export-contentexplorerdata
    - https://learn.microsoft.com/powershell/module/exchange/export-activityexplorerdata

.NOTES
    Version: 1.0.0
    Requires: PowerShell 7+, ExchangeOnlineManagement module 3.2.0+
#>


$moduleRoot = Split-Path -Parent $PSCommandPath
$projectRoot = Split-Path $moduleRoot -Parent
$sectionRoot = Join-Path $moduleRoot "Compl8ExportFunctions"
$sectionFiles = @(
    'Core\01-ScriptVariables.ps1'
    'Core\02-Logging.ps1'
    'Core\03-Connection.ps1'
    'Core\04-Paths.ps1'
    'Core\05-Configuration.ps1'
    'Core\06-ExportHelper.ps1'
    'Providers\01-ContentExplorer.ps1'
    'Providers\02-ActivityExplorer.ps1'
    'Providers\03-DlpAndPolicy.ps1'
    'Providers\04-Labels.ps1'
    'Providers\05-eDiscovery.ps1'
    'Providers\06-Rbac.ps1'
    'Core\07-ConfigurationValidation.ps1'
    'Core\08-ErrorHandling.ps1'
    'Core\09-AuthRecovery.ps1'
    'Core\10-Utility.ps1'
    'Providers\07-Tenant.ps1'
    'Providers\08-ContentExplorer.Aggregates.ps1'
    'Providers\09-ContentExplorer.WorkPlan.ps1'
    'Providers\10-ContentExplorer.RunTracker.ps1'
    'Providers\11-ContentExplorer.DetailExport.ps1'
    'Providers\12-ContentExplorer.Dedup.ps1'
    'Providers\13-ContentExplorer.Telemetry.ps1'
    'Providers\14-ContentExplorer.Progress.ps1'
    'Providers\15-ActivityExplorer.Retry.ps1'
    'Providers\16-ActivityExplorer.RunTracker.ps1'
    'Providers\17-ActivityExplorer.Export.ps1'
    'Orchestrator\01-PrivateHelper.ps1'
    'Orchestrator\02-ContentExplorer.RetryBucket.ps1'
    'Orchestrator\03-WorkerHealth.ps1'
    'Orchestrator\04-PhaseType.ps1'
    'Orchestrator\05-TaskCsv.ps1'
    'Orchestrator\06-FileDrop.ps1'
    'UI\01-Text.ps1'
    'UI\02-Dashboards.ps1'
    'Orchestrator\07-Dispatch.ps1'
    'Orchestrator\08-Watermarks.ps1'
    'Orchestrator\09-ContentExplorer.DispatchCallbacks.ps1'
    'Orchestrator\10-ContentExplorer.AggregatePaging.ps1'
    'Orchestrator\11-ContentExplorer.DetailExportParams.ps1'
)

foreach ($section in $sectionFiles) {
    $sectionPath = Join-Path $sectionRoot $section
    if (-not (Test-Path $sectionPath)) {
        throw "Module section not found: $sectionPath"
    }

    . $sectionPath
}

#region Module Exports

Export-ModuleMember -Function @(
    # Export Directory Path Helpers
    'ConvertTo-SafeDirectoryName',
    'ConvertTo-CsvField',
    'Get-DeterministicNameHash',
    'Get-CoordinationDir',
    'Get-CompletionsDir',
    'Get-WorkerCoordDir',
    'Get-LogsDir',
    'Get-CEDataDir',
    'Get-AEDataDir',
    'Get-CEClassifierDir',
    'Get-AEDayDir',
    'Initialize-ExportDirectories',
    'Write-CEManifest',
    'Write-AEManifest',

    # Logging
    'Initialize-ExportLog',
    'Write-ExportLog',
    'Get-ExportStatistics',
    'Add-ExportCount',

    # Connection
    'Test-ExportPrerequisites',
    'Connect-Compl8Compliance',
    'Disconnect-Compl8Compliance',

    # Configuration
    'Read-JsonConfig',
    'Get-EnabledItems',
    'Get-ContentExplorerSettings',
    'Test-ExportConfiguration',

    # File Export
    'ConvertTo-SerializableObject',
    'Export-ToJsonFile',
    'Export-ToCsvFile',

    # Error Handling
    'Write-ExportErrorLog',
    'Format-ErrorDetail',
    'Get-HttpErrorExplanation',

    # Auth Recovery
    'Invoke-WithAuthRecovery',
    'Invoke-WorkerReconnect',

    # Utility
    'Format-TimeSpan',

    # SIT and Tenant
    'Get-SITsToSkip',
    'Get-Compl8TenantInfo',
    'Get-SitGuidMapping',
    'Get-TrainableClassifiersFromCache',
    'Get-SitNamesFromRulePackXml',
    'Export-SitReferenceSnapshot',

    # Content Explorer - Basic
    'Export-ContentExplorer',

    # Content Explorer - Aggregate Discovery
    'Find-RecentAggregateCsv',
    'Save-AggregateMetadata',
    'Save-ExportSettings',
    'Get-ExportSettings',
    'Resolve-CEPageSize',
    'Resolve-AEFilters',
    'Get-TagNamesFromAggregateCsv',
    'Import-AggregateDataFromCsv',

    # Content Explorer - Work Plan & Export
    'New-ContentExplorerWorkPlan',
    'Export-ContentExplorerWithProgress',

    # Content Explorer - Run Tracker
    'Get-ContentExplorerRunTracker',
    'Save-ContentExplorerRunTracker',

    # Content Explorer - Deduplication
    'Remove-DuplicateContentRecordsV2',

    # Content Explorer - Telemetry & Adaptive Paging
    'New-ContentExplorerTelemetry',
    'Get-AdaptivePageSize',
    'Save-ContentExplorerTelemetry',
    'Get-ContentExplorerTelemetryStats',

    # Content Explorer - Progress
    'Get-ContentExplorerAggregateProgress',
    'Write-ContentExplorerProgress',

    # Content Explorer - Retry Bucket
    'Get-RetryBucketTasks',

    # Activity Explorer - Basic
    'Export-ActivityExplorer',

    # Activity Explorer - Resilient Export
    'Export-ActivityExplorerWithProgress',

    # Activity Explorer - Run Tracker
    'Get-ActivityExplorerRunTracker',
    'Save-ActivityExplorerRunTracker',

    # Activity Explorer - Merge
    'Merge-ActivityExplorerPages',

    # Activity Explorer - Helpers
    'Get-PageErrorMessage',
    'Test-PageHasContent',
    'Add-PartialError',
    'Get-RetryDelay',
    'Invoke-RetryWithBackoff',
    'Find-UnknownActivityTypes',
    'Get-ActivityExplorerFilters',

    # Shared Utility
    'Write-ProgressEntry',

    # Worker Health Monitoring
    'Test-WorkerAlive',
    'Get-WorkerState',

    # Phase/Type I/O
    'Write-ExportPhase',
    'Read-ExportPhase',
    'Write-ExportType',
    'Read-ExportType',

    # Task CSV I/O
    'Write-AETaskCsv',
    'Read-AETaskCsv',
    'Write-TaskCsv',
    'Read-TaskCsv',
    'Set-WorkerParked',
    'Test-WorkerParked',
    'Get-WatermarkPath',
    'Read-Watermarks',
    'Write-Watermarks',
    'Save-WatermarksFromDetailTasks',
    'Write-AggregateDeltaReport',
    'Get-ContentExplorerLocationType',
    'Get-ContentExplorerDetailPageSize',
    'New-ContentExplorerDetailTasks',
    'Get-RoundRobinDetailTaskOrder',
    'Write-RetryTasksCsv',
    'Show-RetryBucketSummary',
    'Write-RemainingTasksCsv',

    # File-Drop Coordination
    'Get-ExportRunSigningKey',
    'ConvertTo-SignedEnvelopeJson',
    'ConvertFrom-SignedEnvelopeJson',
    'Send-WorkerTask',
    'Receive-WorkerTask',
    'Complete-WorkerTask',

    # Text UI Helpers
    'Get-TerminalSize',
    'Format-ProgressBar',
    'Write-BoxTop',
    'Write-BoxBottom',
    'Write-BoxLine',
    'Write-BoxSeparator',
    'Get-BoxInnerWidth',
    'Write-SectionHeader',
    'Write-Banner',
    'Write-DashboardFrame',

    # Dashboard Functions
    'Reset-OrchestratorDashboard',
    'Reset-AEDashboard',
    'Get-ProgressEta',
    'Show-OrchestratorDashboard',
    'Show-AEDashboard',

    # Dispatch Loop Engine
    'Invoke-DispatchLoop',
    'Select-LargestPendingTask',

    # Content Explorer Dispatch Callback Helpers
    'Read-CEDetailSignals',
    'Find-CEDetailTaskMatch',
    'New-CEDetailDispatchPayload',

    # Content Explorer Aggregate Pagination (shared loop)
    'Invoke-CEAggregatePaging',

    # Content Explorer Detail Export Parameter Builder (shared splat for single-terminal loops)
    'Build-CEDetailExportParams',

    # DLP
    'Export-DlpPolicies',
    'Export-SensitiveInfoTypes',

    # Labels
    'Export-SensitivityLabels',
    'Export-RetentionLabels',

    # eDiscovery
    'Export-eDiscoveryCases',

    # RBAC
    'Export-RbacConfiguration'
)

#endregion
