#
# Module manifest for module 'Compl8ExportFunctions'
#
# Generated for: Compl8 Cmdlet Export Tool
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'Compl8ExportFunctions.psm1'

# Version number of this module.
ModuleVersion = '1.0.0'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = 'fd29393d-568a-4a2f-8959-d7e931ad6735'

# Author of this module
Author = 'Compl8'

# Company or vendor of this module
CompanyName = 'Compl8'

# Copyright statement for this module
Copyright = '(c) Compl8. All rights reserved.'

# Description of the functionality provided by this module
Description = 'PowerShell module for the Compl8 Cmdlet Export Tool. Provides functions for connecting to Security & Compliance PowerShell (via ExchangeOnlineManagement), exporting Microsoft Purview compliance configuration data (Content Explorer, Activity Explorer, DLP policies, sensitivity labels, retention labels, eDiscovery, RBAC), and orchestrating multi-terminal parallel exports with file-drop worker coordination.'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '7.4'

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(
    @{
        ModuleName    = 'ExchangeOnlineManagement'
        ModuleVersion = '3.2.0'
    }
)

# Functions to export from this module — derived from the live Import-Module enumeration (126 functions)
FunctionsToExport = @(
    'Add-ExportCount',
    'Add-PartialError',
    'Build-CEDetailExportParams',
    'Complete-WorkerTask',
    'Connect-Compl8Compliance',
    'ConvertFrom-SignedEnvelopeJson',
    'ConvertTo-CsvField',
    'ConvertTo-SafeDirectoryName',
    'ConvertTo-SerializableObject',
    'ConvertTo-SignedEnvelopeJson',
    'Disconnect-Compl8Compliance',
    'Export-ActivityExplorer',
    'Export-ActivityExplorerWithProgress',
    'Export-ContentExplorer',
    'Export-ContentExplorerWithProgress',
    'Export-DlpPolicies',
    'Export-eDiscoveryCases',
    'Export-RbacConfiguration',
    'Export-RetentionLabels',
    'Export-SensitiveInfoTypes',
    'Export-SensitivityLabels',
    'Export-SitReferenceSnapshot',
    'Export-ToCsvFile',
    'Export-ToJsonFile',
    'Find-CEDetailTaskMatch',
    'Find-RecentAggregateCsv',
    'Find-UnknownActivityTypes',
    'Format-ErrorDetail',
    'Format-ProgressBar',
    'Format-TimeSpan',
    'Get-ActivityExplorerFilters',
    'Get-ActivityExplorerRunTracker',
    'Get-AdaptivePageSize',
    'Get-AEDataDir',
    'Get-AEDayDir',
    'Get-BoxInnerWidth',
    'Get-CEClassifierDir',
    'Get-CEDataDir',
    'Get-Compl8TenantInfo',
    'Get-CompletionsDir',
    'Get-ContentExplorerAggregateProgress',
    'Get-ContentExplorerDetailPageSize',
    'Get-ContentExplorerLocationType',
    'Get-ContentExplorerRunTracker',
    'Get-ContentExplorerSettings',
    'Get-ContentExplorerTelemetryStats',
    'Get-CoordinationDir',
    'Get-DeterministicNameHash',
    'Get-EnabledItems',
    'Get-ExportExitCode',
    'Get-ExportRunSigningKey',
    'Get-ExportSettings',
    'Get-ExportStatistics',
    'Get-HttpErrorExplanation',
    'Get-LogsDir',
    'Get-PageErrorMessage',
    'Get-RetryBucketTasks',
    'Get-RetryDelay',
    'Get-RoundRobinDetailTaskOrder',
    'Get-SitGuidMapping',
    'Get-SitNamesFromRulePackXml',
    'Get-SITsToSkip',
    'Get-TagNamesFromAggregateCsv',
    'Get-TerminalSize',
    'Get-TrainableClassifiersFromCache',
    'Get-WatermarkPath',
    'Get-WorkerCoordDir',
    'Get-WorkerState',
    'Import-AggregateDataFromCsv',
    'Initialize-ExportDirectories',
    'Initialize-ExportLog',
    'Invoke-CEAggregatePaging',
    'Invoke-DispatchLoop',
    'Invoke-RetryWithBackoff',
    'Invoke-WithAuthRecovery',
    'Invoke-WorkerReconnect',
    'Merge-ActivityExplorerPages',
    'New-CEDetailDispatchPayload',
    'New-ContentExplorerDetailTasks',
    'New-ContentExplorerTelemetry',
    'New-ContentExplorerWorkPlan',
    'Read-AETaskCsv',
    'Read-CEDetailSignals',
    'Read-ExportPhase',
    'Read-ExportType',
    'Read-JsonConfig',
    'Read-TaskCsv',
    'Read-Watermarks',
    'Receive-WorkerTask',
    'Remove-DuplicateContentRecordsV2',
    'Reset-AEDashboard',
    'Reset-OrchestratorDashboard',
    'Resolve-AEFilters',
    'Resolve-CEPageSize',
    'Save-ActivityExplorerRunTracker',
    'Save-AggregateMetadata',
    'Save-ContentExplorerRunTracker',
    'Save-ContentExplorerTelemetry',
    'Save-ExportSettings',
    'Save-WatermarksFromDetailTasks',
    'Select-LargestPendingTask',
    'Send-WorkerTask',
    'Set-WorkerParked',
    'Show-AEDashboard',
    'Show-OrchestratorDashboard',
    'Show-RetryBucketSummary',
    'Test-ExportConfiguration',
    'Test-ExportPrerequisites',
    'Test-PageHasContent',
    'Test-WorkerAlive',
    'Test-WorkerParked',
    'Write-AEManifest',
    'Write-AETaskCsv',
    'Write-AggregateDeltaReport',
    'Write-Banner',
    'Write-BoxBottom',
    'Write-BoxLine',
    'Write-BoxSeparator',
    'Write-BoxTop',
    'Write-CEManifest',
    'Write-ContentExplorerProgress',
    'Write-DashboardFrame',
    'Write-ExportErrorLog',
    'Write-ExportLog',
    'Write-ExportPhase',
    'Write-ExportType',
    'Write-ProgressEntry',
    'Write-RemainingTasksCsv',
    'Write-RetryTasksCsv',
    'Write-RunSummary',
    'Write-SectionHeader',
    'Write-TaskCsv',
    'Write-Watermarks'
)

# Cmdlets to export from this module
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = @()

# Aliases to export from this module
AliasesToExport = @()

# Private data to pass to the module specified in RootModule/NestedModules.
PrivateData = @{
    PSData = @{
        # Tags applied to this module.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        # ProjectUri = ''

        # ReleaseNotes of this module.
        # ReleaseNotes = ''
    }
}

}
