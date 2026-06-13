#region Script Variables
$script:LogFile = $null
$script:SessionStartTime = $null
$script:ExportStats = @{
    ItemsExported = @{}
    Errors = [System.Collections.ArrayList]::new()
    Warnings = [System.Collections.ArrayList]::new()
}
# Maximum entries retained in ExportStats.Errors and ExportStats.Warnings (oldest dropped when exceeded)
$script:ExportStatsMaxEntries = 500
$script:DashboardLineCount = 0
$script:AEDashboardLineCount = 0
#endregion

