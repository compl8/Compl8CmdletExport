#region Script Variables
$script:LogFile = $null
$script:SessionStartTime = $null
$script:ExportStats = @{
    ItemsExported = @{}
    Errors = [System.Collections.ArrayList]::new()
    Warnings = [System.Collections.ArrayList]::new()
}
$script:DashboardLineCount = 0
$script:AEDashboardLineCount = 0
#endregion

