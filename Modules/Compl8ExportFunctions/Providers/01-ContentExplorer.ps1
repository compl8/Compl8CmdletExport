#region Content Explorer Functions

function Export-ContentExplorer {
    <#
    .SYNOPSIS
        Exports Content Explorer data with automatic pagination.

    .DESCRIPTION
        Exports data classification file details from Microsoft Purview Content Explorer.
        Handles pagination automatically for large datasets.

        Reference: https://learn.microsoft.com/powershell/module/exchange/export-contentexplorerdata

    .PARAMETER TagType
        Type of classifier. Valid values: Retention, SensitiveInformationType, Sensitivity, TrainableClassifier

    .PARAMETER TagName
        Name of the label/classifier to export.

    .PARAMETER Workload
        Location to export from. Valid values: Exchange, EXO, OneDrive, ODB, SharePoint, SPO, Teams

    .PARAMETER PageSize
        Number of records per page (1-10000). Default: 5000

    .PARAMETER ConfidenceLevel
        Filter by confidence level: low, medium, high

    .PARAMETER Aggregate
        Return folder-level aggregates instead of item details.

    .OUTPUTS
        Array of exported records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Retention", "SensitiveInformationType", "Sensitivity", "TrainableClassifier")]
        [string]$TagType,

        [Parameter(Mandatory)]
        [string]$TagName,

        [ValidateSet("Exchange", "EXO", "OneDrive", "ODB", "SharePoint", "SPO", "Teams")]
        [string]$Workload,

        [ValidateRange(1, 10000)]
        [int]$PageSize = 5000,

        [ValidateSet("low", "medium", "high")]
        [string]$ConfidenceLevel,

        [switch]$Aggregate
    )

    $allRecords = [System.Collections.ArrayList]::new()
    $pageNumber = 0
    $totalExported = 0

    # Build parameters
    $exportParams = @{
        TagType = $TagType
        TagName = $TagName
        PageSize = $PageSize
        ErrorAction = 'Stop'
    }

    if ($Workload) { $exportParams['Workload'] = $Workload }
    if ($ConfidenceLevel) { $exportParams['ConfidenceLevel'] = $ConfidenceLevel }
    if ($Aggregate) { $exportParams['Aggregate'] = $true }

    $workloadDisplay = if ($Workload) { $Workload } else { "All" }
    Write-ExportLog -Message "Exporting Content Explorer: $TagType/$TagName (Workload: $workloadDisplay)" -Level Info

    try {
        # Initial query
        $result = Export-ContentExplorerData @exportParams

        # Check if we got results (result[0] contains metadata)
        if ($null -eq $result -or $result.Count -eq 0) {
            Write-ExportLog -Message "  No results returned" -Level Warning
            return @()
        }

        $totalCount = $result[0].TotalCount
        if ($totalCount -eq 0) {
            Write-ExportLog -Message "  No matching records found" -Level Warning
            return @()
        }

        Write-ExportLog -Message "  Total matches: $totalCount" -Level Info

        # Process pages
        do {
            $pageNumber++
            $recordsInPage = $result[0].RecordsReturned

            # Records are in array items 1 onwards (item 0 is metadata)
            if ($recordsInPage -gt 0) {
                $pageRecords = $result[1..$recordsInPage]
                [void]$allRecords.AddRange(@($pageRecords))
                $totalExported += $recordsInPage
            }

            Write-ExportLog -Message "  Page $pageNumber`: $recordsInPage records (Total: $totalExported)" -Level Info

            # Check for more pages
            if ($result[0].MorePagesAvailable -eq $true -or $result[0].MorePagesAvailable -eq "True") {
                $exportParams['PageCookie'] = $result[0].PageCookie
                $result = Export-ContentExplorerData @exportParams
            }
            else {
                break
            }
        } while ($true)

        Add-ExportCount -Category "ContentExplorer_$TagType" -Count $totalExported
        Write-ExportLog -Message "  Completed: $totalExported records exported" -Level Success

        return $allRecords
    }
    catch {
        Write-ExportLog -Message "  Export failed: $($_.Exception.Message)" -Level Error
        return @()
    }
}

#endregion

