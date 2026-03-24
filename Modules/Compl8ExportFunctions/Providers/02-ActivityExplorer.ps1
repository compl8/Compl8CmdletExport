#region Activity Explorer Functions

function Export-ActivityExplorer {
    <#
    .SYNOPSIS
        Exports Activity Explorer data with automatic pagination.

    .DESCRIPTION
        Exports activity data from Microsoft Purview Activity Explorer.
        Handles pagination automatically for large datasets.

        Reference: https://learn.microsoft.com/powershell/module/exchange/export-activityexplorerdata

    .PARAMETER StartTime
        Start of the date range to export.

    .PARAMETER EndTime
        End of the date range to export.

    .PARAMETER OutputFormat
        Output format: Json or Csv. Default: Json

    .PARAMETER PageSize
        Records per page (1-5000). Default: 5000

    .PARAMETER Filters
        Hashtable of filters. Keys are filter names, values are arrays of values.
        Example: @{ Activity = @("DLPRuleMatch"); Workload = @("Exchange", "SharePoint") }

    .OUTPUTS
        Array of activity records.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$StartTime,

        [Parameter(Mandatory)]
        [datetime]$EndTime,

        [ValidateSet("Json", "Csv")]
        [string]$OutputFormat = "Json",

        [ValidateRange(1, 5000)]
        [int]$PageSize = 5000,

        [hashtable]$Filters
    )

    $allRecords = [System.Collections.ArrayList]::new()
    $pageNumber = 0
    $totalExported = 0

    # Format dates for the cmdlet
    $startStr = $StartTime.ToString("MM/dd/yyyy HH:mm:ss")
    $endStr = $EndTime.ToString("MM/dd/yyyy HH:mm:ss")

    # Build parameters
    $exportParams = @{
        StartTime = $startStr
        EndTime = $endStr
        OutputFormat = $OutputFormat
        PageSize = $PageSize
        ErrorAction = 'Stop'
    }

    # Add filters (up to 5 supported by the API)
    if ($Filters -and $Filters.Count -gt 0) {
        if ($Filters.Count -gt 5) {
            Write-Warning "Activity Explorer API supports max 5 filters per category. $($Filters.Count) items enabled — only the first 5 will be used. Disable some items in the config file."
        }
        $filterIndex = 1
        foreach ($filterName in $Filters.Keys) {
            if ($filterIndex -le 5) {
                $filterValues = @($filterName) + @($Filters[$filterName])
                $exportParams["Filter$filterIndex"] = $filterValues
                $filterIndex++
            }
        }
    }

    Write-ExportLog -Message "Exporting Activity Explorer: $startStr to $endStr" -Level Info
    if ($Filters) {
        Write-ExportLog -Message "  Filters: $($Filters.Keys -join ', ')" -Level Info
    }

    try {
        # Initial query
        Write-ExportLog -Message "  Calling Export-ActivityExplorerData..." -Level Info
        $result = Export-ActivityExplorerData @exportParams

        if ($null -eq $result) {
            Write-ExportLog -Message "  Result is null - no data returned from cmdlet" -Level Warning
            return @()
        }

        # Log result object properties for debugging
        Write-ExportLog -Message "  Result type: $($result.GetType().Name)" -Level Info
        Write-ExportLog -Message "  Result properties: $($result.PSObject.Properties.Name -join ', ')" -Level Info

        $totalCount = $result.TotalResultCount
        Write-ExportLog -Message "  TotalResultCount: $totalCount" -Level Info

        if ($null -eq $totalCount -or $totalCount -eq 0 -or $totalCount -eq "0") {
            Write-ExportLog -Message "  No matching activities found (TotalResultCount is 0 or null)" -Level Warning
            return @()
        }

        Write-ExportLog -Message "  Total activities to retrieve: $totalCount" -Level Info

        # Process pages
        do {
            $pageNumber++

            # Parse ResultData (JSON string)
            if ($result.ResultData) {
                Write-ExportLog -Message "  ResultData length: $($result.ResultData.Length) chars" -Level Info
                $pageRecords = $result.ResultData | ConvertFrom-Json
                if ($pageRecords) {
                    $recordCount = @($pageRecords).Count
                    [void]$allRecords.AddRange(@($pageRecords))
                    $totalExported += $recordCount
                    Write-ExportLog -Message "  Page $pageNumber`: $recordCount records (Total: $totalExported)" -Level Info
                }
            }
            else {
                Write-ExportLog -Message "  Page $pageNumber`: ResultData is empty or null" -Level Warning
            }

            # Check for more pages (LastPage property)
            Write-ExportLog -Message "  LastPage: $($result.LastPage), WaterMark: $($result.WaterMark)" -Level Info
            if ($result.LastPage -eq $false) {
                # PageCookie valid for 120 seconds per documentation
                $exportParams['PageCookie'] = $result.WaterMark
                $result = Export-ActivityExplorerData @exportParams
            }
            else {
                break
            }
        } while ($true)

        Add-ExportCount -Category "ActivityExplorer" -Count $totalExported
        Write-ExportLog -Message "  Completed: $totalExported records exported" -Level Success

        return $allRecords
    }
    catch {
        Write-ExportLog -Message "  Export failed: $($_.Exception.Message)" -Level Error
        return @()
    }
}

#endregion

