#region eDiscovery Functions

function Export-eDiscoveryCases {
    <#
    .SYNOPSIS
        Exports eDiscovery cases and searches.

    .OUTPUTS
        Hashtable with Cases and Searches arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Cases = @()
        Searches = @()
    }

    Write-ExportLog -Message "Exporting eDiscovery Cases..." -Level Info

    try {
        # Get Compliance Cases
        Write-ExportLog -Message "  Retrieving compliance cases..." -Level Info
        $cases = Get-ComplianceCase -ErrorAction Stop
        $result.Cases = $cases
        $caseCount = @($cases).Count
        Add-ExportCount -Category "eDiscoveryCases" -Count $caseCount
        Write-ExportLog -Message "  Found $caseCount compliance cases" -Level Success

        # Get Compliance Searches
        Write-ExportLog -Message "  Retrieving compliance searches..." -Level Info
        $searches = Get-ComplianceSearch -ErrorAction Stop
        $result.Searches = $searches
        $searchCount = @($searches).Count
        Add-ExportCount -Category "ComplianceSearches" -Count $searchCount
        Write-ExportLog -Message "  Found $searchCount compliance searches" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

#endregion

