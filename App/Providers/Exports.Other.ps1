function Invoke-eDiscoveryExport {
    Write-ExportLog -Message "`n========== eDiscovery Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        Cases = @()
        Searches = @()
    }

    try {
        $ediscovery = Export-eDiscoveryCases
        $exportResult.Cases = $ediscovery.Cases
        $exportResult.Searches = $ediscovery.Searches
    }
    catch {
        Write-ExportLog -Message "eDiscovery export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "eDiscovery-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        if (@($exportResult.Cases).Count -gt 0) {
            Save-ExportData -Data $exportResult.Cases -Name "eDiscovery-Cases" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.Searches).Count -gt 0) {
            Save-ExportData -Data $exportResult.Searches -Name "eDiscovery-Searches" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

function Invoke-RbacExport {
    Write-ExportLog -Message "`n========== RBAC Export ==========" -Level Info

    $exportResult = @{
        ExportTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        RoleGroups = @()
        Members = @()
    }

    try {
        $rbac = Export-RbacConfiguration
        $exportResult.RoleGroups = $rbac.RoleGroups
        $exportResult.Members = $rbac.Members
    }
    catch {
        Write-ExportLog -Message "RBAC export failed: $($_.Exception.Message)" -Level Error
    }

    if ($OutputFormat -eq "JSON") {
        Save-ExportData -Data $exportResult -Name "RBAC-Config" -Format $OutputFormat -Directory $script:ExportRunDirectory
    }
    else {
        if (@($exportResult.RoleGroups).Count -gt 0) {
            Save-ExportData -Data $exportResult.RoleGroups -Name "RBAC-RoleGroups" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
        if (@($exportResult.Members).Count -gt 0) {
            Save-ExportData -Data $exportResult.Members -Name "RBAC-Members" -Format $OutputFormat -Directory $script:ExportRunDirectory
        }
    }
}

#endregion

