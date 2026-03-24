#region DLP and Policy Functions

function Export-DlpPolicies {
    <#
    .SYNOPSIS
        Exports DLP compliance policies and rules.

    .DESCRIPTION
        Retrieves all DLP policies and their associated rules from Microsoft Purview.

    .OUTPUTS
        Hashtable with Policies and Rules arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Policies = @()
        Rules = @()
    }

    Write-ExportLog -Message "Exporting DLP Policies..." -Level Info

    try {
        # Get DLP Policies
        Write-ExportLog -Message "  Retrieving DLP policies..." -Level Info
        $policies = Get-DlpCompliancePolicy -ErrorAction Stop
        $result.Policies = $policies
        $policyCount = @($policies).Count
        Add-ExportCount -Category "DlpPolicies" -Count $policyCount
        Write-ExportLog -Message "  Found $policyCount DLP policies" -Level Success

        # Get DLP Rules
        Write-ExportLog -Message "  Retrieving DLP rules..." -Level Info
        $rules = Get-DlpComplianceRule -ErrorAction Stop
        $result.Rules = $rules
        $ruleCount = @($rules).Count
        Add-ExportCount -Category "DlpRules" -Count $ruleCount
        Write-ExportLog -Message "  Found $ruleCount DLP rules" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

function Export-SensitiveInfoTypes {
    <#
    .SYNOPSIS
        Exports Sensitive Information Type definitions.

    .OUTPUTS
        Array of SIT definitions.
    #>
    [CmdletBinding()]
    param()

    Write-ExportLog -Message "Exporting Sensitive Information Types..." -Level Info

    try {
        Write-ExportLog -Message "  Retrieving sensitive information types..." -Level Info
        $sits = Get-DlpSensitiveInformationType -ErrorAction Stop
        $sitCount = @($sits).Count
        Add-ExportCount -Category "SensitiveInfoTypes" -Count $sitCount
        Write-ExportLog -Message "  Found $sitCount sensitive information types" -Level Success
        return $sits
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
        return @()
    }
}

#endregion

