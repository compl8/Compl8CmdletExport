#region Label Functions

function Export-SensitivityLabels {
    <#
    .SYNOPSIS
        Exports sensitivity labels and label policies.

    .OUTPUTS
        Hashtable with Labels and Policies arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Labels = @()
        Policies = @()
    }

    Write-ExportLog -Message "Exporting Sensitivity Labels..." -Level Info

    try {
        # Get Labels
        Write-ExportLog -Message "  Retrieving sensitivity labels..." -Level Info
        $labels = Get-Label -ErrorAction Stop
        $result.Labels = $labels
        $labelCount = @($labels).Count
        Add-ExportCount -Category "SensitivityLabels" -Count $labelCount
        Write-ExportLog -Message "  Found $labelCount sensitivity labels" -Level Success

        # Get Label Policies
        Write-ExportLog -Message "  Retrieving label policies..." -Level Info
        $policies = Get-LabelPolicy -ErrorAction Stop
        $result.Policies = $policies
        $policyCount = @($policies).Count
        Add-ExportCount -Category "LabelPolicies" -Count $policyCount
        Write-ExportLog -Message "  Found $policyCount label policies" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

function Export-RetentionLabels {
    <#
    .SYNOPSIS
        Exports retention labels and policies.

    .OUTPUTS
        Hashtable with Labels, Policies, and Rules arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        Labels = @()
        Policies = @()
        Rules = @()
    }

    Write-ExportLog -Message "Exporting Retention Labels..." -Level Info

    try {
        # Get Retention Labels (Compliance Tags)
        Write-ExportLog -Message "  Retrieving retention labels (compliance tags)..." -Level Info
        $labels = Get-ComplianceTag -ErrorAction Stop
        $result.Labels = $labels
        $labelCount = @($labels).Count
        Add-ExportCount -Category "RetentionLabels" -Count $labelCount
        Write-ExportLog -Message "  Found $labelCount retention labels" -Level Success

        # Get Retention Policies
        Write-ExportLog -Message "  Retrieving retention policies..." -Level Info
        $policies = Get-RetentionCompliancePolicy -ErrorAction Stop
        $result.Policies = $policies
        $policyCount = @($policies).Count
        Add-ExportCount -Category "RetentionPolicies" -Count $policyCount
        Write-ExportLog -Message "  Found $policyCount retention policies" -Level Success

        # Get Retention Rules
        Write-ExportLog -Message "  Retrieving retention rules..." -Level Info
        $rules = Get-RetentionComplianceRule -ErrorAction Stop
        $result.Rules = $rules
        $ruleCount = @($rules).Count
        Add-ExportCount -Category "RetentionRules" -Count $ruleCount
        Write-ExportLog -Message "  Found $ruleCount retention rules" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

#endregion

