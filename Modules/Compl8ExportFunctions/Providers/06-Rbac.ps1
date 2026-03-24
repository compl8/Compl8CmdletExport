#region RBAC Functions

function Export-RbacConfiguration {
    <#
    .SYNOPSIS
        Exports RBAC role groups and assignments.

    .OUTPUTS
        Hashtable with RoleGroups and Members arrays.
    #>
    [CmdletBinding()]
    param()

    $result = @{
        RoleGroups = @()
        Members = @()
    }

    Write-ExportLog -Message "Exporting RBAC Configuration..." -Level Info

    try {
        # Get Role Groups
        Write-ExportLog -Message "  Retrieving role groups..." -Level Info
        $roleGroups = Get-RoleGroup -ErrorAction Stop
        $result.RoleGroups = $roleGroups
        $totalGroups = @($roleGroups).Count
        Add-ExportCount -Category "RoleGroups" -Count $totalGroups
        Write-ExportLog -Message "  Found $totalGroups role groups" -Level Success

        # Get Members for each role group with progress
        Write-ExportLog -Message "  Retrieving members for each role group..." -Level Info
        $allMembers = @()
        $currentGroup = 0

        foreach ($rg in $roleGroups) {
            $currentGroup++
            $percentComplete = [math]::Round(($currentGroup / $totalGroups) * 100, 0)
            Write-ExportLog -Message "    [$currentGroup/$totalGroups] ($percentComplete%) Processing: $($rg.Name)" -Level Info

            try {
                $members = Get-RoleGroupMember -Identity $rg.Name -ErrorAction SilentlyContinue
                if ($members) {
                    $memberCount = @($members).Count
                    foreach ($member in $members) {
                        $allMembers += [PSCustomObject]@{
                            RoleGroup = $rg.Name
                            MemberName = $member.Name
                            MemberType = $member.RecipientType
                        }
                    }
                    if ($memberCount -gt 0) {
                        Write-ExportLog -Message "      -> $memberCount members" -Level Info
                    }
                }
            }
            catch {
                Write-ExportLog -Message "      -> Failed to read members: $($_.Exception.Message)" -Level Warning
            }
        }
        $result.Members = $allMembers
        Add-ExportCount -Category "RoleGroupMembers" -Count $allMembers.Count
        Write-ExportLog -Message "  Completed: $($allMembers.Count) total role group members" -Level Success
    }
    catch {
        Write-ExportLog -Message "  Failed: $($_.Exception.Message)" -Level Error
    }

    return $result
}

#endregion

