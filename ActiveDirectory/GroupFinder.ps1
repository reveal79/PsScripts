<#
.SYNOPSIS
Retrieve all Active Directory groups that a specified user, computer, group, or service account belongs to, including nested groups.

.DESCRIPTION
This script recursively retrieves all Active Directory groups for a specified principal (user, computer, group, or service account).
The groups are displayed in an interactive grid view and sorted for better usability.

.SERVICE
Active Directory

.SERVICE TYPE
Group Management

.VERSION
1.0.0

.AUTHOR
Modified by: Don Cook
Original Credit: Brian Reich

.LAST UPDATED
2024-12-30

.PARAMETER dsn
The distinguished name (DN) of the user, computer, group, or service account to query.

.DEPENDENCIES
- Active Directory RSAT tools installed.
- Permissions to query Active Directory.

.EXAMPLE
# Retrieve all groups for a user
.\GroupFinder.ps1

This command prompts for a username and displays all associated groups in a grid view.

.NOTES
- The script recursively resolves nested groups.
- Outputs results in an interactive grid view and optionally exports to a file (future enhancement ready).

#>

# Import Active Directory Module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "The Active Directory module is not installed. Please install RSAT tools for Active Directory." -ForegroundColor Red
    exit
}
Import-Module ActiveDirectory

# Function to retrieve group membership recursively
function Get-ADPrincipalGroupMembershipRecursive {
    <#
    .SYNOPSIS
    Recursively retrieves all group memberships for a specified Active Directory principal.

    .PARAMETER dsn
    The distinguished name (DN) of the principal to query.

    .PARAMETER groups
    An array to store the groups during recursion (default is empty).

    .OUTPUTS
    An array of group objects representing all groups the principal is a member of.

    .NOTES
    This function resolves nested group memberships and avoids duplicates.
    #>

    Param(
        [string]$dsn,
        [array]$groups = @()
    )

    # Retrieve the AD object and its direct group memberships
    $obj = Get-ADObject -Identity $dsn -Properties memberOf -ErrorAction Stop

    foreach ($groupDsn in $obj.memberOf) {
        # Retrieve group details
        $tmpGrp = Get-ADObject -Identity $groupDsn -Properties memberOf -ErrorAction Stop

        # Avoid duplicates and recurse into nested groups
        if (($groups | Where-Object { $_.DistinguishedName -eq $groupDsn }).Count -eq 0) {
            $groups += $tmpGrp
            $groups = Get-ADPrincipalGroupMembershipRecursive -dsn $groupDsn -groups $groups
        }
    }

    return $groups
}

# Main Script Execution
Write-Host "Group Finder Script" -ForegroundColor Cyan
Write-Host "This script retrieves all Active Directory groups for a specified user, including nested groups." -ForegroundColor Green

try {
    # Prompt for username
    $username = Read-Host -Prompt "Enter the username to search (e.g., john.doe)"
    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host "Invalid username. Please try again." -ForegroundColor Yellow
        exit
    }

    # Get the user's distinguished name
    $user = Get-ADUser -Identity $username -Properties DistinguishedName -ErrorAction Stop
    $userDsn = $user.DistinguishedName
    Write-Host "Found user: $($user.Name). Fetching group memberships..." -ForegroundColor Green

    # Retrieve groups
    $groups = Get-ADPrincipalGroupMembershipRecursive -dsn $userDsn

    # Display results
    if ($groups.Count -gt 0) {
        Write-Host "Groups found for ${username}:" -ForegroundColor Green
        $groups | Sort-Object -Property Name | Out-GridView -Title "Groups for $username"
    } else {
        Write-Host "No groups found for $username." -ForegroundColor Yellow
    }
} catch {
    Write-Error "An error occurred: $_"
}

Write-Host "Script execution completed." -ForegroundColor Cyan