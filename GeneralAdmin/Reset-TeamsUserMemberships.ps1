<#
.SYNOPSIS
    Resets a user's Microsoft Teams memberships by removing and re-adding them.

.DESCRIPTION
    This script retrieves a user's current Microsoft Teams memberships and roles, 
    removes them from all teams, and then re-adds them with the same roles.
    
    It helps resolve access or permission issues by effectively refreshing the user's membership.

.NOTES
    Author: Don Cook
    Created: 2025-03-03
    Version: 1.0
    Purpose: Used for troubleshooting Microsoft Teams user access issues.

.REQUIREMENTS
    - Requires the MicrosoftTeams PowerShell module.
    - The user running the script must have the necessary permissions to manage Teams.
    - Must be run in a PowerShell session with administrative privileges.

#>

# Ensure the Microsoft Teams module is installed
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Installing MicrosoftTeams module..."
    Install-Module -Name MicrosoftTeams -Force -AllowClobber
}

# Connect to Microsoft Teams if not already connected
if (-not (Get-Team -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Microsoft Teams..."
    Connect-MicrosoftTeams
}

# Prompt for the user's email (User Principal Name)
$UserPrincipalName = Read-Host "Enter the User Principal Name (email)"

# Step 1: Retrieve user's Teams memberships and roles
Write-Host "Retrieving teams for user $UserPrincipalName..."
$UserTeams = Get-Team | ForEach-Object {
    $team = $_
    $members = Get-TeamUser -GroupId $team.GroupId
    $user = $members | Where-Object { $_.User -eq $UserPrincipalName }

    if ($user) {
        [PSCustomObject]@{
            TeamName  = $team.DisplayName  # Friendly team name
            GroupId   = $team.GroupId      # Team's unique ID
            Role      = $user.Role         # User's role in the team (Member/Owner)
        }
    }
} | Where-Object { $_ }  # Filter out null results

# Check if the user is part of any Teams
if ($UserTeams.Count -eq 0) {
    Write-Host "User is not a member of any Teams."
    exit
}

# Display current teams and roles
Write-Host "`nUser is a member of the following Teams:"
$UserTeams | Format-Table TeamName, Role -AutoSize

# Step 2: Confirm before proceeding
$confirm = Read-Host "Do you want to proceed with removing and re-adding the user? (yes/no)"
if ($confirm -notin @("yes", "y")) {
    Write-Host "Operation canceled."
    exit
}

# Step 3: Remove user from all teams
foreach ($team in $UserTeams) {
    Write-Host "Removing user from $($team.TeamName)..."
    Remove-TeamUser -GroupId $team.GroupId -User $UserPrincipalName
}

Write-Host "User removed from all teams. Waiting before re-adding..."
Start-Sleep -Seconds 5  # Prevent throttling

# Step 4: Re-add user to all teams with the same role
foreach ($team in $UserTeams) {
    Write-Host "Adding user back to $($team.TeamName) as $($team.Role)..."
    Add-TeamUser -GroupId $team.GroupId -User $UserPrincipalName -Role $team.Role
}

Write-Host "`nUser has been successfully removed and re-added to all teams with their original roles."