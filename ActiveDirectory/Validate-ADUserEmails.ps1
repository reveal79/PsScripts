<#
    Script Name: Validate-ADUserEmails.ps1
    Version: v1.1.0
    Author: Don Cook
    Last Updated: 2024-12-30
    Purpose: Checks Active Directory users' email addresses for validity based on a regex pattern.
             Designed for use in single-domain environments.

    Service: Active Directory
    Service Type: User Validation

    Dependencies:
    - Active Directory Module for Windows PowerShell (RSAT tools must be installed)
    - Ensure proper permissions to query AD users.

    Notes:
    - This script only validates email addresses within the current domain.
    - For environments with a forest or global catalog server, modify the script to connect to those resources as needed.
      Example: Use `Get-ADUser` with the `-Server` parameter to specify the target domain or global catalog server.

    Example Usage:
    Run the script on a domain controller or a workstation with RSAT tools installed.
#>

# Import Active Directory module (make sure the RSAT tools are installed)
Import-Module ActiveDirectory

# Define email validation regex pattern (checks basic email format)
$emailPattern = '^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$'

# Get all users with an email address
try {
    $users = Get-ADUser -Filter {mail -like "*@*"} -Property mail
} catch {
    Write-Error "Error retrieving users from Active Directory: $_"
    exit
}

# Check each user's email address for validity
foreach ($user in $users) {
    $email = $user.mail

    # Check if the email is not null and matches the regex pattern
    if ($email -and $email -notmatch $emailPattern) {
        # If the email is invalid, output the user and invalid email
        Write-Host "Invalid email: $email for user: $($user.SamAccountName)"
    }
}

# Placeholder for extending functionality to a forest or global catalog server
# Uncomment and modify the below lines if needed:
# $server = "GlobalCatalogServer.domain.com"
# $users = Get-ADUser -Filter {mail -like "*@*"} -Property mail -Server $server