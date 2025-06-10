<#
.SYNOPSIS
    Retrieves Active Directory user accounts that are active and have an expiration date set.

.DESCRIPTION
    This script queries Active Directory to find enabled user accounts with a set expiration date.
    It then extracts the username, display name, department, and expiration date, sorting the output by department.

.NOTES
    Author: Don Cook
    Created: 2025-03-03
    Version: 1.0
    Purpose: Used for auditing and tracking accounts that are set to expire.

.REQUIREMENTS
    - Active Directory module must be installed (RSAT: Active Directory PowerShell).
    - Requires domain admin or appropriate permissions to query Active Directory.
    - Run this script in a PowerShell session with administrative privileges.
#>

# Import the Active Directory module
Import-Module ActiveDirectory

# Retrieve user accounts that:
# - Have an expiration date set (accountExpires > 0)
# - Are enabled (Ensures the account is active)
$accounts = Get-ADUser -Filter {accountExpires -gt 0 -and Enabled -eq $true} `
    -Properties DisplayName, SamAccountName, AccountExpirationDate, Department |  # Retrieve necessary properties
    Where-Object { $_.AccountExpirationDate -ne $null } |  # Ensure only accounts with a defined expiration date
    Select-Object SamAccountName, DisplayName, Department, AccountExpirationDate |  # Select relevant fields
    Sort-Object Department  # Sort results by department for better organization

# Display the results in a table format with automatic column sizing
$accounts | Format-Table -AutoSize