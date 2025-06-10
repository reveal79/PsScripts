<#
    Script Name: Get-DirectReports.ps1
    Version: v1.0.0
    Author: <Your Name>
    Last Updated: 2024-12-30
    Purpose:
    - Retrieves the direct reports for a given manager in Active Directory based on their sAMAccountName.
    - Outputs a list of users who report to the specified manager.

    Service: Active Directory
    Service Type: User Management

    Dependencies:
    - Active Directory Module for Windows PowerShell (RSAT tools must be installed).
    - Permissions to query AD users.

    Notes:
    - The script queries the current domain. To adapt it for a forest-wide scope, use the `-Server` parameter with a Global Catalog server.
      Example:
      `$gcServer = "GlobalCatalogServer.domain.com"`
      `Get-ADUser -LDAPFilter "(manager=$($managerObj.DistinguishedName))" -Server $gcServer`
    - Ensure the input manager's `sAMAccountName` is accurate to avoid errors.

    Example Usage:
    Run the script and enter the manager's sAMAccountName when prompted.
#>

# Import the ActiveDirectory module if not already loaded
Import-Module ActiveDirectory

# Function to get direct reports
function Get-DirectReports {
    param (
        [string]$ManagerSamAccountName
    )

    # Find the manager object in AD
    try {
        $managerObj = Get-ADUser $ManagerSamAccountName -Properties "DistinguishedName"
    } catch {
        Write-Error "Error retrieving manager object from Active Directory: $_"
        return
    }

    if ($managerObj -eq $null) {
        Write-Host "User with sAMAccountName $ManagerSamAccountName not found in Active Directory." -ForegroundColor Red
        return
    }

    # Fetch all users who report to this manager
    try {
        $directReports = Get-ADUser -LDAPFilter "(manager=$($managerObj.DistinguishedName))" -Properties DisplayName
    } catch {
        Write-Error "Error retrieving direct reports from Active Directory: $_"
        return
    }

    # Output the results
    if ($directReports.Count -eq 0) {
        Write-Host "No direct reports found for the manager with sAMAccountName $ManagerSamAccountName." -ForegroundColor Yellow
    } else {
        Write-Host "The following users report to the manager with sAMAccountName ${ManagerSamAccountName}:" -ForegroundColor Green
        foreach ($report in $directReports) {
            Write-Host "- $($report.DisplayName)"
        }
    }
}

# Prompt the user for the manager's sAMAccountName
$managerSamAccountName = Read-Host "Please enter the sAMAccountName of the manager for whom you want to find direct reports"

# Get and display the list of direct reports
Get-DirectReports -ManagerSamAccountName $managerSamAccountName