<#
    Script Name: Audit-AdminAccounts.ps1
    Version: v2.1.0
    Author: Don Cook
    Last Updated: 2024-12-30
    Purpose:
    - Queries Active Directory for active admin accounts (accounts starting with 'a_').
    - Calculates the number of days since the last logon.
    - Filters accounts inactive for a specified number of days and outputs results.
    - Optionally exports results to a CSV file, ensuring the output directory exists.

    Service: Active Directory
    Service Type: Account Auditing

    Dependencies:
    - Active Directory Module for Windows PowerShell (RSAT tools must be installed).
    - Permissions to query AD users.

    Notes:
    - This script is designed for a single domain. To target all domains in a forest, use the Global Catalog (see instructions below).
    - Ensure that the `$threshold` variable is set to the desired inactivity period.
    - The script assumes admin accounts follow the naming convention 'a_*'.

    Features:
    - Filters inactive accounts based on a custom threshold.
    - Checks for the output directory and prompts the user to create it if missing.
    - Exports results to a CSV file for further analysis.
#>

# Import Active Directory module
Import-Module ActiveDirectory

# Get the current date
$currentDate = Get-Date

# Define inactivity threshold (in days)
$threshold = 90  # Change this value as needed

# Define output directory and file
$outputDirectory = "C:\Reports"
$outputFile = "$outputDirectory\AdminAccountsAudit.csv"

# Check if the output directory exists
if (-not (Test-Path -Path $outputDirectory)) {
    Write-Host "The directory $outputDirectory does not exist." -ForegroundColor Yellow
    $createDir = Read-Host "Do you want to create it? (Y/N)"
    if ($createDir -eq 'Y' -or $createDir -eq 'y') {
        try {
            New-Item -Path $outputDirectory -ItemType Directory -Force
            Write-Host "Directory $outputDirectory created successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to create directory: $_"
            exit
        }
    } else {
        Write-Host "Cannot proceed without a valid output directory. Exiting script." -ForegroundColor Red
        exit
    }
}

# Uncomment and set the Global Catalog server if querying across all domains
# $gcServer = "GlobalCatalogServer.domain.com"

# Query Active Directory for admin accounts starting with 'a_' and that are enabled
try {
    # For single domain:
    $adminAccounts = Get-ADUser -Filter 'SamAccountName -like "a_*" -and Enabled -eq $true' -Properties Name, SamAccountName, lastLogonTimestamp

    # Uncomment the following line if targeting the Global Catalog:
    # $adminAccounts = Get-ADUser -Filter 'SamAccountName -like "a_*" -and Enabled -eq $true' -Properties Name, SamAccountName, lastLogonTimestamp -Server $gcServer
} catch {
    Write-Error "Error retrieving users from Active Directory: $_"
    exit
}

# Prepare results
$results = $adminAccounts | ForEach-Object {
    $lastLogon = [datetime]::FromFileTime($_.lastLogonTimestamp)
    $daysInactive = ($currentDate - $lastLogon).Days

    # Filter accounts based on inactivity threshold
    if ($daysInactive -gt $threshold) {
        [PSCustomObject]@{
            Name = $_.Name
            Username = $_.SamAccountName
            LastLogon = $lastLogon
            DaysInactive = $daysInactive
        }
    }
}

# Output results to the console
if ($results) {
    Write-Host "Admin accounts inactive for more than $threshold days:" -ForegroundColor Yellow
    $results | Format-Table -AutoSize
} else {
    Write-Host "No admin accounts found inactive for more than $threshold days." -ForegroundColor Green
}

# Export results to the CSV file
try {
    if ($results) {
        $results | Export-Csv -Path $outputFile -NoTypeInformation
        Write-Host "Results exported to $outputFile" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to export results to ${outputFile}: $_"
}