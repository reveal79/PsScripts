<#
.SYNOPSIS
Export enabled Active Directory users with password-related details to a CSV file.

.DESCRIPTION
This script retrieves all enabled Active Directory users and exports key information such as 
password last set, expiration date, and other attributes into a structured CSV file. 
It also ensures proper error handling, logging, and dynamic directory management.

.SERVICE
Active Directory

.SERVICE TYPE
User Management

.VERSION
1.0.0

.AUTHOR
Don Cook

.LAST UPDATED
2024-12-30

.PARAMETERS
None.

.DEPENDENCIES
- Active Directory RSAT tools installed.
- Permissions to query Active Directory.

.EXAMPLE
Run the script directly to generate a CSV report at C:\temp\ADUsers_YYYYMMDD.csv.

.NOTES
- The script automatically creates the output directory if it does not exist.
- A log file is maintained in the output directory to track script execution.

#>

# Configuration
$outputPath = "C:\temp\ADUsers_$((Get-Date -Format 'yyyyMMdd')).csv"
$logFile = "C:\temp\ADUsersLog.txt"

# Ensure the Active Directory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found. Please install the RSAT tools for Active Directory." -ForegroundColor Red
    exit
}
Import-Module ActiveDirectory

# Ensure output directory exists
$outputDir = Split-Path -Path $outputPath
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force
}

# Query Active Directory
try {
    Write-Host "Querying Active Directory for enabled users..." -ForegroundColor Yellow
    $users = Get-ADUser -Filter 'Enabled -eq $true' -Properties msDS-UserPasswordExpiryTimeComputed, PasswordLastSet, CannotChangePassword
    Write-Host "Successfully retrieved $($users.Count) users." -ForegroundColor Green
} catch {
    Write-Error "Failed to query Active Directory: $_"
    exit
}

# Process users
$total = $users.Count
$current = 0
$results = $users | ForEach-Object {
    $current++
    Write-Progress -Activity "Processing AD Users" -Status "$current of $total" -PercentComplete (($current / $total) * 100)
    [PSCustomObject]@{
        Username             = $_.SamAccountName
        DisplayName          = $_.DisplayName
        PasswordLastSet      = $_.PasswordLastSet
        PasswordExpires      = [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")
        CannotChangePassword = $_.CannotChangePassword
    }
}

# Export results to CSV
try {
    $results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Export complete. File saved at $outputPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to export results to CSV: $_"
    exit
}

# Log the export
Add-Content -Path $logFile -Value "$(Get-Date): Successfully exported $($results.Count) users to $outputPath"

# Completion message
Write-Host "Script execution completed successfully. Logs saved at $logFile" -ForegroundColor Green