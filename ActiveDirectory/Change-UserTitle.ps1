<#
.SYNOPSIS
Change the title of an Active Directory user and log the changes.

.DESCRIPTION
This script enables administrators to update the title of an Active Directory user account.
It retrieves the user's current title, prompts for a new title, updates the title in Active Directory,
and logs the changes for audit purposes. 

The script securely stores admin credentials for future use in an encrypted format.

.SERVICE
Active Directory

.SERVICE TYPE
User Management

.VERSION
1.0.0

.AUTHOR
Don Cook

.LAST UPDATED
2024-10-07

.DEPENDENCIES
- ActiveDirectory module for managing on-premises Active Directory.

.PARAMETER adminCreds
Specifies the credentials of an administrator account for authenticating against Active Directory.

.PARAMETER username
The SAM account name of the user whose title is being updated.

.PARAMETER logPath
The file path where changes to user titles will be logged.

.EXAMPLE
# Change a user's title in Active Directory
$adminCreds = Get-Credential
$username = "john.doe"
.\Change-UserTitle.ps1

This command updates the user's title in Active Directory and logs the changes to a file.

.NOTES
- The script securely stores admin credentials in an encrypted XML file. If credentials are not found, 
  it prompts for them and saves them for future use.
- Changes to user titles are logged with timestamps for auditing purposes.
#>

# Import the Active Directory module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "The ActiveDirectory module is not installed. Please install RSAT tools for Active Directory." -ForegroundColor Red
    exit
}
Import-Module ActiveDirectory

# Paths
$credPath = "$env:USERPROFILE\adminCreds.xml"
$logPath = "$env:USERPROFILE\ChangeUserTitleLog.txt"

# Function to get or store admin credentials securely
function Get-AdminCredentials {
    if (Test-Path $credPath) {
        Write-Host "Using stored admin credentials..." -ForegroundColor Green
        return Import-Clixml -Path $credPath
    } else {
        Write-Host "Admin credentials not found. Please provide credentials." -ForegroundColor Yellow
        $adminCreds = Get-Credential -Message 'Enter your admin credentials'
        $adminCreds | Export-Clixml -Path $credPath
        Write-Host "Admin credentials securely stored." -ForegroundColor Green
        return $adminCreds
    }
}

# Function to change the title of an Active Directory user
function Change-UserTitle {
    param (
        [string]$username
    )

    try {
        # Search for the user in Active Directory
        $user = Get-ADUser -Filter {SamAccountName -eq $username} -Credential $adminCreds -Properties Title

        if ($user) {
            Write-Host "User found: $($user.Name)" -ForegroundColor Green
            
            # Display the current title
            $oldTitle = $user.Title
            Write-Host "Current Title: $oldTitle"

            # Prompt for the new title
            $newTitle = Read-Host -Prompt 'Enter the new title for the user'

            # Update the title in Active Directory
            Set-ADUser -Identity $user -Title $newTitle -Credential $adminCreds
            Write-Host "Title for $($user.Name) updated to '$newTitle'." -ForegroundColor Green

            # Log the change
            $logEntry = "User: $($user.Name), Old Title: $oldTitle, New Title: $newTitle, Date: $(Get-Date)"
            Add-Content -Path $logPath -Value $logEntry
        } else {
            Write-Host "User not found." -ForegroundColor Red
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
    }
}

# Main Script
Write-Host "Starting Active Directory Title Update Script..." -ForegroundColor Cyan
$adminCreds = Get-AdminCredentials

do {
    # Prompt for the username
    $username = Read-Host -Prompt 'Enter the username of the user to update'
    
    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-Host "Invalid username. Please try again." -ForegroundColor Yellow
        continue
    }

    # Call the function to change the user's title
    Change-UserTitle -username $username

    # Ask if the user wants to process another update
    $continue = Read-Host -Prompt 'Do you want to change the title of another user? (yes/no)'
} while ($continue -eq 'yes')

Write-Host "Script execution completed. Log file: $logPath" -ForegroundColor Cyan