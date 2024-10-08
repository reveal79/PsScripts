<#
.SYNOPSIS
    Change the title of an Active Directory user and log the changes.

.DESCRIPTION
    This script allows administrators to change the title of an Active Directory user account.
    It retrieves the user's current title, prompts for a new title, and then updates the title in Active Directory.
    The script also logs the changes to a file, along with a timestamp.

    If admin credentials are not available, the script will prompt the user to enter them and store the credentials securely 
    in an XML file for future use.

.PARAMETER adminCreds
    The credentials of an administrator account used to authenticate against Active Directory.

.PARAMETER username
    The username (SAM account name) of the user whose title is being updated.

.PARAMETER newTitle
    The new job title to assign to the user.

.PARAMETER logPath
    The file path where changes to user titles will be logged.

.EXAMPLE
    # Change a user's title in Active Directory
    $adminCreds = Get-Credential
    $username = "john.doe"
    .\Change-UserTitle.ps1

    This command will prompt for the user's new title and update it in Active Directory, logging the changes to a file.

.NOTES
    Author: Don Cook
    Date: 2024-10-07

    Modules Required:
      - ActiveDirectory: This module allows managing on-premises Active Directory users.
      
    The script stores admin credentials securely in an XML file. If the credentials are not found, it prompts for them and 
    saves them for future use. Changes made to user titles are logged for audit purposes.
#>

# Import the Active Directory module
Import-Module ActiveDirectory

# Path to the stored credentials file
$credPath = "$env:USERPROFILE\adminCreds.xml"
# Path to the log file
$logPath = "$env:USERPROFILE\ChangeUserTitleLog.txt"

# Function to get admin credentials
function Get-AdminCredentials {
    if (Test-Path $credPath) {
        # Import stored credentials
        return Import-Clixml -Path $credPath
    } else {
        # Prompt for admin credentials
        $adminCreds = Get-Credential -Message 'Enter your admin credentials'
        # Export credentials to file
        $adminCreds | Export-Clixml -Path $credPath
        return $adminCreds
    }
}

# Get admin credentials
$adminCreds = Get-AdminCredentials

# Function to change user title
function Change-UserTitle {
    # Prompt for the username
    $username = Read-Host -Prompt 'Enter the username'

    # Search for the user in Active Directory
    $user = Get-ADUser -Filter {SamAccountName -eq $username} -Credential $adminCreds

    # Check if user was found
    if ($user) {
        Write-Host "User found: $($user.Name)"
        
        # Get the old title
        $oldTitle = $user.Title
        Write-Host "Old Title: $oldTitle"

        # Prompt for the new title
        $newTitle = Read-Host -Prompt 'Enter the new title for the user'

        # Set the new title
        Set-ADUser -Identity $user -Title $newTitle -Credential $adminCreds

        Write-Host "Title for $($user.Name) has been changed to '$newTitle'."
        
        # Log the changes
        $logEntry = "User: $($user.Name), Old Title: $oldTitle, New Title: $newTitle, Date: $(Get-Date)"
        Add-Content -Path $logPath -Value $logEntry
    } else {
        Write-Host "User not found."
    }
}

# Main loop to allow multiple updates
do {
    Change-UserTitle
    $continue = Read-Host -Prompt 'Do you want to change the title of another user? (yes/no)'
} while ($continue -eq 'yes')
