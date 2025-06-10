<#
    Script Name: Get-MailboxRules.ps1
    Version: v1.0.0
    Author: Don Cook
    Last Updated: 2024-12-30
    Purpose:
    - Retrieves inbox rules for a specified mailbox in Exchange Online.
    - Displays rules in a table format, including the rule name, description, and enabled status.

    Service: Office 365
    Service Type: Email Management

    Dependencies:
    - ExchangeOnlineManagement PowerShell module.
    - Modern Authentication (requires credentials prompt).
    - Permissions to access the target mailbox in Exchange Online.

    Notes:
    - This script connects to Exchange Online and retrieves inbox rules for a specified user.
    - Ensure you have the required permissions to view mailbox rules.
    - Disconnects the session after execution to clean up resources.

    Example Usage:
    Run the script, enter the email address of the target mailbox when prompted, and review the displayed rules.
#>

# Install and Import ExchangeOnlineManagement Module
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online (Modern Authentication will prompt for credentials)
try {
    Connect-ExchangeOnline -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit
}

# Prompt for the mailbox email after establishing a connection
$Mailbox = Read-Host -Prompt "Enter the user's email address"

try {
    # Get the inbox rules for the specified mailbox using the EXO V2 cmdlet
    $rules = Get-InboxRule -Mailbox $Mailbox -ErrorAction Stop

    if ($rules) {
        # Display the rules
        Write-Host "Inbox rules for ${Mailbox}:" -ForegroundColor Green
        $rules | Select-Object Name, Description, Enabled | Format-Table -AutoSize
    } else {
        Write-Host "No inbox rules found for $Mailbox." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "An error occurred while retrieving rules: $_"
}
finally {
    # Disconnect the session
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
}