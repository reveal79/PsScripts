<#
.SYNOPSIS
This script checks a user's access permissions to all shared mailboxes in an Exchange Online environment. 

.DESCRIPTION
The script connects to Exchange Online, prompts the user to input an email address, retrieves all shared mailboxes, 
and checks if the specified user has permissions on any of these mailboxes. If permissions are found, 
the script displays the results in a table format. It disconnects from Exchange Online upon completion.

.FEATURES
- Ensures the required module (Exchange Online Management) is installed.
- Suppresses informational and warning messages for cleaner output.
- Retrieves all shared mailboxes in the Exchange Online tenant.
- Checks and lists the user's permissions on each shared mailbox.
- Disconnects from Exchange Online after completing the task.

.REQUIREMENTS
- PowerShell
- Exchange Online Management module installed and configured
- Sufficient permissions to query mailbox data and permissions in Exchange Online

.VERSION
1.1.0
#>

# Function to check and install the required module
function Ensure-Module {
    param (
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "The required module '$ModuleName' is not installed. Attempting to install it..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Module '$ModuleName' installed successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to install the module '$ModuleName'. Ensure you have internet access and permissions to install modules."
            Write-Host "To resolve this issue, manually install the module using: Install-Module -Name $ModuleName -Scope CurrentUser -Force"
            exit
        }
    }
}

# Ensure the Exchange Online Management module is installed
Ensure-Module -ModuleName "ExchangeOnlineManagement"

# Import the module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Suppress informational and warning messages for cleaner output
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ShowBanner:$false -WarningAction SilentlyContinue -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Exchange Online. Ensure your credentials and environment are configured correctly."
    exit
}

# Prompt for the user's email address
$UserEmail = Read-Host -Prompt "Enter the user's email address"

# Get all shared mailboxes
try {
    $SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop
} catch {
    Write-Error "Failed to retrieve shared mailboxes. Ensure you have the necessary permissions."
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Initialize an array to store results
$Results = @()

# Loop through each shared mailbox
foreach ($Mailbox in $SharedMailboxes) {
    try {
        # Get permissions for the user on the current shared mailbox
        $Permissions = Get-MailboxPermission -Identity $Mailbox.Identity -User $UserEmail -ErrorAction SilentlyContinue

        if ($Permissions) {
            # Add the mailbox and permissions to the results array
            $Results += [PSCustomObject]@{
                SharedMailbox       = $Mailbox.Alias
                DisplayName         = $Mailbox.DisplayName
                SharedMailboxEmail  = $Mailbox.PrimarySmtpAddress
                User                = $UserEmail
                AccessRights        = ($Permissions.AccessRights -join ', ')
            }
        }
    } catch {
        Write-Warning "An error occurred while checking permissions for mailbox: $($Mailbox.Identity)"
    }
}

# Check if any results were found
if ($Results.Count -gt 0) {
    # Display the results in a table
    $Results | Format-Table -AutoSize
} else {
    Write-Host "No shared mailbox permissions found for user $UserEmail" -ForegroundColor Yellow
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false -WarningAction SilentlyContinue

# Reset the preference variables if needed
$InformationPreference = 'Continue'
$WarningPreference = 'Continue'