# This script checks for the Exchange Online Management PowerShell module, installs it if necessary,
# connects to Exchange Online, and retrieves the shared mailboxes where a specified user has permissions.
# It outputs the shared mailbox alias, the user, and their access rights.

# Check if the ExchangeOnlineManagement module is installed
# If the module is not found, install it for the current user
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
    # Note: Installing modules may require internet connectivity and admin privileges, depending on policies.
}

# Import the Exchange Online Management module to access necessary cmdlets
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
# This step will prompt for admin credentials to authenticate against the Exchange Online environment.
Connect-ExchangeOnline

# Prompt the user to enter the username to search for permissions.
# Example input: John.Smith@domain.com
$User = Read-Host -Prompt "Enter the username to search (e.g., John.Smith@ecentria.com)"

# Retrieve all shared mailboxes in Exchange Online
# Process each shared mailbox to determine if the specified user has permissions
Get-Mailbox -RecipientTypeDetails SharedMailbox | ForEach-Object {
    # Store the current mailbox object for reference
    $Mailbox = $_
    
    # Retrieve permissions for the specified user on the current mailbox
    # Suppress errors to handle cases where the user does not have permissions on a mailbox
    $Permissions = Get-MailboxPermission -Identity $Mailbox.Alias -User $User -ErrorAction SilentlyContinue
    
    # If permissions exist, create a custom object with mailbox details and user permissions
    if ($Permissions) {
        [PSCustomObject]@{
            SharedMailbox = $Mailbox.Alias   # Alias of the shared mailbox
            User = $User                     # Username being checked
            AccessRights = $Permissions.AccessRights  # Access rights granted to the user
        }
    }
} | Format-Table -AutoSize  # Format the results into a neatly aligned table for readability

# Notes for future users:
# 1. Purpose:
#    - This script helps administrators identify shared mailboxes where a specific user has permissions.
# 2. Prerequisites:
#    - The user running the script must have sufficient permissions in Exchange Online to query mailboxes.
#    - The ExchangeOnlineManagement module must be installed or will be installed by the script.
# 3. Outputs:
#    - A formatted table displaying:
#      - SharedMailbox: The alias of the shared mailbox
#      - User: The username specified during input
#      - AccessRights: The permissions granted to the user (e.g., FullAccess, SendAs).
# 4. Usage Scenarios:
#    - Useful for audits to verify user access to shared mailboxes.
#    - Can help troubleshoot access issues or cleanup unnecessary permissions.
# 5. Authentication:
#    - Ensure you are authenticated to Exchange Online. The script uses `Connect-ExchangeOnline`, which prompts for credentials.
# 6. Error Handling:
#    - Errors are suppressed (`SilentlyContinue`) when a user does not have permissions on a mailbox.
#      This ensures the script continues processing other mailboxes without interruption.

# End of Script
