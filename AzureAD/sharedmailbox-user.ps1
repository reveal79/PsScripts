<#
.SYNOPSIS
    Check User Permissions on Shared Mailboxes in Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online, retrieves all shared mailboxes, and checks whether a specified user has any 
    permissions assigned to those shared mailboxes. It outputs the results in a table format showing the shared mailbox, 
    the user, and their access rights.

.PARAMETER User
    The email address of the user for whom you want to check permissions on shared mailboxes.

.EXAMPLE
    # Connect to Exchange Online and check permissions for "jane.doe@domain.com"
    $User = "jane.doe@domain.com"
    .\Check-SharedMailboxPermissions.ps1
    
    This will retrieve all shared mailboxes and output any permissions the user has on them.

.NOTES
    Author: Don Cook
    Version: 1.0
    Date: 2024-10-07

    This script is useful for system administrators to audit shared mailbox permissions, particularly if you want 
    to quickly find out whether a user has access to any shared mailboxes and the level of access granted.

    The script is helpful for reviewing permissions during access reviews, offboarding processes, or ensuring compliance 
    with mailbox access policies.

#>

# Connect to Exchange Online
Connect-ExchangeOnline

# Specify the user whose permissions you want to check
$User = ""

# Get all shared mailboxes and check if the specified user has permissions
Get-Mailbox -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $Mailbox = $_
    # Get the user's permissions on the shared mailbox
    $Permissions = Get-MailboxPermission -Identity $Mailbox.Alias -User $User -ErrorAction SilentlyContinue
    if ($Permissions -ne $null) {
        # Output the mailbox, user, and their access rights in a custom object
        [PSCustomObject]@{
            SharedMailbox = $Mailbox.Alias
            User = $User
            AccessRights = $Permissions.AccessRights
        }
    }
} | Format-Table -AutoSize  # Format the output as a table
