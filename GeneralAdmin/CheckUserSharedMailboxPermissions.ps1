# Suppress informational messages
$InformationPreference = 'SilentlyContinue'

# Suppress warning messages
$WarningPreference = 'SilentlyContinue'

# Connect to Exchange Online without the banner message and suppress additional messages
Connect-ExchangeOnline -ShowBanner:$false -WarningAction SilentlyContinue

# Prompt for the user's email address
$UserEmail = Read-Host -Prompt "Enter the user's email address"

# Get all shared mailboxes
$SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Initialize an array to store results
$Results = @()

# Loop through each shared mailbox
foreach ($Mailbox in $SharedMailboxes) {
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
}

# Check if any results were found
if ($Results.Count -gt 0) {
    # Display the results in a table
    $Results | Format-Table -AutoSize
} else {
    Write-Host "No shared mailbox permissions found for user $UserEmail"
}

# Disconnect from Exchange Online without confirmation and suppress messages
Disconnect-ExchangeOnline -Confirm:$false -WarningAction SilentlyContinue

# Reset the preference variables if needed
$InformationPreference = 'Continue'
$WarningPreference = 'Continue'
