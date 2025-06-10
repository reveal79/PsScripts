<#
.SYNOPSIS
Retrieve mailbox statistics, folder details, item counts, folder sizes, and inbox rules for a specified user.

.DESCRIPTION
This script connects to Exchange Online to gather detailed mailbox statistics for a specified user. 
It calculates folder sizes, total item counts, and displays mailbox rules along with their status.

.SERVICE
Exchange Online

.SERVICE TYPE
Email Management

.VERSION
1.0.0

.AUTHOR
Don Cook

.LAST UPDATED
2024-12-30

.DEPENDENCIES
- ExchangeOnlineManagement module installed.
- Permissions to query mailbox statistics and rules.

.PARAMETERS
None.

.EXAMPLE
To retrieve mailbox statistics for a specific user:
    `.\MailboxStats.ps1`

.NOTES
- The script requires the user running it to have appropriate permissions for accessing mailbox statistics.
- Data is displayed directly in the console; extend with CSV export if required.

#>

# Import Exchange Online Management module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "The ExchangeOnlineManagement module is not installed. Attempting to install..." -ForegroundColor Yellow
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "Module installed successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install ExchangeOnlineManagement module. Ensure internet connectivity and retry."
        exit
    }
}
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Exchange Online. Ensure your credentials and permissions are valid."
    exit
}

# Prompt for user email address
$UserEmail = Read-Host "Enter the user's email address"

try {
    # Retrieve mailbox statistics
    Write-Host "Retrieving mailbox statistics for $UserEmail..." -ForegroundColor Cyan
    $mailboxStats = Get-MailboxFolderStatistics -Identity $UserEmail -ErrorAction Stop
    $mailboxUsage = Get-MailboxStatistics -Identity $UserEmail -ErrorAction Stop
    $mailboxRules = Get-InboxRule -Mailbox $UserEmail -ErrorAction Stop
} catch {
    Write-Error "Failed to retrieve mailbox data: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

if ($mailboxStats) {
    # Display folder statistics
    $totalFolders = ($mailboxStats | Measure-Object).Count
    Write-Host "Total Number of Unique Folders: $totalFolders" -ForegroundColor Green

    # Retrieve folder-specific statistics
    $inboxStats = $mailboxStats | Where-Object { $_.FolderType -eq "Inbox" }
    $sentItemsStats = $mailboxStats | Where-Object { $_.FolderType -eq "SentItems" }
    $deletedItemsStats = $mailboxStats | Where-Object { $_.FolderType -eq "DeletedItems" }

    # Display Inbox statistics
    if ($inboxStats) {
        Write-Host "Inbox Item Count: $($inboxStats.ItemsInFolder)"
        Write-Host "Inbox Folder Size: $($inboxStats.FolderSize)"
    } else {
        Write-Host "Inbox folder not found."
    }

    # Display Sent Items statistics
    if ($sentItemsStats) {
        Write-Host "Sent Items Count: $($sentItemsStats.ItemsInFolder)"
        Write-Host "Sent Items Folder Size: $($sentItemsStats.FolderSize)"
    } else {
        Write-Host "Sent Items folder not found."
    }

    # Display Deleted Items statistics
    if ($deletedItemsStats) {
        Write-Host "Deleted Items Count: $($deletedItemsStats.ItemsInFolder)"
        Write-Host "Deleted Items Folder Size: $($deletedItemsStats.FolderSize)"
    } else {
        Write-Host "Deleted Items folder not found."
    }

    # Calculate total folder size
    $totalFolderSizeBytes = ($mailboxStats | ForEach-Object {
        if ($_ -and $_.FolderSize -match '\((\d+) bytes\)') {
            [int64]$matches[1]
        } elseif ($_ -and $_.FolderSize -match '^([0-9\.]+)\s*(KB|MB|GB)') {
            $sizeValue = [float]$matches[1]
            switch ($matches[2]) {
                "KB" { $sizeValue * 1KB }
                "MB" { $sizeValue * 1MB }
                "GB" { $sizeValue * 1GB }
                default { 0 }
            }
        } else {
            Write-Host "Unable to parse folder size: $($_.FolderSize)" -ForegroundColor Yellow
            0
        }
    } | Measure-Object -Sum).Sum

    Write-Host "Total Folder Size: $([math]::Round($totalFolderSizeBytes / 1MB, 2)) MB" -ForegroundColor Green
}

# Display mailbox usage statistics
if ($mailboxUsage) {
    Write-Host "Mailbox Usage Statistics:" -ForegroundColor Cyan
    Write-Host "Total Item Count: $($mailboxUsage.ItemCount)"
    Write-Host "Total Mailbox Size: $($mailboxUsage.TotalItemSize)"
}

# Display mailbox rules
if ($mailboxRules) {
    $totalRules = ($mailboxRules | Measure-Object).Count
    $enabledRules = ($mailboxRules | Where-Object { $_.Enabled -eq $true } | Measure-Object).Count
    Write-Host "Total Rules: $totalRules"
    Write-Host "Enabled Rules: $enabledRules"

    $lastRuleCreated = ($mailboxRules | Sort-Object -Property DateCreated -Descending | Select-Object -First 1)
    if ($lastRuleCreated -and $lastRuleCreated.DateCreated) {
        Write-Host "Last Rule Created: $($lastRuleCreated.Name) on $($lastRuleCreated.DateCreated.ToString('yyyy-MM-dd HH:mm:ss'))"
    } else {
        Write-Host "No rules found with a creation date."
    }
} else {
    Write-Host "No inbox rules found for $UserEmail."
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online. Script completed." -ForegroundColor Green