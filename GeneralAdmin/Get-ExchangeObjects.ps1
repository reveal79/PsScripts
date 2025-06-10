<#
.SYNOPSIS
    Inventories Exchange-related objects in Active Directory
.DESCRIPTION
    This script identifies and reports on Exchange-related objects including Distribution Lists,
    mail-enabled users, contacts, and other Exchange attributes still present in on-premises AD.
.NOTES
    Run this script on a domain controller or a machine with the Active Directory module installed
#>

# Import required modules
Import-Module ActiveDirectory

# Set up output file
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputFile = "$env:USERPROFILE\Documents\ExchangeInventory-$timestamp.csv"
$htmlReport = "$env:USERPROFILE\Documents\ExchangeInventory-$timestamp.html"

# Create arrays to store results
$exchangeObjects = @()
$dlCount = 0
$mailEnabledUserCount = 0
$mailEnabledContactCount = 0
$mailEnabledPublicFolderCount = 0
$dynamicDLCount = 0
$resourceMailboxCount = 0
$totalCount = 0

Write-Host "Starting Exchange object inventory..." -ForegroundColor Cyan

# Function to check if an object has Exchange attributes
function Has-ExchangeAttributes {
    param (
        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADObject]$ADObject
    )
    
    $exchangeAttributes = @(
        "homeMDB", "homeMTA", "mailNickname", "proxyAddresses", "mail",
        "legacyExchangeDN", "msExchHomeServerName", "msExchMailboxGuid",
        "msExchRecipientTypeDetails", "msExchRecipientDisplayType"
    )
    
    foreach ($attr in $exchangeAttributes) {
        if ($null -ne $ADObject.$attr) {
            return $true
        }
    }
    
    return $false
}

# Function to get Exchange recipient type as string
function Get-RecipientTypeString {
    param (
        [Parameter(Mandatory=$false)]
        [Int64]$RecipientTypeDetails
    )
    
    if ($null -eq $RecipientTypeDetails) { return "Unknown" }
    
    switch ($RecipientTypeDetails) {
        1 { return "UserMailbox" }
        2 { return "LinkedMailbox" }
        4 { return "SharedMailbox" }
        8 { return "LegacyMailbox" }
        16 { return "RoomMailbox" }
        32 { return "EquipmentMailbox" }
        64 { return "MailContact" }
        128 { return "MailUser" }
        256 { return "MailUniversalDistributionGroup" }
        512 { return "MailNonUniversalGroup" }
        1024 { return "MailUniversalSecurityGroup" }
        2048 { return "DynamicDistributionGroup" }
        4096 { return "PublicFolder" }
        8192 { return "SystemAttendantMailbox" }
        16384 { return "SystemMailbox" }
        32768 { return "MailForestContact" }
        65536 { return "User" }
        131072 { return "Contact" }
        262144 { return "UniversalDistributionGroup" }
        524288 { return "UniversalSecurityGroup" }
        1048576 { return "NonUniversalGroup" }
        2097152 { return "DisabledUser" }
        4194304 { return "MicrosoftExchange" }
        8388608 { return "ArbitrationMailbox" }
        16777216 { return "MailboxPlan" }
        33554432 { return "LinkedUser" }
        default { return "Other ($RecipientTypeDetails)" }
    }
}

try {
    # Get all distribution groups
    Write-Host "Finding distribution lists..." -ForegroundColor Green
    $distributionGroups = Get-ADGroup -Filter {(mail -like "*") -or (proxyAddresses -like "*") -or (msExchRecipientTypeDetails -like "*")} -Properties mail, displayName, proxyAddresses, msExchRecipientTypeDetails, whenChanged, whenCreated, memberOf, member, managedBy
    
    foreach ($group in $distributionGroups) {
        $recipientType = Get-RecipientTypeString -RecipientTypeDetails $group.msExchRecipientTypeDetails
        $memberCount = 0
        
        if ($group.member) {
            $memberCount = $group.member.Count
        }
        
        $exchangeObjects += [PSCustomObject]@{
            Name = $group.Name
            DisplayName = $group.DisplayName
            Type = "Distribution Group"
            SubType = $recipientType
            PrimarySmtpAddress = $group.mail
            MemberCount = $memberCount
            Aliases = ($group.proxyAddresses -join ", ")
            DistinguishedName = $group.DistinguishedName
            WhenCreated = $group.whenCreated
            WhenChanged = $group.whenChanged
            ManagedBy = $group.managedBy
        }
        
        $dlCount++
        $totalCount++
    }
    
    # Get dynamic distribution groups
    Write-Host "Finding dynamic distribution lists..." -ForegroundColor Green
    $dynamicDLs = Get-ADObject -Filter {objectClass -eq "msExchDynamicDistributionList"} -Properties mail, displayName, proxyAddresses, msExchRecipientTypeDetails, whenChanged, whenCreated
    
    foreach ($ddl in $dynamicDLs) {
        $exchangeObjects += [PSCustomObject]@{
            Name = $ddl.Name
            DisplayName = $ddl.DisplayName
            Type = "Dynamic Distribution Group"
            SubType = "DynamicDistributionGroup"
            PrimarySmtpAddress = $ddl.mail
            MemberCount = "Dynamic"
            Aliases = ($ddl.proxyAddresses -join ", ")
            DistinguishedName = $ddl.DistinguishedName
            WhenCreated = $ddl.whenCreated
            WhenChanged = $ddl.whenChanged
            ManagedBy = $ddl.managedBy
        }
        
        $dynamicDLCount++
        $totalCount++
    }
    
    # Get mail-enabled users
    Write-Host "Finding mail-enabled users..." -ForegroundColor Green
    $mailUsers = Get-ADUser -Filter {(mail -like "*") -or (proxyAddresses -like "*")} -Properties mail, displayName, proxyAddresses, msExchRecipientTypeDetails, homeMDB, legacyExchangeDN, whenChanged, whenCreated
    
    foreach ($user in $mailUsers) {
        if (Has-ExchangeAttributes -ADObject $user) {
            $recipientType = Get-RecipientTypeString -RecipientTypeDetails $user.msExchRecipientTypeDetails
            
            $exchangeObjects += [PSCustomObject]@{
                Name = $user.Name
                DisplayName = $user.DisplayName
                Type = "Mail-Enabled User"
                SubType = $recipientType
                PrimarySmtpAddress = $user.mail
                MemberCount = "N/A"
                Aliases = ($user.proxyAddresses -join ", ")
                DistinguishedName = $user.DistinguishedName
                WhenCreated = $user.whenCreated
                WhenChanged = $user.whenChanged
                ManagedBy = "N/A"
            }
            
            $mailEnabledUserCount++
            $totalCount++
            
            # Check for resource mailboxes
            if ($recipientType -eq "RoomMailbox" -or $recipientType -eq "EquipmentMailbox") {
                $resourceMailboxCount++
            }
        }
    }
    
    # Get mail-enabled contacts
    Write-Host "Finding mail-enabled contacts..." -ForegroundColor Green
    $contacts = Get-ADObject -Filter {objectClass -eq "contact" -and (mail -like "*" -or proxyAddresses -like "*")} -Properties mail, displayName, proxyAddresses, msExchRecipientTypeDetails, whenChanged, whenCreated
    
    foreach ($contact in $contacts) {
        if (Has-ExchangeAttributes -ADObject $contact) {
            $recipientType = Get-RecipientTypeString -RecipientTypeDetails $contact.msExchRecipientTypeDetails
            
            $exchangeObjects += [PSCustomObject]@{
                Name = $contact.Name
                DisplayName = $contact.DisplayName
                Type = "Mail-Enabled Contact"
                SubType = $recipientType
                PrimarySmtpAddress = $contact.mail
                MemberCount = "N/A"
                Aliases = ($contact.proxyAddresses -join ", ")
                DistinguishedName = $contact.DistinguishedName
                WhenCreated = $contact.whenCreated
                WhenChanged = $contact.whenChanged
                ManagedBy = "N/A"
            }
            
            $mailEnabledContactCount++
            $totalCount++
        }
    }
    
    # Get mail-enabled public folders
    Write-Host "Finding mail-enabled public folders..." -ForegroundColor Green
    $publicFolders = Get-ADObject -Filter {objectClass -eq "publicFolder" -and (mail -like "*" -or proxyAddresses -like "*")} -Properties mail, displayName, proxyAddresses, msExchRecipientTypeDetails, whenChanged, whenCreated
    
    foreach ($pf in $publicFolders) {
        $exchangeObjects += [PSCustomObject]@{
            Name = $pf.Name
            DisplayName = $pf.DisplayName
            Type = "Mail-Enabled Public Folder"
            SubType = "PublicFolder"
            PrimarySmtpAddress = $pf.mail
            MemberCount = "N/A"
            Aliases = ($pf.proxyAddresses -join ", ")
            DistinguishedName = $pf.DistinguishedName
            WhenCreated = $pf.whenCreated
            WhenChanged = $pf.whenChanged
            ManagedBy = "N/A"
        }
        
        $mailEnabledPublicFolderCount++
        $totalCount++
    }
    
    # Export to CSV
    $exchangeObjects | Export-Csv -Path $outputFile -NoTypeInformation
    
    # Create HTML report
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Objects Inventory Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0099cc; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { background-color: #e6f2ff; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>Exchange Objects Inventory Report</h1>
    <div class="summary">
        <h2>Summary</h2>
        <p>Report generated on $(Get-Date)</p>
        <p>Total Exchange objects found: $totalCount</p>
        <ul>
            <li>Distribution Lists: $dlCount</li>
            <li>Dynamic Distribution Lists: $dynamicDLCount</li>
            <li>Mail-Enabled Users: $mailEnabledUserCount</li>
            <li>Resource Mailboxes (Room/Equipment): $resourceMailboxCount</li>
            <li>Mail-Enabled Contacts: $mailEnabledContactCount</li>
            <li>Mail-Enabled Public Folders: $mailEnabledPublicFolderCount</li>
        </ul>
    </div>
"@

    $htmlFooter = @"
</body>
</html>
"@

    # Create HTML table for Distribution Lists
    $htmlDLs = @"
    <h2>Distribution Lists</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Display Name</th>
            <th>Type</th>
            <th>Primary SMTP Address</th>
            <th>Member Count</th>
            <th>When Created</th>
            <th>When Changed</th>
        </tr>
"@

    $filteredDLs = $exchangeObjects | Where-Object { $_.Type -eq "Distribution Group" }
    foreach ($dl in $filteredDLs) {
        $htmlDLs += @"
        <tr>
            <td>$($dl.Name)</td>
            <td>$($dl.DisplayName)</td>
            <td>$($dl.SubType)</td>
            <td>$($dl.PrimarySmtpAddress)</td>
            <td>$($dl.MemberCount)</td>
            <td>$($dl.WhenCreated)</td>
            <td>$($dl.WhenChanged)</td>
        </tr>
"@
    }
    $htmlDLs += "</table>"

    # Similarly create tables for other object types
    $htmlUsers = @"
    <h2>Mail-Enabled Users</h2>
    <table>
        <tr>
            <th>Name</th>
            <th>Display Name</th>
            <th>Type</th>
            <th>Primary SMTP Address</th>
            <th>When Created</th>
            <th>When Changed</th>
        </tr>
"@

    $filteredUsers = $exchangeObjects | Where-Object { $_.Type -eq "Mail-Enabled User" }
    foreach ($user in $filteredUsers) {
        $htmlUsers += @"
        <tr>
            <td>$($user.Name)</td>
            <td>$($user.DisplayName)</td>
            <td>$($user.SubType)</td>
            <td>$($user.PrimarySmtpAddress)</td>
            <td>$($user.WhenCreated)</td>
            <td>$($user.WhenChanged)</td>
        </tr>
"@
    }
    $htmlUsers += "</table>"

    # Combine all HTML content
    $htmlContent = $htmlHeader + $htmlDLs + $htmlUsers + $htmlFooter
    $htmlContent | Out-File -FilePath $htmlReport -Encoding UTF8

    Write-Host "Inventory completed successfully!" -ForegroundColor Green
    Write-Host "Results exported to:" -ForegroundColor Yellow
    Write-Host "CSV: $outputFile" -ForegroundColor Yellow
    Write-Host "HTML Report: $htmlReport" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "Total Exchange objects found: $totalCount" -ForegroundColor White
    Write-Host "Distribution Lists: $dlCount" -ForegroundColor White
    Write-Host "Dynamic Distribution Lists: $dynamicDLCount" -ForegroundColor White
    Write-Host "Mail-Enabled Users: $mailEnabledUserCount" -ForegroundColor White
    Write-Host "Resource Mailboxes (Room/Equipment): $resourceMailboxCount" -ForegroundColor White
    Write-Host "Mail-Enabled Contacts: $mailEnabledContactCount" -ForegroundColor White
    Write-Host "Mail-Enabled Public Folders: $mailEnabledPublicFolderCount" -ForegroundColor White
}
catch {
    Write-Host "Error occurred: $_" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}