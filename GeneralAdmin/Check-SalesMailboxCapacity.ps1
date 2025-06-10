# Office 365 Sales Mailbox Capacity Monitor Script
# This script will run immediately when executed
# Save as a .ps1 file and run in VS Code or PowerShell

# Set parameters (modify these as needed)
$WarningThresholdPercent = 85
$ExportToCSV = $true  # Set to $false if you don't want CSV export
$CSVPath = ".\SalesMailboxUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Display script header
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " OFFICE 365 SALES MAILBOX MONITOR" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host "Warning Threshold: $WarningThresholdPercent%" -ForegroundColor Cyan
Write-Host "Export to CSV: $ExportToCSV" -ForegroundColor Cyan
if ($ExportToCSV) {
    Write-Host "CSV Path: $CSVPath" -ForegroundColor Cyan
}
Write-Host ""

# Check if Exchange Online module is installed and import it
try {
    Write-Host "Step 1: Checking for Exchange Online module..." -ForegroundColor Green
    if (!(Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Write-Host "Exchange Online Management module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host "Module installed successfully." -ForegroundColor Green
    } else {
        Write-Host "Exchange Online Management module is already installed." -ForegroundColor Green
    }
    
    Write-Host "Importing Exchange Online module..." -ForegroundColor Green
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "Module imported successfully." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to install or import Exchange Online module: $_" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    exit
}

# Connect to Exchange Online
try {
    Write-Host "`nStep 2: Connecting to Exchange Online..." -ForegroundColor Green
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (!$connectionStatus) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
    } else {
        Write-Host "Already connected to Exchange Online." -ForegroundColor Green
    }
} catch {
    Write-Host "ERROR: Failed to connect to Exchange Online: $_" -ForegroundColor Red
    exit
}

# Get all Sales mailboxes (filtering by UserPrincipalName instead of DisplayName)
try {
    Write-Host "`nStep 3: Retrieving Sales mailboxes..." -ForegroundColor Green
    
    # Adjust this filter to match your domain pattern - replace 'domain.com' with your actual domain
    $salesMailboxes = Get-EXOMailbox -Filter "UserPrincipalName -like 'Sales@*'" -ResultSize Unlimited -ErrorAction Stop
    
    if (!$salesMailboxes -or $salesMailboxes.Count -eq 0) {
        Write-Host "WARNING: No mailboxes found with UserPrincipalName starting with 'Sales@'" -ForegroundColor Yellow
        Write-Host "Please modify the filter in the script if needed." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        exit
    }
    
    Write-Host "Found $($salesMailboxes.Count) Sales mailboxes." -ForegroundColor Green
    
    # Display the first few mailboxes found for verification
    Write-Host "Sample mailboxes found:" -ForegroundColor Cyan
    $salesMailboxes | Select-Object -First 5 | Format-Table DisplayName, UserPrincipalName -AutoSize
    
} catch {
    Write-Host "ERROR: Failed to retrieve mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Process each mailbox
try {
    Write-Host "`nStep 4: Processing mailbox statistics..." -ForegroundColor Green
    $results = @()
    $progressCount = 0
    $totalMailboxes = $salesMailboxes.Count
    
    foreach ($mailbox in $salesMailboxes) {
        $progressCount++
        $progressPercentage = [math]::Round(($progressCount / $totalMailboxes) * 100, 2)
        Write-Progress -Activity "Retrieving mailbox statistics" -Status "Processing $progressCount of $totalMailboxes ($progressPercentage%)" -PercentComplete $progressPercentage
        
        # Get mailbox statistics
        $stats = Get-EXOMailboxStatistics -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        
        if ($stats) {
            # Convert sizes to MB for easier reading
            $totalItemSizeMB = if ($stats.TotalItemSize) {
                [math]::Round([double]($stats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2)
            } else { 0 }
            
            # Get mailbox quota
            $quotaDetails = Get-Mailbox -Identity $mailbox.UserPrincipalName | Select-Object IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota
            
            # Parse quota values
            $warningQuotaMB = if ($quotaDetails.IssueWarningQuota -ne "Unlimited") {
                [math]::Round([double]($quotaDetails.IssueWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2)
            } else { "Unlimited" }
            
            $prohibitSendQuotaMB = if ($quotaDetails.ProhibitSendQuota -ne "Unlimited") {
                [math]::Round([double]($quotaDetails.ProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2)
            } else { "Unlimited" }
            
            $prohibitSendReceiveQuotaMB = if ($quotaDetails.ProhibitSendReceiveQuota -ne "Unlimited") {
                [math]::Round([double]($quotaDetails.ProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 2)
            } else { "Unlimited" }
            
            # Calculate percentage used
            $percentUsed = if ($prohibitSendQuotaMB -ne "Unlimited") {
                [math]::Round(($totalItemSizeMB / $prohibitSendQuotaMB) * 100, 2)
            } else { 0 }
            
            # Determine status based on percentage threshold
            $status = if ($percentUsed -ge $WarningThresholdPercent) {
                "WARNING"
            } else {
                "OK"
            }
            
            # Create custom object with relevant information
            $mailboxInfo = [PSCustomObject]@{
                DisplayName           = $mailbox.DisplayName
                PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                UserPrincipalName     = $mailbox.UserPrincipalName
                ItemCount             = $stats.ItemCount
                TotalSizeMB           = $totalItemSizeMB
                WarningQuotaMB        = $warningQuotaMB
                ProhibitSendQuotaMB   = $prohibitSendQuotaMB
                MaxQuotaMB            = $prohibitSendReceiveQuotaMB
                PercentUsed           = $percentUsed
                Status                = $status
                LastLogonTime         = $stats.LastLogonTime
                LastUserActionTime    = $stats.LastUserActionTime
            }
            
            $results += $mailboxInfo
        } else {
            Write-Host "WARNING: Could not retrieve statistics for mailbox: $($mailbox.DisplayName)" -ForegroundColor Yellow
        }
    }
    
    Write-Progress -Activity "Retrieving mailbox statistics" -Completed
    
    # Order the results by percentage used (descending)
    $orderedResults = $results | Sort-Object -Property PercentUsed -Descending
    
    # Display the results in a formatted table
    Write-Host "`nSales Mailbox Usage Report (Warning Threshold: $WarningThresholdPercent%)" -ForegroundColor Cyan
    $orderedResults | Format-Table -Property DisplayName, PrimarySmtpAddress, TotalSizeMB, ProhibitSendQuotaMB, PercentUsed, Status -AutoSize
    
    # Highlight mailboxes approaching capacity
    $warningMailboxes = $orderedResults | Where-Object { $_.PercentUsed -ge $WarningThresholdPercent }
    if ($warningMailboxes) {
        Write-Host "`nMailboxes approaching capacity limit:" -ForegroundColor Yellow
        foreach ($mbx in $warningMailboxes) {
            Write-Host "- $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress)): $($mbx.PercentUsed)% used ($($mbx.TotalSizeMB)MB / $($mbx.ProhibitSendQuotaMB)MB)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nNo mailboxes are approaching capacity limit." -ForegroundColor Green
    }
    
    # Save results to CSV if requested
    if ($ExportToCSV) {
        Write-Host "`nStep 5: Exporting results to CSV: $CSVPath" -ForegroundColor Green
        $orderedResults | Export-Csv -Path $CSVPath -NoTypeInformation
        Write-Host "Export completed successfully." -ForegroundColor Green
        
        # Show file path
        $fullPath = (Get-Item $CSVPath).FullName
        Write-Host "CSV file saved to: $fullPath" -ForegroundColor Green
    }
    
} catch {
    Write-Host "ERROR: An error occurred while processing mailbox information: $_" -ForegroundColor Red
} finally {
    # Disconnect from Exchange Online
    Write-Host "`nStep 6: Disconnecting from Exchange Online..." -ForegroundColor Green
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected successfully." -ForegroundColor Green
}

Write-Host "`nScript execution completed." -ForegroundColor Cyan