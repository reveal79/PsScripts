# Office 365 Shared Mailbox Usage Report
# This script connects to Exchange Online and retrieves shared mailbox usage information

# Connect to Exchange Online
# You'll need to install the ExchangeOnlineManagement module first if you haven't already
# Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber

# Import the module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online (you'll be prompted for credentials)
Connect-ExchangeOnline

# Function to convert bytes to more readable format
function Convert-Size {
    param([long]$Size)
    $units = @('B','KB','MB','GB','TB')
    $unitIndex = 0
    $convertedSize = $Size

    while ($convertedSize -ge 1024 -and $unitIndex -lt ($units.Length - 1)) {
        $convertedSize = $convertedSize / 1024
        $unitIndex++
    }

    return "{0:N2} {1}" -f $convertedSize, $units[$unitIndex]
}

# Function to calculate percentage of quota used
function Get-PercentageUsed {
    param([long]$Used, [long]$Total)
    if ($Total -eq 0) { return 0 }
    return [math]::Round(($Used / $Total) * 100, 2)
}

# Get all shared mailboxes
Write-Host "Retrieving shared mailboxes. This may take a moment..." -ForegroundColor Yellow
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Create an array to store the results
$results = @()

# Process each shared mailbox
$count = 0
$total = $sharedMailboxes.Count

foreach ($mailbox in $sharedMailboxes) {
    $count++
    Write-Progress -Activity "Processing shared mailboxes" -Status "Progress: $count of $total" -PercentComplete (($count / $total) * 100)
    
    # Get mailbox statistics using ExchangeGUID to ensure uniqueness
    try {
        $stats = Get-MailboxStatistics -Identity $mailbox.ExchangeGUID.ToString()
    }
    catch {
        Write-Host "Error getting statistics for mailbox: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        continue
    }
    
    # Get quota information
    $prohibitSendQuota = 0
    if ($mailbox.ProhibitSendQuota -ne "Unlimited") {
        if ($mailbox.ProhibitSendQuota.ToString() -match "\(([0-9,]+) bytes\)") {
            $prohibitSendQuota = [long]($matches[1] -replace ',','')
        }
    }
    
    $prohibitSendReceiveQuota = 0
    if ($mailbox.ProhibitSendReceiveQuota -ne "Unlimited") {
        if ($mailbox.ProhibitSendReceiveQuota.ToString() -match "\(([0-9,]+) bytes\)") {
            $prohibitSendReceiveQuota = [long]($matches[1] -replace ',','')
        }
    }
    
    # Calculate total item size in bytes
    $totalItemSize = 0
    if ($stats.TotalItemSize) {
        # Handle the deserialized ByteQuantifiedSize object
        $sizeString = $stats.TotalItemSize.ToString()
        if ($sizeString -match "\(([0-9,]+) bytes\)") {
            $totalItemSize = [long]($matches[1] -replace ',','')
        }
    }
    
    # Calculate percentage of quota used
    $percentOfSendQuota = Get-PercentageUsed -Used $totalItemSize -Total $prohibitSendQuota
    $percentOfSendReceiveQuota = Get-PercentageUsed -Used $totalItemSize -Total $prohibitSendReceiveQuota
    
    # Determine alert level based on percentage used
    $alertLevel = "Normal"
    if ($percentOfSendQuota -ge 85) { $alertLevel = "Warning" }
    if ($percentOfSendQuota -ge 95) { $alertLevel = "Critical" }
    
    # Create a custom object with the mailbox details
    $resultObj = [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        TotalSize = Convert-Size -Size $totalItemSize
        TotalItems = $stats.ItemCount
        ProhibitSendQuota = if ($mailbox.ProhibitSendQuota -eq "Unlimited") { "Unlimited" } else { Convert-Size -Size $prohibitSendQuota }
        ProhibitSendReceiveQuota = if ($mailbox.ProhibitSendReceiveQuota -eq "Unlimited") { "Unlimited" } else { Convert-Size -Size $prohibitSendReceiveQuota }
        PercentUsed = $percentOfSendQuota
        AlertLevel = $alertLevel
        LastLogonTime = $stats.LastLogonTime
        LastUserActionTime = $stats.LastUserActionTime
    }
    
    # Add the result to the array
    $results += $resultObj
}

# Sort the results by percentage used (descending)
$sortedResults = $results | Sort-Object -Property PercentUsed -Descending

# Display the report
Write-Host "`nShared Mailbox Usage Report" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan

# Display critical alerts
$criticalAlerts = $sortedResults | Where-Object { $_.AlertLevel -eq "Critical" }
if ($criticalAlerts) {
    Write-Host "`nCRITICAL: The following mailboxes have used 95% or more of their quota:" -ForegroundColor Red
    $criticalAlerts | Format-Table -Property DisplayName, PrimarySmtpAddress, TotalSize, ProhibitSendQuota, PercentUsed -AutoSize
}

# Display warning alerts
$warningAlerts = $sortedResults | Where-Object { $_.AlertLevel -eq "Warning" }
if ($warningAlerts) {
    Write-Host "`nWARNING: The following mailboxes have used 85% or more of their quota:" -ForegroundColor Yellow
    $warningAlerts | Format-Table -Property DisplayName, PrimarySmtpAddress, TotalSize, ProhibitSendQuota, PercentUsed -AutoSize
}

# Export the full report to CSV
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
# Use C:\temp\ as specified by the user
$tempFolder = "C:\temp"

# Ensure the directory exists
if (-not (Test-Path -Path $tempFolder)) {
    try {
        New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created directory: $tempFolder" -ForegroundColor Green
    } catch {
        Write-Host "Unable to create $tempFolder. Will try alternate location." -ForegroundColor Yellow
        # Fallback to user's temp directory
        $tempFolder = [System.IO.Path]::GetTempPath()
    }
}

$csvPath = "$tempFolder\SharedMailboxUsageReport-$timestamp.csv"
try {
    $sortedResults | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Report saved to: $csvPath" -ForegroundColor Green
} catch {
    Write-Host "Error writing to $csvPath" -ForegroundColor Red
    Write-Host "Trying system temp directory..." -ForegroundColor Yellow
    
    # Fallback to user's temp directory
    $systemTempFolder = [System.IO.Path]::GetTempPath()
    $csvPath = "$systemTempFolder\SharedMailboxUsageReport-$timestamp.csv"
    try {
        $sortedResults | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Report saved to: $csvPath" -ForegroundColor Green
    } catch {
        Write-Host "Could not save report to any location. CSV export failed." -ForegroundColor Red
        $csvPath = "N/A"
    }
}

Write-Host "`nFull report has been exported to: $csvPath" -ForegroundColor Green
Write-Host "Total shared mailboxes processed: $($results.Count)`n" -ForegroundColor Green

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false