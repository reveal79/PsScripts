# AD-TeleVantage-Phone-Comparison.ps1
#===============================================================================
# Script Name: AD-TeleVantage-Phone-Comparison.ps1
# Created On: April 1, 2025
# Last Modified: April 1, 2025
#
# Description:
#   This script compares Active Directory phone fields with TeleVantage
#   extension and DID information to identify discrepancies and help
#   plan for migration to Teams or other platforms.
#
#   It produces a report showing:
#   - Users with phone numbers in AD and their corresponding TV extensions/DIDs
#   - Active/inactive status of TeleVantage numbers
#   - Forwarding configurations in TeleVantage
#   - Orphaned TeleVantage numbers not associated with AD users
#
# Dependencies:
#   - Active Directory module
#   - System.Data assembly
#   - ImportExcel PowerShell module
#
# Parameters:
#   $server - TeleVantage database server
#   $database - TeleVantage database name
#   $outputFile - Path for the Excel report output
#
# Usage:
#   .\AD-TeleVantage-Phone-Comparison.ps1
#===============================================================================

# Import required modules
Import-Module ActiveDirectory

# Set up the export path
$outputFile = "C:\Temp\AD_TV_Phone_Comparison.xlsx"

# Database connection settings for TeleVantage
$server = "srv-phone-103.usgroup.loc"
$database = "TVDB"

# Step 1: Query Active Directory for users and phone numbers
Write-Host "Retrieving Active Directory user information..."
$adUsers = Get-ADUser -Filter * -Properties DisplayName, telephoneNumber, mobile, ipPhone | 
           Select-Object DisplayName, SamAccountName, telephoneNumber, mobile, ipPhone |
           Where-Object {$_.telephoneNumber -or $_.mobile -or $_.ipPhone} # Only include users with at least one phone field

Write-Host "Retrieved $($adUsers.Count) AD users with phone information"

# Step 2: Query TeleVantage database directly
Write-Host "Retrieving TeleVantage data directly from database..."

# SQL query for TeleVantage data
$query = @"
SELECT 
    es.DIDNumber,
    es.Number AS Extension,
    MAX(cl.StartTime) AS LastAnyCall,
    COUNT(cl.ID) AS TotalCalls,
    SUM(CASE WHEN cl.StartTime >= DATEADD(MONTH, -3, GETDATE()) THEN 1 ELSE 0 END) AS CallsLast3Months,
    SUM(CASE WHEN cl.Direction = 1 THEN 1 ELSE 0 END) AS OutboundCalls,
    SUM(CASE WHEN DATEDIFF(second, cl.StartTime, cl.StopTime) > 30 THEN 1 ELSE 0 END) AS CallsOver30Sec,
    CASE 
        WHEN es.ForwardAddressID IS NOT NULL OR es.DefaultForwardingID IS NOT NULL THEN 'Yes'
        ELSE 'No'
    END AS HasForwarding,
    es.ForwardAddressID,
    es.DefaultForwardingID
FROM ExtensionSettings es
LEFT JOIN CallLog cl 
    ON es.DIDNumber = cl.DIDNumber
    AND cl.StartTime >= DATEADD(YEAR, -3, GETDATE())
WHERE es.DIDNumber IS NOT NULL OR es.Number IS NOT NULL
GROUP BY 
    es.DIDNumber, 
    es.Number, 
    es.ForwardAddressID, 
    es.DefaultForwardingID
ORDER BY CASE WHEN MAX(cl.StartTime) IS NULL THEN 1 ELSE 0 END, MAX(cl.StartTime) DESC;
"@

try {
    # Load required assemblies
    Add-Type -AssemblyName System.Data

    # Build connection string with TrustServerCertificate
    $connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True"
    
    # Create connection and command
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
    
    # Create adapter and dataset
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    
    # Get data
    Write-Host "Executing TeleVantage query..."
    $connection.Open()
    $adapter.Fill($dataset) | Out-Null
    $connection.Close()
    
    # Get results table
    $tvData = $dataset.Tables[0]
    
    Write-Host "✅ Retrieved $($tvData.Rows.Count) records from TeleVantage"
    
    # Clean up the TeleVantage data for easier matching
    $tvUsers = @()
    foreach ($row in $tvData.Rows) {
        $tvUsers += [PSCustomObject]@{
            'DIDNumber' = $row['DIDNumber']
            'Extension' = $row['Extension']
            'LastUsed' = $row['LastAnyCall']
            'TotalCalls' = $row['TotalCalls']
            'CallsLast3Months' = $row['CallsLast3Months']
            'OutboundCalls' = $row['OutboundCalls']
            'CallsOver30Sec' = $row['CallsOver30Sec']
            'HasForwarding' = $row['HasForwarding']
            'ForwardAddressID' = $row['ForwardAddressID']
            'DefaultForwardingID' = $row['DefaultForwardingID']
            'Status' = if ([string]::IsNullOrEmpty($row['LastAnyCall'])) {
                'Inactive'
            } elseif ($row['CallsLast3Months'] -gt 0) {
                'Active'
            } elseif ($row['LastAnyCall'] -ge [DateTime]::Now.AddYears(-1)) {
                'Low Usage'
            } else {
                'Very Low Usage'
            }
        }
    }
}
catch {
    Write-Host "❌ Error retrieving TeleVantage data: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit
}

# Step 3: Create a combined dataset comparing AD and TV
Write-Host "Creating combined dataset for comparison..."
$comparisonData = @()

foreach ($adUser in $adUsers) {
    # Try to find matching TeleVantage user by extension or DID
    $tvUser = $null
    $telephoneMatch = $false
    $mobileMatch = $false
    $ipPhoneMatch = $false
    
    if ($adUser.telephoneNumber) {
        $cleanPhone = $adUser.telephoneNumber -replace '[^0-9]', ''
        $tvUser = $tvUsers | Where-Object {
            $didNumber = $_.DIDNumber
            if ($didNumber) {
                $didNumber -like "*$cleanPhone*"
            } else {
                $false
            }
        }
        if ($tvUser) { $telephoneMatch = $true }
    }
    
    if (-not $tvUser -and $adUser.mobile) {
        $cleanMobile = $adUser.mobile -replace '[^0-9]', ''
        $tvUser = $tvUsers | Where-Object {
            $didNumber = $_.DIDNumber
            if ($didNumber) {
                $didNumber -like "*$cleanMobile*"
            } else {
                $false
            }
        }
        if ($tvUser) { $mobileMatch = $true }
    }
    
    if (-not $tvUser -and $adUser.ipPhone) {
        # Check if ipPhone is an extension
        $tvUser = $tvUsers | Where-Object {$_.Extension -eq $adUser.ipPhone}
        if ($tvUser) { $ipPhoneMatch = $true }
    }
    
    # Create comparison object
    $comparison = [PSCustomObject]@{
        'DisplayName' = $adUser.DisplayName
        'SamAccountName' = $adUser.SamAccountName
        'AD_Phone' = $adUser.telephoneNumber
        'AD_Mobile' = $adUser.mobile
        'AD_ipPhone' = $adUser.ipPhone
        'TV_DIDNumber' = if ($tvUser) { $tvUser.DIDNumber } else { $null }
        'TV_Extension' = if ($tvUser) { $tvUser.Extension } else { $null }
        'TV_LastUsed' = if ($tvUser) { $tvUser.LastUsed } else { $null }
        'TV_TotalCalls' = if ($tvUser) { $tvUser.TotalCalls } else { $null }
        'TV_CallsLast3Months' = if ($tvUser) { $tvUser.CallsLast3Months } else { $null }
        'TV_Status' = if ($tvUser) { $tvUser.Status } else { "Not Found" }
        'TV_HasForwarding' = if ($tvUser) { $tvUser.HasForwarding } else { "No" }
        'PhoneNumberMatches' = if ($telephoneMatch -or $mobileMatch -or $ipPhoneMatch) { "Yes" } else { "No" }
        'MatchSource' = if ($telephoneMatch) { "AD Phone" } elseif ($mobileMatch) { "AD Mobile" } elseif ($ipPhoneMatch) { "AD ipPhone" } else { "None" }
    }
    
    $comparisonData += $comparison
}

# Also check for any TeleVantage numbers not matched to an AD user
Write-Host "Checking for orphaned TeleVantage numbers..."
foreach ($tvUser in $tvUsers) {
    # Skip if this TV user was already matched
    $alreadyMatched = $comparisonData | Where-Object {$_.TV_DIDNumber -eq $tvUser.DIDNumber -and $_.TV_Extension -eq $tvUser.Extension}
    if ($alreadyMatched) { continue }
    
    # Create entry for unmatched TV user
    $comparison = [PSCustomObject]@{
        'DisplayName' = "No AD Match"
        'SamAccountName' = $null
        'AD_Phone' = $null
        'AD_Mobile' = $null
        'AD_ipPhone' = $null
        'TV_DIDNumber' = $tvUser.DIDNumber
        'TV_Extension' = $tvUser.Extension
        'TV_LastUsed' = $tvUser.LastUsed
        'TV_TotalCalls' = $tvUser.TotalCalls
        'TV_CallsLast3Months' = $tvUser.CallsLast3Months
        'TV_Status' = $tvUser.Status
        'TV_HasForwarding' = $tvUser.HasForwarding
        'PhoneNumberMatches' = "No"
        'MatchSource' = "Orphaned TV Number"
    }
    
    $comparisonData += $comparison
}

# Step 4: Export the comparison data
Write-Host "Exporting comparison data to Excel..."
Import-Module ImportExcel
$comparisonData | Export-Excel -Path $outputFile -AutoSize -TableName "PhoneNumberComparison" -WorksheetName "AD-TV Comparison"

Write-Host "✅ Export complete: $outputFile"