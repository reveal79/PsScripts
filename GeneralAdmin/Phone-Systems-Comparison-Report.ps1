# Phone-Systems-Comparison-Report.ps1
#===============================================================================
# Script Name: Phone-Systems-Comparison-Report.ps1
# Created On: April 1, 2025
#
# Description:
#   This script combines data from Active Directory, TeleVantage, and 
#   Microsoft Teams to create a comprehensive phone system comparison report.
#   It identifies discrepancies between systems and helps plan for migration.
#
# Dependencies:
#   - ImportExcel PowerShell module (recommended)
#
# Usage:
#   .\Phone-Systems-Comparison-Report.ps1
#===============================================================================

# Set file paths
$teamsDataPath = "C:\Temp\Teams_Phone_Numbers.csv"
$adTvDataPath = "C:\Temp\AD_TV_Phone_Comparison.xlsx"
$outputPath = "C:\Temp\Phone_Systems_Comparison.xlsx"

Write-Host "Starting phone systems comparison..."

# Step 1: Import the Teams data
Write-Host "Importing Teams phone data..."
if (Test-Path $teamsDataPath) {
    $teamsData = Import-Csv -Path $teamsDataPath
    Write-Host "  Found $($teamsData.Count) Teams users with phone numbers"
}
else {
    Write-Host "Error: Teams data file not found at $teamsDataPath" -ForegroundColor Red
    exit
}

# Step 2: Import the AD-TV comparison data
Write-Host "Importing AD-TV comparison data..."
try {
    # Try to use ImportExcel module if available
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel
        $adTvData = Import-Excel -Path $adTvDataPath
        Write-Host "  Found $($adTvData.Count) records in AD-TV comparison"
    }
    else {
        # Fallback to CSV if previously exported
        $adTvCsvPath = $adTvDataPath -replace "\.xlsx$", ".csv"
        if (Test-Path $adTvCsvPath) {
            $adTvData = Import-Csv -Path $adTvCsvPath
            Write-Host "  Found $($adTvData.Count) records in AD-TV comparison (CSV)"
        }
        else {
            Write-Host "Error: ImportExcel module not found and no CSV version available." -ForegroundColor Red
            Write-Host "Please export AD-TV data to CSV or install ImportExcel module:" -ForegroundColor Yellow
            Write-Host "Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
            exit
        }
    }
}
catch {
    Write-Host "Error importing AD-TV data: $_" -ForegroundColor Red
    exit
}

# Step 3: Create a combined dataset
Write-Host "Creating combined comparison dataset..."

$combinedData = @()

# First pass: Process AD-TV data and look for Teams matches
foreach ($adTvUser in $adTvData) {
    # Try to find matching Teams user
    $teamsUser = $null
    $teamsDIDMatch = $false
    
    # Check for Teams match by SamAccountName/UPN
    if ($adTvUser.SamAccountName) {
        $teamsUser = $teamsData | Where-Object { 
            $_.UserPrincipalName -like "*$($adTvUser.SamAccountName)*" 
        } | Select-Object -First 1
    }
    
    # If no match by name, try matching by phone number
    if (-not $teamsUser -and $adTvUser.TV_DIDNumber) {
        # Clean up the TV DID for comparison
        $tvDID = $adTvUser.TV_DIDNumber -replace '[^0-9]', ''
        
        foreach ($t in $teamsData) {
            $teamsPhone = $t.PhoneNumber -replace '[^0-9]', ''
            
            # Check if Teams phone is contained in TV DID (or vice versa)
            if (($tvDID -and $teamsPhone -and $tvDID.Contains($teamsPhone)) -or 
                ($tvDID -and $teamsPhone -and $teamsPhone.Contains($tvDID))) {
                $teamsUser = $t
                $teamsDIDMatch = $true
                break
            }
        }
    }
    
    # Create comparison record
    $comparison = [PSCustomObject]@{
        'DisplayName' = $adTvUser.DisplayName
        'SamAccountName' = $adTvUser.SamAccountName
        'AD_Phone' = $adTvUser.AD_Phone
        'AD_Mobile' = $adTvUser.AD_Mobile 
        'AD_ipPhone' = $adTvUser.AD_ipPhone
        'TV_DIDNumber' = $adTvUser.TV_DIDNumber
        'TV_Extension' = $adTvUser.TV_Extension
        'TV_Status' = $adTvUser.TV_Status
        'TV_LastUsed' = $adTvUser.TV_LastUsed
        'TV_HasForwarding' = $adTvUser.TV_HasForwarding
        'Teams_PhoneNumber' = if ($teamsUser) { $teamsUser.PhoneNumber } else { $null }
        'Teams_UPN' = if ($teamsUser) { $teamsUser.UserPrincipalName } else { $null }
        'Teams_VoiceEnabled' = if ($teamsUser) { $teamsUser.EnterpriseVoiceEnabled } else { $null }
        'DID_Matches_Teams' = if ($teamsDIDMatch) { "Yes" } else { "No" }
        'MigrationStatus' = if ($teamsUser -and $teamsDIDMatch) {
                                "Already Migrated"
                            } elseif ($teamsUser -and -not $teamsDIDMatch) {
                                "Number Mismatch"
                            } elseif ($adTvUser.TV_Status -eq "Active" -and -not $teamsUser) {
                                "Migration Required"
                            } elseif ($adTvUser.TV_Status -eq "Low Usage" -and -not $teamsUser) {
                                "Evaluate Need"
                            } else {
                                "No Action Needed"
                            }
    }
    
    $combinedData += $comparison
}

# Second pass: Find Teams users not already included
Write-Host "Checking for Teams users not matched to AD/TV records..."

foreach ($teamsUser in $teamsData) {
    # Check if this Teams user was already included
    $existingUser = $combinedData | Where-Object { 
        $_.Teams_UPN -eq $teamsUser.UserPrincipalName 
    }
    
    if (-not $existingUser) {
        # Create a new record for unmatched Teams user
        $comparison = [PSCustomObject]@{
            'DisplayName' = $teamsUser.DisplayName
            'SamAccountName' = $null
            'AD_Phone' = $null
            'AD_Mobile' = $null
            'AD_ipPhone' = $null
            'TV_DIDNumber' = $null
            'TV_Extension' = $null
            'TV_Status' = "Not Found"
            'TV_LastUsed' = $null
            'TV_HasForwarding' = "No"
            'Teams_PhoneNumber' = $teamsUser.PhoneNumber
            'Teams_UPN' = $teamsUser.UserPrincipalName
            'Teams_VoiceEnabled' = $teamsUser.EnterpriseVoiceEnabled
            'DID_Matches_Teams' = "N/A"
            'MigrationStatus' = "Teams Only"
        }
        
        $combinedData += $comparison
    }
}

# Step 4: Export the combined data
Write-Host "Exporting comparison data..."

try {
    # Try to use ImportExcel module if available
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Define Excel formatting
        $excelParams = @{
            Path = $outputPath
            AutoSize = $true
            TableName = "PhoneSystemsComparison"
            WorksheetName = "Phone Systems Comparison"
            FreezeTopRow = $true
            BoldTopRow = $true
            AutoFilter = $true
            ConditionalText = @(
                New-ConditionalText -Text "Already Migrated" -BackgroundColor LightGreen
                New-ConditionalText -Text "Migration Required" -BackgroundColor LightSalmon
                New-ConditionalText -Text "Number Mismatch" -BackgroundColor Yellow
                New-ConditionalText -Text "Evaluate Need" -BackgroundColor LightYellow
                New-ConditionalText -Text "Teams Only" -BackgroundColor LightBlue
                New-ConditionalText -Text "No Action Needed" -BackgroundColor LightGray
            )
        }
        
        $combinedData | Export-Excel @excelParams
        Write-Host "✅ Excel export complete: $outputPath"
    }
    else {
        # Fallback to CSV
        $csvPath = $outputPath -replace "\.xlsx$", ".csv"
        $combinedData | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "✅ CSV export complete: $csvPath"
    }
}
catch {
    Write-Host "Error exporting data: $_" -ForegroundColor Red
    
    # Emergency fallback to simple CSV
    $emergencyCsvPath = "C:\Temp\Phone_Systems_Comparison_Emergency.csv"
    $combinedData | Export-Csv -Path $emergencyCsvPath -NoTypeInformation
    Write-Host "Emergency CSV export complete: $emergencyCsvPath" -ForegroundColor Yellow
}

# Step 5: Generate summary statistics
$stats = [PSCustomObject]@{
    'Total Records' = $combinedData.Count
    'Already Migrated' = ($combinedData | Where-Object { $_.MigrationStatus -eq "Already Migrated" }).Count
    'Migration Required' = ($combinedData | Where-Object { $_.MigrationStatus -eq "Migration Required" }).Count
    'Number Mismatch' = ($combinedData | Where-Object { $_.MigrationStatus -eq "Number Mismatch" }).Count
    'Evaluate Need' = ($combinedData | Where-Object { $_.MigrationStatus -eq "Evaluate Need" }).Count
    'Teams Only' = ($combinedData | Where-Object { $_.MigrationStatus -eq "Teams Only" }).Count
    'No Action Needed' = ($combinedData | Where-Object { $_.MigrationStatus -eq "No Action Needed" }).Count
}

# Display summary
Write-Host "`nMigration Status Summary:"
Write-Host "========================================================"
Write-Host "Total Records:     $($stats.'Total Records')"
Write-Host "Already Migrated:  $($stats.'Already Migrated')"
Write-Host "Migration Required: $($stats.'Migration Required')"
Write-Host "Number Mismatch:   $($stats.'Number Mismatch')"
Write-Host "Evaluate Need:     $($stats.'Evaluate Need')"
Write-Host "Teams Only:        $($stats.'Teams Only')"
Write-Host "No Action Needed:  $($stats.'No Action Needed')"
Write-Host "========================================================"

Write-Host "`nReport complete. Use this data to plan your TeleVantage to Teams migration."