# Teams-Phone-Numbers-Extractor.ps1
#===============================================================================
# Script Name: Teams-Phone-Numbers-Extractor.ps1
# Created On: April 1, 2025
#
# Description:
#   This script extracts phone number assignments from Microsoft Teams
#   to help with comparing existing phone systems during migration.
#
# Dependencies:
#   - MicrosoftTeams PowerShell module
#
# Usage:
#   .\Teams-Phone-Numbers-Extractor.ps1
#===============================================================================

# Set up the export path
$csvOutputFile = "C:\Temp\Teams_Phone_Numbers.csv"

# Ensure the Microsoft Teams module is installed
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Installing MicrosoftTeams module..."
    Install-Module -Name MicrosoftTeams -Force -AllowClobber
}

# Connect to Microsoft Teams with improved error handling
try {
    Write-Host "Connecting to Microsoft Teams..."
    Connect-MicrosoftTeams -ErrorAction Stop
    
    # Get Teams user data
    Write-Host "Retrieving Teams phone number assignments..."
    $teamsUsers = Get-CsOnlineUser -ErrorAction Stop | 
                  Where-Object {$_.LineURI -ne $null} | 
                  Select-Object DisplayName, UserPrincipalName, 
                                @{Name="PhoneNumber"; Expression={$_.LineURI -replace "tel:", ""}}, 
                                EnterpriseVoiceEnabled
    
    # Export to CSV format
    Write-Host "Exporting data to CSV..."
    $teamsUsers | Export-Csv -Path $csvOutputFile -NoTypeInformation
    
    # Output summary
    Write-Host "Found $($teamsUsers.Count) Teams users with phone numbers assigned"
    Write-Host "Data exported to: $csvOutputFile"
    
    # Disconnect from Teams
    Disconnect-MicrosoftTeams
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "If authentication failed, try running this script in STA mode:" -ForegroundColor Yellow
    Write-Host "PowerShell -STA -File $PSCommandPath" -ForegroundColor Yellow
}