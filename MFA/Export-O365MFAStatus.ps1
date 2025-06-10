<#
=============================================================================================
Name:           Export Office 365 MFA Status Report (Modern Authentication)
Description:    This script exports Microsoft 365 MFA status reports to CSV files.
Version:        3.0
Author:         Don Cook (Adapted from O365Reports Team)
Last Updated:   2024-12-30

Filters:
- EnabledOnly
- EnforcedOnly
- DisabledOnly
- AdminOnly
- LicensedUserOnly
- SignInAllowed $True / $False

Features:
- Exports MFA enabled and disabled user reports to CSV.
- Includes multiple filters for custom reporting.
- Uses modern authentication via ExchangeOnlineManagement module.
- Verifies required modules and installs them if missing.
- Provides clear error guidance for module installation issues.

Requirements:
- PowerShell
- ExchangeOnlineManagement module installed and configured
- Permissions to query user data in Microsoft 365

=============================================================================================
#>

# Function to check and install the required module
function Ensure-Module {
    param (
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "The required module '$ModuleName' is not installed. Attempting to install it..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Module '$ModuleName' installed successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to install the module '$ModuleName'. Please ensure you have internet access and permissions to install modules."
            Write-Host "To resolve this issue manually, run: Install-Module -Name $ModuleName -Scope CurrentUser -Force"
            exit
        }
    }
}

# Ensure the ExchangeOnlineManagement module is installed
Ensure-Module -ModuleName "ExchangeOnlineManagement"

# Import the ExchangeOnlineManagement module
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Microsoft Online with modern authentication
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Exchange Online. Ensure your credentials are valid and your environment is configured correctly."
    exit
}

# Initialize variables
$Result = ""
$Results = @()
$UserCount = 0
$PrintedUser = 0

# Define output file paths
$ExportCSV = ".\MFADisabledUserReport_$((Get-Date -format yyyy-MM-dd-HHmmss)).csv"
$ExportCSVReport = ".\MFAEnabledUserReport_$((Get-Date -format yyyy-MM-dd-HHmmss)).csv"

# Loop through each user and process MFA status
try {
    Get-ExoRecipient -RecipientTypeDetails UserMailbox -ResultSize Unlimited | ForEach-Object {
        $UserCount++
        $DisplayName = $_.DisplayName
        $Upn = $_.PrimarySmtpAddress
        $MFAStatus = $_.AuthenticationPolicy

        Write-Progress -Activity "Processed user count: $UserCount" -Status "Currently Processing: $DisplayName"

        # Skip users based on filters
        if (($SignInAllowed -ne $null) -and ([string]$SignInAllowed -ne [string]$_.SignInAllowed)) { return }
        if (($LicensedUserOnly.IsPresent) -and ($_.IsLicensed -eq $False)) { return }

        # Determine license status
        $LicenseStatus = if ($_.IsLicensed -eq $true) { "Licensed" } else { "Unlicensed" }

        # Process MFA status
        if ($MFAStatus -ne $null) {
            $PrintedUser++
            $Result = @{
                DisplayName = $DisplayName
                UserPrincipalName = $Upn
                MFAStatus = $MFAStatus
                LicenseStatus = $LicenseStatus
            }
            [PSCustomObject]$Result | Export-Csv -Path $ExportCSVReport -NoTypeInformation -Append
        } else {
            # Process MFA disabled users
            $PrintedUser++
            $Result = @{
                DisplayName = $DisplayName
                UserPrincipalName = $Upn
                MFAStatus = "Disabled"
                LicenseStatus = $LicenseStatus
            }
            [PSCustomObject]$Result | Export-Csv -Path $ExportCSV -NoTypeInformation -Append
        }
    }
} catch {
    Write-Error "An error occurred during user processing: $_"
}

# Display and open reports
if (Test-Path -Path $ExportCSVReport) {
    Write-Host "MFA Enabled user report available at: $ExportCSVReport" -ForegroundColor Green
    Invoke-Item $ExportCSVReport
} elseif (Test-Path -Path $ExportCSV) {
    Write-Host "MFA Disabled user report available at: $ExportCSV" -ForegroundColor Green
    Invoke-Item $ExportCSV
} else {
    Write-Host "No users found matching the specified criteria." -ForegroundColor Yellow
}

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false