<#
.SYNOPSIS
    Audit Azure AD users for missing Authenticator app-based MFA methods.

.DESCRIPTION
    This PowerShell script identifies Azure AD users who are licensed and active but lack Authenticator app-based MFA methods 
    (PhoneAppOTP or PhoneAppNotification). It collects details such as Display Name, UPN, Department, Title, MFA Status, 
    and current MFA methods, then exports the results to a CSV file for review.

.PARAMETERS
    None

.OUTPUTS
    CSV file saved at C:\temp\UsersWithoutAuthenticatorAppMFA.csv

.NOTES
    Author: Don Cook
    Created: 11/19/2024
    Requires: PowerShell 5.1 or later, MSOnline module

.IMPORTANT
    Ensure you have the necessary permissions to access user data in Azure AD and run the script in an administrative PowerShell session.

#>

# Ensure MSOnline module is available
if (-not (Get-Module -ListAvailable -Name MSOnline)) {
    Write-Host "MSOnline module is required. Installing now..."
    Install-Module MSOnline -Repository PSGallery -AllowClobber -Force
}
Connect-MsolService

# Function to check if user lacks Authenticator app methods
function LacksAuthenticatorAppMethods {
    param($userMethods)
    $hasAuthenticatorApp = $false

    foreach ($method in $userMethods) {
        if ($method.MethodType -eq "PhoneAppOTP" -or $method.MethodType -eq "PhoneAppNotification") {
            $hasAuthenticatorApp = $true
            break
        }
    }

    return -not $hasAuthenticatorApp
}

# Collect users without Authenticator app MFA methods and include Department, Title
$usersWithoutAuthApp = @()
$allUsers = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true -and $_.BlockCredential -eq $false }

foreach ($user in $allUsers) {
    if (LacksAuthenticatorAppMethods -userMethods $user.StrongAuthenticationMethods) {
        $mfaStatus = if ($user.StrongAuthenticationRequirements.State) { $user.StrongAuthenticationRequirements.State } else { "Disabled" }
        $usersWithoutAuthApp += [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Department        = $user.Department
            Title             = $user.Title
            MFAStatus         = $mfaStatus
            MFAMethods        = ($user.StrongAuthenticationMethods | ForEach-Object { $_.MethodType }) -join ', '
        }
    }
}

# Export to CSV
$csvPath = "C:\temp\UsersWithoutAuthenticatorAppMFA.csv"
$usersWithoutAuthApp | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Export complete. File located at: $csvPath"
