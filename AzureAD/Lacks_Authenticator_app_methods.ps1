<#
.SYNOPSIS
    Retrieve and export users without Authenticator App MFA methods in Azure AD.

.DESCRIPTION
    This script checks all licensed users in Azure AD who do not have their credentials blocked and identifies those 
    without Authenticator App methods configured for MFA (such as PhoneAppOTP or PhoneAppNotification). It collects 
    user details, including DisplayName, UserPrincipalName, Department, Title, and MFA status. The results are exported 
    to a CSV file for further review or analysis.

    Use Case:
    This script is useful for IT administrators who need to audit user MFA configurations, specifically focusing on users 
    who lack Authenticator App methods for MFA. This can help in enforcing security policies that require more secure 
    authentication methods, and the exported CSV can be used for compliance or tracking purposes.

.PARAMETER csvPath
    The file path where the results will be exported as a CSV file.

.EXAMPLE
    # Retrieve users without Authenticator App MFA methods and export to CSV
    .\Export-UsersWithoutAuthenticatorAppMFA.ps1

    This command will check all users in Azure AD and export the results to a CSV file named `UsersWithoutAuthenticatorAppMFA.csv` 
    in the C:\temp directory.

.NOTES
    Author: Don Cook
    Date: 2024-10-07

    Modules Required:
      - MSOnline: This module allows querying Azure AD users and their MFA methods.

    The script installs the required MSOnline module if it is not already present and connects to Azure AD.
#>

# Ensure MSOnline module is available
if (-not (Get-Module -ListAvailable -Name MSOnline)) {
    Write-Host "MSOnline module is required. Installing now..."
    Install-Module MSOnline -Repository PSGallery -AllowClobber -Force
}

# Connect to Azure AD
Connect-MsolService

# Function to check if a user lacks Authenticator app methods
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
            MFAMethods        =
