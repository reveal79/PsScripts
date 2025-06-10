<#
    Script Name: Export-UsersWithoutAuthenticatorApp-MSOL.ps1
    Version: v1.1.0
    Author: Don Cook
    Last Updated: 2024-12-30
    Purpose:
    - Identifies licensed users without Authenticator app MFA methods using MSOnline.
    - Includes additional user information such as department and title.
    - Exports the results to a CSV file.

    Service: Office 365
    Service Type: MFA Management

    Dependencies:
    - MSOnline module for querying user and authentication method data.
    - Permissions to query licensed users in Microsoft 365.

    Notes:
    - Retains MSOnline compatibility for environments that require it.
    - Automatically installs the required module if missing.

    Example Usage:
    Run the script directly. The results will be exported to a CSV file located at `C:\temp\UsersWithoutAuthenticatorAppMFA.csv`.
#>

# Ensure the MSOnline module is available
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
            Write-Error "Failed to install the module '$ModuleName'. Ensure you have internet access and permissions to install modules."
            Write-Host "To resolve this issue manually, run: Install-Module -Name $ModuleName -Scope CurrentUser -Force"
            exit
        }
    }
}

# Ensure MSOnline module is installed
Ensure-Module -ModuleName "MSOnline"

# Import the MSOnline module
Import-Module MSOnline -ErrorAction Stop

# Connect to Microsoft Online
try {
    Connect-MsolService
} catch {
    Write-Error "Failed to connect to Microsoft Online. Ensure your credentials are valid and your environment is configured correctly."
    exit
}

# Function to check if a user lacks Authenticator app methods
function LacksAuthenticatorAppMethods {
    param (
        [array]$UserMethods
    )
    $hasAuthenticatorApp = $UserMethods | Where-Object {
        $_.MethodType -eq "PhoneAppOTP" -or $_.MethodType -eq "PhoneAppNotification"
    }
    return -not $hasAuthenticatorApp
}

# Collect users without Authenticator app MFA methods
$usersWithoutAuthApp = @()
try {
    $allUsers = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true -and $_.BlockCredential -eq $false }

    foreach ($user in $allUsers) {
        if (LacksAuthenticatorAppMethods -UserMethods $user.StrongAuthenticationMethods) {
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
} catch {
    Write-Error "An error occurred while retrieving user data: $_"
    exit
}

# Export results to CSV
$csvPath = "C:\temp\UsersWithoutAuthenticatorAppMFA.csv"
try {
    if ($usersWithoutAuthApp.Count -gt 0) {
        $usersWithoutAuthApp | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Export complete. File located at: $csvPath" -ForegroundColor Green
    } else {
        Write-Host "No users found without Authenticator app MFA methods." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Failed to export results to CSV: $_"
}

# Clean up session
Write-Host "Script execution completed successfully." -ForegroundColor Green