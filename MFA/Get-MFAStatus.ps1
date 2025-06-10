<#
    Script Name: Get-MFAStatus-MSOL.ps1
    Version: v2.1.0
    Author: Don Cook
    Last Updated: 2024-12-30
    Purpose:
    - Retrieves detailed MFA and license information for a specified user.
    - Lists all users with MFA enabled via Conditional Access or Portal.

    Service: Office 365
    Service Type: MFA Management

    Dependencies:
    - MSOnline module for querying MFA and license data.
    - Permissions to query user data in Microsoft 365.

    Notes:
    - This script uses the MSOnline module, which provides access to detailed MFA configuration data.
    - Automatically installs the required module if missing.

    Example Usage:
    - Query a specific user: `Get-MFAStatus-MSOL.ps1 -Email user@domain.com`
    - List users with Conditional Access MFA: `Get-MFAStatus-MSOL.ps1 -AllMFAConditional`
    - List users with Portal MFA: `Get-MFAStatus-MSOL.ps1 -AllMFAPortal`
#>

Param(
    [Parameter(Mandatory = $false)]
    [string]$Email,
    [switch]$AllMFAConditional,
    [switch]$AllMFAPortal
)

# Function to ensure the required module is installed
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
            Write-Host "To resolve this issue manually, run: `Install-Module -Name $ModuleName -Scope CurrentUser -Force`."
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

# If no email is provided, prompt the user unless other switches are used
if ([string]::IsNullOrEmpty($Email) -and -not $AllMFAConditional -and -not $AllMFAPortal) {
    $Email = Read-Host -Prompt 'Please provide an email address'
}

# List all users with Conditional Access MFA enabled
if ($AllMFAConditional) {
    $vUsers = Get-MsolUser -All | Where-Object { $_.StrongAuthenticationMethods.MethodType -eq "ConditionalAccess" }
    Write-Host "`nNumber of users with MFA (Conditional) enabled: $($vUsers.Count)" -ForegroundColor Green
    Write-Host "`nList of all users configured with MFA (Conditional):"
    $vUsers | ForEach-Object { Write-Host "$($_.DisplayName) - $($_.UserPrincipalName)" }
    exit
}

# List all users with Portal MFA enabled
if ($AllMFAPortal) {
    $vUsers = Get-MsolUser -All | Where-Object { $_.StrongAuthenticationRequirements.State -ne $null }
    Write-Host "`nNumber of users with MFA (Portal) enabled: $($vUsers.Count)" -ForegroundColor Green
    Write-Host "`nList of all users configured with MFA (Portal):"
    $vUsers | ForEach-Object { Write-Host "$($_.DisplayName) - $($_.UserPrincipalName)" }
    exit
}

# Retrieve user-specific details
if ($Email) {
    $vUser = Get-MsolUser -UserPrincipalName $Email -ErrorAction SilentlyContinue

    if ($vUser) {
        Write-Host "`nUser Details for $Email`n"

        # Self-Service Password Reset (SSPR)
        Write-Host "Self-Service Password Reset (SSPR): " -NoNewline
        if ($vUser.StrongAuthenticationUserDetails) {
            Write-Host -ForegroundColor Green "Enabled"
        } else {
            Write-Host -ForegroundColor Yellow "Not Configured"
        }

        # MFA (Portal)
        Write-Host "MFA (Portal): " -NoNewline
        if ($vUser.StrongAuthenticationRequirements.State) {
            Write-Host -ForegroundColor Yellow "Enabled (Overrides Conditional)"
        } else {
            Write-Host -ForegroundColor Green "Not Configured"
        }

        # MFA (Conditional)
        Write-Host "MFA (Conditional): " -NoNewline
        if ($vUser.StrongAuthenticationMethods.MethodType -eq "ConditionalAccess") {
            Write-Host -ForegroundColor Green "Enabled"
        } else {
            Write-Host -ForegroundColor Yellow "Not Configured"
        }

        # Authentication Methods
        if ($vUser.StrongAuthenticationMethods) {
            Write-Host "`nAuthentication Methods:"
            foreach ($method in $vUser.StrongAuthenticationMethods) {
                Write-Host " - $($method.MethodType) (Default: $($method.IsDefault))"
            }
        }

        # License Details
        Write-Host "`nLicenses applied to the user:"
        foreach ($license in $vUser.Licenses) {
            Write-Host " - $($license.AccountSkuId)"
        }
    } else {
        Write-Host -ForegroundColor Red "[Error]: User $Email could not be found. Check the email address and try again."
        exit
    }
}

# Disconnect session
Write-Host "Script execution completed successfully." -ForegroundColor Green