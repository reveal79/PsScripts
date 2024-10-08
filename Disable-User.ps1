<#
    Script: Reset and Disable User in Azure AD and Active Directory
    Description: This script resets a user's Azure AD password (twice), blocks their sign-in, revokes all sessions, 
                 and disables their account in on-premises Active Directory across multiple domains.
                 
    Modules Required:
      - AzureAD: This module allows managing Azure Active Directory users.
      - ActiveDirectory: This module allows managing on-premises Active Directory users.

    Usage:
    1. Ensure you have the AzureAD and ActiveDirectory PowerShell modules installed.
    2. Set the $ADDomains variable with your Active Directory domains.
    3. Run the script, and when prompted, enter your AD credentials and the UPN (email) of the user whose account 
       you want to reset and disable.
    4. The script will reset the user's password twice in Azure AD, block their sign-in, revoke all Azure AD sessions, 
       and disable their account in Active Directory across the specified domains.
    5. After processing, the script will prompt you if you want to reset another user.

    Example:
    - Domains: @("yourdomain.local", "otherdomain.com")
    - User to reset: john.doe@company.com
    
    Note: The script is designed for hybrid environments where both Azure AD and on-premises AD accounts exist.
#>

# Import necessary modules for interacting with Azure AD and Active Directory
# Make sure these modules are installed on your system.
Import-Module AzureAD
Import-Module ActiveDirectory

# Prompt for Active Directory credentials to perform actions in on-premises AD
$adCreds = Get-Credential

# Function to reset a user's password and block their sign-in in Azure AD
function Reset-AzureADUser {
    param(
        [Parameter(Mandatory=$true)]
        [string]$userPrincipalName
    )

    $userExistsInAzure = $false

    # Check if the user exists in Azure AD
    try {
        $user = Get-AzureADUser -ObjectId $userPrincipalName
        $userExistsInAzure = $true
    }
    catch {
        Write-Host "User not found in Azure AD: $_"
    }

    if ($userExistsInAzure) {
        # Generate two random passwords for the user
        Add-Type -AssemblyName System.Web
        $randomPassword1 = [System.Web.Security.Membership]::GeneratePassword(10, 2)
        $randomPassword2 = [System.Web.Security.Membership]::GeneratePassword(10, 2)
        $newPassword1 = ConvertTo-SecureString $randomPassword1 -AsPlainText -Force
        $newPassword2 = ConvertTo-SecureString $randomPassword2 -AsPlainText -Force

        # Set the first password for the user
        try {
            Set-AzureADUserPassword -ObjectId $userPrincipalName -Password $newPassword1
            Write-Host "First password reset successful"
        }
        catch {
            Write-Host "Error resetting first password: $_"
        }

        # Wait for a few seconds before setting the second password
        Start-Sleep -s 5

        # Set the second password for the user
        try {
            Set-AzureADUserPassword -ObjectId $userPrincipalName -Password $newPassword2
            Write-Host "Second password reset successful"
        }
        catch {
            Write-Host "Error resetting second password: $_"
        }

        # Block the user from signing in to Azure AD
        try {
            Set-AzureADUser -ObjectId $userPrincipalName -AccountEnabled $false
            Write-Host "Sign-in block in Azure AD successful"
        }
        catch {
            Write-Host "Error blocking sign-in in Azure AD: $_"
        }

        # Revoke all sessions in Azure AD
        try {
            Revoke-AzureADUserAllRefreshToken -ObjectId $userPrincipalName
            Write-Host "All sessions revoked successfully in Azure AD"
        }
        catch {
            Write-Host "Error revoking sessions in Azure AD: $_"
        }
    }
}

# Function to disable a user account in Active Directory across multiple domains
function Disable-ADUserAccount {
    param(
        [Parameter(Mandatory=$true)]
        [string]$userPrincipalName,
        [Parameter(Mandatory=$true)]
        [string[]]$ADDomains,
        [Parameter(Mandatory=$true)]
        [pscredential]$adCreds
    )

    $userFoundInAD = $false
    foreach ($domain in $ADDomains) {
        try {
            # Search for the user in the specified AD domain
            $adUser = Get-ADUser -Filter {UserPrincipalName -eq $userPrincipalName} -Server $domain -Credential $adCreds
            if ($adUser) {
                $userFoundInAD = $true
                # Disable the user account if found
                Disable-ADAccount -Identity $adUser -Credential $adCreds -Server $domain
                Write-Host "Active Directory account disabled successfully in domain $domain"
            } else {
                Write-Host "User not found in Active Directory domain $domain"
            }
        }
        catch {
            Write-Host "Error disabling Active Directory account in domain ${domain}: $_"
        }
    }

    if (-not $userFoundInAD) {
        Write-Host "User not found in specified Active Directory domains."
    }
}

# Connect to Azure AD
Connect-AzureAD

# List of Active Directory domains to check
$ADDomains = @("domain1.com", "domain2.com")  # Replace with your AD domains (e.g., yourcompany.local)

# User reset loop
do {
    # Prompt for the user email (User Principal Name)
    $userPrincipalName = Read-Host -Prompt 'Input the email address (UserPrincipalName) of the user you want to reset'

    # Reset Azure AD user
    Reset-AzureADUser -userPrincipalName $userPrincipalName

    # Disable the AD user account in all specified AD domains
    Disable-ADUserAccount -userPrincipalName $userPrincipalName -ADDomains $ADDomains -adCreds $adCreds

    # Prompt to reset another user
    $resetAnother = Read-Host -Prompt 'Do you want to reset another user? (yes/no)'
} while ($resetAnother -eq 'yes')

# Disconnect from Azure AD
Disconnect-AzureAD
