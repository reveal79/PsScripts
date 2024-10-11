# --------------------------------------------
# Script Name: Reset-AzureADUser.ps1
# Purpose: To reset an Azure AD user account in the event of compromise. This script handles:
# - Password resets
# - Blocking sign-ins
# - Revoking active sessions
# - Resetting MFA registrations
# - Removing app-specific passwords (for legacy applications)
#
# Intended Audience: IT administrators responsible for managing Azure Active Directory users,
# especially in security or incident response scenarios where a user account is suspected of being compromised.
#
# How to Use:
# 1. Make sure to have both AzureAD and MSOnline modules installed.
# 2. Run this script with administrative privileges.
# 3. Provide the User Principal Name (email address) of the compromised user when prompted.
# 4. The script will handle resetting the user's account and securing it.
# 5. If multiple users need to be reset, rerun the script for each user to ensure consistency.
#
# Author: Don Cook
# --------------------------------------------

# Import necessary modules
Import-Module AzureAD
Import-Module MSOnline

# Function: Reset-AzureADUser
# This function handles the core steps to reset and secure a compromised Azure AD user account.
# Steps:
# 1. Check if the user exists in Azure AD.
# 2. Reset the user's password twice (as an extra security measure).
# 3. Block the user's sign-in to prevent further access.
# 4. Revoke all active sessions (includes signing the user out of all devices).
# 5. Remove app-specific passwords (for legacy applications using basic auth).
# 6. Reset the user's MFA registration to force re-registration of MFA devices.

function Reset-AzureADUser {
    param(
        [Parameter(Mandatory=$true)]
        [string]$userPrincipalName # Input: User Principal Name (UPN) of the compromised user.
    )

    Write-Host "Starting the reset process for user: $userPrincipalName..."

    # Step 1: Check if user exists in Azure AD
    try {
        $user = Get-AzureADUser -ObjectId $userPrincipalName
    }
    catch {
        Write-Host "User not found: $_"
        return
    }

    # Step 2: Generate two random passwords
    # This is an extra security measure to change the password twice in quick succession.
    Add-Type -AssemblyName System.Web
    $randomPassword1 = [System.Web.Security.Membership]::GeneratePassword(10, 2)
    $randomPassword2 = [System.Web.Security.Membership]::GeneratePassword(10, 2)
    $newPassword1 = ConvertTo-SecureString $randomPassword1 -AsPlainText -Force
    $newPassword2 = ConvertTo-SecureString $randomPassword2 -AsPlainText -Force

    # Step 3: Set the first password
    try {
        Set-AzureADUserPassword -ObjectId $userPrincipalName -Password $newPassword1
        Write-Host "First password reset successful."
    }
    catch {
        Write-Host "Error resetting first password: $_"
        return
    }

    # Step 4: Wait for 5 seconds before resetting the password again
    Start-Sleep -s 5

    # Step 5: Set the second password (extra layer of security)
    try {
        Set-AzureADUserPassword -ObjectId $userPrincipalName -Password $newPassword2
        Write-Host "Second password reset successful."
    }
    catch {
        Write-Host "Error resetting second password: $_"
        return
    }

    # Step 6: Block the user's sign-in to prevent further access
    try {
        Set-AzureADUser -ObjectId $userPrincipalName -AccountEnabled $false
        Write-Host "Sign-in block successful."
    }
    catch {
        Write-Host "Error blocking sign-in: $_"
        return
    }

    # Step 7: Revoke all refresh tokens (this will invalidate any access tokens and force sign-in again)
    try {
        Revoke-AzureADUserAllRefreshToken -ObjectId $userPrincipalName
        Write-Host "All refresh tokens revoked successfully."
    }
    catch {
        Write-Host "Error revoking refresh tokens: $_"
        return
    }

    # Step 8: Revoke all active sessions across devices (force logout on all active devices)
    try {
        Invoke-AzureADSignedInUserSignOut -UserPrincipalName $userPrincipalName
        Write-Host "All active sessions signed out successfully."
    }
    catch {
        Write-Host "Error signing out all active sessions: $_"
        return
    }

    # Step 9: Remove any app-specific passwords (used by legacy apps that don’t support modern authentication)
    try {
        $appPasswords = Get-MsolUser -UserPrincipalName $userPrincipalName | Select-Object -ExpandProperty StrongAuthenticationUserDetails
        if ($appPasswords -ne $null) {
            Remove-MsolStrongAuthenticationMethod -ObjectId $userPrincipalName
            Write-Host "App-specific passwords removed successfully."
        } else {
            Write-Host "No app-specific passwords found."
        }
    }
    catch {
        Write-Host "Error removing app-specific passwords: $_"
        return
    }

    # Step 10: Reset MFA registration to force the user to re-register their MFA devices
    try {
        Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName $userPrincipalName
        Write-Host "MFA registration reset successful."
    }
    catch {
        Write-Host "Error resetting MFA registration: $_"
        return
    }

    Write-Host "User reset complete for: $userPrincipalName"
}

# Step 11: Connect to both Azure AD and MSOnline (required for different operations)
Connect-AzureAD
Connect-MsolService

# Input: Get the user principal name (email address) of the compromised user
$userPrincipalName = Read-Host -Prompt 'Input the email address of the user you want to reset'
Reset-AzureADUser -userPrincipalName $userPrincipalName

# Step 12: Disconnect from Azure AD and MSOnline once the operation is complete
Disconnect-AzureAD
Disconnect-MsolService

Write-Host "Azure AD user reset script completed."
