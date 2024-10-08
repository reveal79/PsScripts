<#
    Script: Export Active Users with SMS or Voice as Primary MFA
    Description:
    This script connects to Microsoft 365, retrieves all active users with Multi-Factor Authentication (MFA) enabled 
    using SMS or voice as their primary method, and exports the result to a CSV file. It ensures that only users who 
    are actively able to sign in (not blocked) and are using specific MFA methods (SMS or Phone App OTP) are included 
    in the export.

    Use Case:
    This script is useful for IT administrators who need to audit users' MFA settings in Microsoft 365, specifically 
    identifying users who have registered SMS or Voice-based MFA methods. It helps in reviewing MFA compliance or 
    in identifying users who may need to switch to more secure MFA methods like Authenticator apps. 
    Additionally, this report can help security teams monitor which users are still using potentially weaker MFA methods.
#>

# Connect to Microsoft 365
Connect-MsolService

# Get active users with MFA enabled (via SMS or Voice) and without sign-in blocked
$users = Get-MsolUser -All | Where-Object { $_.StrongAuthenticationMethods.Count -gt 0 -and $_.BlockCredential -eq $false }

# Iterate through users and find those using SMS or Phone App OTP as MFA methods
$usersWithSMSOrVoiceMFA = foreach ($user in $users) {
    $mfaMethods = $user.StrongAuthenticationMethods | Where-Object { $_.MethodType -eq "OneWaySMS" -or $_.MethodType -eq "PhoneAppOTP" }
    if ($mfaMethods) {
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName  # The UPN of the user
            MFAMethods = $mfaMethods.MethodType -join ', '  # MFA methods used by the user (e.g., SMS or Phone App OTP)
        }
    }
}

# Filter users with SMS or Voice as primary MFA registration method
$usersWithPrimarySMSOrVoiceMFA = $usersWithSMSOrVoiceMFA | Where-Object { $_.MFAMethods -match 'OneWaySMS|PhoneAppOTP' }

# Export the results to a CSV file for reporting purposes
$usersWithPrimarySMSOrVoiceMFA | Export-Csv -Path "MFA_Users2.csv" -NoTypeInformation

# Disconnect from Microsoft 365 after the script execution
#Disconnect-MsolService
