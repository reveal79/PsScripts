# Connect to Microsoft 365
Connect-MsolService

# Get active users with SMS or voice MFA and without sign-in blocked
$users = Get-MsolUser -All | Where-Object { $_.StrongAuthenticationMethods.Count -gt 0 -and $_.BlockCredential -eq $false }
$usersWithSMSOrVoiceMFA = foreach ($user in $users) {
    $mfaMethods = $user.StrongAuthenticationMethods | Where-Object { $_.MethodType -eq "OneWaySMS" -or $_.MethodType -eq "PhoneAppOTP" }
    if ($mfaMethods) {
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            MFAMethods = $mfaMethods.MethodType -join ', '
        }
    }
}

# Filter users with SMS or voice as primary MFA registration
$usersWithPrimarySMSOrVoiceMFA = $usersWithSMSOrVoiceMFA | Where-Object { $_.MFAMethods -match 'OneWaySMS|PhoneAppOTP' }

# Export results to CSV
$usersWithPrimarySMSOrVoiceMFA | Export-Csv -Path "MFA_Users2.csv" -NoTypeInformation

# Disconnect from Microsoft 365
#Disconnect-MsolService
