Param(
    [Parameter(Mandatory=$false)]$Email,
    [switch]$AllMFAConditional,
    [switch]$AllMFAPortal
)

# Ask for the email address if it's not provided as a parameter
if ([string]::IsNullOrEmpty($Email) -and !$AllMFAConditional -and !$AllMFAPortal) {
    $Email = Read-Host -Prompt 'Please provide an email address'
}

# Ensuring the MSOnline module is imported
Import-Module MSOnline

# Connect to MSOL Service
Connect-MsolService

If ($AllMFAConditional) {
    $vUser = Get-MSOLUser -All | Where-Object {$_.StrongAuthenticationMethods -ne $null }

    Write-Host "`nNumber of users with MFA (Conditional) enabled: " $vUser.count
    Write-Host "`nList of all users configured with MFA (Conditional).."
    $vUser | ForEach-Object { Write-Host "$($_.DisplayName) - $($_.UserPrincipalName)" }
    exit
}

If ($AllMFAPortal){
    $vUser = Get-MsolUser -all | where-Object { $_.StrongAuthenticationRequirements.State -ne $null }

    Write-Host "`nNumber of users with MFA (Portal) enabled: " $vUser.count
    Write-Host "`nList of all users configured with MFA (Portal).."
    $vUser | ForEach-Object { Write-Host "$($_.DisplayName) - $($_.UserPrincipalName)" }
    exit
}

$vUser = Get-MsolUser -UserPrincipalName $Email -ErrorAction SilentlyContinue

If ($vUser) {
    Write-Host "`nUser Details for $Email`n"

    Write-Host "Self-Service Password Feature (SSP)..: " -NoNewline;
    If ($vUser.StrongAuthenticationUserDetails) {  
        Write-Host -ForegroundColor Green "Enabled"
    } Else { 
        Write-Host -ForegroundColor Yellow "Not Configured"
    }

    Write-Host "MFA Feature (Portal) ................: " -NoNewline;
    If ($null -ne (($vuser | Select-Object -ExpandProperty StrongAuthenticationRequirements).State)) { 
        Write-Host -ForegroundColor Yellow "Enabled! It overrides Conditional"
    } Else { 
        Write-Host -ForegroundColor Green "Not Configured"
    }

    Write-Host "MFA Feature (Conditional)............: " -NoNewline;
    If ($vUser.StrongAuthenticationMethods){
        Write-Host -ForegroundColor Green "Enabled"
        Write-Host "`nAuthentication Methods:"
        for ($i=0;$i -lt $vuser.StrongAuthenticationMethods.Count;++$i){
            Write-Host "$($vUser.StrongAuthenticationMethods[$i].MethodType) ($($vUser.StrongAuthenticationMethods[$i].IsDefault))"
        }
        Write-Host "`nPhone entered by the end-user:"
        Write-Host "Phone Number.........: " $vuser.StrongAuthenticationUserDetails.PhoneNumber
        Write-Host "Alternative Number...: "$vuser.StrongAuthenticationUserDetails.AlternativePhoneNumber
    } Else{
        Write-Host -ForegroundColor Yellow "Not Configured"
    }

    Write-Host "`nLicense Requirements.................: "
    Write-Host "Licenses applied to the user:"

    $licenseMappings = @{
        "SPB" = "Microsoft 365 Business Standard"
        "POWER_BI_STANDARD" = "Power BI"
        "VISIOCLIENT" = "Visio"
        "PROJECTWORKMANAGEMENT" = "Project"
        "O365_BUSINESS_ESSENTIALS" = "Office 365 E1"
        "ENTERPRISEPACK" = "Office 365 E3"
        "ENTERPRISEPREMIUM" = "Office 365 E5"
        "EMS" = "Enterprise Mobility + Security"
        "DYN365_BUSINESS_PREMIUM" = "Dynamics 365 Business Premium"
        "DYN365_ENTERPRISE_PLAN1" = "Dynamics 365 Enterprise Plan 1"
        "DYN365_ENTERPRISE_PLAN2" = "Dynamics 365 Enterprise Plan 2"
        # Add more license mappings as needed
    }

    $licenseNames = $vuser.Licenses | ForEach-Object { $licenseMappings[$_.AccountSkuId.Split(":")[1]] }
    Write-Host "Licenses applied to the user: $($licenseNames -join ', ')"

    if ($vUser.AltEmailAddresses) {
        Write-Host "Alternate Email Address: $($vUser.AltEmailAddresses -join ', ')"
    }

    if ($vUser.SecurityQuestions -and $vUser.SecurityQuestions.Count -gt 0) {
        Write-Host "Security Questions:"
        foreach ($securityQuestion in $vUser.SecurityQuestions) {
            Write-Host "Question: $($securityQuestion.Question)"
            Write-Host "Answer: $($securityQuestion.Answer)"
        }
    }

    if ($vUser.AppPassword) {
        Write-Host "App Password: $($vUser.AppPassword)"
    }
} Else {
    Write-Host "`n"
    Write-Host -ForegroundColor Red "[Error]: User $Email couldn't be found. Check the email address and try again"
    exit
}
