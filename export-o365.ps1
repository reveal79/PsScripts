<#
.SYNOPSIS
    Retrieve and export a user's Azure AD group memberships, including Office 365 Groups, Distribution Lists, and Security Groups.

.DESCRIPTION
    This script connects to Azure AD and Microsoft Teams, retrieves the specified user's group memberships, 
    and exports the memberships into separate CSV files based on their group type: Office 365 Groups, Distribution Lists, 
    and Security Groups.

    The script first ensures that the required modules (AzureAD and MicrosoftTeams) are installed and then connects 
    to both Azure AD and Microsoft Teams services. It retrieves the user's group memberships and filters them into three 
    categories: Office 365 Groups, Distribution Lists, and Security Groups. The results are exported into separate CSV 
    files for further analysis or documentation purposes.

    Use Case:
    This script is useful for IT administrators who need to audit or document the group memberships of a specific user. 
    It allows for a clear separation of group types, making it easier to review Office 365 Groups, Distribution Lists, 
    and Security Groups individually. It can be particularly helpful for onboarding or offboarding tasks, group management, 
    or access reviews where detailed membership information is required.

.PARAMETER userEmail
    The email address of the user whose group memberships you want to retrieve.

.EXAMPLE
    # Retrieve group memberships for user and export to CSV
    .\Export-UserGroupMemberships.ps1

    This command will prompt for the user's email address, retrieve their group memberships, and export them into 
    separate CSV files based on the group type.

.NOTES
    Author: Don Cook
    Date: 2024-10-07

    Modules Required:
      - AzureAD: This module allows querying Azure Active Directory users and groups.
      - MicrosoftTeams: This module allows managing and retrieving Microsoft Teams-related data.

    The script ensures that both modules are installed and connects to both Azure AD and Microsoft Teams before performing 
    group membership retrieval and exporting the data.
#>

# Check if AzureAD module is installed
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "Installing AzureAD module..." -ForegroundColor Yellow
    Install-Module -Name AzureAD -Scope CurrentUser -Force
}

# Check if MicrosoftTeams module is installed
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Installing MicrosoftTeams module..." -ForegroundColor Yellow
    Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force
}

# Connect to Azure AD
Connect-AzureAD

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Get user's email address
$userEmail = Read-Host -Prompt "Enter the user's email address"

# Get user's ID
$user = Get-AzureADUser -Filter "Mail eq '$userEmail'"
$userId = $user.ObjectId

# Get all group memberships of the user
$groupMemberships = Get-AzureADUserMembership -ObjectId $userId | Where-Object { $_.ObjectType -eq "Group" }

# Filter groups into separate categories
$office365Groups = $groupMemberships | Where-Object { (Get-AzureADGroup -ObjectId $_.ObjectId).SecurityEnabled -eq $false }
$distributionLists = $groupMemberships | Where-Object { (Get-AzureADGroup -ObjectId $_.ObjectId).MailEnabled -eq $true }
$securityGroups = $groupMemberships | Where-Object { (Get-AzureADGroup -ObjectId $_.ObjectId).SecurityEnabled -eq $true -and (Get-AzureADGroup -ObjectId $_.ObjectId).MailEnabled -eq $false }

# Export groups to separate CSV files
$office365Groups | Select-Object DisplayName, ObjectId | Export-Csv -Path "Office365Groups-$userEmail.csv" -NoTypeInformation
$distributionLists | Select-Object DisplayName, ObjectId | Export-Csv -Path "DistributionLists-$userEmail.csv" -NoTypeInformation
$securityGroups | Select-Object DisplayName, ObjectId | Export-Csv -Path "SecurityGroups-$userEmail.csv" -NoTypeInformation

# Output success message
Write-Host "User's Office 365 Groups, Distribution Lists, and Security Groups have been exported to separate CSV files." -ForegroundColor Green
