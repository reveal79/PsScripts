<#
    Script: Restore Deleted OneDrive and Assign Site Collection Admin
    Description: 
    This script is designed to restore a deleted OneDrive site for a user and assign a new Site Collection Admin 
    to the restored site. It checks if the OneDrive site exists in the deleted sites, attempts to restore it if found, 
    and sets the specified user as the Site Collection Admin. 

    Usage:
    1. The script checks if the required "Microsoft.Online.SharePoint.PowerShell" module is installed, 
       installs it if necessary, and imports it.
    2. You will be prompted to enter your tenant admin URL (e.g., https://abc123-admin.sharepoint.com).
    3. After validating the URL, the script connects to SharePoint Online.
    4. You will be prompted to enter the username of the deleted OneDrive user.
    5. The script formulates the OneDrive URL and checks for its existence in deleted sites.
    6. If the deleted site is found, the script attempts to restore it. 
    7. You will then be prompted to enter the email address of the new Site Collection Admin.
    8. Finally, the specified admin will be assigned to the restored OneDrive site.
#>

$moduleName = "Microsoft.Online.SharePoint.PowerShell"

# Check if the module is installed
if (!(Get-Module -ListAvailable -Name $moduleName)) {
    # If not installed, install the module
    Install-Module -Name $moduleName -Force
}

# Import the installed module
Import-Module $moduleName -Force

# Prompt for the tenant admin URL
$adminURL = Read-Host -Prompt "Enter your tenant admin URL (e.g., https://abc123-admin.sharepoint.com)"

# Validate the tenant admin URL format
if ($adminURL -notmatch "https://.+-admin\.sharepoint\.com$") {
    Write-Host "Invalid tenant admin URL format. Please enter a valid URL." -ForegroundColor Red
    exit 1
}

# Extract the tenant name from the admin URL
$tenantName = $adminURL -replace "https://(.+?)-admin.sharepoint.com", '$1'

# Error handling for connecting to SharePoint Online Service
try {
    Connect-SPOService -url $adminURL
    Write-Host "Connected to SharePoint Online successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to SharePoint Online." -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}

# Ask for the username
$userName = Read-Host -Prompt "Enter the username of the deleted OneDrive user (format: first.last@domain.com)"

# Convert username to the required format
$convertedUserName = $userName.Replace(".", "_").Replace("@", "_")

# Formulate the OneDrive URL
$oneDriveURL = "https://$tenantName-my.sharepoint.com/personal/" + $convertedUserName

# Validate the OneDrive URL format
if ($oneDriveURL -notmatch "https://.+-my\.sharepoint\.com/personal/.+$") {
    Write-Host "Invalid OneDrive URL format. Please check the tenant name and username." -ForegroundColor Red
    exit 1
}

# Error handling for getting deleted sites
try {
    # Get all deleted site collections
    $deletedSites = Get-SPODeletedSite -Identity $oneDriveURL

    if ($deletedSites -ne $null) {
        Write-Host "Found deleted OneDrive site for user $userName." -ForegroundColor Green
        Write-Host "Attempting to restore..." -ForegroundColor Green

        # Restore the site
        Restore-SPODeletedSite -Identity $oneDriveURL

        Write-Host "OneDrive site for user $userName has been restored." -ForegroundColor Green
    } else {
        Write-Host "No deleted OneDrive site found for user $userName." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Failed to get or restore the deleted OneDrive site." -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
}

# Prompt for the site collection admin email address
$adminUserEmail = Read-Host -Prompt "Enter the email address of the site collection admin"

# Error handling for setting site collection admin
try {
    # Set user as site collection admin
    Set-SPOUser -Site $oneDriveURL -LoginName $adminUserEmail -IsSiteCollectionAdmin $true
    Write-Host "The user $adminUserEmail has been set as site collection admin for $userName's OneDrive." -ForegroundColor Green
    Write-Host "The URL to share with the new Site Collection Admin is: $oneDriveURL"
    Write-Host "Please copy this URL and give it to the person requesting access we just added"
} catch {
    Write-Host "Failed to set site collection admin for the OneDrive site." -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
}
