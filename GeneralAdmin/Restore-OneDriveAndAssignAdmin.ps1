<#
.SYNOPSIS
Restore a deleted OneDrive site and assign a new Site Collection Admin.

.DESCRIPTION
This script restores a deleted OneDrive site for a specified user and assigns a new Site Collection Admin to the restored site.
It queries the tenant admin URL dynamically, validates inputs, handles errors gracefully, and provides feedback during the process.

.SERVICE
OneDrive for Business

.SERVICE TYPE
Site Collection Management

.VERSION
1.1.0

.AUTHOR
Don Cook

.LAST UPDATED
2024-12-30

.DEPENDENCIES
- Microsoft.Online.SharePoint.PowerShell module for managing SharePoint Online and OneDrive.

.PARAMETER userName
The username (UPN) of the user whose OneDrive site you want to restore, in the format `first.last@domain.com`.

.PARAMETER adminUserEmail
The email address of the new Site Collection Admin who will be granted access to the restored OneDrive site.

.EXAMPLE
# Run the script to restore a deleted OneDrive site and assign a new Site Collection Admin
$userName = "john.doe@company.com"
$adminUserEmail = "admin@company.com"
.\Restore-OneDriveAndAssignAdmin.ps1

.NOTES
- This script dynamically queries the tenant admin URL and ensures all dependencies are installed and configured.
- The Microsoft.Online.SharePoint.PowerShell module is required.
#>

# Define the required module
$moduleName = "Microsoft.Online.SharePoint.PowerShell"

# Ensure the required module is installed
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Host "The required module '$moduleName' is not installed. Installing now..." -ForegroundColor Yellow
    try {
        Install-Module -Name $moduleName -Force -ErrorAction Stop
        Write-Host "Module '$moduleName' installed successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install the module '$moduleName'. Ensure you have internet connectivity and retry."
        exit 1
    }
}
Import-Module $moduleName -Force

# Function to retrieve the tenant admin URL
function Get-TenantAdminUrl {
    try {
        Write-Host "Querying for tenant admin URL..." -ForegroundColor Yellow
        $tenantInfo = Get-SPOTenant
        if ($tenantInfo -and $tenantInfo.AdminCenterUrl) {
            Write-Host "Tenant Admin URL found: $($tenantInfo.AdminCenterUrl)" -ForegroundColor Green
            return $tenantInfo.AdminCenterUrl
        } else {
            Write-Error "Failed to retrieve the tenant admin URL. Ensure you have proper permissions."
            exit 1
        }
    } catch {
        Write-Error "Error querying tenant admin URL: $_"
        exit 1
    }
}

# Get the tenant admin URL dynamically
$adminURL = Get-TenantAdminUrl

# Connect to SharePoint Online
try {
    Connect-SPOService -Url $adminURL -ErrorAction Stop
    Write-Host "Connected to SharePoint Online successfully." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to SharePoint Online. Ensure the admin URL and credentials are correct."
    exit 1
}

# Prompt for the deleted user's username
$userName = Read-Host -Prompt "Enter the username of the deleted OneDrive user (format: first.last@domain.com)"

# Convert the username to OneDrive URL format
$convertedUserName = $userName.Replace(".", "_").Replace("@", "_")
$tenantName = $adminURL -replace "https://(.+?)-admin.sharepoint.com", '$1'
$oneDriveURL = "https://$tenantName-my.sharepoint.com/personal/$convertedUserName"

# Validate OneDrive URL format
if ($oneDriveURL -notmatch "https://.+-my\.sharepoint\.com/personal/.+$") {
    Write-Host "Invalid OneDrive URL format. Please check the tenant name and username." -ForegroundColor Red
    exit 1
}

# Attempt to restore the deleted OneDrive site
try {
    $deletedSite = Get-SPODeletedSite -Identity $oneDriveURL -ErrorAction Stop
    if ($deletedSite) {
        Write-Host "Found deleted OneDrive site for user $userName." -ForegroundColor Green
        Write-Host "Attempting to restore the site..." -ForegroundColor Green

        # Restore the site
        Restore-SPODeletedSite -Identity $oneDriveURL -ErrorAction Stop
        Write-Host "OneDrive site for user $userName has been restored successfully." -ForegroundColor Green
    } else {
        Write-Host "No deleted OneDrive site found for user $userName." -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Error "Failed to get or restore the deleted OneDrive site. Ensure the site exists and try again."
    exit 1
}

# Prompt for the Site Collection Admin email
$adminUserEmail = Read-Host -Prompt "Enter the email address of the new Site Collection Admin"

# Set the new Site Collection Admin
try {
    Set-SPOUser -Site $oneDriveURL -LoginName $adminUserEmail -IsSiteCollectionAdmin $true -ErrorAction Stop
    Write-Host "The user $adminUserEmail has been set as the Site Collection Admin for $userName's OneDrive." -ForegroundColor Green
    Write-Host "Access URL for the new admin: $oneDriveURL" -ForegroundColor Cyan
    Write-Host "Please share this URL with the new Site Collection Admin."
} catch {
    Write-Error "Failed to set the Site Collection Admin for the restored OneDrive site."
    exit 1
}

# Disconnect from SharePoint Online
Disconnect-SPOService
Write-Host "Disconnected from SharePoint Online. Script execution completed successfully." -ForegroundColor Green