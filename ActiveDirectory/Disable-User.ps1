<#
.SYNOPSIS
Reset a user's account in both Azure AD and on-premises Active Directory.

.DESCRIPTION
This script ensures it is run as an administrator, checks for and installs necessary modules, and validates the list of Active Directory domains. If default domains are detected, it prompts the user to update them dynamically or cancel the script.

.SERVICE
Active Directory, Azure AD

.SERVICE TYPE
User Management

.VERSION
1.3.0

.AUTHOR
Don Cook

.LAST UPDATED
2024-12-30

.NOTES
- Ensure proper permissions to reset passwords and disable accounts in both Azure AD and Active Directory.
- Update the $ADDomains array with your specific domains before use.
#>

# Ensure script is run as an administrator
function Ensure-AdminPrivilege {
    if (-not ([bool](New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        Write-Host "This script must be run as an administrator. Please restart PowerShell as Administrator and try again." -ForegroundColor Red
        exit 1
    }
}
Ensure-AdminPrivilege

# Ensure necessary modules are installed
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
            Write-Host "Failed to install the module '$ModuleName'. Please ensure you have internet access and try again." -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Module '$ModuleName' is already installed." -ForegroundColor Green
    }
}

# Check and install required modules
Ensure-Module -ModuleName "AzureAD"
Ensure-Module -ModuleName "ActiveDirectory"

# Import the modules
Import-Module AzureAD -ErrorAction Stop
Import-Module ActiveDirectory -ErrorAction Stop

# Default AD domains (update as necessary)
$ADDomains = @("domain1.local", "domain2.com") # Change these to your actual domains

# Function to validate and update AD domains
function Validate-ADDomains {
    if ($ADDomains -contains "domain1.local" -or $ADDomains -contains "domain2.com") {
        Write-Host "WARNING: The default AD domains are still in use!" -ForegroundColor Yellow

        # Prompt user with a warning message
        $choice = [System.Windows.Forms.MessageBox]::Show(
            "The script is using default AD domains (`domain1.local`, `domain2.com`)." +
            "`nDo you want to update the domains now or cancel to edit the script manually?",
            "Domain Configuration Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        switch ($choice) {
            "Yes" {
                # Prompt for new domains dynamically
                $newDomains = @()
                do {
                    $newDomain = Read-Host -Prompt "Enter an AD domain (leave blank to finish)"
                    if (-not [string]::IsNullOrWhiteSpace($newDomain)) {
                        $newDomains += $newDomain
                    }
                } while (-not [string]::IsNullOrWhiteSpace($newDomain))

                if ($newDomains.Count -eq 0) {
                    Write-Host "No domains provided. Exiting..." -ForegroundColor Red
                    exit 1
                }

                # Update the domain list for this session
                $global:ADDomains = $newDomains
                Write-Host "AD domains updated: $($ADDomains -join ', ')" -ForegroundColor Green
            }
            "No" {
                Write-Host "Please edit the script to update the AD domain list." -ForegroundColor Red
                exit 1
            }
            "Cancel" {
                Write-Host "Operation canceled by user." -ForegroundColor Cyan
                exit 1
            }
        }
    } else {
        Write-Host "AD domains are correctly configured: $($ADDomains -join ', ')" -ForegroundColor Green
    }
}

# Validate AD domains before proceeding
Validate-ADDomains

# Prompt for credentials
$adCreds = Get-Credential -Message "Enter your Active Directory admin credentials"

# Connect to Azure AD
try {
    Connect-AzureAD
    Write-Host "Connected to Azure AD successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Azure AD. Ensure you have the correct credentials and permissions." -ForegroundColor Red
    exit 1
}

# User reset operations can be added below...
Write-Host "Proceeding with user reset operations..." -ForegroundColor Cyan

# Example of operations
Write-Host "This is a placeholder for actual user reset functionality." -ForegroundColor Gray

# Disconnect from Azure AD
Disconnect-AzureAD
Write-Host "Disconnected from Azure AD. Script completed." -ForegroundColor Cyan