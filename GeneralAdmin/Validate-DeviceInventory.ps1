<#
.SYNOPSIS
    Validates device inventory across Active Directory, Entra ID, and Intune.

.DESCRIPTION
    This script identifies discrepancies between Active Directory, Entra ID (Azure AD), 
    and Intune by matching devices based on asset tags extracted from their names.
    
    It helps detect:
    - Devices that exist in one system but not another.
    - Devices that have been renamed but retain the same asset tag.
    - Stale device objects that could interfere with re-enrollment.
    
    Results are displayed in a formatted table and exported to CSV for further review.

.NOTES
    Author: Don Cook
    Created: 2025-03-03
    Version: 1.0
    Purpose: Validation-only script to identify device mismatches before cleanup.

.REQUIREMENTS
    - Active Directory module (`RSAT: Active Directory PowerShell`).
    - Microsoft Graph PowerShell module (`Microsoft.Graph`).
    - Requires permissions to query AD, Entra ID, and Intune.
    - Run in a PowerShell session with administrative privileges.

.OUTPUTS
    - Displays mismatched devices in PowerShell.
    - Exports results to CSV at: `$env:USERPROFILE\Desktop\DeviceValidationResults.csv`.

#>

# Import required modules
Import-Module ActiveDirectory
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.DeviceManagement

# Authenticate with Microsoft Graph (for Entra and Intune queries)
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "Device.Read.All", "DeviceManagementManagedDevices.Read.All"

# Retrieve all Active Directory computers
$ADComputers = Get-ADComputer -Filter * -Property Name, DistinguishedName | Select-Object Name, DistinguishedName

# Retrieve all Entra ID (Azure AD) devices
$EntraDevices = Get-MgDevice | Select-Object DisplayName, Id, DeviceId

# Retrieve all Intune managed devices
$IntuneDevices = Get-MgDeviceManagementManagedDevice | Select-Object DeviceName, Id, ManagedDeviceId, OperatingSystem

# Function: Extract Asset Tag (last part of the name)
function Get-AssetTag {
    param ([string]$ComputerName)
    if ($ComputerName -match "-(\d+)$") {
        return $matches[1]  # Extracts last numeric portion (Asset Tag)
    }
    return $null
}

# Dictionary to store devices by Asset Tag
$DeviceMatches = @{}

# Process Active Directory devices
foreach ($ADComputer in $ADComputers) {
    $AssetTag = Get-AssetTag -ComputerName $ADComputer.Name
    if ($AssetTag) {
        $DeviceMatches[$AssetTag] += [PSCustomObject]@{
            Source       = "Active Directory"
            ComputerName = $ADComputer.Name
            AssetTag     = $AssetTag
        }
    }
}

# Process Entra ID devices
foreach ($EntraDevice in $EntraDevices) {
    $AssetTag = Get-AssetTag -ComputerName $EntraDevice.DisplayName
    if ($AssetTag) {
        $DeviceMatches[$AssetTag] += [PSCustomObject]@{
            Source       = "Entra ID"
            ComputerName = $EntraDevice.DisplayName
            AssetTag     = $AssetTag
        }
    }
}

# Process Intune devices
foreach ($IntuneDevice in $IntuneDevices) {
    $AssetTag = Get-AssetTag -ComputerName $IntuneDevice.DeviceName
    if ($AssetTag) {
        $DeviceMatches[$AssetTag] += [PSCustomObject]@{
            Source       = "Intune"
            ComputerName = $IntuneDevice.DeviceName
            AssetTag     = $AssetTag
        }
    }
}

# Prepare results for analysis
$Results = @()

foreach ($AssetTag in $DeviceMatches.Keys) {
    $Devices = $DeviceMatches[$AssetTag]

    # Extract unique names from different sources
    $ADName = ($Devices | Where-Object { $_.Source -eq "Active Directory" }).ComputerName
    $EntraName = ($Devices | Where-Object { $_.Source -eq "Entra ID" }).ComputerName
    $IntuneName = ($Devices | Where-Object { $_.Source -eq "Intune" }).ComputerName

    # Identify mismatches
    $NameMismatch = ($ADName -ne $EntraName) -or ($ADName -ne $IntuneName) -or ($EntraName -ne $IntuneName)
    $ExistsOnlyIn = @()
    if (-not $ADName) { $ExistsOnlyIn += "AD" }
    if (-not $EntraName) { $ExistsOnlyIn += "Entra" }
    if (-not $IntuneName) { $ExistsOnlyIn += "Intune" }

    # Store results
    $Results += [PSCustomObject]@{
        AssetTag      = $AssetTag
        AD_Computer   = $ADName
        Entra_Computer = $EntraName
        Intune_Computer = $IntuneName
        NameMismatch  = $NameMismatch
        ExistsOnlyIn  = ($ExistsOnlyIn -join ", ")
    }
}

# Display results
$Results | Sort-Object AssetTag | Format-Table -AutoSize

# Export results to CSV for review
$Results | Export-Csv -Path "$env:USERPROFILE\Desktop\DeviceValidationResults.csv" -NoTypeInformation
Write-Host "Results exported to Desktop as DeviceValidationResults.csv"