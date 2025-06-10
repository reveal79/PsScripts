<#
.SYNOPSIS
    Optimized version to validate device inventory across Active Directory, Entra ID, and Intune.

.DESCRIPTION
    - Targets specific AD OUs to reduce query load.
    - Uses Graph API filters to limit Entra/Intune results.
    - Runs queries in parallel using jobs instead of `-Parallel` (fixes PowerShell variable scope issues).
    - Merges results into a single dataset for easy validation.

.NOTES
    Author: Don Cook
    Version: 1.2 (Fixed Parallel Issues)
    Requires: Microsoft.Graph Module, RSAT: Active Directory PowerShell.
#>

# Import required modules
Import-Module ActiveDirectory
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.DeviceManagement

# Authenticate with Microsoft Graph (for Entra and Intune queries)
Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "Device.Read.All", "DeviceManagementManagedDevices.Read.All"

# Function: Extract Asset Tag (last part of the name)
function Get-AssetTag {
    param ([string]$ComputerName)
    if ($ComputerName -match "-(\d+)$") {
        return $matches[1]  # Extracts last numeric portion (Asset Tag)
    }
    return $null
}

# ✅ Optimized: Target a specific AD OU instead of all computers
$ADComputers = Get-ADComputer -SearchBase "OU=Computers,DC=YourDomain,DC=com" `
    -Filter * -Property Name, WhenChanged |
    Where-Object { $_.WhenChanged -gt (Get-Date).AddMonths(-3) } |  # Limit to recently modified
    Select-Object Name

# ✅ Optimized: Use Graph API filters to retrieve only active devices
$EntraDevices = Start-Job -ScriptBlock {
    Import-Module Microsoft.Graph.DeviceManagement
    Get-MgDevice -Filter "AccountEnabled eq true" | Select-Object DisplayName, Id, DeviceId
}

$IntuneDevices = Start-Job -ScriptBlock {
    Import-Module Microsoft.Graph.DeviceManagement
    Get-MgDeviceManagementManagedDevice -Filter "managementState eq 'managed'" | Select-Object DeviceName, Id, ManagedDeviceId, OperatingSystem
}

# Wait for jobs to complete
Wait-Job -Id $EntraDevices.Id, $IntuneDevices.Id

# Retrieve results from jobs
$EntraDevices = Receive-Job -Id $EntraDevices.Id
$IntuneDevices = Receive-Job -Id $IntuneDevices.Id

# Cleanup jobs
Remove-Job -Id $EntraDevices.Id, $IntuneDevices.Id

# Dictionary to store devices by Asset Tag
$DeviceMatches = @{}

# ✅ Process AD devices and store results
foreach ($ADComputer in $ADComputers) {
    $AssetTag = Get-AssetTag -ComputerName $ADComputer.Name
    if ($AssetTag) {
        if (-not $DeviceMatches[$AssetTag]) { $DeviceMatches[$AssetTag] = @() }
        $DeviceMatches[$AssetTag] += [PSCustomObject]@{
            Source       = "Active Directory"
            ComputerName = $ADComputer.Name
            AssetTag     = $AssetTag
        }
    }
}

# ✅ Process Entra ID devices
foreach ($EntraDevice in $EntraDevices) {
    $AssetTag = Get-AssetTag -ComputerName $EntraDevice.DisplayName
    if ($AssetTag) {
        if (-not $DeviceMatches[$AssetTag]) { $DeviceMatches[$AssetTag] = @() }
        $DeviceMatches[$AssetTag] += [PSCustomObject]@{
            Source       = "Entra ID"
            ComputerName = $EntraDevice.DisplayName
            AssetTag     = $AssetTag
        }
    }
}

# ✅ Process Intune devices
foreach ($IntuneDevice in $IntuneDevices) {
    $AssetTag = Get-AssetTag -ComputerName $IntuneDevice.DeviceName
    if ($AssetTag) {
        if (-not $DeviceMatches[$AssetTag]) { $DeviceMatches[$AssetTag] = @() }
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