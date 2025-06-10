<#
    Title: Export AD Distribution Groups to Excel
    Description: This script exports details of Active Directory (AD) distribution groups into individual Excel files, including group members and proxy addresses.
    Author: Don Cook
    Date: 2025-01-02
    Version: 1.0
    Tags: 
        - service: Active Directory
        - service: Reporting
        - task: Export Distribution Groups
        - task: Generate Excel Reports
        - script: PowerShell
    Usage:
        - Ensure the PowerShell session is running as an administrator.
        - Run the script to generate an Excel file for each AD distribution group in the specified directory.
        - Prerequisites:
            * ActiveDirectory module (installed via RSAT on Windows).
            * ImportExcel module (installed via PowerShell Gallery).
    Notes:
        - The script dynamically checks for required modules and installs them if missing.
        - Ensure you have proper permissions to query Active Directory and manage files on the local system.
        - The export directory is defined as `C:\DL_Reports\` but can be modified as needed.
#>

# Function to check if running as administrator
function Is-Administrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check and import a module
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$InstallCommand = $null
    )

    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
        Write-Host "Module '$ModuleName' is not installed. Attempting to install..." -ForegroundColor Yellow
        if (-not (Is-Administrator)) {
            Write-Error "The module '$ModuleName' cannot be installed because this session is not running as administrator. Please restart PowerShell as an administrator and try again."
            return
        }

        if ($InstallCommand) {
            Invoke-Expression $InstallCommand
        } else {
            Write-Error "No install command provided for module '$ModuleName'. Please install it manually."
            return
        }
    }
    Import-Module $ModuleName -ErrorAction Stop
    Write-Host "Module '$ModuleName' has been successfully imported." -ForegroundColor Green
}

# Check and import the Active Directory module
Ensure-Module -ModuleName "ActiveDirectory" `
              -InstallCommand "Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeAllSubFeature"

# Check and import the ImportExcel module
Ensure-Module -ModuleName "ImportExcel" `
              -InstallCommand "Install-Module -Name ImportExcel -Force -Scope CurrentUser"

# Define the directory to store the distribution list (DL) reports
$exportDirectory = "C:\DL_Reports\"
if (-Not (Test-Path -Path $exportDirectory)) {
    New-Item -ItemType Directory -Path $exportDirectory
}

# Retrieve all distribution groups from Active Directory
$groups = Get-ADGroup -Filter 'GroupCategory -eq "Distribution"' -Properties DisplayName, mail, proxyAddresses, member

foreach ($group in $groups) {

    # Retrieve member details for each distribution group
    $members = $group.member | ForEach-Object {
        $member = $_
        if ($member -match '^CN=(?<Name>[^,]+),') {
            $name = $matches['Name']
            if ($member -match '^CN=SystemMailbox') {
                $user = Get-ADObject -Identity $member -Properties DisplayName
                if ($user) {
                    $name = $user.DisplayName
                }
            }
            $name
        }
    }

    # Prepare member details for CSV output
    $membersString = if ($members) {
        ($members | ForEach-Object { "`"$($_)`"" }) -join "`r`n"
    } else {
        "0"  # Default value when there are no members
    }

    # Extract proxy addresses for CSV output
    $proxyAddressesList = $group.proxyAddresses | Where-Object { $_ -cmatch '^smtp:' }
    $proxyAddressesString = if ($proxyAddressesList) {
        ($proxyAddressesList | ForEach-Object { "`"$($_)`"" }) -join "`r`n"
    } else {
        "EMPTY"  # Default value when no proxy addresses are available
    }

    # Determine email field, set as "EMPTY" if not present
    $email = if ($group.mail) { $group.mail } else { "EMPTY" }

    # Calculate total number of members
    $totalMembers = if ($members) { $members.Count } else { 0 }

    # Generate a report name for the Excel file
    $reportName = "{0}_Members({1}).xlsx" -f ($group.Name -replace '[\/:*?"<>|]', ''), $totalMembers

    # Define the path for the Excel export
    $excelPath = Join-Path -Path $exportDirectory -ChildPath $reportName

    # Create a custom object for Excel export
    $excelExport = [PSCustomObject]@{
        "Group Name"     = $group.Name
        "E-Mail"         = $email
        "Members"        = $membersString
        "Proxy Addresses" = $proxyAddressesString
    }

    # Export group's information to an individual Excel file
    $excelExport | Export-Excel -Path $excelPath -BoldTopRow -WorksheetName "GroupInfo"
}