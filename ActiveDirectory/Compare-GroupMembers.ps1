<#
    Script Name: Compare-GroupMembers.ps1
    Version: v1.0.0
    Author: Don Cook
    Last Updated: 2024-12-30
    Purpose:
    - Compares the membership of two Active Directory groups.
    - Identifies members exclusive to each group.
    - Outputs the results in a table format and copies the results to the clipboard.

    Service: Active Directory
    Service Type: Group Management

    Dependencies:
    - Active Directory Module for Windows PowerShell (RSAT tools must be installed).
    - Permissions to query group memberships in Active Directory.

    Notes:
    - The script supports recursive group membership resolution.
    - Ensure both group names entered are valid and exist in Active Directory.

    Example Usage:
    Run the script, enter the names of the two groups when prompted, and review the results.
    The output is displayed in the console and copied to the clipboard for easy sharing.
#>

# Import the Active Directory module if not already loaded
Import-Module ActiveDirectory

# Prompt the user for the names of the two groups
$group1Name = Read-Host "Enter the name of the first group"
$group2Name = Read-Host "Enter the name of the second group"

try {
    # Get the members of each group, including nested group members, and sort them alphabetically
    $group1Members = Get-ADGroupMember -Identity $group1Name -Recursive | Select-Object -ExpandProperty SamAccountName -Unique | Sort-Object
    $group2Members = Get-ADGroupMember -Identity $group2Name -Recursive | Select-Object -ExpandProperty SamAccountName -Unique | Sort-Object
} catch {
    Write-Error "An error occurred while retrieving group memberships: $_"
    exit
}

# Find members that are exclusive to each group
$onlyInGroup1 = $group1Members | Where-Object { $_ -notin $group2Members }
$onlyInGroup2 = $group2Members | Where-Object { $_ -notin $group1Members }

# Prepare the data for output
$dataForOutput = @()
$dataForOutput += "Total members in '$group1Name': $($group1Members.Count)"
$dataForOutput += "Total members in '$group2Name': $($group2Members.Count)"
$dataForOutput += ""
$dataForOutput += "Members in '$group1Name' not in '$group2Name' (Total: $($onlyInGroup1.Count)):"
$onlyInGroup1 | ForEach-Object { $dataForOutput += "- $_" }
$dataForOutput += ""
$dataForOutput += "Members in '$group2Name' not in '$group1Name' (Total: $($onlyInGroup2.Count)):"
$onlyInGroup2 | ForEach-Object { $dataForOutput += "- $_" }

# Convert the data to a string
$stringOutput = $dataForOutput -join "`r`n"

# Output the results in a table format and copy to clipboard
try {
    Write-Output $stringOutput | Out-String | Set-Clipboard
    Write-Host "Results have been copied to the clipboard." -ForegroundColor Green
} catch {
    Write-Error "Failed to copy results to the clipboard."
}

# Additionally display the output in the PowerShell window
Write-Output $stringOutput