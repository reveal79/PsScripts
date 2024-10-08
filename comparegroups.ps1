<#
.SYNOPSIS
    Compare two Active Directory groups and identify unique members in each.

.DESCRIPTION
    This script compares two Active Directory groups, including their nested group members, and identifies 
    the members that are exclusive to each group. It retrieves all members of both groups, compares them, 
    and provides a list of members who are only in one of the groups but not both.

    The results are outputted to the console and copied to the clipboard for convenience.

    Use Case:
    This script is helpful for system administrators when comparing members of two Active Directory groups, 
    particularly when there are nested groups involved. It highlights the unique members in each group, 
    allowing administrators to quickly see differences between the groups, which is useful for access audits, 
    security reviews, or group management tasks.

.PARAMETER group1Name
    The name of the first Active Directory group to compare.

.PARAMETER group2Name
    The name of the second Active Directory group to compare.

.EXAMPLE
    # Compare two groups named "GroupA" and "GroupB"
    .\Compare-ADGroups.ps1

    This command will prompt for the names of the two groups, compare their members, and output the differences.

.NOTES
    Author: Don Cook
    Date: 2024-10-07

    Modules Required:
      - ActiveDirectory: This module allows managing and querying Active Directory groups and members.

    The script outputs the comparison results directly to the PowerShell window and copies the results to the clipboard for easy sharing.
#>

# Prompt the user for the names of the two groups
$group1Name = Read-Host "Enter the name of the first group"
$group2Name = Read-Host "Enter the name of the second group"

# Get the members of each group, including nested group members, and sort them alphabetically
$group1Members = Get-ADGroupMember -Identity $group1Name -Recursive | Select-Object -ExpandProperty SamAccountName -Unique | Sort-Object
$group2Members = Get-ADGroupMember -Identity $group2Name -Recursive | Select-Object -ExpandProperty SamAccountName -Unique | Sort-Object

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
Write-Output $stringOutput | Out-String | Set-Clipboard

# Additionally display the output in the PowerShell window
Write-Output $stringOutput
