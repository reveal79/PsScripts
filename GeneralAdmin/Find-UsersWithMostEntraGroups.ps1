# Find-UsersWithMostEntraGroups.ps1
# Script to find which users have the most security groups assigned in Microsoft Entra ID

param(
    [Parameter(Mandatory=$false)]
    [int]$TopCount = 10,
    
    [Parameter(Mandatory=$false)]
    [string]$UserFilter = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToCSV,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = "C:\Reports\TopEntraGroupUsers.csv"
)

# Check for required modules
function Check-RequiredModules {
    $modules = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Authentication")
    $missingModules = @()
    
    foreach ($module in $modules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "The following required modules are missing:" -ForegroundColor Red
        $missingModules | ForEach-Object { Write-Host "- $_" -ForegroundColor Yellow }
        Write-Host "Please install them using: Install-Module -Name ModuleName -Force" -ForegroundColor Yellow
        return $false
    }
    
    return $true
}

# Connect to Microsoft Graph
function Connect-MgGraphForGroups {
    $requiredScopes = @(
        "User.Read.All",
        "Group.Read.All",
        "Directory.Read.All"
    )
    
    try {
        # Check if already connected with correct scopes
        $currentConnection = Get-MgContext
        $hasRequiredScopes = $true
        
        if ($currentConnection) {
            foreach ($scope in $requiredScopes) {
                if ($currentConnection.Scopes -notcontains $scope) {
                    $hasRequiredScopes = $false
                    break
                }
            }
        }
        
        # Connect if not connected or missing scopes
        if (!$currentConnection -or !$hasRequiredScopes) {
            Write-Host "Connecting to Microsoft Graph API..." -ForegroundColor Green
            Connect-MgGraph -Scopes $requiredScopes
        }
        else {
            Write-Host "Already connected to Microsoft Graph with required permissions." -ForegroundColor Green
        }
        
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# Get users with group counts
function Get-UserGroupCounts {
    param(
        [string]$UserFilter
    )
    
    Write-Host "Retrieving users from Entra ID..." -ForegroundColor Cyan
    
    # Build filter if provided
    $filter = if ($UserFilter) { $UserFilter } else { "UserType eq 'Member'" }
    
    try {
        # Get all users (paginated to handle large directories)
        $allUsers = @()
        $users = Get-MgUser -Filter $filter -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled
        $allUsers += $users
        
        Write-Host "Found $($allUsers.Count) users." -ForegroundColor White
        
        # Process each user to get their group memberships
        $userGroups = @()
        $count = 0
        $total = $allUsers.Count
        
        foreach ($user in $allUsers) {
            $count++
            $percent = [math]::Round(($count / $total) * 100)
            Write-Progress -Activity "Getting group memberships" -Status "Processing user $count of $total" -PercentComplete $percent
            
            try {
                # Get direct group memberships
                $groups = Get-MgUserMemberOf -UserId $user.Id
                
                # Filter to security groups
                $securityGroups = $groups | Where-Object { 
                    $_.AdditionalProperties.securityEnabled -eq $true -or 
                    ($_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group" -and 
                     $_.AdditionalProperties.groupTypes -notcontains "DynamicMembership")
                }
                
                $userGroups += [PSCustomObject]@{
                    UserId = $user.Id
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    Enabled = $user.AccountEnabled
                    GroupCount = $securityGroups.Count
                    Groups = $securityGroups
                }
            }
            catch {
                Write-Host "Error getting groups for user $($user.UserPrincipalName): $_" -ForegroundColor Red
            }
        }
        
        Write-Progress -Activity "Getting group memberships" -Completed
        
        return $userGroups
    }
    catch {
        Write-Host "Error retrieving users: $_" -ForegroundColor Red
        return @()
    }
}

# Main function
function Find-UsersWithMostEntraGroups {
    # Check if Reports directory exists, create if not
    if (!(Test-Path "C:\Reports")) {
        try {
            New-Item -Path "C:\Reports" -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: C:\Reports" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating directory C:\Reports: $_" -ForegroundColor Red
            Write-Host "Please make sure you have permission to create this directory." -ForegroundColor Yellow
            return
        }
    }
    
    # Check modules
    if (!(Check-RequiredModules)) {
        return
    }
    
    # Connect to Graph
    if (!(Connect-MgGraphForGroups)) {
        return
    }
    
    # Get users with group counts
    $userGroups = Get-UserGroupCounts -UserFilter $UserFilter
    
    # Sort and get top users
    $topUsers = $userGroups | Sort-Object -Property GroupCount -Descending | Select-Object -First $TopCount
    
    # Display results
    Write-Host "`n=== TOP $TopCount USERS WITH MOST SECURITY GROUPS ===`n" -ForegroundColor Green
    
    $position = 1
    foreach ($user in $topUsers) {
        $enabledStatus = if ($user.Enabled) { "Enabled" } else { "Disabled" }
        Write-Host "$position. $($user.DisplayName) ($($user.UserPrincipalName)) - $enabledStatus" -ForegroundColor Cyan
        Write-Host "   Group Count: $($user.GroupCount)" -ForegroundColor White
        $position++
    }
    
    # Export to CSV if requested
    if ($ExportToCSV) {
        $exportData = $topUsers | Select-Object DisplayName, UserPrincipalName, Enabled, GroupCount, @{
            Name = 'GroupsList'; 
            Expression = { 
                ($_.Groups | ForEach-Object { 
                    if ($_.AdditionalProperties.displayName) {
                        $_.AdditionalProperties.displayName
                    } else {
                        "Group ID: $($_.Id)"
                    }
                }) -join '; '
            }
        }
        
        try {
            $exportData | Export-Csv -Path $ExportPath -NoTypeInformation
            Write-Host "`nExported results to: $ExportPath" -ForegroundColor Green
        }
        catch {
            Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
        }
    }
    
    # Ask if user wants to see groups for a specific user
    $viewGroups = Read-Host "`nWould you like to see all groups for a specific user from this list? (Y/N)"
    
    if ($viewGroups.ToUpper() -eq "Y") {
        $selectedUser = Read-Host "Enter the user principal name (email) to view groups for"
        
        $userDetail = $userGroups | Where-Object { $_.UserPrincipalName -eq $selectedUser }
        
        if ($userDetail) {
            Write-Host "`nSecurity groups for $selectedUser ($($userDetail.DisplayName)):" -ForegroundColor Green
            
            $groupList = $userDetail.Groups | ForEach-Object {
                [PSCustomObject]@{
                    DisplayName = if ($_.AdditionalProperties.displayName) { $_.AdditionalProperties.displayName } else { "N/A" }
                    Description = if ($_.AdditionalProperties.description) { $_.AdditionalProperties.description } else { "N/A" }
                    Type = if ($_.AdditionalProperties.securityEnabled -eq $true) { "Security" } else { "Distribution" }
                    Source = if ($_.AdditionalProperties.onPremisesSyncEnabled -eq $true) { "On-premises" } else { "Cloud" }
                }
            }
            
            $groupList | Sort-Object -Property DisplayName | Format-Table -AutoSize
            
            # Option to export single user's groups
            $exportSingleUser = Read-Host "Export this user's groups to CSV? (Y/N)"
            if ($exportSingleUser.ToUpper() -eq "Y") {
                $singleUserExportPath = "C:\Reports\${selectedUser}_Groups.csv"
                
                try {
                    $groupList | Export-Csv -Path $singleUserExportPath -NoTypeInformation
                    Write-Host "Exported to: $singleUserExportPath" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error exporting to CSV: $_" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "User $selectedUser not found in the analyzed list." -ForegroundColor Red
        }
    }
    
    return $topUsers
}

# Run the script
Find-UsersWithMostEntraGroups

Write-Host "`nScript completed." -ForegroundColor Green