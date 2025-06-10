# Dynamics 365 Administrator Dependency Checker
# This script identifies dependencies and configurations created by a specific Dynamics 365 administrator
# Run this with appropriate admin permissions before removing user access (4/10/2025)

# Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$UserPrincipalName
)

# Function to check if modules are installed and install if needed
function Ensure-ModuleInstalled {
    param (
        [string]$ModuleName
    )
    
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing $ModuleName module..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
    }
    
    Import-Module $ModuleName -Force
}

# Output folder for reports
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputFolder = ".\DynamicsAdminAudit-$timestamp"
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null

# Ensure required modules are installed
Write-Host "Checking for required PowerShell modules..." -ForegroundColor Cyan
Ensure-ModuleInstalled -ModuleName "ExchangeOnlineManagement"
Ensure-ModuleInstalled -ModuleName "Microsoft.PowerApps.Administration.PowerShell"
Ensure-ModuleInstalled -ModuleName "Microsoft.PowerApps.PowerShell"
Ensure-ModuleInstalled -ModuleName "Microsoft.Xrm.Data.PowerShell"

# Step 1: Connect to services - Manual MFA authentication
try {
    Write-Host "Please connect to Exchange Online in the window that appears..." -ForegroundColor Yellow
    Connect-ExchangeOnline
    Write-Host "Exchange Online connection successful!" -ForegroundColor Green
    
    Write-Host "Please connect to Power Platform Admin in the window that appears..." -ForegroundColor Yellow
    Add-PowerAppsAccount
    Write-Host "Power Platform connection successful!" -ForegroundColor Green
    
    # Get the user details to work with
    $userDetails = Get-User -Identity $UserPrincipalName
    Write-Host "Analyzing dependencies for user: $($userDetails.DisplayName) ($UserPrincipalName)" -ForegroundColor Green
} 
catch {
    Write-Error "Failed to connect to required services: $_"
    exit
}

# Step 2: Check Dynamics 365 environments
Write-Host "Retrieving Dynamics 365 environments..." -ForegroundColor Cyan
$environments = Get-AdminPowerAppEnvironment
$userEnvironments = @()

foreach ($env in $environments) {
    # Get environment roles for the user
    $roles = Get-AdminPowerAppEnvironmentRoleAssignment -EnvironmentName $env.EnvironmentName | 
             Where-Object { $_.PrincipalObjectId -eq $userDetails.ExternalDirectoryObjectId }
    
    if ($roles) {
        $userEnvironments += [PSCustomObject]@{
            EnvironmentName = $env.DisplayName
            EnvironmentId = $env.EnvironmentName
            Roles = $roles.RoleName -join ', '
        }
    }
}

# Export environment information
$userEnvironments | Export-Csv -Path "$outputFolder\UserEnvironments.csv" -NoTypeInformation
Write-Host "Found $($userEnvironments.Count) environments where user has roles" -ForegroundColor Yellow

# Step 3: Check for Power Apps and Flows created by the user
Write-Host "Checking for PowerApps created by user..." -ForegroundColor Cyan
$userApps = @()
$userFlows = @()

foreach ($env in $userEnvironments) {
    # Get apps in this environment
    $apps = Get-AdminPowerApp -EnvironmentName $env.EnvironmentId
    $userCreatedApps = $apps | Where-Object { $_.Owner.Id -eq $userDetails.ExternalDirectoryObjectId }
    
    foreach ($app in $userCreatedApps) {
        $userApps += [PSCustomObject]@{
            AppName = $app.DisplayName
            AppId = $app.AppName
            Environment = $env.EnvironmentName
            LastModified = $app.LastModifiedTime
            Shared = ($app.UserSharedWith.Count -gt 0 -or $app.GroupSharedWith.Count -gt 0)
        }
    }
    
    # Get flows in this environment
    $flows = Get-AdminFlow -EnvironmentName $env.EnvironmentId
    $userCreatedFlows = $flows | Where-Object { $_.CreatedBy.userId -eq $userDetails.ExternalDirectoryObjectId }
    
    foreach ($flow in $userCreatedFlows) {
        $userFlows += [PSCustomObject]@{
            FlowName = $flow.DisplayName
            FlowId = $flow.FlowName
            Environment = $env.EnvironmentName
            Status = $flow.State
            LastModified = $flow.LastModifiedTime
        }
    }
}

# Export apps and flows information
$userApps | Export-Csv -Path "$outputFolder\UserCreatedApps.csv" -NoTypeInformation
$userFlows | Export-Csv -Path "$outputFolder\UserCreatedFlows.csv" -NoTypeInformation
Write-Host "Found $($userApps.Count) PowerApps and $($userFlows.Count) Flows created by user" -ForegroundColor Yellow

# Step 4: Connect to each Dynamics instance and check for customizations
Write-Host "Checking for Dynamics 365 customizations..." -ForegroundColor Cyan
$customizations = @()

foreach ($env in $userEnvironments) {
    try {
        # Connect to this Dynamics instance with MFA
        Write-Host "Please connect to Dynamics 365 environment: $($env.EnvironmentName) in the window that appears..." -ForegroundColor Yellow
        $conn = Connect-CrmOnline -InteractiveMode -Url "https://$($env.EnvironmentName).crm.dynamics.com"
        
        # Get customizations by this user
        $solutions = Get-CrmSolutions
        $userSolutions = $solutions | Where-Object { $_.modifiedby -eq $userDetails.DisplayName }
        
        foreach ($solution in $userSolutions) {
            $customizations += [PSCustomObject]@{
                SolutionName = $solution.friendlyname
                SolutionId = $solution.solutionid
                Environment = $env.EnvironmentName
                Version = $solution.version
                LastModified = $solution.modifiedon
            }
        }
        
        # Get processes/workflows modified by user
        $processes = Get-CrmProcesses
        $userProcesses = $processes | Where-Object { $_.modifiedby -eq $userDetails.DisplayName }
        
        foreach ($process in $userProcesses) {
            $customizations += [PSCustomObject]@{
                ItemType = "Process/Workflow"
                Name = $process.name
                Id = $process.id
                Environment = $env.EnvironmentName
                Status = $process.statecode
                LastModified = $process.modifiedon
            }
        }
    }
    catch {
        Write-Warning "Could not connect to Dynamics environment $($env.EnvironmentName): $_"
    }
}

# Export customizations
$customizations | Export-Csv -Path "$outputFolder\UserCustomizations.csv" -NoTypeInformation
Write-Host "Found $($customizations.Count) customizations by user" -ForegroundColor Yellow

# Step 5: Check for scheduled jobs or automation owned by the user
Write-Host "Checking for scheduled jobs and automation..." -ForegroundColor Cyan
$scheduledJobs = @()

foreach ($env in $userEnvironments) {
    try {
        # Get scheduled jobs in this environment (if we have connected to Dynamics)
        $jobs = Get-CrmRecords -EntityLogicalName "recurringappointmentmaster" -FilterAttribute "ownerid" -FilterOperator "eq" -FilterValue $userDetails.ExternalDirectoryObjectId
        
        foreach ($job in $jobs.CrmRecords) {
            $scheduledJobs += [PSCustomObject]@{
                JobName = $job.subject
                JobId = $job.recurringappointmentmasterid
                Environment = $env.EnvironmentName
                Frequency = $job.recurrencepatterntype
                StartTime = $job.starttime
                EndTime = $job.endtime
            }
        }
    }
    catch {
        Write-Warning "Could not check for scheduled jobs in $($env.EnvironmentName): $_"
    }
}

# Export scheduled jobs
$scheduledJobs | Export-Csv -Path "$outputFolder\UserScheduledJobs.csv" -NoTypeInformation
Write-Host "Found $($scheduledJobs.Count) scheduled jobs owned by user" -ForegroundColor Yellow

# Step 6: Check for security roles and teams created/managed by the user
Write-Host "Checking for security roles and teams..." -ForegroundColor Cyan
$securityItems = @()

foreach ($env in $userEnvironments) {
    try {
        # Get security roles created/modified by this user
        $roles = Get-CrmRecords -EntityLogicalName "role" -FilterAttribute "modifiedby" -FilterOperator "eq" -FilterValue $userDetails.DisplayName
        
        foreach ($role in $roles.CrmRecords) {
            $securityItems += [PSCustomObject]@{
                ItemType = "Security Role"
                Name = $role.name
                Id = $role.roleid
                Environment = $env.EnvironmentName
                LastModified = $role.modifiedon
            }
        }
        
        # Get teams created/modified by this user
        $teams = Get-CrmRecords -EntityLogicalName "team" -FilterAttribute "modifiedby" -FilterOperator "eq" -FilterValue $userDetails.DisplayName
        
        foreach ($team in $teams.CrmRecords) {
            $securityItems += [PSCustomObject]@{
                ItemType = "Team"
                Name = $team.name
                Id = $team.teamid
                Environment = $env.EnvironmentName
                LastModified = $team.modifiedon
            }
        }
    }
    catch {
        Write-Warning "Could not check for security items in $($env.EnvironmentName): $_"
    }
}

# Export security items
$securityItems | Export-Csv -Path "$outputFolder\UserSecurityItems.csv" -NoTypeInformation
Write-Host "Found $($securityItems.Count) security roles and teams modified by user" -ForegroundColor Yellow

# Step 7: Generate summary report
$summaryFile = "$outputFolder\Summary.txt"

$summary = @"
======================================================================
DYNAMICS 365 ADMIN DEPENDENCY REPORT
======================================================================
User: $($userDetails.DisplayName) ($UserPrincipalName)
Date Generated: $(Get-Date)
Access Removal Date: 4/10/2025

SUMMARY OF FINDINGS:
- Environments with Access: $($userEnvironments.Count)
- PowerApps Created: $($userApps.Count)
- Flows Created: $($userFlows.Count)
- Customizations: $($customizations.Count)
- Scheduled Jobs: $($scheduledJobs.Count)
- Security Items: $($securityItems.Count)

RISK ASSESSMENT:
"@

# Simple risk assessment logic
$highRiskItems = @()

# Flag high-risk items
if ($userFlows.Count -gt 0) {
    $activeFlows = $userFlows | Where-Object { $_.Status -eq "Started" }
    if ($activeFlows.Count -gt 0) {
        $highRiskItems += "- CRITICAL: $($activeFlows.Count) active Flows will stop working when user loses access"
    }
}

if ($scheduledJobs.Count -gt 0) {
    $highRiskItems += "- CRITICAL: $($scheduledJobs.Count) scheduled jobs may fail when user loses access"
}

# Add risk assessment to summary
if ($highRiskItems.Count -gt 0) {
    $summary += "`n`nHIGH-RISK ITEMS REQUIRING IMMEDIATE ATTENTION:`n"
    $summary += $highRiskItems -join "`n"
} else {
    $summary += "`n`nNo high-risk items identified. However, please review all reports for detailed information."
}

$summary += @"

======================================================================
RECOMMENDATION:
Review each CSV file in the $outputFolder directory for detailed information.
Consider reassigning ownership of critical items before removing user access.
======================================================================
"@

# Export summary
$summary | Out-File -FilePath $summaryFile
Write-Host "Summary report generated at: $summaryFile" -ForegroundColor Green

# Disconnect from services
Write-Host "Disconnecting from services..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

# Display final report location
Write-Host "Audit complete! All reports have been saved to the $outputFolder folder." -ForegroundColor Green
Write-Host "IMPORTANT: Review these reports before removing user access on 4/10/2025" -ForegroundColor Red
Write-Host @"

USAGE INSTRUCTIONS:
1. Run this script with:
   .\Dynamics-Admin-Dependency-Checker.ps1 -UserPrincipalName "admin@yourcompany.com"

2. Complete the MFA authentication in each window that appears

3. Review all CSV files in the output directory to identify items that need reassignment
"@ -ForegroundColor Yellow