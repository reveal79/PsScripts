# M365 Inactive Users Audit Tool - Chart Fix version
# This script helps audit and validate inactive users in Microsoft 365 environment
<#
.SYNOPSIS
    Audits inactive users in Microsoft 365 and generates comprehensive reports.

.DESCRIPTION
    This script connects to Microsoft 365, retrieves user activity data,
    identifies inactive users, and generates detailed reports in CSV, Excel,
    and HTML formats. It's designed to help administrators identify license
    cost-saving opportunities and improve security by managing inactive accounts.

.PARAMETER Period
    Specifies the time period to analyze. Valid values are D7, D30, D90, and D180.
    Default is D30 (30 days).

.PARAMETER InactiveThreshold
    Number of days without activity to consider a user inactive. Default is 30 days.

.PARAMETER OutputPath
    Directory where reports will be saved. Default is a folder named with the current date.

.PARAMETER GenerateHTML
    Switch to generate HTML report. Default is true.

.PARAMETER GenerateExcel
    Switch to generate Excel report. Default is true.

.PARAMETER GenerateCSV
    Switch to generate CSV report. Default is true.

.PARAMETER Help
    Displays this help information.

.EXAMPLE
    .\M365InactiveUsersChartFix.ps1
    Runs with default parameters (30-day period, 30-day threshold)

.EXAMPLE
    .\M365InactiveUsersChartFix.ps1 -Period D90 -InactiveThreshold 45
    Checks for users inactive for 45 days or more, using 90 days of activity data

.EXAMPLE
    .\M365InactiveUsersChartFix.ps1 -OutputPath "C:\Reports\M365Audit"
    Saves all reports to the specified directory

.NOTES
    Author: Your Organization
    Last Updated: May 15, 2025
    Requires: Microsoft.Graph PowerShell module, ImportExcel, PSWriteHTML
#>

param(
    [Parameter(HelpMessage="Specify the time period to analyze (D7, D30, D90, D180)")]
    [ValidateSet("D7", "D30", "D90", "D180")]
    [string]$Period = "D30",
    
    [Parameter(HelpMessage="Days without activity to consider a user inactive")]
    [int]$InactiveThreshold = 30,
    
    [Parameter(HelpMessage="Directory where reports will be saved")]
    [string]$OutputPath = ".\M365Audit_$(Get-Date -Format 'yyyyMMdd')",
    
    [Parameter(HelpMessage="Generate HTML report")]
    [switch]$GenerateHTML = $true,
    
    [Parameter(HelpMessage="Generate Excel report")]
    [switch]$GenerateExcel = $true,
    
    [Parameter(HelpMessage="Generate CSV report")]
    [switch]$GenerateCSV = $true,
    
    [Parameter(HelpMessage="Display help information")]
    [switch]$Help
)

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
    Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    try {
        Write-Host "Connecting to Microsoft Graph interactively..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes "Reports.Read.All", "Directory.Read.All", "User.Read.All"
        
        $context = Get-MgContext
        Write-Host "Successfully connected as: $($context.Account)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# Function to get inactive users report as JSON
function Get-InactiveUsersReport {
    param (
        [string]$Period,
        [string]$OutputPath
    )
    
    try {
        Write-Host "Fetching Office 365 active user details for period '$Period' as JSON..." -ForegroundColor Cyan
        
        # Create a temporary file path for the JSON
        $tempFilePath = Join-Path -Path $OutputPath -ChildPath "TempUserActivityReport.json"
        
        # Use Invoke-MgGraphRequest to get the data in JSON format
        $reportUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='$Period')"
        $headers = @{
            "Accept" = "application/json"
        }
        
        # First try to get JSON directly
        try {
            $jsonResponse = Invoke-MgGraphRequest -Uri $reportUri -Method GET -Headers $headers -OutputType PSObject
            Write-Host "Successfully retrieved JSON data directly" -ForegroundColor Green
            
            # Save the JSON for reference
            $jsonResponse | ConvertTo-Json -Depth 10 | Out-File -FilePath $tempFilePath -Encoding UTF8
            
            return $jsonResponse
        }
        catch {
            Write-Host "JSON direct request failed: $_" -ForegroundColor Yellow
            Write-Host "Falling back to CSV download and conversion..." -ForegroundColor Yellow
            
            # If direct JSON retrieval fails, fall back to CSV download and convert to JSON
            $csvFilePath = Join-Path -Path $OutputPath -ChildPath "TempUserActivityReport.csv"
            Invoke-MgGraphRequest -Uri $reportUri -Method GET -OutputFilePath $csvFilePath
            
            Write-Host "Downloaded CSV report to $csvFilePath" -ForegroundColor Green
            
            # Check if file exists and has content
            if (Test-Path $csvFilePath) {
                # Import the CSV and convert to JSON-like objects
                $csvData = Import-Csv -Path $csvFilePath
                
                # Save the converted data as JSON
                $csvData | ConvertTo-Json -Depth 10 | Out-File -FilePath $tempFilePath -Encoding UTF8
                
                Write-Host "Converted CSV to JSON format at $tempFilePath" -ForegroundColor Green
                Write-Host "Successfully processed data for $($csvData.Count) users" -ForegroundColor Green
                
                return $csvData
            }
            else {
                throw "Failed to download report to $csvFilePath"
            }
        }
    }
    catch {
        Write-Host "Error fetching user activity report: $_" -ForegroundColor Red
        return $null
    }
}

# Function to process user data and identify inactive users
function Get-EnhancedUserDetails {
    param (
        [array]$Users
    )
    
    try {
        Write-Host "Enhancing user data and identifying inactive users..." -ForegroundColor Cyan
        
        $enhancedUsers = @()
        $total = $Users.Count
        $current = 0
        
        # First ensure we have the right property names by checking the first user
        if ($Users.Count -gt 0) {
            $firstUser = $Users[0]
            $propNames = $firstUser.PSObject.Properties.Name
            Write-Host "Found properties: $($propNames -join ', ')" -ForegroundColor Yellow
        }
        
        foreach ($user in $Users) {
            $current++
            
            # Only show progress every 50 users to reduce console output
            if ($current % 50 -eq 0) {
                Write-Progress -Activity "Processing users" -Status "Processing $current of $total" -PercentComplete (($current / $total) * 100)
                Write-Host "Processed $current of $total users..." -ForegroundColor Gray
            }
            
            # Extract property values using dynamic approach to handle different property names
            $upnProp = $propNames | Where-Object { $_ -like "*Principal Name*" } | Select-Object -First 1
            $userPrincipalName = if ($upnProp) { $user.$upnProp } else { "" }
            
            $displayNameProp = $propNames | Where-Object { $_ -like "*Display Name*" } | Select-Object -First 1
            $displayName = if ($displayNameProp) { $user.$displayNameProp } else { "" }
            
            $isDeletedProp = $propNames | Where-Object { $_ -like "*Is Deleted*" } | Select-Object -First 1
            $isDeleted = if ($isDeletedProp) { $user.$isDeletedProp -eq $true -or $user.$isDeletedProp -eq "True" } else { $false }
            
            $deletedDateProp = $propNames | Where-Object { $_ -like "*Deleted Date*" } | Select-Object -First 1
            $deletedDate = if ($deletedDateProp) { $user.$deletedDateProp } else { "" }
            
            $assignedProductsProp = $propNames | Where-Object { $_ -like "*Assigned Products*" } | Select-Object -First 1
            $assignedProducts = if ($assignedProductsProp) { $user.$assignedProductsProp } else { "" }
            
            $reportRefreshDateProp = $propNames | Where-Object { $_ -like "*Report Refresh Date*" } | Select-Object -First 1
            $reportRefreshDate = if ($reportRefreshDateProp) { $user.$reportRefreshDateProp } else { "" }
            
            # Get activity dates for different services
            $activityDates = @()
            $serviceLastActivity = @{}
            $serviceHasLicense = @{}
            
            # Services to check
            $services = @("Exchange", "OneDrive", "SharePoint", "Teams", "Yammer", "Skype For Business")
            
            foreach ($service in $services) {
                # Check for activity date
                $activityProp = $propNames | Where-Object { $_ -like "*$service*Activity Date*" } | Select-Object -First 1
                $activityDate = if ($activityProp -and -not [string]::IsNullOrEmpty($user.$activityProp)) { $user.$activityProp } else { $null }
                
                $serviceLastActivity[$service] = $activityDate
                
                if ($activityDate) {
                    try {
                        $parsedDate = [DateTime]::Parse($activityDate)
                        $activityDates += $parsedDate
                    } catch {
                        # Skip if date parsing fails
                    }
                }
                
                # Check license status
                $licenseProp = $propNames | Where-Object { $_ -like "*Has $service License*" } | Select-Object -First 1
                $hasLicense = if ($licenseProp) { $user.$licenseProp -eq $true -or $user.$licenseProp -eq "True" } else { $false }
                
                $serviceHasLicense[$service] = $hasLicense
            }
            
            # Calculate most recent activity and days since
            $lastActivityDate = if ($activityDates.Count -gt 0) { ($activityDates | Sort-Object -Descending)[0] } else { $null }
            $daysSinceActivity = if ($lastActivityDate) { (Get-Date) - $lastActivityDate } else { $null }
            $daysSinceActivityCount = if ($daysSinceActivity) { $daysSinceActivity.Days } else { $null }
            
            # Determine if user is inactive based on threshold
            $isInactive = if ($daysSinceActivityCount -and $daysSinceActivityCount -gt $InactiveThreshold) { $true } else { $false }
            
            # Create enhanced user object
            $enhancedUser = [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                DisplayName = $displayName
                IsActive = $null -ne $lastActivityDate
                LastActivityDate = $lastActivityDate
                DaysSinceActivity = $daysSinceActivityCount
                IsInactive = $isInactive
                IsDeleted = $isDeleted
                DeletedDate = $deletedDate
                AssignedProducts = $assignedProducts
                ReportRefreshDate = $reportRefreshDate
                HasExchangeLicense = $serviceHasLicense["Exchange"]
                HasOneDriveLicense = $serviceHasLicense["OneDrive"]
                HasSharePointLicense = $serviceHasLicense["SharePoint"]
                HasTeamsLicense = $serviceHasLicense["Teams"]
                HasYammerLicense = $serviceHasLicense["Yammer"]
                HasSkypeLicense = $serviceHasLicense["Skype For Business"]
                ExchangeLastActivityDate = $serviceLastActivity["Exchange"]
                OneDriveLastActivityDate = $serviceLastActivity["OneDrive"]
                SharePointLastActivityDate = $serviceLastActivity["SharePoint"]
                TeamsLastActivityDate = $serviceLastActivity["Teams"]
                YammerLastActivityDate = $serviceLastActivity["Yammer"]
                SkypeLastActivityDate = $serviceLastActivity["Skype For Business"]
                LicensedServicesCount = ($serviceHasLicense.Values | Where-Object { $_ -eq $true }).Count
                ActiveServicesCount = ($serviceLastActivity.Values | Where-Object { $_ -ne $null }).Count
            }
            
            $enhancedUsers += $enhancedUser
        }
        
        Write-Progress -Activity "Processing users" -Completed
        Write-Host "Successfully processed data for $($enhancedUsers.Count) users" -ForegroundColor Green
        
        return $enhancedUsers
    }
    catch {
        Write-Host "Error processing user data: $_" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
        return $null
    }
}

# Function to generate CSV report
function Export-ToCSV {
    param (
        [array]$Users,
        [string]$OutputPath
    )
    
    try {
        $csvPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.csv"
        Write-Host "Exporting CSV report to $csvPath..." -ForegroundColor Cyan
        $Users | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "CSV report exported successfully" -ForegroundColor Green
        return $csvPath
    }
    catch {
        Write-Host "Error exporting CSV report: $_" -ForegroundColor Red
        return $null
    }
}

# Function to generate Excel report
function Export-ToExcel {
    param (
        [array]$Users,
        [string]$OutputPath
    )
    
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
        Install-Module ImportExcel -Force -AllowClobber
    }
    
    try {
        $excelPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.xlsx"
        Write-Host "Exporting Excel report to $excelPath..." -ForegroundColor Cyan
        
        $excelParams = @{
            Path = $excelPath
            AutoSize = $true
            TableName = "InactiveUsers"
            WorksheetName = "Inactive Users"
            TableStyle = "Medium6"
        }
        
        $Users | Export-Excel @excelParams
        
        # Create summary worksheet
        $summaryData = [PSCustomObject]@{
            'Total Users' = $Users.Count
            'Inactive Users' = ($Users | Where-Object { $_.IsInactive -eq $true }).Count
            'Deleted Users' = ($Users | Where-Object { $_.IsDeleted -eq $true }).Count
            'Inactive & Licensed' = ($Users | Where-Object { $_.IsInactive -eq $true -and $_.LicensedServicesCount -gt 0 }).Count
            'Days Threshold' = $InactiveThreshold
            'Report Date' = Get-Date
            'Period' = $Period
        }
        
        $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -TableName "Summary" -TableStyle "Medium6"
        
        # Create worksheets filtered by active status
        $Users | Where-Object { $_.IsInactive -eq $true } | 
            Export-Excel -Path $excelPath -WorksheetName "Inactive Users" -AutoSize -TableName "InactiveUsers" -TableStyle "Medium6"
        
        $Users | Where-Object { $_.IsInactive -eq $true -and $_.LicensedServicesCount -gt 0 } | 
            Export-Excel -Path $excelPath -WorksheetName "Inactive Licensed" -AutoSize -TableName "InactiveLicensed" -TableStyle "Medium6"
        
        Write-Host "Excel report exported successfully" -ForegroundColor Green
        return $excelPath
    }
    catch {
        Write-Host "Error exporting Excel report: $_" -ForegroundColor Red
        return $null
    }
}

# Function to generate HTML report with fixed chart generation
function Export-ToHTML {
    param (
        [array]$Users,
        [string]$OutputPath
    )
    
    if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
        Write-Host "PSWriteHTML module not found. Installing..." -ForegroundColor Yellow
        Install-Module PSWriteHTML -Force -AllowClobber
    }
    
    try {
        $htmlPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"
        Write-Host "Generating HTML report to $htmlPath..." -ForegroundColor Cyan
        
        Import-Module PSWriteHTML
        
        # Prepare data for charts - pre-calculate all values
        # Get inactive and active user counts
        [int]$inactiveCount = ($Users | Where-Object { $_.IsInactive -eq $true }).Count
        [int]$activeCount = ($Users | Where-Object { $_.IsInactive -eq $false }).Count
        
        # Count inactive licensed users
        [int]$inactiveLicensedCount = ($Users | Where-Object { 
            $_.IsInactive -eq $true -and $_.LicensedServicesCount -gt 0 
        }).Count
        
        # Calculate percentage
        [decimal]$inactivePercentage = 0
        if ($Users.Count -gt 0) {
            $inactivePercentage = [math]::Round(($inactiveCount / $Users.Count) * 100, 2)
        }
        
        # Count deleted users
        [int]$deletedCount = ($Users | Where-Object { $_.IsDeleted -eq $true }).Count
        
        # Calculate potential savings
        [int]$potentialSavings = $inactiveLicensedCount * 20
        [int]$annualSavings = $potentialSavings * 12
        
        # For charts, create data arrays
        # For pie chart: active vs inactive
        $pieChartData = @([PSCustomObject]@{
            Name = "Active"
            Count = $activeCount
        },
        [PSCustomObject]@{
            Name = "Inactive"
            Count = $inactiveCount
        })
        
        # Service data
        $exchangeActive = ($Users | Where-Object { $_.HasExchangeLicense -eq $true -and -not [string]::IsNullOrEmpty($_.ExchangeLastActivityDate) }).Count
        $exchangeInactive = ($Users | Where-Object { $_.HasExchangeLicense -eq $true -and [string]::IsNullOrEmpty($_.ExchangeLastActivityDate) }).Count
        
        $oneDriveActive = ($Users | Where-Object { $_.HasOneDriveLicense -eq $true -and -not [string]::IsNullOrEmpty($_.OneDriveLastActivityDate) }).Count
        $oneDriveInactive = ($Users | Where-Object { $_.HasOneDriveLicense -eq $true -and [string]::IsNullOrEmpty($_.OneDriveLastActivityDate) }).Count
        
        $sharepointActive = ($Users | Where-Object { $_.HasSharePointLicense -eq $true -and -not [string]::IsNullOrEmpty($_.SharePointLastActivityDate) }).Count
        $sharepointInactive = ($Users | Where-Object { $_.HasSharePointLicense -eq $true -and [string]::IsNullOrEmpty($_.SharePointLastActivityDate) }).Count
        
        $teamsActive = ($Users | Where-Object { $_.HasTeamsLicense -eq $true -and -not [string]::IsNullOrEmpty($_.TeamsLastActivityDate) }).Count
        $teamsInactive = ($Users | Where-Object { $_.HasTeamsLicense -eq $true -and [string]::IsNullOrEmpty($_.TeamsLastActivityDate) }).Count
        
        # For service chart: create data structure for stacked bar chart
        $serviceChartData = @(
            [PSCustomObject]@{ Name = "Exchange"; Active = $exchangeActive; Inactive = $exchangeInactive },
            [PSCustomObject]@{ Name = "OneDrive"; Active = $oneDriveActive; Inactive = $oneDriveInactive },
            [PSCustomObject]@{ Name = "SharePoint"; Active = $sharepointActive; Inactive = $sharepointInactive },
            [PSCustomObject]@{ Name = "Teams"; Active = $teamsActive; Inactive = $teamsInactive }
        )
        
        # Create user collections
        $inactiveUsers = $Users | Where-Object { $_.IsInactive -eq $true } | Sort-Object -Property DaysSinceActivity -Descending
        $inactiveLicensedUsers = $Users | Where-Object { $_.IsInactive -eq $true -and $_.LicensedServicesCount -gt 0 } | Sort-Object -Property DaysSinceActivity -Descending
        
        # Create summary data
        $summaryTable = @(
            [PSCustomObject]@{ 'Metric' = 'Total Users'; 'Value' = $Users.Count },
            [PSCustomObject]@{ 'Metric' = 'Active Users'; 'Value' = $activeCount },
            [PSCustomObject]@{ 'Metric' = 'Inactive Users'; 'Value' = $inactiveCount },
            [PSCustomObject]@{ 'Metric' = 'Inactive Percentage'; 'Value' = "$inactivePercentage%" },
            [PSCustomObject]@{ 'Metric' = 'Deleted Users'; 'Value' = $deletedCount },
            [PSCustomObject]@{ 'Metric' = 'Inactive & Licensed'; 'Value' = $inactiveLicensedCount },
            [PSCustomObject]@{ 'Metric' = 'Potential Annual Savings'; 'Value' = "$($annualSavings)" }
        )
        
        # Create HTML without using the problematic chart generation
        New-HTML -TitleText "Microsoft 365 Inactive Users Audit Report" -FilePath $htmlPath {
            New-HTMLHeader {
                New-HTMLText -Text "Microsoft 365 Inactive Users Audit" -Color Black -Alignment center -FontSize 30
                New-HTMLText -Text "Generated on $(Get-Date) | Period: $Period | Inactive Threshold: $InactiveThreshold days" -Color Gray -Alignment center
            }
            
            # Summary section with plain table and no charts
            New-HTMLSection -HeaderText "Executive Summary" -CanCollapse {
                New-HTMLPanel {
                    New-HTMLText -Text "This report provides an analysis of inactive users in your Microsoft 365 environment." -FontSize 14
                    New-HTMLTable -DataTable $summaryTable
                    
                    # Display manually crafted visual representation using divs instead of charts
                    New-HTMLPanel {
                        New-HTMLText -Text "<div style='margin-top:20px'><strong>User Status Distribution:</strong> $activeCount Active, $inactiveCount Inactive</div>" -Color Black
                        
                        # Manual progress bar for active vs inactive
                        $activePercent = [math]::Round(($activeCount / $Users.Count) * 100, 0)
                        $inactivePercent = 100 - $activePercent
                        
                        $htmlBar = @"
                        <div style="margin-top:10px;width:100%;background-color:#f3f3f3;border-radius:4px;overflow:hidden">
                            <div style="width:$activePercent%;height:24px;background-color:#91D4F5;float:left;text-align:center;color:#333;font-weight:bold">Active $activePercent%</div>
                            <div style="width:$inactivePercent%;height:24px;background-color:#0078D4;float:left;text-align:center;color:white;font-weight:bold">Inactive $inactivePercent%</div>
                        </div>
                        <div style="clear:both"></div>
"@
                        New-HTMLText -Text $htmlBar
                    }
                }
            }
            
            # Service Usage section with tabular data
            New-HTMLSection -HeaderText "Service Usage" -CanCollapse {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $serviceChartData
                    
                    # Add manual bar chart visualization for services
                    $serviceHtml = "<div style='margin-top:20px'><strong>Service Usage Visualization:</strong></div>"
                    
                    foreach ($service in $serviceChartData) {
                        $totalUsers = $service.Active + $service.Inactive
                        if ($totalUsers -gt 0) {
                            $activePercent = [math]::Round(($service.Active / $totalUsers) * 100, 0)
                            $inactivePercent = 100 - $activePercent
                            
                            $serviceHtml += @"
                            <div style="margin-top:15px;"><strong>$($service.Name)</strong></div>
                            <div style="width:100%;background-color:#f3f3f3;border-radius:4px;overflow:hidden">
                                <div style="width:$activePercent%;height:20px;background-color:#91D4F5;float:left;text-align:center;color:#333;font-weight:bold">$($service.Active)</div>
                                <div style="width:$inactivePercent%;height:20px;background-color:#0078D4;float:left;text-align:center;color:white;font-weight:bold">$($service.Inactive)</div>
                            </div>
                            <div style="clear:both"></div>
"@
                        }
                    }
                    
                    New-HTMLText -Text $serviceHtml
                }
            }
            
            # Inactive users table
            New-HTMLSection -HeaderText "Inactive Users ($inactiveCount)" -CanCollapse {
                New-HTMLTable -DataTable $inactiveUsers
            }
            
            # Inactive licensed users
            New-HTMLSection -HeaderText "Inactive Licensed Users ($inactiveLicensedCount)" -CanCollapse {
                New-HTMLTable -DataTable $inactiveLicensedUsers
            }
            
            # Recommendations
            New-HTMLSection -HeaderText "Recommendations" -CanCollapse {
                New-HTMLList {
                    New-HTMLListItem -Text "Review and consider disabling accounts inactive for more than $InactiveThreshold days"
                    New-HTMLListItem -Text "Remove licenses from inactive accounts to reduce costs (potential annual savings: $($annualSavings))"
                    New-HTMLListItem -Text "Implement a regular review process for inactive accounts"
                    New-HTMLListItem -Text "Consider implementing an automated deprovisioning workflow"
                }
            }
        }
        
        Write-Host "HTML report generated successfully" -ForegroundColor Green
        return $htmlPath
    }
    catch {
        Write-Host "Error generating HTML report: $_" -ForegroundColor Red
        Write-Host "Falling back to basic HTML..." -ForegroundColor Yellow
        
        # Fallback to simple HTML
        try {
            $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 Inactive Users Audit Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; text-align: center; }
        h2 { color: #333; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f2f2f2; }
        tr:hover {background-color: #f5f5f5;}
        .summary { background-color: #f9f9f9; padding: 15px; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>Microsoft 365 Inactive Users Audit Report</h1>
    <p>Generated on $(Get-Date) | Period: $Period | Inactive Threshold: $InactiveThreshold days</p>
    
    <div class="summary">
        <h2>Executive Summary</h2>
        <p>This report provides an analysis of inactive users in your Microsoft 365 environment.</p>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Total Users</td><td>$($Users.Count)</td></tr>
            <tr><td>Active Users</td><td>$activeCount</td></tr>
            <tr><td>Inactive Users</td><td>$inactiveCount</td></tr>
            <tr><td>Inactive Percentage</td><td>$inactivePercentage%</td></tr>
            <tr><td>Inactive & Licensed</td><td>$inactiveLicensedCount</td></tr>
            <tr><td>Potential Annual Savings</td><td>$($annualSavings)</td></tr>
        </table>
    </div>
    
    <h2>Recommendations</h2>
    <ul>
        <li>Review and consider disabling accounts inactive for more than $InactiveThreshold days</li>
        <li>Remove licenses from inactive accounts to reduce costs (potential annual savings: $$annualSavings)</li>
        <li>Implement a regular review process for inactive accounts</li>
        <li>Consider implementing an automated deprovisioning workflow</li>
    </ul>
</body>
</html>
"@
            
            $htmlPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"
            $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
            Write-Host "Basic HTML report generated successfully" -ForegroundColor Green
            return $htmlPath
        }
        catch {
            Write-Host "Error generating fallback HTML report: $_" -ForegroundColor Red
            return $null
        }
    }
}

# Main execution
try {
    # Check if help was requested
    if ($Help) {
        Get-Help $MyInvocation.MyCommand.Path -Detailed
        return
    }
    
    # Display script banner and information
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "  Microsoft 365 Inactive Users Audit Tool" -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "Period: $Period | Inactive Threshold: $InactiveThreshold days" -ForegroundColor Cyan
    Write-Host "Reports will be saved to: $OutputPath" -ForegroundColor Cyan
    Write-Host "CSV: $GenerateCSV | Excel: $GenerateExcel | HTML: $GenerateHTML" -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host ""

    # Connect to Microsoft Graph
    if (-not (Connect-ToMicrosoftGraph)) {
        Write-Host "Failed to connect to Microsoft Graph. Exiting script." -ForegroundColor Red
        return
    }

    # Get inactive users report
    $usersReport = Get-InactiveUsersReport -Period $Period -OutputPath $OutputPath
    if (-not $usersReport) {
        Write-Host "Failed to get user activity report. Exiting script." -ForegroundColor Red
        return
    }

    # Enhance user details
    $enhancedUsers = Get-EnhancedUserDetails -Users $usersReport
    if (-not $enhancedUsers) {
        Write-Host "Failed to process user data. Exiting script." -ForegroundColor Red
        return
    }

    # Export reports
    $reports = @()

    if ($GenerateCSV) {
        $csvPath = Export-ToCSV -Users $enhancedUsers -OutputPath $OutputPath
        if ($csvPath) {
            $reports += "CSV Report: $csvPath"
        }
    }

    if ($GenerateExcel) {
        $excelPath = Export-ToExcel -Users $enhancedUsers -OutputPath $OutputPath
        if ($excelPath) {
            $reports += "Excel Report: $excelPath"
        }
    }

    if ($GenerateHTML) {
        $htmlPath = Export-ToHTML -Users $enhancedUsers -OutputPath $OutputPath
        if ($htmlPath) {
            $reports += "HTML Report: $htmlPath"
        }
    }

    # Summary
    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "  Microsoft 365 Inactive Users Audit Completed" -ForegroundColor Green
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "Total users processed: $($enhancedUsers.Count)" -ForegroundColor Green
    
    $inactiveCount = ($enhancedUsers | Where-Object { $_.IsInactive -eq $true }).Count
    $inactivePercent = if ($enhancedUsers.Count -gt 0) { [math]::Round(($inactiveCount / $enhancedUsers.Count) * 100, 1) } else { 0 }
    Write-Host "Inactive users: $inactiveCount ($inactivePercent%)" -ForegroundColor Green
    
    $inactiveLicensedCount = ($enhancedUsers | Where-Object { $_.IsInactive -eq $true -and $_.LicensedServicesCount -gt 0 }).Count
    if ($inactiveLicensedCount -gt 0) {
        $potentialSavings = $inactiveLicensedCount * 20 * 12  # $20 per license per month × 12 months
        Write-Host "Inactive licensed users: $inactiveLicensedCount (Est. annual savings: $$potentialSavings)" -ForegroundColor Green
    }
    
    Write-Host "Reports generated:" -ForegroundColor Green
    foreach ($report in $reports) {
        Write-Host " - $report" -ForegroundColor Cyan
    }
    Write-Host "=====================================================" -ForegroundColor Green
    
    # Show how to access the reports
    if ($reports.Count -gt 0) {
        Write-Host ""
        Write-Host "To view the HTML report directly, run:" -ForegroundColor Yellow
        if ($GenerateHTML) {
            $htmlReportPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"
            Write-Host "start '$htmlReportPath'" -ForegroundColor Yellow
        }
        
        Write-Host ""
        Write-Host "For advanced options, run:" -ForegroundColor Yellow
        Write-Host "Get-Help $($MyInvocation.MyCommand.Path) -Detailed" -ForegroundColor Yellow
        Write-Host ""
    }

    # Open HTML report if generated
    if ($GenerateHTML -and (Test-Path -Path (Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"))) {
        $htmlPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"
        Write-Host "Opening HTML report in default browser..." -ForegroundColor Cyan
        Start-Process $htmlPath
    }
} catch {
    Write-Host "Critical error in main script execution: $_" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}