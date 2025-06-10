# M365 Inactive Users Audit Tool
# This script helps audit and validate inactive users in Microsoft 365 environment

# Requirements:
# - Microsoft Graph PowerShell SDK (Install-Module Microsoft.Graph)
# - ImportExcel module for Excel output (Install-Module ImportExcel)
# - PSWriteHTML module for HTML reports (Install-Module PSWriteHTML)

param(
    [Parameter()]
    [ValidateSet("D7", "D30", "D90", "D180")]
    [string]$Period = "D90",
    
    [Parameter()]
    [string]$OutputPath = ".\M365Audit_$(Get-Date -Format 'yyyyMMdd')",
    
    [Parameter()]
    [switch]$Interactive = $true,
    
    [Parameter()]
    [string]$ClientId,
    
    [Parameter()]
    [string]$TenantId,
    
    [Parameter()]
    [string]$CertificateThumbprint,
    
    [Parameter()]
    [switch]$GenerateHTML = $true,
    
    [Parameter()]
    [switch]$GenerateExcel = $true,
    
    [Parameter()]
    [switch]$GenerateCSV = $true,
    
    [Parameter()]
    [switch]$ValidateLicenses = $true,
    
    [Parameter()]
    [switch]$ValidateGroups = $true,
    
    [Parameter()]
    [int]$InactiveThreshold = 45
)

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory | Out-Null
    Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    try {
        if ($Interactive) {
            Write-Host "Connecting to Microsoft Graph interactively..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes "Reports.Read.All", "Directory.Read.All", "User.Read.All", "AuditLog.Read.All", "Organization.Read.All"
        }
        else {
            if (-not $ClientId -or -not $TenantId -or -not $CertificateThumbprint) {
                throw "For non-interactive mode, you must provide ClientId, TenantId, and CertificateThumbprint."
            }
            
            Write-Host "Connecting to Microsoft Graph using certificate authentication..." -ForegroundColor Cyan
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        }
        
        $context = Get-MgContext
        Write-Host "Successfully connected as: $($context.Account)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# Function to get inactive users report
function Get-InactiveUsersReport {
    param (
        [string]$Period
    )
    
    try {
        Write-Host "Fetching Office 365 active user details for period '$Period'..." -ForegroundColor Cyan
        $reportUri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='$Period')"
        $reportData = Invoke-MgGraphRequest -Uri $reportUri -Method GET
        
        # Convert the report content from bytes to string and parse CSV
        $csvContent = [System.Text.Encoding]::UTF8.GetString($reportData.Content)
        $users = ConvertFrom-Csv $csvContent
        
        Write-Host "Successfully retrieved data for $($users.Count) users" -ForegroundColor Green
        return $users
    }
    catch {
        Write-Host "Error fetching user activity report: $_" -ForegroundColor Red
        return $null
    }
}

# Function to get additional user details
function Get-EnhancedUserDetails {
    param (
        [array]$Users
    )
    
    try {
        Write-Host "Enhancing user data with additional details..." -ForegroundColor Cyan
        $enhancedUsers = @()
        $total = $Users.Count
        $current = 0
        
        foreach ($user in $Users) {
            $current++
            Write-Progress -Activity "Retrieving additional user details" -Status "Processing $current of $total" -PercentComplete (($current / $total) * 100)
            
            # Get additional user details from Microsoft Graph
            $userPrincipalName = $user.UserPrincipalName
            
            if ([string]::IsNullOrEmpty($userPrincipalName)) {
                continue
            }
            
            try {
                $mgUser = Get-MgUser -UserId $userPrincipalName -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, Department, AccountEnabled, CreatedDateTime, AssignedLicenses
                
                # Calculate days since last activity
                $lastActivityDate = if ($user.LastActivityDate) { [DateTime]::Parse($user.LastActivityDate) } else { $null }
                $daysSinceActivity = if ($lastActivityDate) { (Get-Date) - $lastActivityDate } else { $null }
                
                # Get assigned licenses
                $assignedLicenses = @()
                if ($ValidateLicenses -and $mgUser.AssignedLicenses) {
                    foreach ($license in $mgUser.AssignedLicenses) {
                        $licensePlan = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $license.SkuId }
                        if ($licensePlan) {
                            $assignedLicenses += $licensePlan.SkuPartNumber
                        }
                    }
                }
                
                # Get group memberships
                $groupMemberships = @()
                if ($ValidateGroups) {
                    $groups = Get-MgUserMemberOf -UserId $mgUser.Id
                    foreach ($group in $groups) {
                        if ($group.AdditionalProperties.displayName) {
                            $groupMemberships += $group.AdditionalProperties.displayName
                        }
                    }
                }
                
                # Create enhanced user object
                $enhancedUser = [PSCustomObject]@{
                    UserPrincipalName   = $userPrincipalName
                    DisplayName         = $mgUser.DisplayName
                    IsActive            = -not [string]::IsNullOrEmpty($user.LastActivityDate)
                    LastActivityDate    = $lastActivityDate
                    DaysSinceActivity   = if ($daysSinceActivity) { $daysSinceActivity.Days } else { $null }
                    IsInactive          = if ($daysSinceActivity -and $daysSinceActivity.Days -gt $InactiveThreshold) { $true } else { $false }
                    AccountEnabled      = $mgUser.AccountEnabled
                    CreatedDateTime     = $mgUser.CreatedDateTime
                    DeletedDate         = $user.DeletedDate
                    ExchangeLastActivityDate = $user.ExchangeLastActivityDate
                    OneDriveLastActivityDate = $user.OneDriveLastActivityDate
                    SharePointLastActivityDate = $user.SharePointLastActivityDate
                    TeamsLastActivityDate = $user.TeamsLastActivityDate
                    YammerLastActivityDate = $user.YammerLastActivityDate
                    AssignedProducts    = $user.AssignedProducts
                    AssignedLicenses    = $assignedLicenses -join "; "
                    LicenseCount        = $assignedLicenses.Count
                    GroupMemberships    = $groupMemberships -join "; "
                    GroupCount          = $groupMemberships.Count
                    Department          = $mgUser.Department
                    JobTitle            = $mgUser.JobTitle
                    ReportRefreshDate   = $user.ReportRefreshDate
                }
                
                $enhancedUsers += $enhancedUser
            }
            catch {
                Write-Host "Error processing user ${userPrincipalName}: $_" -ForegroundColor Yellow
            }
        }
        
        Write-Progress -Activity "Retrieving additional user details" -Completed
        Write-Host "Successfully enhanced data for $($enhancedUsers.Count) users" -ForegroundColor Green
        return $enhancedUsers
    }
    catch {
        Write-Host "Error enhancing user details: $_" -ForegroundColor Red
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
        
        $Users | Export-Excel @excelParams -PassThru
        
        # Create summary worksheet
        $summaryData = [PSCustomObject]@{
            'Total Users' = $Users.Count
            'Inactive Users' = ($Users | Where-Object { $_.IsInactive -eq $true }).Count
            'Disabled Accounts' = ($Users | Where-Object { $_.AccountEnabled -eq $false }).Count
            'Inactive & Licensed' = ($Users | Where-Object { $_.IsInactive -eq $true -and $_.LicenseCount -gt 0 }).Count
            'Days Threshold' = $InactiveThreshold
            'Report Date' = Get-Date
            'Period' = $Period
        }
        
        $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -TableName "Summary" -TableStyle "Medium6"
        
        # Create pivot tables for analysis
        $excel = Open-ExcelPackage -Path $excelPath
        
        # Department analysis
        $pivotTableParams = @{
            PivotTableName = "DepartmentAnalysis"
            Address = $excel.Workbook.Worksheets["Inactive Users"].Cells["A1"].Address
            SourceWorksheet = "Inactive Users"
            PivotRows = @("Department")
            PivotData = @{"UserPrincipalName" = "Count"; "IsInactive" = "Count"}
            PivotColumns = @("IsInactive")
        }
        
        Add-PivotTable @pivotTableParams -PassThru -ExcelPackage $excel
        
        # License analysis
        $pivotTableParams = @{
            PivotTableName = "LicenseAnalysis"
            Address = $excel.Workbook.Worksheets["Inactive Users"].Cells["A1"].Address
            SourceWorksheet = "Inactive Users"
            PivotRows = @("AssignedLicenses")
            PivotData = @{"UserPrincipalName" = "Count"}
            PivotColumns = @("IsInactive")
        }
        
        Add-PivotTable @pivotTableParams -PassThru -ExcelPackage $excel
        
        Close-ExcelPackage $excel
        
        Write-Host "Excel report exported successfully" -ForegroundColor Green
        return $excelPath
    }
    catch {
        Write-Host "Error exporting Excel report: $_" -ForegroundColor Red
        return $null
    }
}

# Function to generate HTML report
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
        
        # Calculate summary stats
        $totalUsers = $Users.Count
        $inactiveUsers = ($Users | Where-Object { $_.IsInactive -eq $true }).Count
        $inactivePercentage = if ($totalUsers -gt 0) { [math]::Round(($inactiveUsers / $totalUsers) * 100, 2) } else { 0 }
        $disabledAccounts = ($Users | Where-Object { $_.AccountEnabled -eq $false }).Count
        $inactiveLicensed = ($Users | Where-Object { $_.IsInactive -eq $true -and $_.LicenseCount -gt 0 }).Count
        $potentialSavings = $inactiveLicensed * 20  # Assuming $20 per license on average
        
        # Department data for charts
        $departmentData = $Users | Group-Object -Property Department | Select-Object @{Name='Name'; Expression={$_.Name}}, @{Name='Count'; Expression={$_.Count}}
        
        # Activity data for timelines
        $activityData = @(
            [PSCustomObject]@{
                Service = "Exchange"
                InactiveCount = ($Users | Where-Object { [string]::IsNullOrEmpty($_.ExchangeLastActivityDate) }).Count
                ActiveCount = ($Users | Where-Object { -not [string]::IsNullOrEmpty($_.ExchangeLastActivityDate) }).Count
            },
            [PSCustomObject]@{
                Service = "OneDrive"
                InactiveCount = ($Users | Where-Object { [string]::IsNullOrEmpty($_.OneDriveLastActivityDate) }).Count
                ActiveCount = ($Users | Where-Object { -not [string]::IsNullOrEmpty($_.OneDriveLastActivityDate) }).Count
            },
            [PSCustomObject]@{
                Service = "SharePoint"
                InactiveCount = ($Users | Where-Object { [string]::IsNullOrEmpty($_.SharePointLastActivityDate) }).Count
                ActiveCount = ($Users | Where-Object { -not [string]::IsNullOrEmpty($_.SharePointLastActivityDate) }).Count
            },
            [PSCustomObject]@{
                Service = "Teams"
                InactiveCount = ($Users | Where-Object { [string]::IsNullOrEmpty($_.TeamsLastActivityDate) }).Count
                ActiveCount = ($Users | Where-Object { -not [string]::IsNullOrEmpty($_.TeamsLastActivityDate) }).Count
            },
            [PSCustomObject]@{
                Service = "Yammer"
                InactiveCount = ($Users | Where-Object { [string]::IsNullOrEmpty($_.YammerLastActivityDate) }).Count
                ActiveCount = ($Users | Where-Object { -not [string]::IsNullOrEmpty($_.YammerLastActivityDate) }).Count
            }
        )
        
        # Generate the HTML report
        New-HTML -TitleText "Microsoft 365 Inactive Users Audit Report" -Online -FilePath $htmlPath {
            New-HTMLHeader {
                New-HTMLText -Text "Microsoft 365 Inactive Users Audit" -Color Black -Alignment center -FontSize 30
                New-HTMLText -Text "Generated on $(Get-Date) | Period: $Period | Inactive Threshold: $InactiveThreshold days" -Color Gray -Alignment center
            }
            
            # Summary section
            New-HTMLSection -HeaderText "Executive Summary" -CanCollapse {
                New-HTMLPanel {
                    New-HTMLText -Text "This report provides an analysis of inactive users in your Microsoft 365 environment." -FontSize 14
                    
                    # Key metrics
                    New-HTMLTable -DataTable @(
                        [PSCustomObject]@{
                            'Metric' = 'Total Users'
                            'Value' = $totalUsers
                        },
                        [PSCustomObject]@{
                            'Metric' = 'Inactive Users'
                            'Value' = $inactiveUsers
                        },
                        [PSCustomObject]@{
                            'Metric' = 'Inactive Percentage'
                            'Value' = "$inactivePercentage%"
                        },
                        [PSCustomObject]@{
                            'Metric' = 'Disabled Accounts'
                            'Value' = $disabledAccounts
                        },
                        [PSCustomObject]@{
                            'Metric' = 'Inactive & Licensed'
                            'Value' = $inactiveLicensed
                        },
                        [PSCustomObject]@{
                            'Metric' = 'Potential Annual Savings'
                            'Value' = "$$($potentialSavings * 12)"
                        }
                    ) -HideFooter
                }
            }
            
            # Dashboard charts
            New-HTMLSection -HeaderText "Dashboard" -CanCollapse {
                New-HTMLPanel {
                    New-HTMLChart {
                        New-ChartPie -Name "User Status" -Value @($inactiveUsers, $totalUsers - $inactiveUsers) -Label @("Inactive", "Active")
                    }
                    
                    New-HTMLChart {
                        New-ChartPie -Name "Account Status" -Value @($disabledAccounts, $totalUsers - $disabledAccounts) -Label @("Disabled", "Enabled")
                    }
                    
                    New-HTMLChart {
                        New-ChartBarOptions -Type Bar
                        foreach ($item in $activityData) {
                            New-ChartBar -Name $item.Service -Value @($item.ActiveCount, $item.InactiveCount) -Label @('Active', 'Inactive')
                        }
                    }
                }
            }
            
            # Department analysis
            New-HTMLSection -HeaderText "Department Analysis" -CanCollapse {
                New-HTMLPanel {
                    New-HTMLChart {
                        New-ChartBarOptions -Type Bar
                        foreach ($dept in ($departmentData | Sort-Object -Property Count -Descending | Select-Object -First 10)) {
                            New-ChartBar -Name $dept.Name -Value $dept.Count
                        }
                    }
                }
            }
            
            # Inactive users table
            New-HTMLSection -HeaderText "Inactive Users" -CanCollapse {
                New-HTMLTable -DataTable ($Users | Where-Object { $_.IsInactive -eq $true } | Sort-Object -Property DaysSinceActivity -Descending) -Filtering -Searching -PagingOptions @(10, 25, 50, 100) -ScrollX
            }
            
            # All users data table
            New-HTMLSection -HeaderText "All Users" -CanCollapse {
                New-HTMLTable -DataTable $Users -Filtering -Searching -PagingOptions @(10, 25, 50, 100) -ScrollX
            }
            
            # Recommendations
            New-HTMLSection -HeaderText "Recommendations" -CanCollapse {
                New-HTMLPanel {
                    New-HTMLList {
                        New-HTMLListItem -Text "Review and consider disabling accounts inactive for more than $InactiveThreshold days"
                        New-HTMLListItem -Text "Remove licenses from inactive accounts to reduce costs"
                        New-HTMLListItem -Text "Implement a regular review process for inactive accounts"
                        New-HTMLListItem -Text "Consider implementing an automated deprovisioning workflow"
                        New-HTMLListItem -Text "Review department-specific inactivity patterns for targeted training"
                    }
                }
            }
        }
        
        Write-Host "HTML report generated successfully" -ForegroundColor Green
        return $htmlPath
    }
    catch {
        Write-Host "Error generating HTML report: $_" -ForegroundColor Red
        return $null
    }
}

# Main execution
Write-Host "Starting Microsoft 365 Inactive Users Audit..." -ForegroundColor Green
Write-Host "Period: $Period | Inactive Threshold: $InactiveThreshold days" -ForegroundColor Green

# Connect to Microsoft Graph
if (-not (Connect-ToMicrosoftGraph)) {
    Write-Host "Failed to connect to Microsoft Graph. Exiting script." -ForegroundColor Red
    return
}

# Get inactive users report
$usersReport = Get-InactiveUsersReport -Period $Period
if (-not $usersReport) {
    Write-Host "Failed to get user activity report. Exiting script." -ForegroundColor Red
    return
}

# Enhance user details
$enhancedUsers = Get-EnhancedUserDetails -Users $usersReport
if (-not $enhancedUsers) {
    Write-Host "Failed to enhance user details. Exiting script." -ForegroundColor Red
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
Write-Host "`nMicrosoft 365 Inactive Users Audit Completed" -ForegroundColor Green
Write-Host "Total users processed: $($enhancedUsers.Count)" -ForegroundColor Green
Write-Host "Inactive users: $(($enhancedUsers | Where-Object { $_.IsInactive -eq $true }).Count)" -ForegroundColor Green
Write-Host "Reports generated:" -ForegroundColor Green
foreach ($report in $reports) {
    Write-Host "- $report" -ForegroundColor Cyan
}

# Open HTML report if generated
if ($GenerateHTML -and (Test-Path -Path (Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"))) {
    $htmlPath = Join-Path -Path $OutputPath -ChildPath "InactiveUsers.html"
    Write-Host "Opening HTML report in default browser..." -ForegroundColor Cyan
    Start-Process $htmlPath
}
