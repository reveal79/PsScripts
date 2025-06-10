# Microsoft 365 App Permissions Audit Script - Simplified Version
# This script identifies applications with high-privilege permissions in your tenant

# Ensure we have the right modules
$graphModules = Get-Module Microsoft.Graph* -ListAvailable

if ($graphModules) {
    Write-Host "Found the following Microsoft Graph modules:" -ForegroundColor Green
    $graphModules | Group-Object Name | ForEach-Object {
        $latestVersion = $_.Group | Sort-Object Version -Descending | Select-Object -First 1
        Write-Host "- $($latestVersion.Name) v$($latestVersion.Version)" -ForegroundColor Green
    }
    
    # Ensure we have the required modules for application management
    $requiredModules = @(
        "Microsoft.Graph.Applications", 
        "Microsoft.Graph.Authentication"
    )
    
    foreach ($module in $requiredModules) {
        if (-not ($graphModules | Where-Object { $_.Name -eq $module })) {
            Write-Host "Installing required module: $module" -ForegroundColor Yellow
            Install-Module $module -Scope CurrentUser -Force
        } else {
            # Import the latest version of each module
            $latestVersion = $graphModules | Where-Object { $_.Name -eq $module } | Sort-Object Version -Descending | Select-Object -First 1
            Import-Module $module -RequiredVersion $latestVersion.Version -Force
        }
    }
} else {
    Write-Host "Microsoft Graph modules not found. Installing required modules..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All"
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Define high-privilege permissions to watch for with detailed descriptions
$highPrivilegePermissions = @(
    # Microsoft Graph Permissions
    @{
        Id="1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"; 
        Name="Application.ReadWrite.All"; 
        Description="Full access to manage applications";
        DetailedDescription="Allows the app to create, read, update and delete any app in your directory.";
        RiskLevel="Critical";
        BusinessImpact="Can create rogue applications, modify other apps' permissions, and gain unauthorized access to company data."
    },
    @{
        Id="62a82d76-70ea-41e2-9197-370581804d09"; 
        Name="Group.ReadWrite.All"; 
        Description="Read and write all groups";
        DetailedDescription="Allows the app to create groups, read all group properties and memberships, update group properties and memberships, and delete groups.";
        RiskLevel="High";
        BusinessImpact="Can access all Teams content, add/remove members from security groups, read all files shared in all groups."
    },
    @{
        Id="06da0dbc-49e2-44d2-8312-53f166ab848a"; 
        Name="Directory.ReadWrite.All"; 
        Description="Read and write directory data";
        DetailedDescription="Allows the app to read and write all data in your organization's directory, such as users, groups, and applications.";
        RiskLevel="Critical";
        BusinessImpact="Effectively grants admin rights to the entire directory. Can manage all users, reset passwords, and modify security groups."
    },
    @{
        Id="741f803b-c850-494e-b5df-cde7c675a1ca"; 
        Name="Directory.Read.All"; 
        Description="Read directory data";
        DetailedDescription="Allows the app to read data in your organization's directory.";
        RiskLevel="Medium";
        BusinessImpact="Access to all user information, group memberships, and org structure."
    },
    @{
        Id="7b2449af-6ccd-4f4d-9f78-e550c193f0d1"; 
        Name="Files.ReadWrite.All"; 
        Description="Read and write files in all site collections";
        DetailedDescription="Allows the app to read, create, update, and delete all files in all site collections.";
        RiskLevel="High";
        BusinessImpact="Can access, modify, and delete ANY file stored in SharePoint/OneDrive across the entire organization."
    },
    @{
        Id="75359482-378d-4052-8f01-80520e7db3cd"; 
        Name="Files.Read.All"; 
        Description="Read files in all site collections";
        DetailedDescription="Allows the app to read all files in all site collections.";
        RiskLevel="High";
        BusinessImpact="Can access ANY file stored in SharePoint/OneDrive across the entire organization."
    },
    @{
        Id="82895abf-7a98-4f37-8519-1f0be6274800"; 
        Name="Sites.Selected"; 
        Description="Limited access to specific SharePoint sites";
        DetailedDescription="Allows the app to access only specific SharePoint sites that an admin explicitly approves.";
        RiskLevel="Low";
        BusinessImpact="Limited to accessing only specific sites that are explicitly approved by an admin. Recommended for most apps that need SharePoint access."
    }
)

# Initialize results array
$results = @()

# Get all service principals (apps) in the tenant
Write-Host "Retrieving all service principals..." -ForegroundColor Yellow
try {
    $servicePrincipals = Get-MgServicePrincipal -All
    Write-Host "Retrieved $($servicePrincipals.Count) service principals." -ForegroundColor Green
} catch {
    Write-Host "Error retrieving service principals: $_" -ForegroundColor Red
    # Fallback to listing applications if service principals fail
    try {
        Write-Host "Trying to retrieve applications instead..." -ForegroundColor Yellow
        $applications = Get-MgApplication -All
        Write-Host "Retrieved $($applications.Count) applications." -ForegroundColor Green
        # Create a simplified output structure
        $servicePrincipals = $applications | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                Id = $_.Id
                AppId = $_.AppId
                PublisherName = "Unknown" # Application objects don't have publisher info directly
                AccountEnabled = $true    # Assuming enabled by default
                SignInAudience = $_.SignInAudience
                CreatedDateTime = $_.CreatedDateTime
                Tags = $_.Tags
                Notes = "Retrieved via Application API"
            }
        }
    } catch {
        Write-Host "Error retrieving applications: $_" -ForegroundColor Red
        Write-Host "This might be a permissions issue. Please ensure you have the Application.Read.All and Directory.Read.All permissions." -ForegroundColor Yellow
        exit
    }
}

$totalApps = $servicePrincipals.Count
$currentApp = 0

foreach ($sp in $servicePrincipals) {
    $currentApp++
    Write-Progress -Activity "Analyzing applications" -Status "Processing $($sp.DisplayName)" -PercentComplete (($currentApp / $totalApps) * 100)
    
    # Get the app roles (application permissions) assigned to this service principal
    $appRoleAssignments = @()
    try {
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Error fetching app role assignments for $($sp.DisplayName): $_"
    }
    
    # Check if any high-privilege permissions are assigned
    $highPrivPermissionsFound = @()
    
    foreach ($assignment in $appRoleAssignments) {
        $permission = $highPrivilegePermissions | Where-Object { $_.Id -eq $assignment.AppRoleId }
        
        if ($permission) {
            $resourceSp = Get-MgServicePrincipal -ServicePrincipalId $assignment.ResourceId -ErrorAction SilentlyContinue
            $resourceName = if ($resourceSp) { $resourceSp.DisplayName } else { "Unknown Resource" }
            
            $highPrivPermissionsFound += "$($permission.Name) ($resourceName)"
        }
    }
    
    # Add to results if high-privilege permissions were found
    if ($highPrivPermissionsFound.Count -gt 0) {
        $highestRiskLevel = "Low"
        $riskScoreNumeric = 0
        
        # Calculate risk score based on permissions
        foreach ($permName in $highPrivPermissionsFound) {
            $permInfo = $permName -split " \("
            $actualPermName = $permInfo[0]
            $permDetail = $highPrivilegePermissions | Where-Object { $_.Name -eq $actualPermName }
            
            if ($permDetail) {
                # Update highest risk level based on permission risk
                if ($permDetail.RiskLevel -eq "Critical") {
                    if ($highestRiskLevel -ne "Critical") { $highestRiskLevel = "Critical" }
                    $riskScoreNumeric += 100
                }
                elseif ($permDetail.RiskLevel -eq "High") {
                    if ($highestRiskLevel -ne "Critical" -and $highestRiskLevel -ne "High") { $highestRiskLevel = "High" }
                    $riskScoreNumeric += 50
                }
                elseif ($permDetail.RiskLevel -eq "Medium") {
                    if ($highestRiskLevel -ne "Critical" -and $highestRiskLevel -ne "High" -and $highestRiskLevel -ne "Medium") { $highestRiskLevel = "Medium" }
                    $riskScoreNumeric += 20
                }
                elseif ($permDetail.RiskLevel -eq "Low") {
                    $riskScoreNumeric += 5
                }
            }
        }
        
        # Apply multipliers for certain app characteristics
        if (-not $sp.PublisherName) {
            # Higher risk for apps without verified publishers
            $riskScoreNumeric *= 1.5
        }
        
        if ($sp.SignInAudience -eq "AzureADMultipleOrgs") {
            # Higher risk for multi-tenant apps
            $riskScoreNumeric *= 1.2
        }
        
        $results += [PSCustomObject]@{
            "ApplicationName" = $sp.DisplayName
            "ApplicationId" = $sp.AppId
            "Publisher" = $sp.PublisherName
            "HighPrivilegePermissions" = ($highPrivPermissionsFound -join ", ")
            "CreatedDateTime" = $sp.CreatedDateTime
            "IsServicePrincipalEnabled" = $sp.AccountEnabled
            "SignInAudience" = $sp.SignInAudience
            "Tags" = ($sp.Tags -join ", ")
            "Notes" = $sp.Notes
            "RiskLevel" = $highestRiskLevel
            "RiskScore" = [math]::Round($riskScoreNumeric)
            "BusinessImpact" = ($highPrivPermissionsFound | ForEach-Object {
                $permInfo = $_ -split " \("
                $actualPermName = $permInfo[0]
                $perm = $highPrivilegePermissions | Where-Object { $_.Name -eq $actualPermName }
                if ($perm) { "- $($actualPermName): $($perm.BusinessImpact)" }
            }) -join "`n"
        }
    }
}

Write-Progress -Activity "Analyzing applications" -Completed

# Check for applications with Sites.Selected permissions
Write-Host "Checking for applications with Sites.Selected permissions..." -ForegroundColor Yellow

$sitesSelectedPermId = "82895abf-7a98-4f37-8519-1f0be6274800"
$appsWithSitesSelected = @()

foreach ($sp in $servicePrincipals) {
    $hasPermission = $false
    
    try {
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
        $hasPermission = $appRoleAssignments | Where-Object { $_.AppRoleId -eq $sitesSelectedPermId } | Select-Object -First 1
    } catch {
        Write-Warning "Error checking Sites.Selected for $($sp.DisplayName): $_"
    }
    
    if ($hasPermission) {
        # Simplified version without trying to get specific sites (which might cause errors)
        $appsWithSitesSelected += [PSCustomObject]@{
            "ApplicationName" = $sp.DisplayName
            "ApplicationId" = $sp.AppId
            "SitesWithAccess" = "Permission granted, site list unavailable"
            "PermissionType" = "Read (assumed)"
        }
    }
}

# Ensure we have the right modules for Excel formatting
$excelModuleInstalled = $false
try {
    # Check if ImportExcel module is installed
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module for better report formatting..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force
        $excelModuleInstalled = $true
    } else {
        $excelModuleInstalled = $true
    }
} catch {
    Write-Host "Could not install ImportExcel module. Will use basic CSV output instead." -ForegroundColor Yellow
    $excelModuleInstalled = $false
}

# Get the current date for the filename
$dateStr = Get-Date -Format "yyyyMMdd-HHmmss"

# Generate HTML report first so we can reference it
$htmlReport = "AppPermissionsEducationalReport-$dateStr.html"
$criticalApps = $results | Where-Object { $_.RiskLevel -eq "Critical" }
$highRiskApps = $results | Where-Object { $_.RiskLevel -eq "High" }
$otherApps = $results | Where-Object { $_.RiskLevel -ne "Critical" -and $_.RiskLevel -ne "High" }

# [HTML report generation code remains the same]

$htmlContent | Out-File -FilePath $htmlReport
Write-Host "Exported educational HTML report to $htmlReport" -ForegroundColor Green

# Export results to Excel with formatting if module is available, otherwise fall back to CSV
if ($excelModuleInstalled) {
    $excelFile = "AppPermissionsAudit-$dateStr.xlsx"
    
    # Add hyperlink to HTML report in the Excel data
    $resultsWithLinks = $results | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name "ReportLink" -Value "=HYPERLINK(`"$htmlReport`",`"View Details`")" -Force
        $_
    }
    
    # Create Excel package - use the safe version of business impact for Excel
    $excelPackage = $resultsWithLinks | Select-Object ApplicationName, Publisher, HighPrivilegePermissions, RiskLevel, RiskScore, IsServicePrincipalEnabled, CreatedDateTime, @{N='BusinessImpact';E={$_.BusinessImpactSafe}}, ReportLink | 
        Export-Excel -Path $excelFile -WorksheetName "High Privilege Apps" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -PassThru

    # Add formatting based on risk level
    $worksheet = $excelPackage.Workbook.Worksheets["High Privilege Apps"]
    
    # Find the RiskLevel column
    $riskLevelCol = $null
    for ($i = 1; $i -le $worksheet.Dimension.End.Column; $i++) {
        if ($worksheet.Cells[1, $i].Value -eq "RiskLevel") {
            $riskLevelCol = $i
            break
        }
    }
    
    if ($riskLevelCol -ne $null) {
        # Add conditional formatting for risk levels - safely
        try {
            $lastRow = $worksheet.Dimension.End.Row
            if ($lastRow -gt 1) { # Make sure we have data rows
                Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "Critical" -BackgroundColor "#FFCCCC" -PassThru | Out-Null
                Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "High" -BackgroundColor "#FFEBCC" -PassThru | Out-Null
                Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "Medium" -BackgroundColor "#FFFFCC" -PassThru | Out-Null
                Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "Low" -BackgroundColor "#E6F5FF" -PassThru | Out-Null
            }
        } catch {
            Write-Host "Could not add conditional formatting. Continuing with basic formatting." -ForegroundColor Yellow
        }
    }
    
    # Add a dashboard worksheet
    try {
        $dashboardSheet = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "Dashboard" -Activate
        
        # Set up the dashboard header
        $dashboardSheet.Cells["A1"].Value = "APPLICATION PERMISSIONS AUDIT DASHBOARD"
        $dashboardSheet.Cells["A1"].Style.Font.Size = 16
        $dashboardSheet.Cells["A1"].Style.Font.Bold = $true
        $dashboardSheet.Cells["A1:J1"].Merge = $true
        
        # Add summary statistics
        $dashboardSheet.Cells["A3"].Value = "SUMMARY"
        $dashboardSheet.Cells["A3"].Style.Font.Bold = $true
        $dashboardSheet.Cells["A3:J3"].Merge = $true
        
        $dashboardSheet.Cells["A4"].Value = "Total Applications Analyzed:"
        $dashboardSheet.Cells["B4"].Value = $totalApps
        
        $dashboardSheet.Cells["A5"].Value = "Applications with High-Privilege Permissions:"
        $dashboardSheet.Cells["B5"].Value = $results.Count
        
        $dashboardSheet.Cells["A6"].Value = "Critical Risk Applications:"
        $dashboardSheet.Cells["B6"].Value = $criticalApps.Count
        $dashboardSheet.Cells["B6"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $dashboardSheet.Cells["B6"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 204, 204))
        
        $dashboardSheet.Cells["A7"].Value = "High Risk Applications:"
        $dashboardSheet.Cells["B7"].Value = $highRiskApps.Count
        $dashboardSheet.Cells["B7"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $dashboardSheet.Cells["B7"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 204))
        
        $dashboardSheet.Cells["A8"].Value = "Applications Using Sites.Selected (Best Practice):"
        $dashboardSheet.Cells["B8"].Value = $appsWithSitesSelected.Count
        
        # Add link to HTML report
        $dashboardSheet.Cells["A10"].Value = "View Detailed HTML Report:"
        $dashboardSheet.Cells["B10"].Value = "Open Report"
        $dashboardSheet.Cells["B10"].Hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $htmlReport, "Click to open detailed HTML report"
        $dashboardSheet.Cells["B10"].Style.Font.UnderLine = $true
        $dashboardSheet.Cells["B10"].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        
        # Display top risky applications
        $dashboardSheet.Cells["A12"].Value = "TOP RISKY APPLICATIONS"
        $dashboardSheet.Cells["A12"].Style.Font.Bold = $true
        $dashboardSheet.Cells["A12:J12"].Merge = $true
        
        $dashboardSheet.Cells["A13"].Value = "Application Name"
        $dashboardSheet.Cells["B13"].Value = "Publisher"
        $dashboardSheet.Cells["C13"].Value = "Risk Level"
        $dashboardSheet.Cells["D13"].Value = "Risk Score"
        $dashboardSheet.Cells["E13"].Value = "Status"
        $dashboardSheet.Cells["F13"].Value = "Permissions"
        $dashboardSheet.Cells["A13:F13"].Style.Font.Bold = $true
        
        # Get top 10 risky apps by risk score
        $topRiskyApps = $results | Sort-Object -Property RiskScore -Descending | Select-Object -First 10
        $row = 14
        foreach ($app in $topRiskyApps) {
            $dashboardSheet.Cells["A$row"].Value = $app.ApplicationName
            $dashboardSheet.Cells["B$row"].Value = $app.Publisher
            $dashboardSheet.Cells["C$row"].Value = $app.RiskLevel
            $dashboardSheet.Cells["D$row"].Value = $app.RiskScore
            $dashboardSheet.Cells["E$row"].Value = if ($app.IsServicePrincipalEnabled) { "ENABLED" } else { "Disabled" }
            $dashboardSheet.Cells["F$row"].Value = $app.HighPrivilegePermissions
            
            # Color the row based on risk level
            if ($app.RiskLevel -eq "Critical") {
                $dashboardSheet.Cells["A$row:F$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $dashboardSheet.Cells["A$row:F$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 204, 204))
            } elseif ($app.RiskLevel -eq "High") {
                $dashboardSheet.Cells["A$row:F$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $dashboardSheet.Cells["A$row:F$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 204))
            }
            
            $row++
        }
        
        # Add Action Buttons section
        $dashboardSheet.Cells["A$($row+2)"].Value = "RECOMMENDED ACTIONS"
        $dashboardSheet.Cells["A$($row+2)"].Style.Font.Bold = $true
        $dashboardSheet.Cells["A$($row+2):J$($row+2)"].Merge = $true
        
        $recommendations = @(
            "1. Review all critical and high-risk applications immediately",
            "2. Replace Files.Read.All with Sites.Selected where possible",
            "3. Document business justification for high-privilege permissions",
            "4. Disable applications with dangerous permissions that are not needed",
            "5. Implement regular permission audits as part of your security program"
        )
        
        $recRow = $row + 3
        foreach ($rec in $recommendations) {
            $dashboardSheet.Cells["A$recRow"].Value = $rec
            $dashboardSheet.Cells["A$recRow:F$recRow"].Merge = $true
            $recRow++
        }
    } catch {
        Write-Host "Could not create dashboard sheet. Continuing with basic sheet." -ForegroundColor Yellow
    }
    
    # Save and close the Excel file
    Close-ExcelPackage $excelPackage -Show
    
    Write-Host "Exported interactive Excel report to $excelFile" -ForegroundColor Green
} else {
    # Fall back to CSV if Excel module not available
    $outputFile = "AppPermissionsAudit-$dateStr.csv"
    $results | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Exported permissions report to $outputFile" -ForegroundColor Green
}

# Export Sites.Selected apps
$sitesSelectedFile = "SitesSelectedApps-$dateStr.csv"
$appsWithSitesSelected | Export-Csv -Path $sitesSelectedFile -NoTypeInformation
Write-Host "Exported Sites.Selected permissions report to $sitesSelectedFile" -ForegroundColor Green

$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 Application Permissions Audit</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
        h1 { color: #0078D4; border-bottom: 1px solid #0078D4; padding-bottom: 10px; }
        h2 { color: #0078D4; margin-top: 30px; }
        h3 { margin-top: 25px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 30px; }
        th, td { text-align: left; padding: 12px; }
        th { background-color: #0078D4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .critical { background-color: #FFCCCC; }
        .high { background-color: #FFEBCC; }
        .medium { background-color: #FFFFCC; }
        .low { background-color: #E6F5FF; }
        .permission-table { margin-left: 20px; width: 95%; }
        .permission-name { font-weight: bold; }
        .business-impact { font-style: italic; color: #D83B01; }
        .section { background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .summary-box { background-color: #e6f0ff; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 5px solid #0078D4; }
        .recommendation { background-color: #e9f4ea; padding: 15px; border-radius: 5px; margin: 20px 0; border-left: 5px solid #107C10; }
    </style>
</head>
<body>
    <h1>Microsoft 365 Application Permissions Audit</h1>
    <p>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm")</p>
    
    <div class="summary-box">
        <h2>Executive Summary</h2>
        <p><b>Total applications analyzed:</b> $totalApps</p>
        <p><b>Applications with high-privilege permissions:</b> $($results.Count)</p>
        <p><b>Critical risk applications:</b> $($criticalApps.Count)</p>
        <p><b>High risk applications:</b> $($highRiskApps.Count)</p>
        <p><b>Applications with Sites.Selected (recommended) permissions:</b> $($appsWithSitesSelected.Count)</p>
    </div>
    
    <div class="section">
        <h2>Understanding Permission Risks</h2>
        <p>Application permissions in Microsoft 365 can pose significant security risks if not properly managed. This report highlights applications with high-privilege permissions that could potentially be misused.</p>
        
        <h3>Risk Levels Explained:</h3>
        <ul>
            <li><b>Critical:</b> Permissions that grant broad administrative capabilities with potential for catastrophic data breaches or tenant-wide compromise</li>
            <li><b>High:</b> Permissions that grant access to sensitive data across multiple services or teams</li>
            <li><b>Medium:</b> Permissions that grant read access to potentially sensitive information</li>
            <li><b>Low:</b> Permissions with limited scope or impact</li>
        </ul>
        
        <h3>Key Permission Descriptions</h3>
        <table class="permission-table">
            <tr>
                <th>Permission</th>
                <th>Risk Level</th>
                <th>What It Can Do</th>
            </tr>
"@

foreach ($perm in ($highPrivilegePermissions | Sort-Object -Property RiskLevel -Descending)) {
    $riskClass = $perm.RiskLevel.ToLower()
    $htmlContent += @"
        <tr class="$riskClass">
            <td class="permission-name">$($perm.Name)</td>
            <td>$($perm.RiskLevel)</td>
            <td>$($perm.DetailedDescription)<p class="business-impact">Business Impact: $($perm.BusinessImpact)</p></td>
        </tr>
"@
}

$htmlContent += @"
        </table>
    </div>
    
    <div class="recommendation">
        <h2>Best Practices & Recommendations</h2>
        <ul>
            <li><b>Use the principle of least privilege</b>: Applications should only have the minimum permissions necessary to function.</li>
            <li><b>Replace broad permissions with specific ones</b>: Instead of Files.Read.All, consider using Sites.Selected for specific SharePoint sites access.</li>
            <li><b>Review permissions regularly</b>: Conduct regular audits of application permissions, especially for critical and high-risk applications.</li>
            <li><b>Document business justification</b>: Maintain documentation for why each application needs its assigned permissions.</li>
            <li><b>Implement application access lifecycle</b>: Remove or disable unused applications and monitor for sudden permission changes.</li>
            <li><b>Use Conditional Access for applications</b>: Apply additional security policies for applications with high-privilege permissions.</li>
        </ul>
    </div>
"@

if ($criticalApps.Count -gt 0) {
    $htmlContent += @"
    <h2>Critical Risk Applications</h2>
    <p>These applications have permissions that could potentially allow full tenant compromise or significant data exfiltration:</p>
    <table>
        <tr>
            <th>Application Name</th>
            <th>Publisher</th>
            <th>Permissions</th>
            <th>Risk Score</th>
        </tr>
"@

    foreach ($app in $criticalApps) {
        $htmlContent += @"
        <tr class="critical">
            <td>$($app.ApplicationName)</td>
            <td>$($app.Publisher)</td>
            <td>$($app.HighPrivilegePermissions)</td>
            <td>$($app.RiskScore)</td>
        </tr>
"@
    }
    
    $htmlContent += @"
    </table>
"@
}

if ($highRiskApps.Count -gt 0) {
    $htmlContent += @"
    <h2>High Risk Applications</h2>
    <p>These applications have permissions that provide broad access to sensitive data:</p>
    <table>
        <tr>
            <th>Application Name</th>
            <th>Publisher</th>
            <th>Permissions</th>
            <th>Risk Score</th>
        </tr>
"@

    foreach ($app in $highRiskApps) {
        $htmlContent += @"
        <tr class="high">
            <td>$($app.ApplicationName)</td>
            <td>$($app.Publisher)</td>
            <td>$($app.HighPrivilegePermissions)</td>
            <td>$($app.RiskScore)</td>
        </tr>
"@
    }
    
    $htmlContent += @"
    </table>
"@
}

if ($otherApps.Count -gt 0) {
    $htmlContent += @"
    <h2>Other Applications with Privileged Permissions</h2>
    <table>
        <tr>
            <th>Application Name</th>
            <th>Publisher</th>
            <th>Permissions</th>
            <th>Risk Level</th>
            <th>Risk Score</th>
        </tr>
"@

    foreach ($app in $otherApps) {
        $riskClass = $app.RiskLevel.ToLower()
        $htmlContent += @"
        <tr class="$riskClass">
            <td>$($app.ApplicationName)</td>
            <td>$($app.Publisher)</td>
            <td>$($app.HighPrivilegePermissions)</td>
            <td>$($app.RiskLevel)</td>
            <td>$($app.RiskScore)</td>
        </tr>
"@
    }
    
    $htmlContent += @"
    </table>
"@
}

if ($appsWithSitesSelected.Count -gt 0) {
    $htmlContent += @"
    <h2>Applications Using Recommended Sites.Selected Permission</h2>
    <p>These applications are following best practices by using Sites.Selected permission instead of broad Files.Read.All permissions:</p>
    <table>
        <tr>
            <th>Application Name</th>
            <th>Sites With Access</th>
            <th>Permission Type</th>
        </tr>
"@

    foreach ($app in $appsWithSitesSelected) {
        $htmlContent += @"
        <tr>
            <td>$($app.ApplicationName)</td>
            <td>$($app.SitesWithAccess)</td>
            <td>$($app.PermissionType)</td>
        </tr>
"@
    }
    
    $htmlContent += @"
    </table>
"@
}

$htmlContent += @"
    <div class="section">
        <h2>Next Steps</h2>
        <ol>
            <li>Review all critical and high-risk applications to confirm they require these permissions</li>
            <li>Consider migrating applications from broad permissions (like Files.Read.All) to more specific ones (like Sites.Selected)</li>
            <li>Document business justification for any application that requires high-privilege permissions</li>
            <li>Implement regular permission audits as part of your security program</li>
        </ol>
    </div>
    
    <p><small>Generated by Microsoft 365 Application Permissions Audit Script. For more information, contact your Microsoft 365 administrator.</small></p>
</body>
</html>
"@

$htmlContent | Out-File -FilePath $htmlReport
Write-Host "Exported educational HTML report to $htmlReport" -ForegroundColor Green

# Summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total applications in tenant: $totalApps" -ForegroundColor White
Write-Host "Applications with high-privilege permissions: $($results.Count)" -ForegroundColor Yellow
Write-Host "Critical risk applications: $($criticalApps.Count)" -ForegroundColor Red
Write-Host "High risk applications: $($highRiskApps.Count)" -ForegroundColor Yellow
Write-Host "Applications with Sites.Selected permissions: $($appsWithSitesSelected.Count)" -ForegroundColor Green

Write-Host "`nReports generated:" -ForegroundColor Cyan
Write-Host "- CSV Report: $outputFile" -ForegroundColor White
Write-Host "- Sites.Selected Report: $sitesSelectedFile" -ForegroundColor White
Write-Host "- Educational HTML Report: $htmlReport" -ForegroundColor White

# Provide remediation suggestions
Write-Host "`n=== REMEDIATION SUGGESTIONS ===" -ForegroundColor Cyan

# Check for critical applications
$criticalApps = $results | Where-Object { $_.RiskLevel -eq "Critical" }
if ($criticalApps.Count -gt 0) {
    Write-Host "`nCRITICAL: Found $($criticalApps.Count) applications with critical risk permissions" -ForegroundColor Red
    Write-Host "These applications have permissions that could potentially compromise your entire tenant:" -ForegroundColor Red
    $criticalApps | ForEach-Object {
        $enabledStatus = if ($_.IsServicePrincipalEnabled) { "ENABLED" } else { "Disabled" }
        $statusColor = if ($_.IsServicePrincipalEnabled) { "Red" } else { "Green" }
        Write-Host "- $($_.ApplicationName) ($($_.ApplicationId)) - Status: " -ForegroundColor Red -NoNewline
        Write-Host "$enabledStatus" -ForegroundColor $statusColor
        Write-Host "  Permissions: $($_.HighPrivilegePermissions)" -ForegroundColor Red
    }
    
    Write-Host "`nRemediation steps for critical applications:" -ForegroundColor Yellow
    Write-Host "1. Review each critical application immediately" -ForegroundColor White
    Write-Host "2. Verify business need for these permissions" -ForegroundColor White
    Write-Host "3. Consider replacing with more limited permissions" -ForegroundColor White
    Write-Host "4. Disable applications that don't need these permissions using:" -ForegroundColor White
    Write-Host "   Update-MgServicePrincipal -ServicePrincipalId `$spId -AccountEnabled `$false" -ForegroundColor White
}

# Add function to disable service principals
function Disable-RiskyApplication {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApplicationId,
        [Parameter(Mandatory = $false)]
        [string]$Reason = "High-risk permissions detected"
    )

    try {
        # First, find the service principal ID from the app ID
        $sp = Get-MgServicePrincipal -Filter "appId eq '$ApplicationId'"
        
        if (-not $sp) {
            Write-Host "Could not find service principal for application ID: $ApplicationId" -ForegroundColor Yellow
            return $false
        }
        
        # Disable the service principal
        Update-MgServicePrincipal -ServicePrincipalId $sp.Id -AccountEnabled:$false
        
        # Log the action
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp - Disabled application: $($sp.DisplayName) ($ApplicationId) - Reason: $Reason"
        $logEntry | Out-File -Append -FilePath "AppPermissionsAudit-Actions.log"
        
        Write-Host "Successfully disabled application: $($sp.DisplayName) ($ApplicationId)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error disabling application $ApplicationId : $_" -ForegroundColor Red
        return $false
    }
}

# Offer to disable critical applications
if ($criticalApps.Count -gt 0) {
    $enabledCriticalApps = $criticalApps | Where-Object { $_.IsServicePrincipalEnabled -eq $true }
    
    if ($enabledCriticalApps.Count -gt 0) {
        Write-Host "`nWould you like to disable any critical risk applications? (Y/N)" -ForegroundColor Yellow
        $disableResponse = Read-Host
        
        if ($disableResponse.ToUpper() -eq "Y") {
            Write-Host "`nSelect applications to disable:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $enabledCriticalApps.Count; $i++) {
                Write-Host "[$i] $($enabledCriticalApps[$i].ApplicationName) ($($enabledCriticalApps[$i].ApplicationId))" -ForegroundColor White
            }
            Write-Host "[A] All critical applications" -ForegroundColor White
            Write-Host "[C] Cancel" -ForegroundColor White
            
            $selection = Read-Host "Enter selection"
            
            if ($selection.ToUpper() -eq "A") {
                Write-Host "Disabling all critical applications..." -ForegroundColor Yellow
                foreach ($app in $enabledCriticalApps) {
                    Disable-RiskyApplication -ApplicationId $app.ApplicationId -Reason "Critical risk permissions"
                }
            }
            elseif ($selection.ToUpper() -eq "C") {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
            }
            elseif ([int]::TryParse($selection, [ref]$null) -and [int]$selection -ge 0 -and [int]$selection -lt $enabledCriticalApps.Count) {
                $appToDisable = $enabledCriticalApps[[int]$selection]
                Disable-RiskyApplication -ApplicationId $appToDisable.ApplicationId -Reason "Critical risk permissions"
            }
            else {
                Write-Host "Invalid selection." -ForegroundColor Red
            }
        }
    }
}

# Check for applications with Files.Read.All or Files.ReadWrite.All that could use Sites.Selected
$broadFileAccessApps = $results | Where-Object { 
    $_.HighPrivilegePermissions -match "Files\.Read\.All|Files\.ReadWrite\.All" -and 
    $appsWithSitesSelected.ApplicationId -notcontains $_.ApplicationId
}

if ($broadFileAccessApps.Count -gt 0) {
    Write-Host "`nWARNING: Found $($broadFileAccessApps.Count) applications with broad file access that could use Sites.Selected instead" -ForegroundColor Yellow
    Write-Host "These applications have access to ALL files in your tenant:" -ForegroundColor Yellow
    $broadFileAccessApps | ForEach-Object {
        Write-Host "- $($_.ApplicationName) ($($_.ApplicationId))" -ForegroundColor Yellow
    }
    
    Write-Host "`nRemediation steps to restrict file access:" -ForegroundColor Yellow
    Write-Host "1. Remove the broad Files.Read.All or Files.ReadWrite.All permissions" -ForegroundColor White
    Write-Host "2. Add the Sites.Selected permission instead" -ForegroundColor White
    Write-Host "3. Grant access only to specific sites using:" -ForegroundColor White
    Write-Host "   Grant-MgServicePrincipalSitePermission -ServicePrincipalId `$spId -SiteId `$siteId -PermissionType 'Read'" -ForegroundColor White
}

# Check for multi-tenant applications with high permissions
$multiTenantHighRiskApps = $results | Where-Object { 
    $_.SignInAudience -eq "AzureADMultipleOrgs" -and 
    ($_.RiskLevel -eq "Critical" -or $_.RiskLevel -eq "High")
}

if ($multiTenantHighRiskApps.Count -gt 0) {
    Write-Host "`nWARNING: Found $($multiTenantHighRiskApps.Count) multi-tenant applications with high or critical permissions" -ForegroundColor Yellow
    Write-Host "Multi-tenant apps with high permissions pose additional risk:" -ForegroundColor Yellow
    $multiTenantHighRiskApps | ForEach-Object {
        Write-Host "- $($_.ApplicationName) ($($_.ApplicationId))" -ForegroundColor Yellow
    }
    
    Write-Host "`nRemediation steps for multi-tenant applications:" -ForegroundColor Yellow
    Write-Host "1. Verify the necessity of multi-tenant configuration" -ForegroundColor White
    Write-Host "2. Consider changing to single-tenant if possible" -ForegroundColor White
    Write-Host "3. Implement Conditional Access policies for these applications" -ForegroundColor White
}

Write-Host "`n=== USAGE INSTRUCTIONS ===" -ForegroundColor Cyan
Write-Host "This report can be used for the following purposes:" -ForegroundColor White
Write-Host "1. Security Audits: Identify applications with excessive permissions" -ForegroundColor White
Write-Host "2. Developer Education: Share the HTML report to help developers understand permission risks" -ForegroundColor White
Write-Host "3. Presentations: Use the CSV data in Excel for presentations on app security" -ForegroundColor White
Write-Host "4. Compliance Documentation: Document your regular permission reviews" -ForegroundColor White
Write-Host "5. Risk Remediation: Prioritize fixing applications with critical and high-risk permissions" -ForegroundColor White

# Also add a menu option at the end of the script to take immediate actions
Write-Host "`n=== TAKE ACTION ===" -ForegroundColor Cyan
Write-Host "What would you like to do next?" -ForegroundColor White
Write-Host "1. Disable risky applications" -ForegroundColor White
Write-Host "2. View applications by risk level" -ForegroundColor White
Write-Host "3. Export application details to JSON (with more data)" -ForegroundColor White
Write-Host "4. Exit" -ForegroundColor White

$menuSelection = Read-Host "Enter selection"

switch ($menuSelection) {
    "1" {
        Write-Host "`nSelect applications to view:" -ForegroundColor Cyan
        Write-Host "1. Critical risk applications" -ForegroundColor Red
        Write-Host "2. High risk applications" -ForegroundColor Yellow
        Write-Host "3. All applications with privileged permissions" -ForegroundColor White
        Write-Host "4. Back to main menu" -ForegroundColor White
        
        $riskSelection = Read-Host "Enter selection"
        
        $appsToShow = @()
        switch ($riskSelection) {
            "1" { $appsToShow = $results | Where-Object { $_.RiskLevel -eq "Critical" } }
            "2" { $appsToShow = $results | Where-Object { $_.RiskLevel -eq "High" } }
            "3" { $appsToShow = $results }
            "4" { return }
            default { 
                Write-Host "Invalid selection." -ForegroundColor Red
                return
            }
        }
        
        $enabledApps = $appsToShow | Where-Object { $_.IsServicePrincipalEnabled -eq $true }
        
        if ($enabledApps.Count -eq 0) {
            Write-Host "No enabled applications found with the selected risk level." -ForegroundColor Yellow
            return
        }
        
        Write-Host "`nSelect applications to disable:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $enabledApps.Count; $i++) {
            Write-Host "[$i] $($enabledApps[$i].ApplicationName) ($($enabledApps[$i].ApplicationId))" -ForegroundColor White
        }
        Write-Host "[A] All applications in this category" -ForegroundColor White
        Write-Host "[C] Cancel" -ForegroundColor White
        
        $selection = Read-Host "Enter selection"
        
        if ($selection.ToUpper() -eq "A") {
            Write-Host "Disabling all selected applications..." -ForegroundColor Yellow
            foreach ($app in $enabledApps) {
                Disable-RiskyApplication -ApplicationId $app.ApplicationId -Reason "High-risk permissions"
            }
        }
        elseif ($selection.ToUpper() -eq "C") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
        }
        elseif ([int]::TryParse($selection, [ref]$null) -and [int]$selection -ge 0 -and [int]$selection -lt $enabledApps.Count) {
            $appToDisable = $enabledApps[[int]$selection]
            Disable-RiskyApplication -ApplicationId $appToDisable.ApplicationId -Reason "High-risk permissions"
        }
        else {
            Write-Host "Invalid selection." -ForegroundColor Red
        }
    }
    "2" {
        Write-Host "`nApplications by risk level:" -ForegroundColor Cyan
        
        Write-Host "`nCritical Risk ($($results | Where-Object { $_.RiskLevel -eq 'Critical' } | Measure-Object).Count):" -ForegroundColor Red
        $results | Where-Object { $_.RiskLevel -eq "Critical" } | ForEach-Object {
            $enabledStatus = if ($_.IsServicePrincipalEnabled) { "ENABLED" } else { "Disabled" }
            $statusColor = if ($_.IsServicePrincipalEnabled) { "Red" } else { "Green" }
            Write-Host "- $($_.ApplicationName) ($($_.ApplicationId)) - Status: " -NoNewline
            Write-Host "$enabledStatus" -ForegroundColor $statusColor
        }
        
        Write-Host "`nHigh Risk ($($results | Where-Object { $_.RiskLevel -eq 'High' } | Measure-Object).Count):" -ForegroundColor Yellow
        $results | Where-Object { $_.RiskLevel -eq "High" } | ForEach-Object {
            $enabledStatus = if ($_.IsServicePrincipalEnabled) { "ENABLED" } else { "Disabled" }
            $statusColor = if ($_.IsServicePrincipalEnabled) { "Yellow" } else { "Green" }
            Write-Host "- $($_.ApplicationName) ($($_.ApplicationId)) - Status: " -NoNewline
            Write-Host "$enabledStatus" -ForegroundColor $statusColor
        }
        
        Write-Host "`nMedium & Low Risk ($($results | Where-Object { $_.RiskLevel -eq 'Medium' -or $_.RiskLevel -eq 'Low' } | Measure-Object).Count):" -ForegroundColor Green
        $results | Where-Object { $_.RiskLevel -eq "Medium" -or $_.RiskLevel -eq "Low" } | ForEach-Object {
            $enabledStatus = if ($_.IsServicePrincipalEnabled) { "ENABLED" } else { "Disabled" }
            Write-Host "- $($_.ApplicationName) ($($_.ApplicationId)) - Status: $enabledStatus"
        }
    }
    "3" {
        $jsonExportFile = "AppPermissionsDetailedAudit-$dateStr.json"
        # Export to JSON with more details for deeper analysis
        $results | ConvertTo-Json -Depth 4 | Out-File -FilePath $jsonExportFile
        Write-Host "Detailed application data exported to $jsonExportFile" -ForegroundColor Green
    }
    "4" {
        Write-Host "Exiting script." -ForegroundColor Green
    }
    default {
        Write-Host "Invalid selection." -ForegroundColor Red
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Green