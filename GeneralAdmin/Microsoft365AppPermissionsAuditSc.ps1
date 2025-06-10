# Microsoft 365 App Permissions Audit Script
# This script identifies applications with high-privilege permissions in your tenant
# and exports the results to Excel files and an HTML report

# Install required modules if not already installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph*)) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Repository PSGallery -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}

# Connect to Microsoft Graph with admin privileges
Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All" -NoWelcome
Write-Host "Connected to Microsoft Graph" -ForegroundColor Green

# Define high-privilege permissions to watch for with detailed descriptions
$highPrivilegePermissions = @(
    # Microsoft Graph Permissions
    @{
        Id="1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"; 
        Name="Application.ReadWrite.All"; 
        Description="Full access to manage applications";
        DetailedDescription="Allows the app to create, read, update and delete any app in your directory. Enables full control of app definitions and app assignments to users. Apps with this permission can update their own permissions to access your data.";
        RiskLevel="Critical";
        BusinessImpact="Can create rogue applications, modify other apps' permissions, and gain unauthorized access to company data."
    },
    @{
        Id="62a82d76-70ea-41e2-9197-370581804d09"; 
        Name="Group.ReadWrite.All"; 
        Description="Read and write all groups";
        DetailedDescription="Allows the app to create groups, read all group properties and memberships, update group properties and memberships, and delete groups. Also allows the app to read and write group calendar, conversations, files, and other group content.";
        RiskLevel="High";
        BusinessImpact="Can access all Teams content, add/remove members from security groups, read all files shared in all groups, and access all conversations."
    },
    @{
        Id="06da0dbc-49e2-44d2-8312-53f166ab848a"; 
        Name="Directory.ReadWrite.All"; 
        Description="Read and write directory data";
        DetailedDescription="Allows the app to read and write all data in your organization's directory, such as users, groups, and applications. The app can perform any user, group or application management operation.";
        RiskLevel="Critical";
        BusinessImpact="Effectively grants admin rights to the entire directory. Can manage all users, reset passwords, and modify security groups."
    },
    @{
        Id="c5366453-9fb0-48a5-a156-24f0c49a4b84"; 
        Name="User.ReadWrite.All"; 
        Description="Read and write all user profiles";
        DetailedDescription="Allows the app to read and modify the full profiles of all users in the organization including sensitive profile data like employee ID, manager chain, and office location.";
        RiskLevel="High";
        BusinessImpact="Can update user information, modify profile data, and potentially reset passwords for any user in the organization."
    },
    @{
        Id="741f803b-c850-494e-b5df-cde7c675a1ca"; 
        Name="Directory.Read.All"; 
        Description="Read directory data";
        DetailedDescription="Allows the app to read data in your organization's directory, such as users, groups and apps, but not to make changes.";
        RiskLevel="Medium";
        BusinessImpact="Access to all user information, group memberships, and org structure. Contains sensitive information about company personnel."
    },
    @{
        Id="df021288-bdef-4463-88db-98f22de89214"; 
        Name="Channel.ReadBasic.All"; 
        Description="Read all channel names and descriptions";
        DetailedDescription="Allows the app to read the names, descriptions, and membership list of all channels in all teams.";
        RiskLevel="Low";
        BusinessImpact="Can see all Teams channel names which might contain sensitive project information or confidential team identifiers."
    },
    @{
        Id="5b567255-7703-4780-807c-7be8301ae99b"; 
        Name="Group.Read.All"; 
        Description="Read all groups";
        DetailedDescription="Allows the app to read group properties and memberships, and read conversations and calendar events for all groups.";
        RiskLevel="Medium";
        BusinessImpact="Can see all security group memberships, distribution lists, and all Teams group content."
    },
    @{
        Id="7b2449af-6ccd-4f4d-9f78-e550c193f0d1"; 
        Name="Files.ReadWrite.All"; 
        Description="Read and write files in all site collections";
        DetailedDescription="Allows the app to read, create, update, and delete all files in all site collections, including OneDrive accounts, team sites, and SharePoint sites.";
        RiskLevel="High";
        BusinessImpact="Can access, modify, and delete ANY file stored in SharePoint/OneDrive across the entire organization. Data exfiltration risk."
    },
    @{
        Id="75359482-378d-4052-8f01-80520e7db3cd"; 
        Name="Files.Read.All"; 
        Description="Read files in all site collections";
        DetailedDescription="Allows the app to read all files in all site collections, including OneDrive accounts, team sites, and SharePoint sites.";
        RiskLevel="High";
        BusinessImpact="Can access ANY file stored in SharePoint/OneDrive across the entire organization, including confidential documents."
    },
    @{
        Id="a3410be2-8e48-4f32-8454-c29a7465209d"; 
        Name="Calendars.ReadWrite"; 
        Description="Read and write calendars";
        DetailedDescription="Allows the app to read events in all calendars that the user can access, including delegate and shared calendars. Also allows creation of new events.";
        RiskLevel="Medium";
        BusinessImpact="Can read sensitive meeting information and create calendar events that could appear to come from users."
    },
    @{
        Id="bf7b1a76-6e77-406b-b258-bf5c7720e98f"; 
        Name="Team.ReadBasic.All"; 
        Description="Read the names and descriptions of all teams";
        DetailedDescription="Allows the app to list and read basic properties of all teams the authenticated user can access.";
        RiskLevel="Low";
        BusinessImpact="Can see all team names which might contain confidential project information."
    },
    @{
        Id="dbaae8cf-10b5-4b86-a4a1-f871c94c6695"; 
        Name="TeamSettings.ReadWrite.All"; 
        Description="Read and change all teams' settings";
        DetailedDescription="Allows the app to read and modify all team's settings that the authenticated user can access.";
        RiskLevel="High";
        BusinessImpact="Can modify teams settings, add/remove channels, and change team-wide communication policies."
    },
    @{
        Id="82895abf-7a98-4f37-8519-1f0be6274800"; 
        Name="Sites.Selected"; 
        Description="Limited access to specific SharePoint sites";
        DetailedDescription="Allows the app to access only specific SharePoint sites that an admin explicitly approves. A more secure alternative to Files.Read.All or Files.ReadWrite.All when an app only needs access to specific sites.";
        RiskLevel="Low";
        BusinessImpact="Limited to accessing only specific sites that are explicitly approved by an admin. Recommended for most apps that need SharePoint access."
    }
)

# Initialize results array
$results = @()

# Get all service principals (apps) in the tenant
Write-Host "Retrieving all service principals..." -ForegroundColor Yellow
try {
    $servicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop
    if (-not $servicePrincipals) {
        throw "No service principals retrieved. Check Graph permissions or tenant data."
    }
} catch {
    Write-Error "Failed to retrieve service principals: $_"
    exit
}
$totalApps = $servicePrincipals.Count

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
            "AppOwnerOrganizationId" = $sp.AppOwnerOrganizationId
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
        # Check which specific sites this app has access to
        try {
            $sitePermissions = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/sitesSelected" -ErrorAction SilentlyContinue
            
            $appsWithSitesSelected += [PSCustomObject]@{
                "ApplicationName" = $sp.DisplayName
                "ApplicationId" = $sp.AppId
                "SitesWithAccess" = ($sitePermissions.value.site.displayName -join ", ")
                "SiteUrls" = ($sitePermissions.value.site.webUrl -join ", ")
                "PermissionType" = ($sitePermissions.value.allowedRoles -join ", ")
            }
        } catch {
            Write-Warning "Error fetching site permissions for $($sp.DisplayName): $_"
            
            $appsWithSitesSelected += [PSCustomObject]@{
                "ApplicationName" = $sp.DisplayName
                "ApplicationId" = $sp.AppId
                "SitesWithAccess" = "Error retrieving sites"
                "SiteUrls" = "Error retrieving sites"
                "PermissionType" = "Unknown"
            }
        }
    }
}

# Get the current date for the filename
$dateStr = Get-Date -Format "yyyyMMdd-HHmmss"

# Export results to formatted Excel file instead of CSV
$excelFile = "AppPermissionsAudit-$dateStr.xlsx"

# Create Excel package for main report
$excelPackage = $results | Export-Excel -Path $excelFile -WorksheetName "High Privilege Apps" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -PassThru

# Format the Excel file
$worksheet = $excelPackage.Workbook.Worksheets["High Privilege Apps"]

# Add conditional formatting for risk levels
$lastRow = $worksheet.Dimension.End.Row
$riskLevelColumn = $null

# Find the RiskLevel column
for ($col = 1; $col -le $worksheet.Dimension.End.Column; $col++) {
    if ($worksheet.Cells[1, $col].Value -eq "RiskLevel") {
        $riskLevelColumn = $col
        break
    }
}

if ($riskLevelColumn) {
    # Add conditional formatting for risk levels
    $criticalRule = Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "Critical" -BackgroundColor "#FFCCCC" -PassThru
    $highRule = Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "High" -BackgroundColor "#FFEBCC" -PassThru
    $mediumRule = Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "Medium" -BackgroundColor "#FFFFCC" -PassThru
    $lowRule = Add-ConditionalFormatting -WorkSheet $worksheet -Range "A2:Z$lastRow" -RuleType ContainsText -ConditionValue "Low" -BackgroundColor "#E6F5FF" -PassThru
    
    # Make the entire row colorized based on risk level
    $ruleRange = "A2:Z$lastRow"
    Add-ConditionalFormatting -WorkSheet $worksheet -Range $ruleRange -RuleType Expression -ConditionValue "`$`$`$$riskLevelColumn2=`"Critical`"" -BackgroundColor "#FFF0F0" -PassThru
    Add-ConditionalFormatting -WorkSheet $worksheet -Range $ruleRange -RuleType Expression -ConditionValue "`$`$`$$riskLevelColumn2=`"High`"" -BackgroundColor "#FFF5EB" -PassThru
    Add-ConditionalFormatting -WorkSheet $worksheet -Range $ruleRange -RuleType Expression -ConditionValue "`$`$`$$riskLevelColumn2=`"Medium`"" -BackgroundColor "#FFFFF0" -PassThru
    Add-ConditionalFormatting -WorkSheet $worksheet -Range $ruleRange -RuleType Expression -ConditionValue "`$`$`$$riskLevelColumn2=`"Low`"" -BackgroundColor "#F0F8FF" -PassThru
}

# Add legend to the worksheet
$legendRow = $lastRow + 2
$worksheet.Cells[$legendRow, 1].Value = "Risk Level Legend:"
$worksheet.Cells[$legendRow, 1].Style.Font.Bold = $true

$worksheet.Cells[$legendRow + 1, 1].Value = "Critical"
$worksheet.Cells[$legendRow + 1, 1].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$worksheet.Cells[$legendRow + 1, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 204, 204))

$worksheet.Cells[$legendRow + 2, 1].Value = "High"
$worksheet.Cells[$legendRow + 2, 1].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$worksheet.Cells[$legendRow + 2, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 204))

$worksheet.Cells[$legendRow + 3, 1].Value = "Medium"
$worksheet.Cells[$legendRow + 3, 1].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$worksheet.Cells[$legendRow + 3, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 255, 204))

$worksheet.Cells[$legendRow + 4, 1].Value = "Low"
$worksheet.Cells[$legendRow + 4, 1].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$worksheet.Cells[$legendRow + 4, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(230, 245, 255))

$worksheet.Cells[$legendRow, 2].Value = "Permission Impact and Recommended Actions:"
$worksheet.Cells[$legendRow, 2].Style.Font.Bold = $true

$worksheet.Cells[$legendRow + 1, 2].Value = "Critical: Immediate review required. These permissions could compromise your entire tenant."
$worksheet.Cells[$legendRow + 2, 2].Value = "High: Urgent review needed. These permissions provide extensive access to sensitive data."
$worksheet.Cells[$legendRow + 3, 2].Value = "Medium: Review recommended. These permissions provide significant data access."
$worksheet.Cells[$legendRow + 4, 2].Value = "Low: Monitor usage. These permissions have limited scope or impact."

# Add a permissions explanation worksheet
$permissionsData = $highPrivilegePermissions | Select-Object Name, RiskLevel, DetailedDescription, BusinessImpact | Sort-Object RiskLevel -Descending
$permissionsWorksheet = $permissionsData | Export-Excel -Path $excelFile -WorksheetName "Permission Explanations" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -PassThru

# Format the permissions worksheet
$permSheet = $permissionsWorksheet.Workbook.Worksheets["Permission Explanations"]
$permLastRow = $permSheet.Dimension.End.Row

# Add conditional formatting for permissions risk levels
Add-ConditionalFormatting -WorkSheet $permSheet -Range "A2:D$permLastRow" -RuleType ContainsText -ConditionValue "Critical" -BackgroundColor "#FFCCCC" -PassThru
Add-ConditionalFormatting -WorkSheet $permSheet -Range "A2:D$permLastRow" -RuleType ContainsText -ConditionValue "High" -BackgroundColor "#FFEBCC" -PassThru
Add-ConditionalFormatting -WorkSheet $permSheet -Range "A2:D$permLastRow" -RuleType ContainsText -ConditionValue "Medium" -BackgroundColor "#FFFFCC" -PassThru
Add-ConditionalFormatting -WorkSheet $permSheet -Range "A2:D$permLastRow" -RuleType ContainsText -ConditionValue "Low" -BackgroundColor "#E6F5FF" -PassThru

# Create a summary dashboard worksheet
$summarySheet = Add-Worksheet -ExcelPackage $permissionsWorksheet -WorksheetName "Dashboard" -Activate

# Set up the dashboard
$summarySheet.Cells["A1"].Value = "APPLICATION PERMISSIONS AUDIT DASHBOARD"
$summarySheet.Cells["A1"].Style.Font.Size = 16
$summarySheet.Cells["A1"].Style.Font.Bold = $true
$summarySheet.Cells["A1"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$summarySheet.Cells["A1:F1"].Merge = $true

$summarySheet.Cells["A3"].Value = "SUMMARY"
$summarySheet.Cells["A3"].Style.Font.Bold = $true
$summarySheet.Cells["A3"].Style.Font.Size = 12
$summarySheet.Cells["A3:F3"].Merge = $true
$summarySheet.Cells["A3:F3"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$summarySheet.Cells["A3:F3"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0, 120, 212))
$summarySheet.Cells["A3:F3"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$summarySheet.Cells["A4"].Value = "Total Applications:"
$summarySheet.Cells["B4"].Value = $totalApps
$summarySheet.Cells["A5"].Value = "Applications with High-Privilege Permissions:"
$summarySheet.Cells["B5"].Value = $results.Count
$summarySheet.Cells["A6"].Value = "Critical Risk Applications:"
$summarySheet.Cells["B6"].Value = ($results | Where-Object { $_.RiskLevel -eq 'Critical' } | Measure-Object).Count
$summarySheet.Cells["A7"].Value = "High Risk Applications:"
$summarySheet.Cells["B7"].Value = ($results | Where-Object { $_.RiskLevel -eq 'High' } | Measure-Object).Count
$summarySheet.Cells["A8"].Value = "Medium Risk Applications:"
$summarySheet.Cells["B8"].Value = ($results | Where-Object { $_.RiskLevel -eq 'Medium' } | Measure-Object).Count
$summarySheet.Cells["A9"].Value = "Low Risk Applications:"
$summarySheet.Cells["B9"].Value = ($results | Where-Object { $_.RiskLevel -eq 'Low' } | Measure-Object).Count

$summarySheet.Cells["A6:B6"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$summarySheet.Cells["A6:B6"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 204, 204))
$summarySheet.Cells["A7:B7"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$summarySheet.Cells["A7:B7"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 204))

$summarySheet.Cells["A11"].Value = "TOP RISKY APPLICATIONS"
$summarySheet.Cells["A11"].Style.Font.Bold = $true
$summarySheet.Cells["A11"].Style.Font.Size = 12
$summarySheet.Cells["A11:F11"].Merge = $true
$summarySheet.Cells["A11:F11"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$summarySheet.Cells["A11:F11"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0, 120, 212))
$summarySheet.Cells["A11:F11"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$summarySheet.Cells["A12"].Value = "Application Name"
$summarySheet.Cells["B12"].Value = "Publisher"
$summarySheet.Cells["C12"].Value = "Risk Level"
$summarySheet.Cells["D12"].Value = "Risk Score"
$summarySheet.Cells["E12"].Value = "Permissions"
$summarySheet.Cells["A12:E12"].Style.Font.Bold = $true

# Get top 10 risky apps by risk score
$topRiskyApps = $results | Sort-Object -Property RiskScore -Descending | Select-Object -First 10
$row = 13
foreach ($app in $topRiskyApps) {
    $summarySheet.Cells["A$row"].Value = $app.ApplicationName
    $summarySheet.Cells["B$row"].Value = $app.Publisher
    $summarySheet.Cells["C$row"].Value = $app.RiskLevel
    $summarySheet.Cells["D$row"].Value = $app.RiskScore
    $summarySheet.Cells["E$row"].Value = $app.HighPrivilegePermissions
    
    # Color the row based on risk level
    if ($app.RiskLevel -eq "Critical") {
        $summarySheet.Cells["A$row:E$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $summarySheet.Cells["A$row:E$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 204, 204))
    } elseif ($app.RiskLevel -eq "High") {
        $summarySheet.Cells["A$row:E$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $summarySheet.Cells["A$row:E$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 235, 204))
    }
    
    $row++
}

$summarySheet.Cells["A23"].Value = "SITES.SELECTED - BEST PRACTICE APPLICATIONS"
$summarySheet.Cells["A23"].Style.Font.Bold = $true
$summarySheet.Cells["A23"].Style.Font.Size = 12
$summarySheet.Cells["A23:E23"].Merge = $true
$summarySheet.Cells["A23:E23"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$summarySheet.Cells["A23:E23"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0, 120, 212))
$summarySheet.Cells["A23:E23"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$summarySheet.Cells["A24"].Value = "Application Name"
$summarySheet.Cells["B24"].Value = "Sites With Access"
$summarySheet.Cells["C24"].Value = "Permission Type"
$summarySheet.Cells["A24:C24"].Style.Font.Bold = $true

$row = 25
foreach ($app in $appsWithSitesSelected) {
    $summarySheet.Cells["A$row"].Value = $app.ApplicationName
    $summarySheet.Cells["B$row"].Value = $app.SitesWithAccess
    $summarySheet.Cells["C$row"].Value = $app.PermissionType
    $row++
}

$summarySheet.Cells["A35"].Value = "RECOMMENDATIONS"
$summarySheet.Cells["A35"].Style.Font.Bold = $true
$summarySheet.Cells["A35"].Style.Font.Size = 12
$summarySheet.Cells["A35:F35"].Merge = $true
$summarySheet.Cells["A35:F35"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$summarySheet.Cells["A35:F35"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(0, 120, 212))
$summarySheet.Cells["A35:F35"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$recommendations = @(
    "Review all critical and high-risk applications immediately",
    "Replace Files.Read.All with Sites.Selected where possible",
    "Document business justification for high-privilege permissions",
    "Implement regular permission audits as part of your security program",
    "Apply conditional access policies to applications with sensitive permissions",
    "Monitor application permission changes through alerts or regular audits",
    "Consider implementing Just-In-Time access for critical applications",
    "Train developers on permission best practices and security implications"
)

$row = 36
foreach ($rec in $recommendations) {
    $summarySheet.Cells["A$row"].Value = "- $rec"
    $summarySheet.Cells["A$row:F$row"].Merge = $true
    $row++
}

# Auto-size all columns
foreach ($ws in $permissionsWorksheet.Workbook.Worksheets) {
    $ws.Cells.AutoFitColumns()
}

# Save and close the Excel file
Close-ExcelPackage $permissionsWorksheet

# Export Sites.Selected apps to a separate Excel file
$sitesSelectedExcel = "SitesSelectedApps-$dateStr.xlsx"
$sitesSelectedPackage = $appsWithSitesSelected | Export-Excel -Path $sitesSelectedExcel -WorksheetName "Sites.Selected Apps" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter -PassThru
Close-ExcelPackage $sitesSelectedPackage

# Generate an educational HTML report
$htmlReport = "AppPermissionsEducationalReport-$dateStr.html"
$criticalApps = $results | Where-Object { $_.RiskLevel -eq "Critical" }
$highRiskApps = $results | Where-Object { $_.RiskLevel -eq "High" }
$otherApps = $results | Where-Object { $_.RiskLevel -ne "Critical" -and $_.RiskLevel -ne "High" }

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
    <p>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm") for $(Get-MgContext).TenantId</p>
    
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
            <th>Publisher</th>
            <th>Sites With Access</th>
            <th>Permission Type</th>
        </tr>
"@

    foreach ($app in $appsWithSitesSelected) {
        $htmlContent += @"
        <tr>
            <td>$($app.ApplicationName)</td>
            <td>$(($results | Where-Object { $_.ApplicationId -eq $app.ApplicationId } | Select-Object -First 1).Publisher)</td>
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
            <li>Use the downloadable CSV files for more detailed analysis</li>
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
Write-Host "Critical risk applications: $($results | Where-Object { $_.RiskLevel -eq 'Critical' } | Measure-Object).Count" -ForegroundColor Red
Write-Host "High risk applications: $($results | Where-Object { $_.RiskLevel -eq 'High' } | Measure-Object).Count" -ForegroundColor Yellow
Write-Host "Applications with Sites.Selected permissions: $($appsWithSitesSelected.Count)" -ForegroundColor Green

Write-Host "`nReports generated:" -ForegroundColor Cyan
Write-Host "- Excel Dashboard Report: $excelFile" -ForegroundColor White
Write-Host "- Sites.Selected Excel Report: $sitesSelectedExcel" -ForegroundColor White
Write-Host "- Educational HTML Report: $htmlReport" -ForegroundColor White
Write-Host "`nOpen the Excel Dashboard for an executive summary and the HTML report for detailed information." -ForegroundColor Yellow

# Provide remediation suggestions
Write-Host "`n=== REMEDIATION SUGGESTIONS ===" -ForegroundColor Cyan

# Check for critical applications
$criticalApps = $results | Where-Object { $_.RiskLevel -eq "Critical" }
if ($criticalApps.Count -gt 0) {
    Write-Host "`nCRITICAL: Found $($criticalApps.Count) applications with critical risk permissions" -ForegroundColor Red
    Write-Host "These applications have permissions that could potentially compromise your entire tenant:" -ForegroundColor Red
    $criticalApps | ForEach-Object {
        Write-Host "- $($_.ApplicationName) ($($_.ApplicationId))" -ForegroundColor Red
        Write-Host "  Permissions: $($_.HighPrivilegePermissions)" -ForegroundColor Red
    }
    
    Write-Host "`nRemediation steps for critical applications:" -ForegroundColor Yellow
    Write-Host "1. Review each critical application immediately" -ForegroundColor White
    Write-Host "2. Verify business need for these permissions" -ForegroundColor White
    Write-Host "3. Consider replacing with more limited permissions" -ForegroundColor White
    Write-Host "4. To modify permissions, use the Azure Portal or run:" -ForegroundColor White
    Write-Host "   Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId `$spId -AppRoleAssignmentId `$assignmentId" -ForegroundColor White
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
Write-Host "2. Developer Education: Share the Permission Explanations tab to help developers understand permission risks" -ForegroundColor White
Write-Host "3. Presentations: Use the Dashboard tab for executive presentations on app security" -ForegroundColor White
Write-Host "4. Compliance Documentation: Document your regular permission reviews" -ForegroundColor White
Write-Host "5. Risk Remediation: Prioritize fixing applications with critical and high-risk permissions" -ForegroundColor White

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Green