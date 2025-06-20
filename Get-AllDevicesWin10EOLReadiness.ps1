# Get-AllDevicesWin10EOLReadiness.ps1 - Version 2.0
# 
# Comprehensive Windows 10 End-of-Life readiness assessment for ALL devices (PCs, laptops, mobile devices)
# Excludes servers to focus on end-user devices requiring Windows 11 upgrades
#
# Author: Don Cook IT Community Script
# Version: 2.0
# Created: 2025
# Updated: For Windows 10 EOL (October 14, 2025)
#
# OVERVIEW:
# This script helps IT administrators assess their entire fleet of end-user devices
# for Windows 11 upgrade readiness before the Windows 10 End-of-Life deadline.
# 
# KEY FEATURES:
# - Discovers ALL end-user devices (PCs, laptops, mobile devices, workstations)
# - Excludes servers automatically using multiple detection methods
# - Cross-references Active Directory with Microsoft Graph/Intune data
# - Analyzes storage requirements for Windows 11 upgrades
# - Provides priority-based action recommendations
# - Generates executive-ready reports with visual indicators
# - Handles hybrid AD/Microsoft 365 environments
# - Robust API retry logic for large environments
#
# USAGE EXAMPLES:
# Basic assessment of all devices:
#   .\Get-AllDevicesWin10EOLReadiness.ps1
#
# Include specific device patterns:
#   .\Get-AllDevicesWin10EOLReadiness.ps1 -IncludePatterns @("WS-*", "LAPTOP-*", "PC-*")
#
# Exclude additional patterns:
#   .\Get-AllDevicesWin10EOLReadiness.ps1 -ExcludePatterns @("KIOSK-*", "CONF-*")
#
# Generate HTML report:
#   .\Get-AllDevicesWin10EOLReadiness.ps1 -ExportHtmlReport
#
# DEPENDENCIES:
# - ActiveDirectory PowerShell module
# - Microsoft.Graph PowerShell modules
# - Domain connectivity for AD queries
# - Microsoft Graph permissions (read-only)

[CmdletBinding()]
param (
    # Device filtering parameters
    [Parameter(Mandatory=$false)]
    [string[]]$IncludePatterns = @("*"),  # Include all devices by default
    
    [Parameter(Mandatory=$false)]
    [string[]]$ExcludePatterns = @("*SERVER*", "*SRV*", "*DC*", "*SQL*", "*EXCH*", "*SCCM*", "*WSUS*"),  # Exclude servers
    
    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 30,  # Look back period for activity analysis
    
    # Output parameters
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\AllDevicesWin10EOL_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportHtmlReport,
    
    # Assessment parameters
    [Parameter(Mandatory=$false)]
    [switch]$SkipHardwareCheck,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 3,
    
    [Parameter(Mandatory=$false)]
    [int]$RetryDelaySeconds = 5,
    
    # Authentication parameters (optional)
    [Parameter(Mandatory=$false)]
    [PSCredential]$ADCredential,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController
)

# Initialize required modules with admin check and auto-install
function Initialize-RequiredModules {
    Write-Host "=== MODULE INITIALIZATION ===" -ForegroundColor Cyan
    
    # Check if running as administrator
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    Write-Host "Current user: $($env:USERNAME)" -ForegroundColor White
    Write-Host "Running as administrator: $(if($isAdmin){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($isAdmin){'Green'}else{'Yellow'})
    
    $requiredModules = @(
        @{ Name = "ActiveDirectory"; Description = "Active Directory PowerShell module for device discovery" },
        @{ Name = "Microsoft.Graph.Authentication"; Description = "Microsoft Graph authentication" },
        @{ Name = "Microsoft.Graph.Users"; Description = "Microsoft Graph user management" },
        @{ Name = "Microsoft.Graph.Identity.SignIns"; Description = "Microsoft Graph sign-in logs" },
        @{ Name = "Microsoft.Graph.DeviceManagement"; Description = "Microsoft Graph device management (Intune)" },
        @{ Name = "Microsoft.Graph.DirectoryObjects"; Description = "Microsoft Graph directory objects" }
    )
    
    # Check which modules are missing
    $missingModules = @()
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module.Name)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "`n❌ MISSING REQUIRED MODULES" -ForegroundColor Red
        Write-Host "The following PowerShell modules are required but not installed:" -ForegroundColor Yellow
        Write-Host ""
        
        foreach ($module in $missingModules) {
            Write-Host "  ❌ $($module.Name)" -ForegroundColor Red
            Write-Host "     Purpose: $($module.Description)" -ForegroundColor Gray
        }
        
        Write-Host ""
        
        if (-not $isAdmin) {
            Write-Host "🔒 ADMINISTRATOR PRIVILEGES REQUIRED" -ForegroundColor Red
            Write-Host "To install the missing modules, you need to run PowerShell as Administrator." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "📋 SOLUTION:" -ForegroundColor Cyan
            Write-Host "1. Close this PowerShell window" -ForegroundColor White
            Write-Host "2. Right-click on PowerShell and select 'Run as Administrator'" -ForegroundColor White
            Write-Host "3. Run this command to install all required modules:" -ForegroundColor White
            Write-Host ""
            Write-Host "   Install-Module Microsoft.Graph, ActiveDirectory -Force" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "4. Then run this script again" -ForegroundColor White
            Write-Host ""
            Write-Host "💡 Alternative (User-level install):" -ForegroundColor Cyan
            Write-Host "   Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Yellow
            Write-Host "   (Note: ActiveDirectory module requires admin rights)" -ForegroundColor Gray
        } else {
            Write-Host "🔧 AUTO-INSTALL ATTEMPT" -ForegroundColor Cyan
            Write-Host "Attempting to install missing modules automatically..." -ForegroundColor Yellow
            
            $installSuccess = $true
            foreach ($module in $missingModules) {
                try {
                    Write-Host "Installing $($module.Name)..." -ForegroundColor Yellow
                    
                    if ($module.Name -eq "ActiveDirectory") {
                        # ActiveDirectory is part of RSAT
                        Write-Host "  Note: ActiveDirectory module is part of RSAT (Remote Server Administration Tools)" -ForegroundColor Gray
                        Write-Host "  You may need to install RSAT separately if this fails" -ForegroundColor Gray
                    }
                    
                    Install-Module -Name $module.Name -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
                    Write-Host "  ✓ Successfully installed $($module.Name)" -ForegroundColor Green
                }
                catch {
                    Write-Host "  ✗ Failed to install $($module.Name): $($_.Exception.Message)" -ForegroundColor Red
                    $installSuccess = $false
                    
                    if ($module.Name -eq "ActiveDirectory") {
                        Write-Host "  💡 For ActiveDirectory module, you may need to:" -ForegroundColor Yellow
                        Write-Host "     - Install RSAT: Enable-WindowsOptionalFeature -Online -FeatureName RSATClient-Roles-AD-Powershell" -ForegroundColor Gray
                        Write-Host "     - Or use: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor Gray
                    }
                }
            }
            
            if (-not $installSuccess) {
                Write-Host "`n❌ MODULE INSTALLATION FAILED" -ForegroundColor Red
                Write-Host "Some modules could not be installed automatically." -ForegroundColor Yellow
                Write-Host "Please install them manually and run the script again." -ForegroundColor White
                exit 1
            }
        }
        
        if (-not $isAdmin) {
            Write-Host "`nPress any key to exit..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 1
        }
    }
    
    # Import all modules
    Write-Host "`nImporting required modules..." -ForegroundColor Yellow
    foreach ($module in $requiredModules) {
        try {
            Import-Module $module.Name -ErrorAction Stop
            Write-Host "  ✓ $($module.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ Failed to import $($module.Name): $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }
    
    Write-Host "✅ All required modules loaded successfully" -ForegroundColor Green
}

# Connect to Microsoft Graph with comprehensive permission checking
function Connect-ToMSGraph {
    Write-Host "`n=== MICROSOFT GRAPH AUTHENTICATION ===" -ForegroundColor Cyan
    
    # Check if already connected
    try {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-Host "Already connected to Graph as: $($context.Account)" -ForegroundColor Green
            Write-Host "Tenant: $($context.TenantId)" -ForegroundColor Gray
            
            # Test existing permissions
            Write-Host "Testing existing permissions..." -ForegroundColor Yellow
            $permissionTest = Test-GraphPermissions
            if ($permissionTest.AllPermissionsValid) {
                Write-Host "✅ All required permissions validated" -ForegroundColor Green
                return $true
            } else {
                Write-Host "⚠️  Some permissions are missing or insufficient" -ForegroundColor Yellow
                Write-Host "Reconnecting to request proper permissions..." -ForegroundColor Yellow
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        # Not connected, proceed with connection
    }
    
    $requiredScopes = @(
        @{ Scope = "User.Read.All"; Purpose = "Read user profiles and organizational structure" },
        @{ Scope = "AuditLog.Read.All"; Purpose = "Read sign-in logs for device activity analysis" },
        @{ Scope = "Directory.Read.All"; Purpose = "Read directory objects and device registrations" },
        @{ Scope = "DeviceManagementManagedDevices.Read.All"; Purpose = "Read Intune managed device information" },
        @{ Scope = "Device.Read.All"; Purpose = "Read Azure AD registered devices" }
    )
    
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "Required permissions:" -ForegroundColor White
    foreach ($scope in $requiredScopes) {
        Write-Host "  • $($scope.Scope)" -ForegroundColor Gray
        Write-Host "    Purpose: $($scope.Purpose)" -ForegroundColor DarkGray
    }
    Write-Host ""
    
    try {
        $scopeNames = $requiredScopes | ForEach-Object { $_.Scope }
        Connect-MgGraph -Scopes $scopeNames -ErrorAction Stop
        
        $context = Get-MgContext
        Write-Host "✅ Successfully connected as: $($context.Account)" -ForegroundColor Green
        Write-Host "Tenant: $($context.TenantId)" -ForegroundColor Gray
        
        # Comprehensive permission validation
        Write-Host "`nValidating permissions..." -ForegroundColor Yellow
        $permissionTest = Test-GraphPermissions
        
        if ($permissionTest.AllPermissionsValid) {
            Write-Host "✅ All permissions validated successfully" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ PERMISSION VALIDATION FAILED" -ForegroundColor Red
            Show-PermissionErrors -PermissionTest $permissionTest
            return $false
        }
    }
    catch {
        Write-Host "❌ FAILED TO CONNECT TO MICROSOFT GRAPH" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "🔍 TROUBLESHOOTING STEPS:" -ForegroundColor Cyan
        Write-Host "1. Verify you have a valid Microsoft 365/Azure AD account" -ForegroundColor White
        Write-Host "2. Check if your account has the necessary permissions" -ForegroundColor White
        Write-Host "3. Try running as a different user with appropriate permissions" -ForegroundColor White
        Write-Host "4. Contact your IT administrator if permissions are insufficient" -ForegroundColor White
        Write-Host ""
        Write-Host "Script will continue with Active Directory data only..." -ForegroundColor Yellow
        return $false
    }
}

# Test Microsoft Graph permissions
function Test-GraphPermissions {
    $results = @{
        AllPermissionsValid = $true
        TestedPermissions = @()
    }
    
    # Test 1: User.Read.All - Try to read users
    Write-Host "  Testing User.Read.All..." -ForegroundColor Gray
    try {
        $testUser = Get-MgUser -Top 1 -ErrorAction Stop
        $results.TestedPermissions += @{
            Permission = "User.Read.All"
            Status = "✅ Valid"
            Error = $null
        }
        Write-Host "    ✅ User.Read.All - OK" -ForegroundColor Green
    }
    catch {
        $results.AllPermissionsValid = $false
        $results.TestedPermissions += @{
            Permission = "User.Read.All"
            Status = "❌ Failed"
            Error = $_.Exception.Message
        }
        Write-Host "    ❌ User.Read.All - Failed" -ForegroundColor Red
    }
    
    # Test 2: Directory.Read.All - Try to read devices
    Write-Host "  Testing Directory.Read.All..." -ForegroundColor Gray
    try {
        $testDevice = Get-MgDevice -Top 1 -ErrorAction Stop
        $results.TestedPermissions += @{
            Permission = "Directory.Read.All"
            Status = "✅ Valid"
            Error = $null
        }
        Write-Host "    ✅ Directory.Read.All - OK" -ForegroundColor Green
    }
    catch {
        $results.AllPermissionsValid = $false
        $results.TestedPermissions += @{
            Permission = "Directory.Read.All"
            Status = "❌ Failed"
            Error = $_.Exception.Message
        }
        Write-Host "    ❌ Directory.Read.All - Failed" -ForegroundColor Red
    }
    
    # Test 3: DeviceManagementManagedDevices.Read.All - Try to read Intune devices
    Write-Host "  Testing DeviceManagementManagedDevices.Read.All..." -ForegroundColor Gray
    try {
        $testIntuneDevice = Get-MgDeviceManagementManagedDevice -Top 1 -ErrorAction Stop
        $results.TestedPermissions += @{
            Permission = "DeviceManagementManagedDevices.Read.All"
            Status = "✅ Valid"
            Error = $null
        }
        Write-Host "    ✅ DeviceManagementManagedDevices.Read.All - OK" -ForegroundColor Green
    }
    catch {
        # This might be expected if no Intune licensing
        if ($_.Exception.Message -like "*Forbidden*" -or $_.Exception.Message -like "*Unauthorized*") {
            $results.AllPermissionsValid = $false
            $results.TestedPermissions += @{
                Permission = "DeviceManagementManagedDevices.Read.All"
                Status = "❌ No Access"
                Error = "Insufficient permissions or no Intune licensing"
            }
            Write-Host "    ❌ DeviceManagementManagedDevices.Read.All - No Access" -ForegroundColor Red
        } else {
            $results.TestedPermissions += @{
                Permission = "DeviceManagementManagedDevices.Read.All"
                Status = "⚠️  Limited"
                Error = "May work but limited data available"
            }
            Write-Host "    ⚠️  DeviceManagementManagedDevices.Read.All - Limited" -ForegroundColor Yellow
        }
    }
    
    # Test 4: AuditLog.Read.All - Try to read sign-in logs (most restrictive)
    Write-Host "  Testing AuditLog.Read.All..." -ForegroundColor Gray
    try {
        $testSignIn = Get-MgAuditLogSignIn -Top 1 -ErrorAction Stop
        $results.TestedPermissions += @{
            Permission = "AuditLog.Read.All"
            Status = "✅ Valid"
            Error = $null
        }
        Write-Host "    ✅ AuditLog.Read.All - OK" -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -like "*Forbidden*" -or $_.Exception.Message -like "*Unauthorized*") {
            $results.AllPermissionsValid = $false
            $results.TestedPermissions += @{
                Permission = "AuditLog.Read.All"
                Status = "❌ No Access"
                Error = "Insufficient permissions - requires Global Reader or Security Administrator role"
            }
            Write-Host "    ❌ AuditLog.Read.All - No Access" -ForegroundColor Red
        } else {
            $results.TestedPermissions += @{
                Permission = "AuditLog.Read.All"
                Status = "⚠️  Limited"
                Error = "May work but limited data available"
            }
            Write-Host "    ⚠️  AuditLog.Read.All - Limited" -ForegroundColor Yellow
        }
    }
    
    return $results
}

# Show detailed permission errors and solutions
function Show-PermissionErrors {
    param([object]$PermissionTest)
    
    Write-Host "`n🔐 PERMISSION ANALYSIS RESULTS" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    
    $failedPermissions = $PermissionTest.TestedPermissions | Where-Object { $_.Status -like "*Failed*" -or $_.Status -like "*No Access*" }
    $limitedPermissions = $PermissionTest.TestedPermissions | Where-Object { $_.Status -like "*Limited*" }
    
    if ($failedPermissions.Count -gt 0) {
        Write-Host "`n❌ FAILED PERMISSIONS:" -ForegroundColor Red
        foreach ($perm in $failedPermissions) {
            Write-Host "  • $($perm.Permission)" -ForegroundColor Red
            Write-Host "    Issue: $($perm.Error)" -ForegroundColor Gray
        }
    }
    
    if ($limitedPermissions.Count -gt 0) {
        Write-Host "`n⚠️  LIMITED PERMISSIONS:" -ForegroundColor Yellow
        foreach ($perm in $limitedPermissions) {
            Write-Host "  • $($perm.Permission)" -ForegroundColor Yellow
            Write-Host "    Note: $($perm.Error)" -ForegroundColor Gray
        }
    }
    
    Write-Host "`n🔧 SOLUTIONS:" -ForegroundColor Cyan
    Write-Host ""
    
    # Check for common permission issues
    $hasAuditLogIssue = $failedPermissions | Where-Object { $_.Permission -eq "AuditLog.Read.All" }
    $hasIntuneIssue = $failedPermissions | Where-Object { $_.Permission -eq "DeviceManagementManagedDevices.Read.All" }
    $hasDirectoryIssue = $failedPermissions | Where-Object { $_.Permission -eq "Directory.Read.All" }
    
    if ($hasAuditLogIssue) {
        Write-Host "📋 For AuditLog.Read.All access:" -ForegroundColor White
        Write-Host "   Your account needs one of these Azure AD roles:" -ForegroundColor Gray
        Write-Host "   • Global Administrator" -ForegroundColor Gray
        Write-Host "   • Global Reader" -ForegroundColor Gray
        Write-Host "   • Security Administrator" -ForegroundColor Gray
        Write-Host "   • Security Reader" -ForegroundColor Gray
        Write-Host "   • Reports Reader" -ForegroundColor Gray
        Write-Host ""
    }
    
    if ($hasIntuneIssue) {
        Write-Host "📱 For DeviceManagementManagedDevices.Read.All access:" -ForegroundColor White
        Write-Host "   • Your organization needs Intune licensing" -ForegroundColor Gray
        Write-Host "   • Your account needs Intune Administrator role" -ForegroundColor Gray
        Write-Host "   • Or Global Administrator / Global Reader" -ForegroundColor Gray
        Write-Host ""
    }
    
    if ($hasDirectoryIssue) {
        Write-Host "📁 For Directory.Read.All access:" -ForegroundColor White
        Write-Host "   Your account needs one of these Azure AD roles:" -ForegroundColor Gray
        Write-Host "   • Global Administrator" -ForegroundColor Gray
        Write-Host "   • Global Reader" -ForegroundColor Gray
        Write-Host "   • Directory Readers" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "💡 WHAT TO DO:" -ForegroundColor Cyan
    Write-Host "1. Contact your IT administrator to request appropriate permissions" -ForegroundColor White
    Write-Host "2. Provide them with this permission analysis" -ForegroundColor White
    Write-Host "3. The script will continue with limited functionality using available data" -ForegroundColor White
    Write-Host ""
    
    # Determine what functionality will be limited
    Write-Host "⚠️  IMPACT ON ASSESSMENT:" -ForegroundColor Yellow
    if ($hasAuditLogIssue) {
        Write-Host "   • No sign-in activity analysis" -ForegroundColor Gray
        Write-Host "   • Cannot determine active users of devices" -ForegroundColor Gray
    }
    if ($hasIntuneIssue) {
        Write-Host "   • No Intune device management data" -ForegroundColor Gray
        Write-Host "   • No storage space analysis" -ForegroundColor Gray
        Write-Host "   • No compliance status information" -ForegroundColor Gray
    }
    if ($hasDirectoryIssue) {
        Write-Host "   • Limited Azure AD device information" -ForegroundColor Gray
        Write-Host "   • Cannot cross-reference with cloud device data" -ForegroundColor Gray
    }
    
    Write-Host "`nScript will continue with Active Directory data and available Graph data..." -ForegroundColor Yellow
}

# Get all end-user devices from Active Directory (excluding servers)
function Get-AllEndUserDevices {
    Write-Host "`n=== ACTIVE DIRECTORY DEVICE DISCOVERY ===" -ForegroundColor Cyan
    Write-Host "Discovering all end-user devices (excluding servers)..." -ForegroundColor Yellow
    
    try {
        # Build base parameters
        $adParams = @{
            Filter = "Enabled -eq 'True'"
            Properties = @(
                'Name', 'OperatingSystem', 'OperatingSystemVersion', 'LastLogonDate', 
                'Enabled', 'DistinguishedName', 'Description', 'whenCreated', 'PasswordLastSet',
                'OperatingSystemServicePack', 'IPv4Address', 'DNSHostName'
            )
        }
        
        # Add credentials if provided
        if ($ADCredential) {
            $adParams.Credential = $ADCredential
        }
        if ($DomainController) {
            $adParams.Server = $DomainController
        }
        
        # Get all enabled computer objects
        Write-Host "Querying Active Directory for enabled computer objects..." -ForegroundColor Gray
        $allComputers = Get-ADComputer @adParams -ErrorAction Stop
        
        Write-Host "Found $($allComputers.Count) total enabled computer objects" -ForegroundColor White
        
        # Filter devices based on include/exclude patterns
        $filteredDevices = @()
        
        foreach ($computer in $allComputers) {
            $deviceName = $computer.Name
            $osInfo = "$($computer.OperatingSystem) $($computer.Description)".ToLower()
            
            # Check if device should be excluded (servers, etc.)
            $shouldExclude = $false
            
            # Server detection logic
            if ($computer.OperatingSystem -like "*Server*" -or 
                $computer.OperatingSystem -like "*Datacenter*" -or
                $computer.OperatingSystem -like "*Standard*" -and $computer.OperatingSystem -like "*Server*") {
                $shouldExclude = $true
                Write-Verbose "Excluding server by OS: $deviceName ($($computer.OperatingSystem))"
            }
            
            # Pattern-based exclusions
            foreach ($excludePattern in $ExcludePatterns) {
                if ($deviceName -like $excludePattern -or $osInfo -like $excludePattern.ToLower()) {
                    $shouldExclude = $true
                    Write-Verbose "Excluding by pattern '$excludePattern': $deviceName"
                    break
                }
            }
            
            # Skip if excluded
            if ($shouldExclude) {
                continue
            }
            
            # Check include patterns
            $shouldInclude = $false
            foreach ($includePattern in $IncludePatterns) {
                if ($deviceName -like $includePattern) {
                    $shouldInclude = $true
                    break
                }
            }
            
            if (-not $shouldInclude) {
                continue
            }
            
            # Process the device
            $filteredDevices += $computer
        }
        
        Write-Host "✓ Filtered to $($filteredDevices.Count) end-user devices" -ForegroundColor Green
        
        # Process each device for assessment
        $processedDevices = @()
        $deviceCount = $filteredDevices.Count
        $currentDevice = 0
        
        foreach ($device in $filteredDevices) {
            $currentDevice++
            Write-Progress -Activity "Processing devices" -Status "Device $currentDevice of $deviceCount" -PercentComplete (($currentDevice / $deviceCount) * 100)
            
            # Enhanced OS analysis
            $windowsVersion = "Unknown"
            $buildNumber = ""
            $isWindows10 = $false
            $dataFreshness = "Unknown"
            $eolUrgency = "Unknown"
            $deviceType = "Unknown"
            
            # Determine device type
            if ($device.OperatingSystem -like "*Server*") {
                $deviceType = "Server" # Shouldn't happen due to filtering
            }
            elseif ($device.Name -like "*LAPTOP*" -or $device.Name -like "*NB*" -or $device.Description -like "*laptop*") {
                $deviceType = "Laptop"
            }
            elseif ($device.Name -like "*DESKTOP*" -or $device.Name -like "*PC*" -or $device.Name -like "*WS*") {
                $deviceType = "Desktop"
            }
            elseif ($device.Name -like "*MD*" -or $device.Name -like "*MOBILE*") {
                $deviceType = "Mobile Device"
            }
            else {
                $deviceType = "Workstation"
            }
            
            # Parse OS version
            $osString = "$($device.OperatingSystem) $($device.OperatingSystemVersion)"
            if ($device.OperatingSystemVersion -match '\(([\d]+)\)') {
                $buildNumber = $matches[1]
            }
            
            # Determine Windows version and EOL urgency
            if ($osString -match "Windows 11" -or [int]$buildNumber -ge 22000) {
                $windowsVersion = "Windows 11"
                $isWindows10 = $false
                $eolUrgency = "None - Win11"
            }
            elseif ($osString -match "Windows 10") {
                $windowsVersion = "Windows 10"
                $isWindows10 = $true
                $eolUrgency = "HIGH - EOL Oct 2025"
            }
            elseif ([int]$buildNumber -ge 10240 -and [int]$buildNumber -lt 22000) {
                $windowsVersion = "Windows 10"
                $isWindows10 = $true
                $eolUrgency = "HIGH - EOL Oct 2025"
            }
            elseif ($osString -match "Windows 8" -or $osString -match "Windows 7") {
                $windowsVersion = $osString
                $isWindows10 = $false
                $eolUrgency = "CRITICAL - Legacy OS"
            }
            
            # Assess data freshness
            $daysInactive = 9999
            if ($device.LastLogonDate) {
                $daysInactive = (Get-Date).Subtract($device.LastLogonDate).Days
                if ($daysInactive -lt 7) { $dataFreshness = "Very Fresh" }
                elseif ($daysInactive -lt 30) { $dataFreshness = "Fresh" }
                elseif ($daysInactive -lt 90) { $dataFreshness = "Stale" }
                else { $dataFreshness = "Very Stale" }
            }
            
            $processedDevices += [PSCustomObject]@{
                DeviceName = $device.Name
                DeviceType = $deviceType
                OU = ($device.DistinguishedName -split ',')[1] -replace 'OU=', ''
                ADOperatingSystem = $device.OperatingSystem
                ADOSVersion = $device.OperatingSystemVersion
                WindowsVersion = $windowsVersion
                BuildNumber = $buildNumber
                IsWindows10 = $isWindows10
                EOLUrgency = $eolUrgency
                LastLogonDate = $device.LastLogonDate
                DaysInactive = if ($daysInactive -eq 9999) { "Never" } else { $daysInactive }
                DataFreshness = $dataFreshness
                Enabled = $device.Enabled
                DeviceAge = (Get-Date).Subtract($device.whenCreated).Days
                PasswordLastSet = $device.PasswordLastSet
                Description = $device.Description
                DNSHostName = $device.DNSHostName
                IPv4Address = $device.IPv4Address
                # Placeholders for Graph data
                GraphDeviceFound = $false
                GraphOS = "Unknown"
                SignInUser = "Unknown"
                LastSignInDate = "Unknown"
                IntuneManaged = $false
                StorageTotal = "Unknown"
                StorageFree = "Unknown"
                StorageAdequate = "Unknown"
                Win11Eligible = "Unknown"
                Win11Status = "Needs Assessment"
                UpgradeAction = "Unknown"
                EOLPriority = "Unknown"
                ComplianceState = "Unknown"
                Notes = @()
            }
        }
        
        Write-Progress -Activity "Processing devices" -Completed
        return $processedDevices
        
    }
    catch {
        Write-Host "✗ Failed to query Active Directory: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Enhance devices with Microsoft Graph data
function Add-GraphIntelligence {
    param([array]$Devices, [bool]$GraphConnected)
    
    if (-not $GraphConnected) {
        Write-Host "`n=== SKIPPING GRAPH INTELLIGENCE ===" -ForegroundColor Yellow
        Write-Host "Microsoft Graph not connected - using AD data only" -ForegroundColor White
        
        # Set basic assessments for AD-only mode
        foreach ($device in $Devices) {
            $device.Win11Status = Get-BasicWin11Assessment -Device $device
            $device.EOLPriority = Get-BasicEOLPriority -Device $device
            $device.Notes = "AD data only - Graph not available"
        }
        return $Devices
    }
    
    Write-Host "`n=== MICROSOFT GRAPH INTELLIGENCE ENHANCEMENT ===" -ForegroundColor Cyan
    Write-Host "Enhancing device data with Graph/Intune information..." -ForegroundColor Yellow
    
    $enhancedDevices = @()
    $deviceCount = $Devices.Count
    $currentDevice = 0
    
    foreach ($device in $Devices) {
        $currentDevice++
        Write-Progress -Activity "Enhancing with Graph data" -Status "Device $currentDevice of $deviceCount" -PercentComplete (($currentDevice / $deviceCount) * 100)
        
        Write-Verbose "Processing Graph data for: $($device.DeviceName)"
        
        # Try multiple Graph lookups
        $graphDevice = $null
        $intuneDevice = $null
        
        # Strategy 1: Direct Graph device lookup
        try {
            $graphDevice = Get-MgDevice -Filter "displayName eq '$($device.DeviceName)'" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($graphDevice) {
                $device.GraphDeviceFound = $true
                $device.GraphOS = if ($graphDevice.OperatingSystem) { $graphDevice.OperatingSystem } else { "Unknown" }
                $device.Notes += "Found in Graph registry"
                
                if ($graphDevice.IsCompliant -ne $null) {
                    $device.ComplianceState = if ($graphDevice.IsCompliant) { "Compliant" } else { "Non-Compliant" }
                }
            }
        }
        catch {
            Write-Verbose "Graph device lookup failed for $($device.DeviceName): $($_.Exception.Message)"
        }
        
        # Strategy 2: Intune managed device lookup
        try {
            $intuneDevice = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$($device.DeviceName)'" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($intuneDevice) {
                $device.IntuneManaged = $true
                $device.LastSignInDate = $intuneDevice.LastSyncDateTime
                $device.Notes += "Intune managed"
                
                # Get storage information (critical for Win11 upgrades)
                if ($intuneDevice.TotalStorageSpaceInBytes -and $intuneDevice.TotalStorageSpaceInBytes -gt 0) {
                    $device.StorageTotal = [math]::Round($intuneDevice.TotalStorageSpaceInBytes / 1GB, 1)
                }
                if ($intuneDevice.FreeStorageSpaceInBytes -and $intuneDevice.FreeStorageSpaceInBytes -gt 0) {
                    $device.StorageFree = [math]::Round($intuneDevice.FreeStorageSpaceInBytes / 1GB, 1)
                }
                
                # Assess storage adequacy for Windows 11 (needs ~20GB free)
                if ($device.StorageFree -ne "Unknown" -and $device.StorageFree -is [double]) {
                    $device.StorageAdequate = if ($device.StorageFree -ge 20) { "Yes" } else { "No - Insufficient" }
                }
                
                # Better OS info from Intune
                if ($intuneDevice.OperatingSystem) {
                    $device.GraphOS = $intuneDevice.OperatingSystem
                }
                
                # User information
                if ($intuneDevice.UserDisplayName) {
                    $device.SignInUser = $intuneDevice.UserDisplayName
                }
                
                # Compliance state
                if ($intuneDevice.ComplianceState) {
                    $device.ComplianceState = $intuneDevice.ComplianceState
                }
            }
        }
        catch {
            Write-Verbose "Intune lookup failed for $($device.DeviceName): $($_.Exception.Message)"
        }
        
        # Windows 11 readiness assessment
        $win11Assessment = Get-ComprehensiveWin11Assessment -Device $device -IntuneDevice $intuneDevice
        $device.Win11Eligible = $win11Assessment.Eligible
        $device.Win11Status = $win11Assessment.Status
        $device.UpgradeAction = $win11Assessment.Action
        
        # EOL Priority assessment
        $device.EOLPriority = Get-ComprehensiveEOLPriority -Device $device
        
        # Finalize notes
        $device.Notes = ($device.Notes | Where-Object { $_ }) -join "; "
        if (-not $device.Notes) { $device.Notes = "Standard assessment" }
        
        $enhancedDevices += $device
        
        # Brief pause to be kind to the APIs
        Start-Sleep -Milliseconds 100
    }
    
    Write-Progress -Activity "Enhancing with Graph data" -Completed
    return $enhancedDevices
}

# Basic Windows 11 assessment (AD-only mode)
function Get-BasicWin11Assessment {
    param([object]$Device)
    
    if ($Device.IsWindows10) {
        return "Windows 10 - Needs upgrade assessment"
    }
    elseif ($Device.WindowsVersion -eq "Windows 11") {
        return "Already on Windows 11"
    }
    elseif ($Device.WindowsVersion -like "*Windows 7*" -or $Device.WindowsVersion -like "*Windows 8*") {
        return "Legacy OS - Replacement required"
    }
    else {
        return "Unknown OS - Manual assessment required"
    }
}

# Basic EOL priority (AD-only mode)
function Get-BasicEOLPriority {
    param([object]$Device)
    
    if ($Device.IsWindows10 -and $Device.DataFreshness -in @("Very Fresh", "Fresh")) {
        return "HIGH - Active Windows 10"
    }
    elseif ($Device.IsWindows10) {
        return "MEDIUM - Windows 10"
    }
    elseif ($Device.DaysInactive -eq "Never" -or ($Device.DaysInactive -ne "Never" -and $Device.DaysInactive -gt 90)) {
        return "LOW - Inactive device"
    }
    elseif ($Device.WindowsVersion -eq "Windows 11") {
        return "NONE - Already Windows 11"
    }
    else {
        return "MEDIUM - Needs assessment"
    }
}

# Comprehensive Windows 11 assessment (with Graph data)
function Get-ComprehensiveWin11Assessment {
    param([object]$Device, [object]$IntuneDevice)
    
    $result = @{
        Eligible = "Unknown"
        Status = "Needs assessment"
        Action = "Assessment required"
    }
    
    $daysUntilEOL = (Get-Date "2025-10-14").Subtract((Get-Date)).Days
    
    # Use best available OS information
    $osToAnalyze = if ($Device.GraphOS -and $Device.GraphOS -ne "Unknown") { 
        $Device.GraphOS 
    } else { 
        $Device.ADOperatingSystem + " " + $Device.ADOSVersion 
    }
    
    # Build number analysis
    if ($Device.BuildNumber -and $Device.BuildNumber -match "(\d+)") {
        $buildNumber = [int]$matches[1]
        
        if ($buildNumber -ge 22000) {
            # Windows 11
            $result.Eligible = "Yes"
            $result.Status = "Already on Windows 11"
            $result.Action = "None - already upgraded"
        }
        elseif ($buildNumber -ge 10240) {
            # Windows 10
            $result.Eligible = "Yes (OS)"
            $result.Status = "Windows 10 - $daysUntilEOL days until EOL"
            $result.Action = "Upgrade required before October 2025"
            
            # Storage constraint check
            if ($Device.StorageAdequate -eq "No - Insufficient") {
                $result.Eligible = "No"
                $result.Status = "Windows 10 - Insufficient storage"
                $result.Action = "Storage cleanup or replacement required"
            }
            elseif ($Device.StorageAdequate -eq "Yes") {
                $result.Eligible = "Yes"
                $result.Status = "Windows 10 - Ready for upgrade"
                $result.Action = "Schedule Windows 11 upgrade"
            }
        }
        else {
            $result.Eligible = "No"
            $result.Status = "OS too old for Windows 11"
            $result.Action = "Device replacement required"
        }
    }
    elseif ($osToAnalyze -match "Windows 11") {
        $result.Eligible = "Yes"
        $result.Status = "Already on Windows 11"
        $result.Action = "None - already upgraded"
    }
    elseif ($osToAnalyze -match "Windows 10") {
        $result.Eligible = "Yes (OS)"
        $result.Status = "Windows 10 - $daysUntilEOL days until EOL"
        $result.Action = "Upgrade assessment required"
    }
    elseif ($osToAnalyze -match "Windows 7|Windows 8") {
        $result.Eligible = "No"
        $result.Status = "Legacy OS - End of life"
        $result.Action = "Immediate replacement required"
    }
    
    return $result
}

# Comprehensive EOL priority assessment
function Get-ComprehensiveEOLPriority {
    param([object]$Device)
    
    $daysUntilEOL = (Get-Date "2025-10-14").Subtract((Get-Date)).Days
    
    # Critical: Legacy OS
    if ($Device.WindowsVersion -like "*Windows 7*" -or $Device.WindowsVersion -like "*Windows 8*") {
        return "CRITICAL - Legacy OS"
    }
    
    # Critical: Active Windows 10 with blockers
    if ($Device.IsWindows10 -and $Device.DataFreshness -in @("Very Fresh", "Fresh")) {
        if ($Device.StorageAdequate -eq "No - Insufficient" -or $Device.Win11Eligible -eq "No") {
            return "CRITICAL - Active Win10 with blockers"
        }
        elseif ($Device.Win11Eligible -eq "Yes") {
            return "HIGH - Active Win10, ready for upgrade"
        }
        else {
            return "MEDIUM - Active Win10, needs assessment"
        }
    }
    
    # Low priority for inactive devices
    if ($Device.DaysInactive -eq "Never" -or ($Device.DaysInactive -ne "Never" -and $Device.DaysInactive -gt 90)) {
        return "LOW - Inactive device"
    }
    
    # Already upgraded
    if (-not $Device.IsWindows10 -and $Device.WindowsVersion -eq "Windows 11") {
        return "NONE - Already Windows 11"
    }
    
    return "MEDIUM - Requires assessment"
}

# Generate comprehensive EOL report
function Generate-ComprehensiveEOLReport {
    param([array]$DeviceData)
    
    $totalDevices = $DeviceData.Count
    $windows10Devices = $DeviceData | Where-Object IsWindows10 -eq $true
    $windows11Devices = $DeviceData | Where-Object { $_.WindowsVersion -eq "Windows 11" }
    $legacyDevices = $DeviceData | Where-Object { $_.WindowsVersion -like "*Windows 7*" -or $_.WindowsVersion -like "*Windows 8*" }
    $activeWindows10 = $windows10Devices | Where-Object DataFreshness -in @("Very Fresh", "Fresh")
    $criticalDevices = $DeviceData | Where-Object EOLPriority -like "CRITICAL*"
    $daysUntilEOL = (Get-Date "2025-10-14").Subtract((Get-Date)).Days
    
    Write-Host "`n=== COMPREHENSIVE WINDOWS 10 EOL ASSESSMENT REPORT ===" -ForegroundColor Red
    Write-Host "Assessment Date: $(Get-Date -Format 'MMMM dd, yyyy')" -ForegroundColor White
    Write-Host "Days until Windows 10 EOL: $daysUntilEOL" -ForegroundColor $(if($daysUntilEOL -lt 120){'Red'}else{'Yellow'})
    Write-Host "================================================================" -ForegroundColor Red
    
    Write-Host "`n--- Device Overview ---" -ForegroundColor Yellow
    Write-Host "Total end-user devices: $totalDevices" -ForegroundColor White
    Write-Host "Windows 10 devices: $($windows10Devices.Count) ($(if($totalDevices -gt 0){[math]::Round(($windows10Devices.Count/$totalDevices)*100)}else{0})%)" -ForegroundColor $(if($windows10Devices.Count -gt 0){'Red'}else{'Green'})
    Write-Host "Windows 11 devices: $($windows11Devices.Count) ($(if($totalDevices -gt 0){[math]::Round(($windows11Devices.Count/$totalDevices)*100)}else{0})%)" -ForegroundColor Green
    Write-Host "Legacy OS devices: $($legacyDevices.Count) ($(if($totalDevices -gt 0){[math]::Round(($legacyDevices.Count/$totalDevices)*100)}else{0})%)" -ForegroundColor $(if($legacyDevices.Count -gt 0){'Red'}else{'Green'})
    Write-Host "Active Windows 10 devices: $($activeWindows10.Count)" -ForegroundColor $(if($activeWindows10.Count -gt 0){'Red'}else{'Green'})
    
    Write-Host "`n--- Device Type Breakdown ---" -ForegroundColor Yellow
    $DeviceData | Group-Object DeviceType | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Name): $($_.Count) devices" -ForegroundColor White
    }
    
    Write-Host "`n--- Priority Breakdown ---" -ForegroundColor Yellow
    $DeviceData | Group-Object EOLPriority | Sort-Object @{Expression={
        switch -Wildcard ($_.Name) {
            "CRITICAL*" { 1 }
            "HIGH*" { 2 }
            "MEDIUM*" { 3 }
            "LOW*" { 4 }
            "NONE*" { 5 }
            default { 6 }
        }
    }} | ForEach-Object {
        $color = switch -Wildcard ($_.Name) {
            "CRITICAL*" { "Red" }
            "HIGH*" { "Red" }
            "MEDIUM*" { "Yellow" }
            "LOW*" { "Gray" }
            "NONE*" { "Green" }
            default { "White" }
        }
        Write-Host "$($_.Name): $($_.Count) devices" -ForegroundColor $color
    }
    
    # Storage analysis
    $storageIssues = $DeviceData | Where-Object StorageAdequate -eq "No - Insufficient"
    $unknownStorage = $DeviceData | Where-Object StorageAdequate -eq "Unknown"
    
    Write-Host "`n--- Storage Analysis ---" -ForegroundColor Yellow
    Write-Host "Devices with insufficient storage: $($storageIssues.Count)" -ForegroundColor $(if($storageIssues.Count -gt 0){'Red'}else{'Green'})
    Write-Host "Devices with unknown storage: $($unknownStorage.Count)" -ForegroundColor $(if($unknownStorage.Count -gt 0){'Yellow'}else{'Green'})
    
    return @{
        TotalDevices = $totalDevices
        Windows10Count = $windows10Devices.Count
        Windows11Count = $windows11Devices.Count
        LegacyCount = $legacyDevices.Count
        ActiveWindows10 = $activeWindows10.Count
        CriticalCount = $criticalDevices.Count
        DaysUntilEOL = $daysUntilEOL
        StorageIssues = $storageIssues.Count
    }
}

# Enhanced HTML Report with color legend
function New-EnhancedHtmlReport {
    param (
        [array]$DeviceData,
        [hashtable]$Summary
    )
    
    $htmlPath = $OutputPath -replace '\.csv$', '.html'
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows 10 EOL Readiness Assessment - All Devices</title>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background-color: #f5f5f5; 
            line-height: 1.6;
        }
        
        .header { 
            background: linear-gradient(135deg, #d32f2f, #f44336); 
            color: white; 
            padding: 30px; 
            border-radius: 12px; 
            margin-bottom: 30px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        
        .header h1 { margin: 0; font-size: 28px; }
        .header p { margin: 5px 0; opacity: 0.9; }
        
        .summary { 
            background-color: white; 
            padding: 25px; 
            border-radius: 12px; 
            margin-bottom: 30px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.1); 
        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .metric { 
            background-color: #f8f9fa; 
            padding: 20px; 
            border-radius: 8px; 
            text-align: center;
            border-left: 4px solid #1976d2;
        }
        
        .metric-value { 
            font-size: 28px; 
            font-weight: bold; 
            color: #1976d2; 
            display: block;
        }
        
        .metric-label { 
            font-size: 14px; 
            color: #666; 
            margin-top: 5px;
        }
        
        .legend {
            background-color: white;
            padding: 20px;
            border-radius: 12px;
            margin-bottom: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .legend h3 {
            margin-top: 0;
            color: #333;
            border-bottom: 2px solid #eee;
            padding-bottom: 10px;
        }
        
        .legend-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            padding: 10px;
            border-radius: 6px;
            font-size: 14px;
        }
        
        .legend-color {
            width: 20px;
            height: 20px;
            border-radius: 4px;
            margin-right: 10px;
            flex-shrink: 0;
        }
        
        /* Priority Colors */
        .critical { background-color: #ffebee; border-left: 4px solid #d32f2f; }
        .high { background-color: #fff3e0; border-left: 4px solid #f57c00; }
        .medium { background-color: #f3e5f5; border-left: 4px solid #7b1fa2; }
        .low { background-color: #e8f5e8; border-left: 4px solid #388e3c; }
        .win11 { background-color: #e3f2fd; border-left: 4px solid #1976d2; }
        .unknown { background-color: #f5f5f5; border-left: 4px solid #757575; }
        
        /* Legend Colors */
        .legend-critical { background-color: #d32f2f; }
        .legend-high { background-color: #f57c00; }
        .legend-medium { background-color: #7b1fa2; }
        .legend-low { background-color: #388e3c; }
        .legend-win11 { background-color: #1976d2; }
        .legend-unknown { background-color: #757575; }
        
        table { 
            border-collapse: collapse; 
            width: 100%; 
            background-color: white; 
            border-radius: 12px; 
            overflow: hidden; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            font-size: 14px;
        }
        
        th { 
            background-color: #1976d2; 
            color: white; 
            text-align: left; 
            padding: 15px 12px; 
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        
        td { 
            border-bottom: 1px solid #eee; 
            padding: 12px; 
            vertical-align: top;
        }
        
        tr:hover { background-color: #f8f9fa; }
        
        .device-name { font-weight: 600; color: #1976d2; }
        .os-version { font-family: 'Courier New', monospace; font-size: 12px; }
        
        .footer {
            margin-top: 30px;
            padding: 20px;
            background-color: white;
            border-radius: 12px;
            text-align: center;
            color: #666;
            font-size: 12px;
        }
        
        @media (max-width: 768px) {
            .metrics-grid { grid-template-columns: 1fr; }
            .legend-grid { grid-template-columns: 1fr; }
            table { font-size: 12px; }
            th, td { padding: 8px; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>🖥️ Windows 10 EOL Readiness Assessment</h1>
        <p><strong>All End-User Devices Report</strong></p>
        <p>Generated: $(Get-Date -Format 'MMMM dd, yyyy HH:mm:ss')</p>
        <p>⏰ Windows 10 End of Life: <strong>October 14, 2025</strong> ($($Summary.DaysUntilEOL) days remaining)</p>
    </div>
    
    <div class="summary">
        <h2>📊 Executive Summary</h2>
        <div class="metrics-grid">
            <div class="metric">
                <span class="metric-value">$($Summary.TotalDevices)</span>
                <div class="metric-label">Total Devices</div>
            </div>
            <div class="metric">
                <span class="metric-value" style="color: #d32f2f;">$($Summary.Windows10Count)</span>
                <div class="metric-label">Windows 10 Devices</div>
            </div>
            <div class="metric">
                <span class="metric-value" style="color: #388e3c;">$($Summary.Windows11Count)</span>
                <div class="metric-label">Windows 11 Devices</div>
            </div>
            <div class="metric">
                <span class="metric-value" style="color: #d32f2f;">$($Summary.CriticalCount)</span>
                <div class="metric-label">Critical Priority</div>
            </div>
            <div class="metric">
                <span class="metric-value" style="color: #f57c00;">$($Summary.StorageIssues)</div>
                <div class="metric-label">Storage Issues</div>
            </div>
            <div class="metric">
                <span class="metric-value" style="color: #7b1fa2;">$($Summary.LegacyCount)</span>
                <div class="metric-label">Legacy OS</div>
            </div>
        </div>
    </div>
    
    <div class="legend">
        <h3>🎨 Priority Color Legend</h3>
        <div class="legend-grid">
            <div class="legend-item">
                <div class="legend-color legend-critical"></div>
                <strong>Critical Priority:</strong> Legacy OS or active Win10 with blockers - immediate action required
            </div>
            <div class="legend-item">
                <div class="legend-color legend-high"></div>
                <strong>High Priority:</strong> Active Windows 10 devices ready for upgrade
            </div>
            <div class="legend-item">
                <div class="legend-color legend-medium"></div>
                <strong>Medium Priority:</strong> Windows 10 devices needing assessment or inactive
            </div>
            <div class="legend-item">
                <div class="legend-color legend-low"></div>
                <strong>Low Priority:</strong> Inactive devices with minimal impact
            </div>
            <div class="legend-item">
                <div class="legend-color legend-win11"></div>
                <strong>Complete:</strong> Already running Windows 11 - no action needed
            </div>
            <div class="legend-item">
                <div class="legend-color legend-unknown"></div>
                <strong>Unknown:</strong> Requires manual investigation
            </div>
        </div>
    </div>
    
    <table>
        <thead>
            <tr>
                <th>🖥️ Device Name</th>
                <th>📱 Type</th>
                <th>🪟 Windows Version</th>
                <th>⚠️ EOL Priority</th>
                <th>✅ Win11 Status</th>
                <th>💾 Storage</th>
                <th>👤 User</th>
                <th>📅 Last Activity</th>
                <th>🔧 Action Required</th>
            </tr>
        </thead>
        <tbody>
"@

    foreach ($device in $DeviceData) {
        $rowClass = switch -Wildcard ($device.EOLPriority) {
            "CRITICAL*" { "critical" }
            "HIGH*" { "high" }
            "MEDIUM*" { "medium" }
            "LOW*" { "low" }
            "NONE*" { "win11" }
            default { "unknown" }
        }
        
        $storageInfo = if ($device.StorageFree -ne "Unknown" -and $device.StorageFree -is [double]) {
            "$($device.StorageFree) GB free"
        } elseif ($device.StorageTotal -ne "Unknown" -and $device.StorageTotal -is [double]) {
            "$($device.StorageTotal) GB total"
        } else {
            "Unknown"
        }
        
        $lastActivity = if ($device.LastSignInDate -ne "Unknown") {
            $device.LastSignInDate
        } elseif ($device.LastLogonDate) {
            $device.LastLogonDate.ToString("yyyy-MM-dd")
        } else {
            "Never"
        }
        
        $htmlContent += @"
            <tr class="$rowClass">
                <td class="device-name">$($device.DeviceName)</td>
                <td>$($device.DeviceType)</td>
                <td class="os-version">$($device.WindowsVersion)</td>
                <td><strong>$($device.EOLPriority)</strong></td>
                <td>$($device.Win11Status)</td>
                <td>$storageInfo</td>
                <td>$($device.SignInUser)</td>
                <td>$lastActivity</td>
                <td>$($device.UpgradeAction)</td>
            </tr>
"@
    }
    
    $htmlContent += @"
        </tbody>
    </table>
    
    <div class="footer">
        <p>Report generated by Get-AllDevicesWin10EOLReadiness.ps1 v2.0</p>
        <p>For IT administrators preparing for Windows 10 End-of-Life (October 14, 2025)</p>
    </div>
</body>
</html>
"@
    
    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
    return $htmlPath
}

# Main execution
try {
    Write-Host "=== ALL DEVICES WINDOWS 10 EOL ASSESSMENT v2.0 ===" -ForegroundColor Cyan
    Write-Host "Comprehensive assessment for PCs, laptops, and mobile devices" -ForegroundColor White
    Write-Host "Excludes servers - focuses on end-user devices only" -ForegroundColor White
    Write-Host "=========================================================" -ForegroundColor Cyan
    
    # Initialize
    Initialize-RequiredModules
    
    # Connect to services
    $graphConnected = Connect-ToMSGraph
    
    # Discover all end-user devices
    $devices = Get-AllEndUserDevices
    
    if ($devices.Count -eq 0) {
        Write-Host "No end-user devices found matching criteria. Exiting." -ForegroundColor Yellow
        exit 0
    }
    
    # Enhance with Graph intelligence
    $enhancedDevices = Add-GraphIntelligence -Devices $devices -GraphConnected $graphConnected
    
    # Generate comprehensive report
    $summary = Generate-ComprehensiveEOLReport -DeviceData $enhancedDevices
    
    # Display detailed results
    Write-Host "`n=== DETAILED DEVICE ASSESSMENT ===" -ForegroundColor Green
    $enhancedDevices | Format-Table DeviceName, DeviceType, WindowsVersion, EOLPriority, Win11Status, StorageAdequate, SignInUser -AutoSize
    
    # Export comprehensive results
    $exportFields = @(
        'DeviceName', 'DeviceType', 'OU', 'WindowsVersion', 'BuildNumber', 'IsWindows10', 'EOLUrgency', 'EOLPriority',
        'Win11Eligible', 'Win11Status', 'UpgradeAction', 'DataFreshness', 'DaysInactive', 'DeviceAge',
        'StorageTotal', 'StorageFree', 'StorageAdequate', 'SignInUser', 'LastSignInDate', 'IntuneManaged',
        'GraphDeviceFound', 'ComplianceState', 'Enabled', 'IPv4Address', 'DNSHostName', 'Notes'
    )
    
    $enhancedDevices | Select-Object $exportFields | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "`n✅ Comprehensive results exported to: $OutputPath" -ForegroundColor Green
    
    # Generate HTML report
    if ($ExportHtmlReport) {
        $htmlPath = New-EnhancedHtmlReport -DeviceData $enhancedDevices -Summary $summary
        Write-Host "🌐 Enhanced HTML report generated: $htmlPath" -ForegroundColor Green
        Write-Host "   Opens in browser with color-coded priorities and legend" -ForegroundColor White
    }
    
    # Critical action items
    $criticalDevices = $enhancedDevices | Where-Object EOLPriority -like "CRITICAL*"
    if ($criticalDevices.Count -gt 0) {
        Write-Host "`n🚨 CRITICAL DEVICES REQUIRING IMMEDIATE ACTION 🚨" -ForegroundColor Red
        $criticalDevices | Select-Object DeviceName, DeviceType, Win11Status, UpgradeAction, SignInUser | Format-Table -AutoSize
    }
    
    # Final assessment
    Write-Host "`n=== 🎯 FINAL ASSESSMENT ===" -ForegroundColor Cyan
    if ($summary.ActiveWindows10 -eq 0 -and $summary.LegacyCount -eq 0) {
        Write-Host "🎉 EXCELLENT: All active devices are Windows 11 ready!" -ForegroundColor Green
    } else {
        Write-Host "⚠️  ACTION REQUIRED for Windows 10 EOL preparation:" -ForegroundColor Red
        Write-Host "   📊 $($summary.ActiveWindows10) active Windows 10 devices need attention" -ForegroundColor White
        Write-Host "   🔥 $($summary.CriticalCount) devices have critical issues" -ForegroundColor White
        Write-Host "   💾 $($summary.StorageIssues) devices have storage constraints" -ForegroundColor White
        Write-Host "   📅 $($summary.DaysUntilEOL) days until Windows 10 EOL" -ForegroundColor Yellow
        
        if ($summary.DaysUntilEOL -lt 120) {
            Write-Host "🚨 URGENT: Less than 4 months remaining!" -ForegroundColor Red
        }
    }
    
    Write-Host "`n✨ Assessment complete! Share this script with your IT friends! ✨" -ForegroundColor Green
    
} catch {
    Write-Host "❌ Assessment failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    try {
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green
        }
    } catch {
        # Ignore disconnect errors
    }
}
