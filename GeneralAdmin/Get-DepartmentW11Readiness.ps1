# Get-DepartmentW11Readiness.ps1
#
# A PowerShell script to inventory Windows devices used by employees reporting to a specific manager 
# and assess their Windows 11 upgrade readiness.
#
# This script builds on Get-DepartmentDevices.ps1 and adds Windows 11 readiness assessment.
#
# OVERVIEW:
# This script creates a comprehensive Windows 11 readiness report by:
# - Retrieving all direct reports of a specified manager
# - Querying Microsoft Graph API for Windows sign-in events
# - Identifying device information from those sign-in events
# - Assessing Windows 11 compatibility for each device
# - Producing formatted reports (console, CSV, and optional HTML)
#
# USAGE:
# Basic usage with interactive prompts:
#   .\Get-DepartmentW11Readiness.ps1
#
# Specify manager and generate HTML report:
#   .\Get-DepartmentW11Readiness.ps1 -ManagerUpn "manager@company.com" -ExportHtmlReport
#
# Custom lookback periods:
#   .\Get-DepartmentW11Readiness.ps1 -ManagerUpn "manager@company.com" -DaysBack 5 -MaxDaysBack 10
#
# DEPENDENCIES:
# - Microsoft Graph PowerShell modules:
#   - Microsoft.Graph.Users
#   - Microsoft.Graph.Identity.SignIns
#   - Microsoft.Graph.Authentication
#   - Microsoft.Graph.DeviceManagement

param (
    [Parameter(Mandatory=$false)]
    [string]$ManagerUpn,
    
    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 7,  # Default to 7 days
    
    [Parameter(Mandatory=$false)]
    [int]$MaxDaysBack = 14,  # Changed from 30 to 14 days maximum
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\Win11Readiness_$(Get-Date -Format 'yyyyMMdd').csv",
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportCsvOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportHtmlReport,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 3,
    
    [Parameter(Mandatory=$false)]
    [int]$RetryDelaySeconds = 5,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipHardwareCheck
)

# Check and install required modules
function Initialize-RequiredModules {
    $requiredModules = @(
        "Microsoft.Graph.Users", 
        "Microsoft.Graph.Identity.SignIns", 
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.DeviceManagement"
    )
    
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "Module $module is not installed. Installing now..." -ForegroundColor Yellow
            Install-Module -Name $module -Scope CurrentUser -Force
        }
        Import-Module $module
    }
}

# Connect to Microsoft Graph with required permissions
function Connect-ToMSGraph {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    
    $requiredScopes = @(
        "User.Read.All",  
        "AuditLog.Read.All",
        "Directory.Read.All",
        "DeviceManagementManagedDevices.Read.All"
    )
    
    try {
        Connect-MgGraph -Scopes $requiredScopes
        Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# Get all direct reports for a manager
function Get-DirectReports {
    param (
        [string]$ManagerUpn,
        [bool]$IncludeManager = $false
    )
    
    Write-Host "Looking up manager: $ManagerUpn" -ForegroundColor Yellow
    
    try {
        # Find the manager
        $manager = Get-MgUser -Filter "userPrincipalName eq '$ManagerUpn' or mail eq '$ManagerUpn'" -Property "Id,DisplayName,UserPrincipalName"
        
        if (!$manager) {
            Write-Host "Manager not found with UPN/email: $ManagerUpn" -ForegroundColor Red
            return $null
        }
        
        Write-Host "Found manager: $($manager.DisplayName)" -ForegroundColor Green
        Write-Host "Retrieving direct reports for manager: $($manager.DisplayName)" -ForegroundColor Yellow
        
        # Get direct reports
        $directReports = Get-MgUserDirectReport -UserId $manager.Id
        $users = @()
        
        foreach ($report in $directReports) {
            $reportDetails = Get-MgUser -UserId $report.Id -Property "Id,UserPrincipalName,DisplayName,Department,Mail,JobTitle"
            $users += $reportDetails
        }
        
        # Include the manager if requested
        if ($IncludeManager) {
            $managerDetails = Get-MgUser -UserId $manager.Id -Property "Id,UserPrincipalName,DisplayName,Department,Mail,JobTitle"
            $users += $managerDetails
            Write-Host "Including manager in the list of users to process" -ForegroundColor Green
        }
        
        Write-Host "Found $($users.Count) users to process" -ForegroundColor Green
        return $users
    }
    catch {
        Write-Host "Error retrieving direct reports: $_" -ForegroundColor Red
        return $null
    }
}

# Get sign-in logs for a specific user with retries for timeouts
function Get-UserSignInsWithRetry {
    param (
        [string]$UserId,
        [string]$UserUpn,
        [int]$DaysBack,
        [int]$MaxRetries,
        [int]$RetryDelaySeconds,
        [switch]$MinimalMode = $false,
        [switch]$FinalPassMode = $false
    )
    
    # Make sure all dates are in UTC
    $currentDateUtc = (Get-Date).ToUniversalTime()
    $startDateUtc = $currentDateUtc.AddDays(-$DaysBack)
    $success = $false
    $result = $null
    
    # Track the strategies attempted
    $attemptedStrategies = @()
    
    # Format the date for Graph API (ISO 8601 in UTC)
    $formatDate = {
        param($date)
        return $date.ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    
    # Try different approaches progressively - from most specific to most general
    $strategies = @(
        @{ Description = "Windows Sign In app name, last 24 hours"; Filter = "userId eq '$UserId' and appDisplayName eq 'Windows Sign In' and createdDateTime ge $(&$formatDate $currentDateUtc.AddHours(-24))"; Top = 2 },
        @{ Description = "Windows Sign In app name, last week"; Filter = "userId eq '$UserId' and appDisplayName eq 'Windows Sign In' and createdDateTime ge $(&$formatDate $currentDateUtc.AddDays(-7))"; Top = 2 },
        @{ Description = "Any recent sign-in, last 24 hours"; Filter = "userId eq '$UserId' and status/errorCode eq 0 and createdDateTime ge $(&$formatDate $currentDateUtc.AddHours(-24))"; Top = 1 }
    )
    
    # For minimal mode (admin accounts), use only one simplified strategy
            if ($MinimalMode) {
        $strategies = @(
            @{ Description = "Any recent sign-in with device, last 3 days"; Filter = "userId eq '$UserId' and status/errorCode eq 0 and createdDateTime ge $(&$formatDate $currentDateUtc.AddDays(-3))"; Top = 2 }
        )
    }
    
    # For final pass mode, use a more limited set of strategies focusing on efficiency
    if ($FinalPassMode) {
        $strategies = @(
            @{ Description = "Windows Sign In app name, last week"; Filter = "userId eq '$UserId' and appDisplayName eq 'Windows Sign In' and createdDateTime ge $(&$formatDate $currentDateUtc.AddDays(-7))"; Top = 2 },
            @{ Description = "Any recent sign-in, last 24 hours"; Filter = "userId eq '$UserId' and status/errorCode eq 0 and createdDateTime ge $(&$formatDate $currentDateUtc.AddHours(-24))"; Top = 1 }
        )
    }
    else {
        # Default strategies remain unchanged
    }
    
    # Try each strategy until one works
    foreach ($strategy in $strategies) {
        Write-Host "    Trying strategy: $($strategy.Description)" -ForegroundColor Yellow
        $attemptedStrategies += $strategy.Description
        
        $attempt = 0
        while ($attempt -lt $MaxRetries -and -not $success) {
            $attempt++
            try {
                if ($attempt -gt 1) {
                    Write-Host "      Retry attempt $attempt of $MaxRetries..." -ForegroundColor Yellow
                }
                
                # Set a timeout policy for the request
                $previousTimeoutPolicy = [System.Net.ServicePointManager]::FindServicePoint("https://graph.microsoft.com").ConnectionLeaseTimeout
                try {
                    # Set a 30 second timeout
                    [System.Net.ServicePointManager]::FindServicePoint("https://graph.microsoft.com").ConnectionLeaseTimeout = 30000
                    
                    # Use a very small batch size to avoid timeouts
                    $result = Get-MgAuditLogSignIn -Filter $strategy.Filter -Top $strategy.Top -All:$false -ErrorAction Stop
                    
                    if ($result -and $result.Count -gt 0) {
                        # Check if any results have device details
                        $withDeviceDetails = $result | Where-Object { $_.DeviceDetail -and $_.DeviceDetail.DisplayName }
                        
                        if ($withDeviceDetails -and $withDeviceDetails.Count -gt 0) {
                            Write-Host "      Success! Found $($withDeviceDetails.Count) sign-ins with device details." -ForegroundColor Green
                            $success = $true
                            $result = $withDeviceDetails  # Return only the ones with device details
                            # Set metadata for the result
                            $result | Add-Member -NotePropertyName "StrategyUsed" -NotePropertyValue $strategy.Description -Force
                            $result | Add-Member -NotePropertyName "StrategiesAttempted" -NotePropertyValue ($attemptedStrategies -join ', ') -Force
                            $result | Add-Member -NotePropertyName "AttemptNumber" -NotePropertyValue $attempt -Force
                            break  # Exit the while loop
                        }
                        else {
                            Write-Host "      Found sign-ins but none had device details. Continuing to next strategy." -ForegroundColor Yellow
                        }
                    }
                    else {
                        Write-Host "      Query succeeded but returned no results." -ForegroundColor Yellow
                    }
                }
                finally {
                    # Restore original timeout
                    [System.Net.ServicePointManager]::FindServicePoint("https://graph.microsoft.com").ConnectionLeaseTimeout = $previousTimeoutPolicy
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                
                # Different handling based on error type
                if ($errorMessage -like "*task was canceled*") {
                    Write-Warning "      Timeout occurred. Trying a different approach..."
                    break  # Skip remaining retries with this strategy
                }
                else {
                    Write-Warning "      API request failed: $errorMessage. Retrying in $RetryDelaySeconds seconds..."
                    Start-Sleep -Seconds $RetryDelaySeconds
                    $RetryDelaySeconds = [Math]::Min(30, $RetryDelaySeconds * 2)  # Exponential backoff with a cap
                }
            }
        }
        
        if ($success) {
            break  # Exit the foreach loop if we got data
        }
        
        # Give the API a short break between strategies
        Start-Sleep -Seconds 1
    }
    
    # If all strategies failed, try a direct device lookup as a last resort
    if (-not $success) {
        Write-Host "    All sign-in strategies failed. Attempting direct device lookup..." -ForegroundColor Yellow
        $attemptedStrategies += "Direct device lookup"
        
        try {
            # Try to see if the user has any registered devices in Entra ID
            $devices = Get-MgUserRegisteredDevice -UserId $UserId -ErrorAction Stop
            if ($devices -and $devices.Count -gt 0) {
                Write-Host "      Success! Found $($devices.Count) registered devices for user." -ForegroundColor Green
                
                # Format the device data to match our expected structure
                $signInResults = @()
                foreach ($device in $devices) {
                    $fakeSignIn = [PSCustomObject]@{
                        CreatedDateTime = (Get-Date)
                        DeviceDetail = [PSCustomObject]@{
                            DisplayName = $device.DisplayName
                            OperatingSystem = "Unknown (Direct Lookup)"
                            TrustType = "Registered Device"
                            IsCompliant = $null
                            IsManagedDevice = $null
                        }
                        Location = [PSCustomObject]@{
                            City = ""
                        }
                        IPAddress = ""
                        StrategyUsed = "Direct device lookup"
                        StrategiesAttempted = $attemptedStrategies -join ', '
                        AttemptNumber = 1
                    }
                    $signInResults += $fakeSignIn
                }
                return $signInResults  # Return the devices found and exit function
            }
            else {
                Write-Host "      No registered devices found." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "      Direct device lookup failed: $_" -ForegroundColor Red
        }
    }
    
    # Add metadata even if no results
    if (-not $result) {
        $result = @{
            StrategiesAttempted = $attemptedStrategies -join ', '
        }
    }
    
    return $result
}

# Process sign-in logs to extract device information
function Process-SignInLogs {
    param (
        [array]$Users,
        [int]$DaysBack,
        [int]$MaxRetries,
        [int]$RetryDelaySeconds
    )
    
    $results = @()
    $retryQueue = @()  # Users that need a second chance
    $userCount = $Users.Count
    $currentUser = 0
    
    # Create a hashtable to track processed users
    $processedUsers = @{}
    
    Write-Host "`n=== PROCESSING USERS (First Pass) ===" -ForegroundColor Cyan
    
    # First pass - process all users, queue up failures for retry
    foreach ($user in $Users) {
        $currentUser++
        Write-Progress -Activity "Processing users" -Status "User $currentUser of $userCount" -PercentComplete (($currentUser / $userCount) * 100)
        
        # Track this user as processed
        $processedUsers[$user.Id] = $true
        
        Write-Host "Processing user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Cyan
        
        # Check for admin account patterns
        $isAdminAccount = $false
        if ($user.DisplayName -like "Admin*" -or 
            $user.UserPrincipalName -like "a_*" -or 
            $user.UserPrincipalName -like "svc_*" -or
            $user.UserPrincipalName -like "admin*") {
            
            $isAdminAccount = $true
            Write-Host "  Detected admin/service account pattern. Using minimal search." -ForegroundColor Yellow
            
            # For admin accounts, just add to results with minimal searching
            $adminDeviceInfo = [PSCustomObject]@{
                UserDisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Department = $user.Department
                JobTitle = $user.JobTitle
                DeviceName = "Admin account - No device expected"
                OperatingSystem = ""
                TrustType = ""
                IsCompliant = ""
                IsManagedDevice = ""
                LastSignIn = ""
                City = ""
                IPAddress = ""
                Note = "Admin account detected; minimal search performed"
                Win11_Eligible = "N/A"
                Win11_Status = "N/A"
                Upgrade_Action = "N/A"
            }
            
            $results += $adminDeviceInfo
            
            # Try a minimal check just to see if there is any device (but don't waste time with extensive searches)
            $quickSignIns = Get-UserSignInsWithRetry -UserId $user.Id -UserUpn $user.UserPrincipalName -DaysBack 3 -MaxRetries 1 -RetryDelaySeconds $RetryDelaySeconds -MinimalMode $true
            
            if ($quickSignIns -and ($quickSignIns.Count -gt 0) -and ($quickSignIns -is [array] -or $quickSignIns -is [System.Collections.ArrayList])) {
                # We found devices for the admin account - process them
                $processedDevices = Process-DeviceInformation -User $user -SignIns $quickSignIns
                $results += $processedDevices
            }
            
            # Skip to next user
            continue
        }
        
        # Try to get sign-in logs with device details (for non-admin accounts)
        $signIns = Get-UserSignInsWithRetry -UserId $user.Id -UserUpn $user.UserPrincipalName -DaysBack $DaysBack -MaxRetries $MaxRetries -RetryDelaySeconds $RetryDelaySeconds
        
        if ($signIns -and ($signIns.Count -gt 0)) {
            # Filter for sign-ins with device details (should be redundant now but kept for safety)
            $signInsWithDeviceDetails = $signIns | Where-Object { $_.DeviceDetail -and $_.DeviceDetail.DisplayName }
            
            if ($signInsWithDeviceDetails -and ($signInsWithDeviceDetails.Count -gt 0)) {
                # Process device information
                $processedDevices = Process-DeviceInformation -User $user -SignIns $signInsWithDeviceDetails
                $results += $processedDevices
            }
            else {
                Write-Host "  No sign-ins with device details found" -ForegroundColor Yellow
                $retryQueue += $user
            }
        }
        else {
            Write-Host "  No Windows sign-ins found or error occurred" -ForegroundColor Yellow
            $retryQueue += $user
        }
        
        # Give the Graph API a break between users
        Start-Sleep -Milliseconds 500
    }
    
    # Second pass - retry with extended days for users with no results
    if ($retryQueue.Count -gt 0) {
        Write-Host "`n=== RETRY QUEUE (Second Pass) ===" -ForegroundColor Cyan
        
        $currentRetry = 0
        $retryCount = $retryQueue.Count
        
        Write-Host "Retrying $retryCount users with limited search approach..." -ForegroundColor Yellow
        
        foreach ($user in $retryQueue) {
            $currentRetry++
            Write-Progress -Activity "Processing retry queue" -Status "User $currentRetry of $retryCount" -PercentComplete (($currentRetry / $retryCount) * 100)
            
            Write-Host "Retrying user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Cyan
            
            # Try limited search with 5-day lookback
            Write-Host "  Trying Windows Sign In limited search (5 days)..." -ForegroundColor Yellow
            
            $signIns = Get-UserSignInsWithRetry -UserId $user.Id -UserUpn $user.UserPrincipalName -DaysBack 5 -MaxRetries 1 -RetryDelaySeconds $RetryDelaySeconds -FinalPassMode $true
            
            if ($signIns -and ($signIns.Count -gt 0) -and ($signIns -is [array] -or $signIns -is [System.Collections.ArrayList])) {
                # Process device information
                $processedDevices = Process-DeviceInformation -User $user -SignIns $signIns -IsExtendedSearch $true -ExtendedDays 5
                $results += $processedDevices
                continue
            }
            
            # If Windows Sign In search fails, try direct device lookup as a last resort
            try {
                Write-Host "  Trying direct device lookup as fallback..." -ForegroundColor Yellow
                $devices = Get-MgUserRegisteredDevice -UserId $user.Id -ErrorAction Stop
                
                if ($devices -and $devices.Count -gt 0) {
                    Write-Host "  Found $($devices.Count) registered devices for user via direct lookup, but these may not be reliable." -ForegroundColor Yellow
                    
                    # Process and add these devices with a warning note
                    foreach ($device in $devices) {
                        $deviceInfo = [PSCustomObject]@{
                            UserDisplayName = $user.DisplayName
                            UserPrincipalName = $user.UserPrincipalName
                            Department = $user.Department
                            JobTitle = $user.JobTitle
                            DeviceName = $device.DisplayName
                            OperatingSystem = "Unknown (Direct Lookup)"
                            TrustType = "Registered Device"
                            IsCompliant = $null
                            IsManagedDevice = $null
                            LastSignIn = "Unknown"
                            City = ""
                            IPAddress = ""
                            Note = "Found via direct device lookup - may need verification as device registrations are not always reliable"
                            Win11_Eligible = "Unknown"
                            Win11_Status = "Needs assessment"
                            Upgrade_Action = "Manual investigation required"
                        }
                        $results += $deviceInfo
                        Write-Host "    Device: $($device.DisplayName) (verify this registration)" -ForegroundColor Yellow
                    }
                    
                    # Continue since we found something, even if potentially unreliable
                    continue
                }
                else {
                    Write-Host "  No registered devices found via direct lookup." -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "  Direct device lookup failed: $_" -ForegroundColor Yellow
            }
            
            # If we get here, no devices were found through any method
            try {
                # Get account status to add to note
                $detailedUser = Get-MgUser -UserId $user.Id -Property AccountEnabled, OfficeLocation, UsageLocation -ErrorAction Stop
                
                $noDeviceInfo = [PSCustomObject]@{
                    UserDisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    Department = $user.Department
                    JobTitle = $user.JobTitle
                    DeviceName = "Needs manual lookup"
                    OperatingSystem = ""
                    TrustType = ""
                    IsCompliant = ""
                    IsManagedDevice = ""
                    LastSignIn = ""
                    City = $detailedUser.UsageLocation
                    IPAddress = ""
                    Note = "Account " + $(if($detailedUser.AccountEnabled){"enabled"}else{"disabled"}) + 
                           ". No devices found via Windows Sign In or device registration. Needs manual investigation."
                    Win11_Eligible = "Unknown"
                    Win11_Status = "No device found"
                    Upgrade_Action = "Manual investigation required"
                }
            }
            catch {
                $noDeviceInfo = [PSCustomObject]@{
                    UserDisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    Department = $user.Department
                    JobTitle = $user.JobTitle
                    DeviceName = "Needs manual lookup"
                    OperatingSystem = ""
                    TrustType = ""
                    IsCompliant = ""
                    IsManagedDevice = ""
                    LastSignIn = ""
                    City = ""
                    IPAddress = ""
                    Note = "Limited search unsuccessful. Needs manual investigation."
                    Win11_Eligible = "Unknown"
                    Win11_Status = "No device found"
                    Upgrade_Action = "Manual investigation required"
                }
            }
            
            $results += $noDeviceInfo
        }
    }
    
    return $results
}

# Helper function to process device information from sign-ins
function Process-DeviceInformation {
    param(
        [object]$User,
        [array]$SignIns,
        [bool]$IsExtendedSearch = $false,
        [int]$ExtendedDays = 0
    )
    
    $deviceResults = @()
    
    # Group by device to get unique devices
    $deviceSignIns = $SignIns | Group-Object -Property { $_.DeviceDetail.DisplayName }
    
    Write-Host "  Found $($deviceSignIns.Count) devices for user" -ForegroundColor Green
    
    foreach ($device in $deviceSignIns) {
        # Get most recent sign-in for this device
        $signIn = $device.Group | Sort-Object CreatedDateTime -Descending | Select-Object -First 1
        
        $note = ""
        if ($IsExtendedSearch) {
            $note = "Found in extended search ($ExtendedDays days)"
        }
        
        # Include strategy information in the note
        $strategyInfo = ""
        if ($signIn.PSObject.Properties.Name -contains "StrategyUsed") {
            $strategyInfo = "Strategy: $($signIn.StrategyUsed)"
            if ($signIn.PSObject.Properties.Name -contains "AttemptNumber" -and $signIn.AttemptNumber -gt 1) {
                $strategyInfo += " (Attempt $($signIn.AttemptNumber))"
            }
        }
        
        # Add strategy attempts if available
        $strategyAttempts = ""
        if ($signIn.PSObject.Properties.Name -contains "StrategiesAttempted") {
            $strategyAttempts = "Strategies tried: $($signIn.StrategiesAttempted)"
        }
        
        # Combine notes
        if ($strategyInfo -ne "") {
            if ($note -ne "") {
                $note += ". $strategyInfo"
            }
            else {
                $note = $strategyInfo
            }
        }
        
        if ($strategyAttempts -ne "") {
            if ($note -ne "") {
                $note += ". $strategyAttempts"
            }
            else {
                $note = $strategyAttempts
            }
        }
        
        # Assess Windows 11 readiness
        $win11Status = Assess-Win11Readiness -OSVersion $signIn.DeviceDetail.OperatingSystem -DeviceName $signIn.DeviceDetail.DisplayName
        
        $deviceInfo = [PSCustomObject]@{
            UserDisplayName = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            Department = $User.Department
            JobTitle = $User.JobTitle
            DeviceName = $signIn.DeviceDetail.DisplayName
            OperatingSystem = $signIn.DeviceDetail.OperatingSystem
            TrustType = $signIn.DeviceDetail.TrustType
            IsCompliant = $signIn.DeviceDetail.IsCompliant
            IsManagedDevice = $signIn.DeviceDetail.IsManagedDevice
            LastSignIn = $signIn.CreatedDateTime
            City = $signIn.Location.City
            IPAddress = $signIn.IPAddress
            Note = $note
            Win11_Eligible = $win11Status.Eligible
            Win11_Status = $win11Status.Status
            Upgrade_Action = $win11Status.Action
        }
        
        $deviceResults += $deviceInfo
        Write-Host "    Device: $($deviceInfo.DeviceName) - $($deviceInfo.OperatingSystem) - Win11: $($deviceInfo.Win11_Status)" -ForegroundColor White
    }
    
    return $deviceResults
}

# Function to assess Windows 11 readiness based on OS version
function Assess-Win11Readiness {
    param (
        [string]$OSVersion,
        [string]$DeviceName
    )
    
    # Initialize result object
    $result = @{
        Eligible = "Unknown"
        Status = "Needs assessment"
        Action = "Hardware assessment required"
    }
    
    # Check for null or empty OS version
    if ([string]::IsNullOrWhiteSpace($OSVersion)) {
        $result.Status = "Unknown OS"
        $result.Action = "Cannot determine - missing OS information"
        return $result
    }
    
    # Simple build number check for Windows 10 vs Windows 11
    if ($OSVersion -match "10\.0\.(\d+)") {
        $mainBuild = $Matches[1]
        
        if ([int]$mainBuild -ge 22000) {
            # Windows 11
            $result.Eligible = "Yes"
            $result.Status = "Already on Windows 11"
            $result.Action = "None - already on Windows 11"
        }
        else {
            # Windows 10
            $result.Eligible = "Yes (OS)"
            $result.Status = "Windows 10 - Eligible for upgrade"
            $result.Action = "Hardware compatibility check required"
            
            # If not skipping hardware check, attempt to get detailed info
            if (-not $script:SkipHardwareCheck) {
                # Get additional hardware details if possible
                try {
                    # Try to find device in Intune
                    $managedDevice = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$DeviceName'" -ErrorAction SilentlyContinue
                    
                    if ($managedDevice) {
                        # Add hardware check
                        $hardwareEligible = Test-Win11HardwareCompatibility -Device $managedDevice
                        
                        if ($hardwareEligible.IsCompatible) {
                            $result.Eligible = "Yes"
                            $result.Status = "Windows 10 with compatible hardware"
                            $result.Action = "Eligible for upgrade"
                        }
                        else {
                            $result.Eligible = "No"
                            $result.Status = "Hardware not compatible: $($hardwareEligible.Reason)"
                            $result.Action = "Hardware upgrade required: $($hardwareEligible.Reason)"
                        }
                    }
                }
                catch {
                    # If hardware check fails, just leave the OS assessment
                    Write-Host "      Warning: Could not check hardware compatibility for $DeviceName. $_" -ForegroundColor Yellow
                }
            }
            else {
                # Skip hardware check as requested, keep the OS assessment only
                # No action needed - already set up with OS assessment values
            }
        }
    }
    # Handle generic "Windows" with no version info
    elseif ($OSVersion -match "Windows") {
        # Assume it's most likely Windows 10
        $result.Eligible = "Yes (assumed)"
        $result.Status = "Windows detected - specific version unknown"
        $result.Action = "Detailed assessment required"
        
        # Try to get more detailed info about the device
        if (-not $script:SkipHardwareCheck) {
            try {
                # Try to find device in Intune
                $managedDevice = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$DeviceName'" -ErrorAction SilentlyContinue
                
                if ($managedDevice -and $managedDevice.OperatingSystem) {
                    # We got better OS info from device management
                    if ($managedDevice.OperatingSystem -match "Windows 11") {
                        $result.Eligible = "Yes"
                        $result.Status = "Already on Windows 11 (from device record)"
                        $result.Action = "None - already on Windows 11"
                    }
                    elseif ($managedDevice.OperatingSystem -match "Windows 10") {
                        $result.Eligible = "Yes (OS)"
                        $result.Status = "Windows 10 (from device record)"
                        $result.Action = "Hardware compatibility check required"
                        
                        # Add hardware check
                        $hardwareEligible = Test-Win11HardwareCompatibility -Device $managedDevice
                        
                        if ($hardwareEligible.IsCompatible) {
                            $result.Eligible = "Yes"
                            $result.Status = "Windows 10 with compatible hardware"
                            $result.Action = "Eligible for upgrade"
                        }
                        else {
                            $result.Eligible = "No"
                            $result.Status = "Hardware not compatible: $($hardwareEligible.Reason)"
                            $result.Action = "Hardware upgrade required: $($hardwareEligible.Reason)"
                        }
                    }
                }
            }
            catch {
                # If lookup fails, keep the assumption
                Write-Host "      Warning: Could not get additional device info for $DeviceName. $_" -ForegroundColor Yellow
            }
        }
    }
    else {
        # Unexpected version format
        $result.Eligible = "Unknown"
        $result.Status = "Unrecognized OS version format"
        $result.Action = "Manual assessment required"
    }
    
    return $result
}

# Function to check Windows 11 hardware compatibility
function Test-Win11HardwareCompatibility {
    param (
        [object]$Device
    )
    
    # Windows 11 minimum requirements
    $win11Requirements = @{
        MinProcessorCores = 2
        MinRAM = 4  # GB
        MinStorage = 64  # GB
        RequiresTPM = $true
        MinTPMVersion = "2.0"
        RequiresSecureBoot = $true
    }
    
    # Initialize result
    $result = @{
        IsCompatible = $true
        Reason = ""
    }
    
    # Check processor cores
    if ($Device.ProcessorCoreCount -gt 0 -and $Device.ProcessorCoreCount -lt $win11Requirements.MinProcessorCores) {
        $result.IsCompatible = $false
        $result.Reason = "Insufficient CPU cores ($($Device.ProcessorCoreCount), need $($win11Requirements.MinProcessorCores))"
    }
    
    # Check RAM
    if ($Device.PhysicalMemoryInBytes -gt 0) {
        $ramGB = [math]::Round($Device.PhysicalMemoryInBytes / 1GB, 1)
        if ($ramGB -lt $win11Requirements.MinRAM) {
            $result.IsCompatible = $false
            $result.Reason = "Insufficient RAM ($ramGB GB, need $($win11Requirements.MinRAM) GB)"
        }
    }
    
    # Check storage
    if ($Device.TotalStorageSpaceInBytes -gt 0) {
        $storageGB = [math]::Round($Device.TotalStorageSpaceInBytes / 1GB)
        if ($storageGB -lt $win11Requirements.MinStorage) {
            $result.IsCompatible = $false
            $result.Reason = "Insufficient storage ($storageGB GB, need $($win11Requirements.MinStorage) GB)"
        }
    }
    
    # Check TPM version
    if ($win11Requirements.RequiresTPM -and $Device.psobject.Properties.Name -contains "TpmSpecificationVersion") {
        if ([string]::IsNullOrEmpty($Device.TpmSpecificationVersion)) {
            $result.Reason = "TPM status unknown"
        }
        elseif ($Device.TpmSpecificationVersion -eq "Not Present") {
            $result.IsCompatible = $false
            $result.Reason = "TPM not present"
        }
        else {
            # Check TPM version
            try {
                $tpmVersionNum = [decimal]::Parse($Device.TpmSpecificationVersion)
                $minTpmVersionNum = [decimal]::Parse($win11Requirements.MinTPMVersion)
                
                if ($tpmVersionNum -lt $minTpmVersionNum) {
                    $result.IsCompatible = $false
                    $result.Reason = "TPM version too low ($($Device.TpmSpecificationVersion), need $($win11Requirements.MinTPMVersion))"
                }
            }
            catch {
                # If parsing fails, just continue
            }
        }
    }
    
    # Check Secure Boot
    if ($win11Requirements.RequiresSecureBoot -and $Device.psobject.Properties.Name -contains "IsSecureBootEnabled") {
        if ($Device.IsSecureBootEnabled -eq $false) {
            $result.IsCompatible = $false
            $result.Reason = "Secure Boot not enabled"
        }
    }
    
    return $result
}

# Generate HTML report
function New-HtmlReport {
    param (
        [array]$DeviceData,
        [string]$ManagerUpn,
        [int]$TotalUsers,
        [int]$UsersWithDevices,
        [int]$UsersNoDevices,
        [int]$TotalDevices,
        [int]$Win11ReadyDevices,
        [int]$AlreadyOnWin11Devices,
        [int]$NotReadyDevices
    )
    
    $htmlPath = $OutputPath -replace '\.csv$', '.html'
    
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows 11 Readiness Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { margin-top: 30px; background-color: #e6f2ff; padding: 15px; border-radius: 5px; }
        .warning { color: orange; }
        .success { color: green; }
        .error { color: red; }
        .win11ready { background-color: #e6ffe6; }
        .win11notready { background-color: #ffebe6; }
        .win11already { background-color: #e6f9ff; }
    </style>
</head>
<body>
    <h1>Windows 11 Readiness Report</h1>
    <p>Manager: $ManagerUpn</p>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <table>
        <tr>
            <th>User</th>
            <th>Email</th>
            <th>Department</th>
            <th>Device Name</th>
            <th>Operating System</th>
            <th>Last Sign-In</th>
            <th>Win11 Status</th>
            <th>Upgrade Action</th>
            <th>Notes</th>
        </tr>
"@

    $htmlRows = $DeviceData | ForEach-Object {
        $rowClass = ""
        
        if ($_.Win11_Status -eq "Already on Windows 11") {
            $rowClass = ' class="win11already"'
        }
        elseif ($_.Win11_Eligible -eq "Yes") {
            $rowClass = ' class="win11ready"'
        }
        elseif ($_.Win11_Eligible -eq "No" -or $_.DeviceName -like "No *") {
            $rowClass = ' class="win11notready"'
        }
        
        "<tr$rowClass>
            <td>$($_.UserDisplayName)</td>
            <td>$($_.UserPrincipalName)</td>
            <td>$($_.Department)</td>
            <td>$($_.DeviceName)</td>
            <td>$($_.OperatingSystem)</td>
            <td>$($_.LastSignIn)</td>
            <td>$($_.Win11_Status)</td>
            <td>$($_.Upgrade_Action)</td>
            <td>$($_.Note)</td>
        </tr>"
    }

    $htmlSummary = @"
    <div class="summary">
        <h2>Summary</h2>
        <p>Total users processed: $TotalUsers</p>
        <p>Users with devices found: $UsersWithDevices</p>
        <p>Users with no devices found: $UsersNoDevices</p>
        <p>Total devices found: $TotalDevices</p>
        <h3>Windows 11 Readiness</h3>
        <p>Devices already on Windows 11: $AlreadyOnWin11Devices</p>
        <p>Devices ready for Windows 11: $Win11ReadyDevices</p>
        <p>Devices not ready for Windows 11: $NotReadyDevices</p>
    </div>
</body>
</html>
"@

    $htmlContent = $htmlHeader + ($htmlRows -join '') + $htmlSummary
    $htmlContent | Out-File -FilePath $htmlPath -Encoding utf8
    
    return $htmlPath
}

# Prompt for manager email with validation
function Get-ValidatedManagerUpn {
    $isValid = $false
    $managerUpn = ""
    
    while (-not $isValid) {
        $managerUpn = Read-Host -Prompt "Enter manager's email address"
        
        if ([string]::IsNullOrEmpty($managerUpn)) {
            Write-Host "No manager email provided. Please enter a valid email address." -ForegroundColor Red
            continue
        }
        
        # Basic email format validation
        if ($managerUpn -notmatch "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
            Write-Host "Invalid email format. Please enter a valid email address." -ForegroundColor Red
            continue
        }
        
        # Ask for confirmation to catch typos
        $confirmation = Read-Host -Prompt "Is '$managerUpn' correct? (Y/N)"
        if ($confirmation -like "Y*") {
            $isValid = $true
        }
    }
    
    return $managerUpn
}

# Main execution
try {
    # Check and install required modules
    Initialize-RequiredModules
    
    # Prompt for manager email if not provided
    if ([string]::IsNullOrEmpty($ManagerUpn)) {
        $ManagerUpn = Get-ValidatedManagerUpn
    }
    else {
        # Validate the provided email parameter
        if ($ManagerUpn -notmatch "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
            Write-Host "Warning: The provided email '$ManagerUpn' doesn't appear to be in a valid format." -ForegroundColor Yellow
            $confirmEmail = Read-Host -Prompt "Do you want to continue with this email anyway? (Y/N)"
            if ($confirmEmail -notlike "Y*") {
                $ManagerUpn = Get-ValidatedManagerUpn
            }
        }
    }
    
    # Confirm days back
    $confirmDays = Read-Host -Prompt "Initially look back $DaysBack days for sign-in data? (Y/N or enter new number)"
    if ($confirmDays -match "^\d+$") {
        $DaysBack = [int]$confirmDays
        Write-Host "Setting initial search period to $DaysBack days" -ForegroundColor Yellow
    }
    elseif ($confirmDays -like "N*") {
        $newDays = Read-Host -Prompt "Enter number of days to initially look back"
        if ($newDays -match "^\d+$") {
            $DaysBack = [int]$newDays
            Write-Host "Setting initial search period to $DaysBack days" -ForegroundColor Yellow
        }
    }
    
    # Confirm maximum days back
    $confirmMaxDays = Read-Host -Prompt "Maximum look back period for retry searches: $MaxDaysBack days? (Y/N or enter new number)"
    if ($confirmMaxDays -match "^\d+$") {
        $MaxDaysBack = [int]$confirmMaxDays
        Write-Host "Setting maximum search period to $MaxDaysBack days" -ForegroundColor Yellow
    }
    elseif ($confirmMaxDays -like "N*") {
        $newMaxDays = Read-Host -Prompt "Enter maximum number of days to look back during retries"
        if ($newMaxDays -match "^\d+$") {
            $MaxDaysBack = [int]$newMaxDays
            Write-Host "Setting maximum search period to $MaxDaysBack days" -ForegroundColor Yellow
        }
    }
    
    # Prompt for export options if not specified
    if (-not $PSBoundParameters.ContainsKey('ExportHtmlReport')) {
        $htmlChoice = Read-Host -Prompt "Generate HTML report? (Y/N)"
        if ($htmlChoice -like "Y*") {
            $ExportHtmlReport = $true
            Write-Host "Will generate HTML report" -ForegroundColor Yellow
        }
    }
    
    # Prompt for detailed hardware check
    if (-not $PSBoundParameters.ContainsKey('SkipHardwareCheck')) {
        $skipHardwareChoice = Read-Host -Prompt "Skip detailed hardware compatibility check? (Y/N)"
        if ($skipHardwareChoice -like "Y*") {
            $SkipHardwareCheck = $true
            Write-Host "Will skip detailed hardware compatibility check" -ForegroundColor Yellow
        }
    }

    # Prompt for including manager
    $includeManager = $false
    $managerChoice = Read-Host -Prompt "Include manager's devices in the report? (Y/N)"
    if ($managerChoice -like "Y*") {
        $includeManager = $true
        Write-Host "Will include manager's devices" -ForegroundColor Yellow
    }
    
    # Connect to Microsoft Graph
    $connected = Connect-ToMSGraph
    
    if (!$connected) {
        Write-Host "Failed to connect to Microsoft Graph. Exiting script." -ForegroundColor Red
        exit 1
    }
    
    # Get users reporting to the specified manager
    $managersDirectReports = Get-DirectReports -ManagerUpn $ManagerUpn -IncludeManager:$includeManager
    
    if ($managersDirectReports -and ($managersDirectReports.Count -gt 0)) {
        # Process sign-in logs
        $userDevices = Process-SignInLogs -Users $managersDirectReports -DaysBack $DaysBack -MaxRetries $MaxRetries -RetryDelaySeconds $RetryDelaySeconds
        
        # Calculate summary statistics
        $usersWithDevices = ($userDevices | Where-Object { $_.DeviceName -notlike "No *" }).UserPrincipalName | Select-Object -Unique | Measure-Object | Select-Object -ExpandProperty Count
        $totalDevicesFound = ($userDevices | Where-Object { $_.DeviceName -notlike "No *" }).Count
        $usersNoDevices = $managersDirectReports.Count - $usersWithDevices
        
        # Calculate Windows 11 readiness statistics
        $alreadyOnWin11Devices = ($userDevices | Where-Object { $_.Win11_Status -eq "Already on Windows 11" }).Count
        $win11ReadyDevices = ($userDevices | Where-Object { $_.Win11_Eligible -eq "Yes" -and $_.Win11_Status -ne "Already on Windows 11" }).Count
        $notReadyDevices = ($userDevices | Where-Object { $_.Win11_Eligible -eq "No" -or $_.DeviceName -like "No *" }).Count
        
        # Process output based on parameters
        if (!$ExportCsvOnly) {
            # Display results as table in console
            Write-Host "`n=== WINDOWS 11 READINESS RESULTS ===" -ForegroundColor Cyan
            $userDevices | Format-Table -Property UserDisplayName, DeviceName, OperatingSystem, Win11_Status, Upgrade_Action -AutoSize
        }
        
        # Export results to CSV file
        $userDevices | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Host "Results exported to CSV: $OutputPath" -ForegroundColor Green
        
        # Generate HTML report if requested
        if ($ExportHtmlReport) {
            $htmlPath = New-HtmlReport -DeviceData $userDevices -ManagerUpn $ManagerUpn `
                -TotalUsers $managersDirectReports.Count -UsersWithDevices $usersWithDevices -UsersNoDevices $usersNoDevices `
                -TotalDevices $totalDevicesFound -Win11ReadyDevices $win11ReadyDevices -AlreadyOnWin11Devices $alreadyOnWin11Devices `
                -NotReadyDevices $notReadyDevices
            Write-Host "HTML report exported to: $htmlPath" -ForegroundColor Green
        }
        
        # Display summary
        Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
        Write-Host "Total users processed: $($managersDirectReports.Count)" -ForegroundColor Green
        Write-Host "Users with devices found: $usersWithDevices" -ForegroundColor Green
        Write-Host "Users with no devices found: $usersNoDevices" -ForegroundColor Yellow
        Write-Host "Total devices found: $totalDevicesFound" -ForegroundColor Green
        
        Write-Host "`n=== WINDOWS 11 READINESS ===" -ForegroundColor Cyan
        Write-Host "Devices already on Windows 11: $alreadyOnWin11Devices" -ForegroundColor Green
        Write-Host "Devices ready for Windows 11: $win11ReadyDevices" -ForegroundColor Green
        Write-Host "Devices not ready for Windows 11: $notReadyDevices" -ForegroundColor Yellow
    }
    else {
        Write-Host "No users found reporting to the specified manager. Exiting script." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
}
finally {
    # Disconnect from Microsoft Graph
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green
}