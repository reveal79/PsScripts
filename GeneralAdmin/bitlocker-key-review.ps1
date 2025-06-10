<#
.SYNOPSIS
    Gets a comprehensive list of devices with BitLocker enabled and keys stored in Microsoft Entra ID.
.DESCRIPTION
    This script retrieves BitLocker recovery keys from Microsoft Entra ID and matches them with 
    detailed device information, handling cases where the initial device lookup fails.
.NOTES
    Requires the Microsoft Graph PowerShell SDK modules.
#>

# Error handling wrapper
try {
    # Step 1: Import required modules with error handling
    Write-Host "Importing Microsoft Graph modules..." -ForegroundColor Cyan
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    
    # Step 2: Connect to Microsoft Graph with required permissions
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "BitLockerKey.Read.All", "Device.Read.All", "Directory.Read.All" -ErrorAction Stop
    
    # Get current tenant ID from the connection
    $Context = Get-MgContext
    $TenantId = $Context.TenantId
    Write-Host "Connected to tenant: $TenantId" -ForegroundColor Green
    
    # Step 3: Define output path
    $OutputPath = "C:\temp\bitlocker_report.csv"
    
    # Step 4: Get ALL BitLocker recovery keys
    Write-Host "Retrieving BitLocker recovery keys..." -ForegroundColor Cyan
    $BitLockerKeys = Get-MgInformationProtectionBitlockerRecoveryKey -All -ErrorAction Stop | 
        ForEach-Object {
            try {
                # Get the actual recovery key
                $recoveryKey = (Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $_.Id -Property key -ErrorAction Stop).key
                
                [PSCustomObject]@{
                    KeyId = $_.Id
                    DeviceId = $_.DeviceId
                    CreatedDateTime = $_.CreatedDateTime
                    VolumeType = $_.VolumeType
                    Key = $recoveryKey
                }
            }
            catch {
                Write-Warning "Could not retrieve key for ID: $($_.Id). Error: $_"
                # Return object with null key if we couldn't retrieve it
                [PSCustomObject]@{
                    KeyId = $_.Id
                    DeviceId = $_.DeviceId
                    CreatedDateTime = $_.CreatedDateTime
                    VolumeType = $_.VolumeType
                    Key = $null
                }
            }
        }
    
    Write-Host "Retrieved $($BitLockerKeys.Count) BitLocker keys" -ForegroundColor Green
    
    # Step 5: Get all device information using advanced filtering
    Write-Host "Retrieving comprehensive device information..." -ForegroundColor Cyan
    
    # Get all devices from Entra ID
    $AllDevices = Get-MgDevice -All -ErrorAction Stop | 
        Select-Object Id, DeviceId, DisplayName, OperatingSystem, OperatingSystemVersion, 
                      TrustType, ApproximateLastSignInDateTime, IsCompliant, AccountEnabled,
                      Manufacturer, Model
    
    Write-Host "Retrieved information for $($AllDevices.Count) devices" -ForegroundColor Green
    
    # Step 6: Create a device lookup table for faster searches
    $DeviceLookup = @{}
    foreach ($device in $AllDevices) {
        # Store by both Id and DeviceId for faster lookups
        $DeviceLookup[$device.Id] = $device
        if ($device.DeviceId) {
            $DeviceLookup[$device.DeviceId] = $device
        }
    }
    
    # Step 7: Join data and create comprehensive report with improved device matching
    Write-Host "Generating comprehensive report..." -ForegroundColor Cyan
    $Results = @()
    
    foreach ($key in $BitLockerKeys) {
        # First try to find device by DeviceId directly from our lookup table
        $device = $null
        
        if ($key.DeviceId -and $DeviceLookup.ContainsKey($key.DeviceId)) {
            $device = $DeviceLookup[$key.DeviceId]
        }
        
        # If direct lookup fails, try searching by ID using Graph API filtering
        if (-not $device -and $key.DeviceId) {
            try {
                # Try both Id and DeviceId properties in the filter
                $device = Get-MgDevice -Filter "id eq '$($key.DeviceId)'" -ErrorAction SilentlyContinue
                
                if (-not $device) {
                    # Try using deviceId property
                    $device = Get-MgDevice -Filter "deviceId eq '$($key.DeviceId)'" -ErrorAction SilentlyContinue
                }
                
                # If found, add to lookup table for future references
                if ($device) {
                    $DeviceLookup[$device.Id] = $device
                    $DeviceLookup[$device.DeviceId] = $device
                }
            }
            catch {
                Write-Warning "Could not find device with ID: $($key.DeviceId). Error: $_"
            }
        }
        
        # Create result object with device information if found
        $resultObj = [PSCustomObject]@{
            KeyId = $key.KeyId
            RecoveryKey = $key.Key
            CreatedDateTime = $key.CreatedDateTime
            VolumeType = $key.VolumeType
            KeyAge = if ($key.CreatedDateTime) { 
                [math]::Round((New-TimeSpan -Start $key.CreatedDateTime -End (Get-Date)).TotalDays, 0) 
            } else { 
                "Unknown" 
            }
            DeviceId = $key.DeviceId
            DisplayName = if ($device) { $device.DisplayName } else { "Unknown" }
            OperatingSystem = if ($device) { $device.OperatingSystem } else { "Unknown" }
            OSVersion = if ($device) { $device.OperatingSystemVersion } else { "Unknown" }
            Manufacturer = if ($device) { $device.Manufacturer } else { "Unknown" }
            Model = if ($device) { $device.Model } else { "Unknown" }
            LastSignIn = if ($device) { $device.ApproximateLastSignInDateTime } else { "Unknown" }
            JoinType = if ($device) { $device.TrustType } else { "Unknown" }
            IsCompliant = if ($device) { $device.IsCompliant } else { "Unknown" }
            AccountEnabled = if ($device) { $device.AccountEnabled } else { "Unknown" }
        }
        
        $Results += $resultObj
    }
    
    # Step 8: Export to CSV
    Write-Host "Exporting report to $OutputPath..." -ForegroundColor Cyan
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -ErrorAction Stop
    
    # Step 9: Display summary and additional diagnostic info
    Write-Host "Report successfully generated!" -ForegroundColor Green
    Write-Host "Total devices with BitLocker keys: $($Results | Select-Object -Unique DeviceId | Measure-Object | Select-Object -ExpandProperty Count)" -ForegroundColor Green
    Write-Host "Total BitLocker keys: $($Results.Count)" -ForegroundColor Green
    
    # Diagnostic information
    Write-Host "`nDiagnostic Information:" -ForegroundColor Yellow
    Write-Host "Total devices in Entra ID: $($AllDevices.Count)" -ForegroundColor Yellow
    Write-Host "Devices with BitLocker keys where device details were found: $(($Results | Where-Object { $_.DisplayName -ne 'Unknown' } | Select-Object -Unique DeviceId | Measure-Object).Count)" -ForegroundColor Yellow
    Write-Host "Devices with BitLocker keys where device details were not found: $(($Results | Where-Object { $_.DisplayName -eq 'Unknown' } | Select-Object -Unique DeviceId | Measure-Object).Count)" -ForegroundColor Yellow
    Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
    
    # Optional: Open the CSV file
    Write-Host "Would you like to open the report now? (Y/N)" -ForegroundColor Yellow
    $openFile = Read-Host
    if ($openFile -eq "Y" -or $openFile -eq "y") {
        Invoke-Item $OutputPath
    }
    
} catch {
    Write-Error "An error occurred: $_"
    
    # Additional error diagnosis
    if ($_.Exception.Message -like "*unauthorized*" -or $_.Exception.Message -like "*access denied*") {
        Write-Host "This appears to be a permissions issue. Please ensure you have the appropriate Entra ID role (e.g., Cloud Device Administrator or Helpdesk Administrator)." -ForegroundColor Red
    }
    elseif ($_.Exception.Message -like "*not found*") {
        Write-Host "Module or command not found. Ensure you have the Microsoft Graph PowerShell SDK installed:" -ForegroundColor Red
        Write-Host "Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    }
    
    # Exit with error code
    exit 1
}