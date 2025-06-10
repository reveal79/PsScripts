# Import required modules
Import-Module ActiveDirectory

# Define paths
$inputCsv = "C:\Scripts\Monday.csv"
$outputCsv = "C:\Scripts\Monday-Updated.csv"

# Read the CSV file
$users = Import-Csv -Path $inputCsv

# Process each user
foreach ($user in $users) {
    # Get the email address from the CSV
    $email = $user.Email
    
    # Skip if email is empty
    if ([string]::IsNullOrEmpty($email)) {
        Write-Warning "Empty email found, skipping record"
        continue
    }
    
    # Check if user has "Activated" status but isn't found or is disabled in AD
    $adUserStatus = "Not Found"
    
    # Query AD for the user based on email address
    $adUser = Get-ADUser -Filter "EmailAddress -eq '$email'" -Properties Department, Title, Manager, Enabled
    
    # If user found in AD, update the fields
    if ($adUser) {
        # Check if the AD account is enabled or disabled
        if ($adUser.Enabled -eq $true) {
            $adUserStatus = "Enabled"
        } else {
            $adUserStatus = "Disabled"
        }
        
        # Get manager's display name if manager exists
        $managerName = ""
        if ($adUser.Manager) {
            try {
                # The Manager property contains the Distinguished Name, so we need to query AD again
                $managerDN = $adUser.Manager
                $manager = Get-ADUser -Identity $managerDN -Properties DisplayName
                $managerName = $manager.DisplayName
            } catch {
                Write-Warning "Could not retrieve manager for $email. Error: $_"
            }
        }
        
        # Update user object with AD information
        $user | Add-Member -NotePropertyName 'Department' -NotePropertyValue $adUser.Department -Force
        $user | Add-Member -NotePropertyName 'Job Title' -NotePropertyValue $adUser.Title -Force
        $user | Add-Member -NotePropertyName 'Manager' -NotePropertyValue $managerName -Force
        $user | Add-Member -NotePropertyName 'AD Status' -NotePropertyValue $adUserStatus -Force
        
        # Flag if user shows as Activated in monday.com but is Disabled in AD
        if ($user.'User Status' -like "Activated*" -and $adUserStatus -eq "Disabled") {
            $user | Add-Member -NotePropertyName 'Status Mismatch' -NotePropertyValue "Active in Monday but Disabled in AD" -Force
        } else {
            $user | Add-Member -NotePropertyName 'Status Mismatch' -NotePropertyValue "" -Force
        }
        
        Write-Host "Updated information for $email (AD Status: $adUserStatus)"
    } else {
        Write-Warning "User with email $email not found in Active Directory"
        
        # Add empty properties for consistency
        $user | Add-Member -NotePropertyName 'Department' -NotePropertyValue "" -Force
        $user | Add-Member -NotePropertyName 'Job Title' -NotePropertyValue "" -Force
        $user | Add-Member -NotePropertyName 'Manager' -NotePropertyValue "" -Force
        $user | Add-Member -NotePropertyName 'AD Status' -NotePropertyValue "Not Found" -Force
        
        # Flag if user shows as Activated in monday.com but doesn't exist in AD
        if ($user.'User Status' -like "Activated*") {
            $user | Add-Member -NotePropertyName 'Status Mismatch' -NotePropertyValue "Active in Monday but Not Found in AD" -Force
        } else {
            $user | Add-Member -NotePropertyName 'Status Mismatch' -NotePropertyValue "" -Force
        }
    }
}

# Define the desired properties and their order for output
$properties = @(
    'Email',
    'Department',
    'Job Title',
    'Manager',
    'User Status',
    'User Type',
    'Last active',
    'AD Status',
    'Status Mismatch'
)

# Export the updated list to CSV with the specified column order
$users | Select-Object $properties | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Host "Process completed. Updated data saved to $outputCsv"