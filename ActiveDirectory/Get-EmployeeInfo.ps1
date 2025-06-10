<#
    Script Name: Get-EmployeeInfo.ps1
    Version: v1.0.0
    Author: Don Cook (Adapted from Mike Kanakos)
    Last Updated: 2024-12-30
    Purpose:
    - Retrieves detailed Active Directory account information for a single user.
    - Combines commonly needed information for helpdesk or IT troubleshooting purposes.

    Service: Active Directory
    Service Type: User Management

    Dependencies:
    - Active Directory Module for Windows PowerShell (RSAT tools must be installed).
    - Permissions to query AD user attributes and group memberships.

    Notes:
    - This script retrieves user information and group memberships, formats it into two custom objects, and outputs the results.
    - Includes error handling for invalid or non-existent user accounts.

    Example Usage:
    PS C:\Scripts> Get-EmployeeInfo -UserName Joe_Smith
    Returns detailed account information for the user Joe_Smith.

#>

Function Get-EmployeeInfo {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$UserName
    )

    # Import Active Directory Module
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Error "Active Directory module is not available. Please ensure RSAT tools are installed."
        return
    }
    Import-Module ActiveDirectory -ErrorAction Stop

    # Retrieve user information
    try {
        $Employee = Get-ADUser $UserName -Properties *, 'msDS-UserPasswordExpiryTimeComputed'
    } catch {
        Write-Error "Failed to retrieve information for user $UserName: $_"
        return
    }

    # Retrieve user group memberships
    try {
        $Member = Get-ADUser $UserName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    } catch {
        Write-Warning "Failed to retrieve group memberships for $UserName. Continuing without group membership data."
        $Member = @()
    }

    # Retrieve manager's sAMAccountName
    try {
        $Manager = if ($Employee.manager) {
            (Get-ADUser $Employee.manager).SamAccountName
        } else {
            "No Manager Assigned"
        }
    } catch {
        Write-Warning "Failed to retrieve manager information for $UserName. Continuing without manager data."
        $Manager = "Unknown"
    }

    # Calculate password expiration date
    try {
        $PasswordExpiry = [datetime]::FromFileTime($Employee.'msDS-UserPasswordExpiryTimeComputed')
    } catch {
        $PasswordExpiry = "Unavailable"
    }

    # Create custom objects for account info and status
    $AccountInfo = [PSCustomObject]@{
        FirstName    = $Employee.GivenName
        LastName     = $Employee.Surname
        Title        = $Employee.Title
        Department   = $Employee.Department
        Membership   = $Member
        Manager      = $Manager
        City         = $Employee.City
        UserName     = $Employee.SamAccountName
        DisplayName  = $Employee.DisplayName
        EmailAddress = $Employee.EmailAddress
        OfficePhone  = $Employee.OfficePhone
        MobilePhone  = $Employee.MobilePhone
    }

    $AccountStatus = [PSCustomObject]@{
        PasswordExpired       = $Employee.PasswordExpired
        AccountLockedOut      = $Employee.LockedOut
        LockOutTime           = $Employee.AccountLockoutTime
        AccountEnabled        = $Employee.Enabled
        AccountExpirationDate = $Employee.AccountExpirationDate
        PasswordLastSet       = $Employee.PasswordLastSet
        PasswordExpireDate    = $PasswordExpiry
    }

    # Output the results
    $AccountInfo
    $AccountStatus
} # End of Function