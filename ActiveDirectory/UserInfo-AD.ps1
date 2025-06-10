<#
.SYNOPSIS
User Information Retrieval GUI for Active Directory.

.DESCRIPTION
This script provides a GUI to query user details from Active Directory, allowing ITO members to fetch user information like password expiry, department, title, and more.

.FEATURES
- Dynamic domain controller discovery and storage.
- Provides password reset instructions for expired accounts.
- Logs actions for auditing and debugging.
- Handles missing prerequisites and permissions gracefully.

.NOTES
Author: Don Cook
Last Updated: 2024-12-30
Version: 1.0.0

.REQUIREMENTS
- Active Directory RSAT tools installed.
- Permissions to query Active Directory.

.EXAMPLE
Run the script directly to launch the GUI.

#>

# Ensure the Active Directory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module not found. Please install the RSAT tools for Active Directory." -ForegroundColor Red
    exit
}
Import-Module ActiveDirectory

# Configuration Directory and Log File
$configDir = "$env:ProgramData\ITO_Scripts"
if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory
}
$domainControllerFile = "$configDir\domain_controllers.txt"
$logFile = "$configDir\UserInfoLog.txt"

# Load previously used domain controllers or discover dynamically
$domainControllers = @()
if (Test-Path $domainControllerFile) {
    $domainControllers = Get-Content -Path $domainControllerFile
} else {
    $domainControllers = (Get-ADDomainController -Filter *).HostName
    $domainControllers | Out-File -FilePath $domainControllerFile
}

# Add-Type for GUI components
Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Text = "User Information"

# Username Input
$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Location = New-Object System.Drawing.Size(10, 10)
$labelUsername.Size = New-Object System.Drawing.Size(100, 20)
$labelUsername.Text = "Username:"
$form.Controls.Add($labelUsername)

$inputUsername = New-Object System.Windows.Forms.TextBox
$inputUsername.Location = New-Object System.Drawing.Size(110, 10)
$inputUsername.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($inputUsername)

# Domain Controller Input
$labelDomainController = New-Object System.Windows.Forms.Label
$labelDomainController.Location = New-Object System.Drawing.Size(10, 40)
$labelDomainController.Size = New-Object System.Drawing.Size(120, 20)
$labelDomainController.Text = "Domain Controller:"
$form.Controls.Add($labelDomainController)

$inputDomainController = New-Object System.Windows.Forms.ComboBox
$inputDomainController.Location = New-Object System.Drawing.Size(130, 40)
$inputDomainController.Size = New-Object System.Drawing.Size(230, 20)
$inputDomainController.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$inputDomainController.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$inputDomainController.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$inputDomainController.Items.AddRange($domainControllers)
$form.Controls.Add($inputDomainController)

# Labels for User Info
$labelName = New-Object System.Windows.Forms.Label
$labelName.Location = New-Object System.Drawing.Size(10, 70)
$labelName.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($labelName)

$labelExpiry = New-Object System.Windows.Forms.Label
$labelExpiry.Location = New-Object System.Drawing.Size(10, 100)
$labelExpiry.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($labelExpiry)

$labelPasswordLastSet = New-Object System.Windows.Forms.Label
$labelPasswordLastSet.Location = New-Object System.Drawing.Size(10, 130)
$labelPasswordLastSet.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($labelPasswordLastSet)

# Reset Password Button
$buttonResetPassword = New-Object System.Windows.Forms.Button
$buttonResetPassword.Location = New-Object System.Drawing.Size(10, 160)
$buttonResetPassword.Size = New-Object System.Drawing.Size(150, 20)
$buttonResetPassword.Text = "Reset Password"
$buttonResetPassword.Visible = $false
$form.Controls.Add($buttonResetPassword)

$buttonResetPassword.Add_Click({
    $resetInstructions = "Please follow these instructions to reset your password:" +
                         "`n1. Log in to the password reset portal at https://account.activedirectory.windowsazure.com/ChangePassword.aspx" +
                         "`n2. It requires you to login with your email address and password and MFA prompt." +
                         "`n3. Follow the prompts to reset your password." +
                         "`n4. When you change your password make sure if you are using the VPN, you disconnect and re-connect with your NEW PASSWORD." +
                         "`n5. Lock your computer using Windows Key + L and log in with your NEW PASSWORD."
    [System.Windows.Forms.Clipboard]::SetText($resetInstructions)
    [System.Windows.Forms.MessageBox]::Show("Password reset instructions copied to clipboard.", "Password Reset Instructions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Get Info Button
$buttonGetInfo = New-Object System.Windows.Forms.Button
$buttonGetInfo.Location = New-Object System.Drawing.Size(370, 10)
$buttonGetInfo.Size = New-Object System.Drawing.Size(100, 20)
$buttonGetInfo.Text = "&Get Info"
$form.Controls.Add($buttonGetInfo)

$buttonGetInfo.Add_Click({
    try {
        $user = $inputUsername.Text
        $domainController = $inputDomainController.Text
        if ([string]::IsNullOrWhiteSpace($user)) { throw "Username cannot be empty." }
        if ([string]::IsNullOrWhiteSpace($domainController)) { throw "Domain Controller cannot be empty." }

        $userinfo = Get-ADUser -Server $domainController -Identity $user -Properties msDS-UserPasswordExpiryTimeComputed, PasswordLastSet | Select-Object Name, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, PasswordLastSet

        if ($null -eq $userinfo) { throw "User not found." }

        $labelName.Text = "Name: $($userinfo.Name)"
        $labelExpiry.Text = "Password Expiry: $($userinfo.ExpiryDate)"
        $labelPasswordLastSet.Text = "Password Last Set: $($userinfo.PasswordLastSet)"

        if ($userinfo.ExpiryDate -lt (Get-Date)) { $buttonResetPassword.Visible = $true } else { $buttonResetPassword.Visible = $false }

        # Save the domain controller if new
        if ($inputDomainController.Items -notcontains $domainController) {
            $inputDomainController.Items.Add($domainController)
            $inputDomainController.Items | Out-File -FilePath $domainControllerFile
        }

        Add-Content -Path $logFile -Value "$(Get-Date): Queried user $user on domain controller $domainController"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Clear Button
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Location = New-Object System.Drawing.Size(370, 40)
$buttonClear.Size = New-Object System.Drawing.Size(100, 20)
$buttonClear.Text = "&Clear"
$form.Controls.Add($buttonClear)

$buttonClear.Add_Click({
    $inputUsername.Text = ""
    $inputDomainController.Text = ""
    $labelName.Text = ""
    $labelExpiry.Text = ""
    $labelPasswordLastSet.Text = ""
    $buttonResetPassword.Visible = $false
})

# Close Button
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Location = New-Object System.Drawing.Size(370, 70)
$buttonClose.Size = New-Object System.Drawing.Size(100, 20)
$buttonClose.Text = "&Close"
$form.Controls.Add($buttonClose)

$buttonClose.Add_Click({ $form.Close() })

# Show the Form
$form.ShowDialog()