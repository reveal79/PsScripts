<#
    Script: Active Directory User Info Form with Password Reset Instructions and Domain Selection
    Description:
    This script creates a graphical user interface (GUI) form in PowerShell using Windows Forms. 
    It allows you to input a username and optionally a domain controller or global catalog server, retrieve the user’s information 
    from the specified Active Directory domain, and perform various actions such as copying user details, resetting 
    their password, or clearing the form. 

    Key Features:
    - Retrieves a user’s full name, password expiry date, and last password set date from Active Directory.
    - Supports searching across different domains by allowing input for a domain controller or global catalog server.
    - Provides password reset instructions if the password is expired, which can be copied to the clipboard.
    - Copies user information (name, expiry date, and password last set) to the clipboard.
    - Includes options to clear the form and close the application.

    Use Case:
    This script is useful for system administrators who need to search users across a forest or specific domain controllers
    quickly. The script simplifies user data retrieval by providing a user-friendly interface and supports 
    multi-domain environments or global catalog searches.
    
    How to Use:
    - Input the username in the "Username" field.
    - Optionally, enter the Domain Controller (or Global Catalog server) in the "Domain Controller" field. 
      - If you don't provide a domain controller, the script will default to the domain that your machine is currently connected to.
      - Example: You could specify `DC1.yourdomain.com` for domain controller searches, or `gc1.globalcatalog.com` for a global catalog search.
#>

# Add required assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create the main form window
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(600, 500)  # Set form size
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen  # Center form on the screen
$form.Text = "User Information (Cross-Domain Search)"  # Title of the form

# Create label and textbox for entering the username
$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Location = New-Object System.Drawing.Size(10, 10)
$labelUsername.Size = New-Object System.Drawing.Size(100, 20)
$labelUsername.Text = "Username:"

$inputUsername = New-Object System.Windows.Forms.TextBox
$inputUsername.Location = New-Object System.Drawing.Size(110, 10)
$inputUsername.Size = New-Object System.Drawing.Size(200, 20)

# Create label and textbox for entering the domain controller or global catalog server
$labelDomain = New-Object System.Windows.Forms.Label
$labelDomain.Location = New-Object System.Drawing.Size(10, 40)
$labelDomain.Size = New-Object System.Drawing.Size(150, 20)
$labelDomain.Text = "Domain Controller (optional):"

$inputDomain = New-Object System.Windows.Forms.TextBox
$inputDomain.Location = New-Object System.Drawing.Size(160, 40)
$inputDomain.Size = New-Object System.Drawing.Size(250, 20)

# Add username label and input textbox to the form
$form.Controls.Add($labelUsername)
$form.Controls.Add($inputUsername)
$form.Controls.Add($labelDomain)
$form.Controls.Add($inputDomain)

# Create labels to display user information after retrieval
$labelName = New-Object System.Windows.Forms.Label
$labelName.Location = New-Object System.Drawing.Size(10, 70)
$labelName.Size = New-Object System.Drawing.Size(250, 20)

$labelExpiry = New-Object System.Windows.Forms.Label
$labelExpiry.Location = New-Object System.Drawing.Size(10, 100)
$labelExpiry.Size = New-Object System.Drawing.Size(250, 20)

$labelPasswordLastSet = New-Object System.Windows.Forms.Label
$labelPasswordLastSet.Location = New-Object System.Drawing.Size(10, 130)
$labelPasswordLastSet.Size = New-Object System.Drawing.Size(250, 20)

# Create a button to display password reset instructions
$buttonResetPassword = New-Object System.Windows.Forms.Button
$buttonResetPassword.Location = New-Object System.Drawing.Size(10, 160)
$buttonResetPassword.Size = New-Object System.Drawing.Size(150, 20)
$buttonResetPassword.Text = "Reset Password"

# Define action for password reset button (copies instructions to clipboard)
$buttonResetPassword.Add_Click({
    $resetInstructions = "Please follow these instructions to reset your password:" +
                         "`n1. Log in to the password reset portal at https://account.activedirectory.windowsazure.com/ChangePassword.aspx" +
                         "`n2. It requires you to login with your email address and password and MFA prompt" +
                         "`n3. Follow the prompts to reset your password." +
                         "`n4. When you change your password make sure if you are using the VPN, you disconnect and re-connect with your NEW PASSWORD." +
                         "`n5. You can then lock your computer/laptop using the windows key + L to take you to the sign-in screen and login with your NEW PASSWORD"

    [System.Windows.Forms.Clipboard]::SetText($resetInstructions)  # Copy the instructions to the clipboard
    [System.Windows.Forms.MessageBox]::Show("The password reset instructions have been copied to your clipboard.", "Password Reset Instructions", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$buttonResetPassword.Visible = $false  # Hide reset password button by default

# Create button to copy user information to clipboard
$buttonCopy = New-Object System.Windows.Forms.Button
$buttonCopy.Location = New-Object System.Drawing.Size(10, 190)
$buttonCopy.Size = New-Object System.Drawing.Size(150, 20)
$buttonCopy.Text = "Copy Information"

# Define action for the copy button (copies user info to clipboard)
$buttonCopy.Add_Click({
    $information = "Name: " + $labelName.Text + "`n" +
                   "Expiry date: " + $labelExpiry.Text + "`n" +
                   "Password last set: " + $labelPasswordLastSet.Text

    [System.Windows.Forms.Clipboard]::SetText($information)  # Copy the user information to clipboard
    [System.Windows.Forms.MessageBox]::Show("The user information has been copied to your clipboard.", "Information Copied", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

$buttonCopy.Visible = $false  # Hide copy button by default

# Create button to retrieve user information from Active Directory
$buttonGetInfo = New-Object System.Windows.Forms.Button
$buttonGetInfo.Location = New-Object System.Drawing.Size(320, 10)
$buttonGetInfo.Size = New-Object System.Drawing.Size(100, 20)
$buttonGetInfo.Text = "Get Info"

# Define action for the get info button (fetches user info from AD)
$buttonGetInfo.Add_Click({
    $user = $inputUsername.Text
    $domainController = $inputDomain.Text

    # If a domain controller is specified, use it; otherwise, use the default domain
    if ($domainController) {
        # Search using the specified domain controller or global catalog
        $userinfo = Get-ADUser $user -Server $domainController -Properties msDS-UserPasswordExpiryTimeComputed, PasswordLastSet | Select-Object Name, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, PasswordLastSet
    } else {
        # Search using the current domain the machine is connected to
        $userinfo = Get-ADUser $user -Properties msDS-UserPasswordExpiryTimeComputed, PasswordLastSet | Select-Object Name, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, PasswordLastSet
    }

    # Display user information on the form
    $labelName.Text = "Name: " + $userinfo.Name
    $labelExpiry.Text = "Expiry date: " + $userinfo.ExpiryDate
    $labelPasswordLastSet.Text = "Password last set: " + $userinfo.PasswordLastSet

    # Show the reset password button if the password has expired
    if ($userinfo.ExpiryDate -lt (Get-Date)) {
        $buttonResetPassword.Visible = $true
    }
    else {
        $buttonResetPassword.Visible = $false
    }

    $buttonCopy.Visible = $true  # Show the copy button after fetching user information
})

# Create button to clear the form
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Location = New-Object System.Drawing.Size(320, 40)
$buttonClear.Size = New-Object System.Drawing.Size(100, 20)
$buttonClear.Text = "Clear"

# Define action for the clear button (resets all fields)
$buttonClear.Add_Click({
    $inputUsername.Text = ""
    $inputDomain.Text = ""
    $labelName.Text = ""
    $labelExpiry.Text = ""
    $labelPasswordLastSet.Text = ""
    $buttonResetPassword.Visible = $false
    $buttonCopy.Visible = $false
})

# Create button to close the form
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Location = New-Object System.Drawing.Size(320, 70)
$buttonClose.Size = New-Object System.Drawing.Size(100, 20)
$buttonClose.Text = "Close"

# Define action for the close button (closes the form)
$buttonClose.Add_Click({
    $form.Close()
    [System.Windows.Forms.Application]::Exit()
})

# Add all controls (buttons and labels) to the form
$form.Controls.Add($buttonGetInfo)
$form.Controls.Add($buttonClear)
$form.Controls.Add($buttonClose)
$form.Controls.Add($labelName)
$form.Controls.Add($labelExpiry)
$form.Controls.Add($labelPasswordLastSet)
$form.Controls.Add($buttonResetPassword)
$form.Controls.Add($buttonCopy)

# Show the form
$form.ShowDialog()
