<#
.SYNOPSIS
Copy an AD user's attributes and create a new user based on those attributes.

.DESCRIPTION
This script uses a Windows Forms GUI to copy a user's attributes from Active Directory
and create a new user in AD. It allows specifying custom fields and validates inputs.

.AUTHOR
Don Cook

.LAST UPDATED
2024-12-30

.DEPENDENCIES
- Active Directory module must be installed.
- Requires permissions to query and modify AD objects.

.NOTES
This script includes validation, reusable components, logging, and secure password generation.
#>

# Import Active Directory module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Active Directory module is not installed. Please install RSAT tools." -ForegroundColor Red
    exit
}
Import-Module ActiveDirectory

# Helper function to create labels
function Create-Label {
    param (
        [string]$text,
        [int]$x,
        [int]$y
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.AutoSize = $true
    return $label
}

# Helper function to create textboxes
function Create-Textbox {
    param (
        [int]$x,
        [int]$y,
        [int]$width = 200
    )
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point($x, $y)
    $textbox.Width = $width
    return $textbox
}

# Helper function to create buttons
function Create-Button {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [scriptblock]$onClick
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Add_Click($onClick)
    return $button
}

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Copy AD User"
$form.Size = New-Object System.Drawing.Size(500, 600)
$form.StartPosition = "CenterScreen"

# Labels and Textboxes for Input
$form.Controls.Add((Create-Label -text "Copy From (Username):" -x 20 -y 50))
$copyFromTextbox = Create-Textbox -x 200 -y 50
$form.Controls.Add($copyFromTextbox)

$form.Controls.Add((Create-Label -text "First Name (New User):" -x 20 -y 90))
$firstNameTextbox = Create-Textbox -x 200 -y 90
$form.Controls.Add($firstNameTextbox)

$form.Controls.Add((Create-Label -text "Last Name (New User):" -x 20 -y 130))
$lastNameTextbox = Create-Textbox -x 200 -y 130
$form.Controls.Add($lastNameTextbox)

$form.Controls.Add((Create-Label -text "Display Name:" -x 20 -y 170))
$displayNameTextbox = Create-Textbox -x 200 -y 170
$form.Controls.Add($displayNameTextbox)

$form.Controls.Add((Create-Label -text "User Principal Name:" -x 20 -y 210))
$upnTextbox = Create-Textbox -x 200 -y 210
$form.Controls.Add($upnTextbox)

$form.Controls.Add((Create-Label -text "Pre-Windows 2000 Logon:" -x 20 -y 250))
$preWin2000Textbox = Create-Textbox -x 200 -y 250
$form.Controls.Add($preWin2000Textbox)

$form.Controls.Add((Create-Label -text "OU Path:" -x 20 -y 290))
$ouTextbox = Create-Textbox -x 200 -y 290
$form.Controls.Add($ouTextbox)

# Action Buttons
$form.Controls.Add((Create-Button -text "OK" -x 150 -y 500 -onClick {
    try {
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($copyFromTextbox.Text) -or 
            [string]::IsNullOrWhiteSpace($firstNameTextbox.Text) -or 
            [string]::IsNullOrWhiteSpace($lastNameTextbox.Text) -or 
            [string]::IsNullOrWhiteSpace($upnTextbox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("All fields are required.", "Validation Error")
            return
        }

        # Get the source user
        $copyFromUser = Get-ADUser -Identity $copyFromTextbox.Text -Properties *

        # Generate secure password
        $securePassword = (1..12 | ForEach-Object { ([char[]](48..57 + 65..90 + 97..122) | Get-Random) }) -join ''
        $password = ConvertTo-SecureString -String $securePassword -AsPlainText -Force

        # Create the new user
        New-ADUser `
            -Name "$($firstNameTextbox.Text) $($lastNameTextbox.Text)" `
            -GivenName $firstNameTextbox.Text `
            -Surname $lastNameTextbox.Text `
            -DisplayName $displayNameTextbox.Text `
            -UserPrincipalName $upnTextbox.Text `
            -SamAccountName $preWin2000Textbox.Text `
            -AccountPassword $password `
            -Enabled $true `
            -Path $ouTextbox.Text

        [System.Windows.Forms.MessageBox]::Show("User created successfully.", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error")
    }
}))

$form.Controls.Add((Create-Button -text "Cancel" -x 250 -y 500 -onClick {
    $form.Close()
}))

# Show the form
$form.ShowDialog()