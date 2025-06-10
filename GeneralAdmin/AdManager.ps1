Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Main Form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "AD User Manager - e:Centria"
$Form.Size = New-Object System.Drawing.Size(600, 800)
$Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#2a5c6e")

# Define Log File Path
$LogFilePath = "C:\Logs\ADUserManager.log"

# Test Mode Flag
$global:TestMode = $true # Default to Test Mode

# Ensure Log Directory Exists
if (-not (Test-Path -Path (Split-Path -Path $LogFilePath))) {
    New-Item -ItemType Directory -Path (Split-Path -Path $LogFilePath) -Force | Out-Null
}

# Create Label Function
function Create-Label {
    param($Text, $X, $Y, $FontSize, $ForeColor)
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Text
    $Label.AutoSize = $true
    $Label.Location = New-Object System.Drawing.Point($X, $Y)
    $Label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $FontSize)
    $Label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($ForeColor)
    $Label.BackColor = $Form.BackColor
    $Form.Controls.Add($Label)
    return $Label
}

# Create TextBox Function
function Create-TextBox {
    param($X, $Y, $Width)
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point($X, $Y)
    $TextBox.Width = $Width
    $Form.Controls.Add($TextBox)
    return $TextBox
}

# Create Button Function
function Create-Button {
    param($Text, $X, $Y, $Width, $Height, $BackColor, $ClickEvent)
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = $Text
    $Button.Location = New-Object System.Drawing.Point($X, $Y)
    $Button.Size = New-Object System.Drawing.Size($Width, $Height)
    $Button.BackColor = [System.Drawing.ColorTranslator]::FromHtml($BackColor)
    $Button.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
    $Button.Add_Click($ClickEvent)
    $Form.Controls.Add($Button)
    return $Button
}

# Function for Auto-Dismissing Message Box
function Show-AutoDismissMessage {
    param (
        [string]$Message,
        [int]$Timeout = 5,
        [string]$Title = "Notification"
    )

    $wshell = New-Object -ComObject WScript.Shell
    $result = $wshell.Popup($Message, $Timeout, $Title, 0)
}

# Logging Function with Multi-Line Support, File Logging
function Log-Message {
    param (
        [string]$Message,
        [string]$Type = "INFO" # INFO, WARN, ERROR
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $FormattedMessage = "[$Type] ${Timestamp}: $Message`r`n"  # Ensure consistent formatting

    # Append the formatted message to the Output Logs
    $OutputLogs.AppendText("$FormattedMessage")

    # Append the message to the log file
    Add-Content -Path $LogFilePath -Value $FormattedMessage

    # Auto-scroll to the latest entry
    $OutputLogs.SelectionStart = $OutputLogs.Text.Length
    $OutputLogs.ScrollToCaret()
}

# Function to Pull Jira Ticket
function Pull-JiraTicket {
    param (
        [string]$TicketID
    )

    Log-Message -Message "--- Starting new Jira ticket action ---" -Type "INFO"

    if (-not $TicketID) {
        Show-AutoDismissMessage -Message "No Jira ticket ID provided." -Title "Error" -Timeout 5
        Log-Message -Message "Jira ticket ID not provided. Process terminated." -Type "ERROR"
        return
    }

    if ($global:TestMode) {
        # Simulated data for testing
        if ($TicketID -eq "FAKE-1234") {
            $Firstname.Text = "John"
            $Lastname.Text = "Doe"
            $ManagerEmail.Text = "manager@example.com"
            $HREmail.Text = "hr@example.com"

            # Populate logs
            Log-Message -Message "[TEST MODE] Simulated Jira ticket pull for Ticket ID: $TicketID.`r`nFirstname: John`r`nLastname: Doe`r`nManager Email: manager@example.com`r`nHR Email: hr@example.com"

            Show-AutoDismissMessage -Message "[TEST MODE] Jira ticket FAKE-1234 pulled successfully!" -Title "Success" -Timeout 5
        } else {
            $Firstname.Clear()
            $Lastname.Clear()
            $ManagerEmail.Clear()
            $HREmail.Clear()

            Log-Message -Message "[TEST MODE] No data found for Jira ticket ID: $TicketID." -Type "ERROR"
            Show-AutoDismissMessage -Message "[TEST MODE] Failed to pull ticket details for Ticket ID: $TicketID." -Title "Error" -Timeout 5
        }
    } else {
        # Production logic using JiraPS
        try {
            Import-Module JiraPS -ErrorAction Stop

            $jiraCredentials = Get-Credential -Message "Enter your Jira username and password"
            Set-JiraConfigServer -Server "https://jira.ecentria.tools"
            New-JiraSession -Credential $jiraCredentials

            $JiraIssue = Get-JiraIssue -Key $TicketID -ErrorAction Stop

            $Firstname.Text = $JiraIssue.CustomFields.FirstName
            $Lastname.Text = $JiraIssue.CustomFields.LastName
            $ManagerEmail.Text = $JiraIssue.Reporter.EmailAddress
            $HREmail.Text = $JiraIssue.CustomFields.HREmail

            Log-Message -Message "Jira ticket pull for Ticket ID: $TicketID.`r`nFirstname: $($Firstname.Text)`r`nLastname: $($Lastname.Text)`r`nManager Email: $($ManagerEmail.Text)`r`nHR Email: $($HREmail.Text)"

            Show-AutoDismissMessage -Message "Jira ticket $TicketID pulled successfully!" -Title "Success" -Timeout 5
        } catch {
            $Firstname.Clear()
            $Lastname.Clear()
            $ManagerEmail.Clear()
            $HREmail.Clear()

            Log-Message -Message "Failed to pull Jira ticket details for Ticket ID: $TicketID. Error: $_" -Type "ERROR"
            Show-AutoDismissMessage -Message "Failed to pull ticket details for Ticket ID: $TicketID." -Title "Error" -Timeout 5
        }
    }
}

# User Information Section
Create-Label "User Information" 20 20 12 "#FFFFFF"
$Firstname = Create-TextBox 150 100 200
Create-Label "Firstname*" 20 100 10 "#FFFFFF"

$Lastname = Create-TextBox 150 140 200
Create-Label "Lastname*" 20 140 10 "#FFFFFF"

$ManagerEmail = Create-TextBox 150 180 200
Create-Label "Manager Email" 20 180 10 "#FFFFFF"

$HREmail = Create-TextBox 150 220 200
Create-Label "HR Email" 20 220 10 "#FFFFFF"

# Office 365 Needed Checkbox
Create-Label "Office 365 Needed" 20 260 10 "#FFFFFF"
$Office365Checkbox = New-Object System.Windows.Forms.CheckBox
$Office365Checkbox.Location = New-Object System.Drawing.Point(150, 260)
$Office365Checkbox.Text = "Yes"
$Office365Checkbox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
$Form.Controls.Add($Office365Checkbox)

# Output Logs
Create-Label "Output Logs" 20 320 12 "#FFFFFF"
$OutputLogs = New-Object System.Windows.Forms.TextBox
$OutputLogs.Multiline = $true
$OutputLogs.ScrollBars = "Vertical"
$OutputLogs.Location = New-Object System.Drawing.Point(20, 360)
$OutputLogs.Size = New-Object System.Drawing.Size(540, 200)
$OutputLogs.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
$Form.Controls.Add($OutputLogs)

# Add Jira Ticket Field to GUI
$JiraTicket = Create-TextBox 150 60 200
Create-Label "Jira Ticket*" 20 60 10 "#FFFFFF"

# Add Pull Info Button
Create-Button "Pull Info" 370 60 100 30 "#d0caca" {
    Pull-JiraTicket -TicketID $JiraTicket.Text
}

# Mode Display Label
$ModeLabel = Create-Label "Mode: TEST MODE" 20 280 12 "#00FF00"

# Buttons
$ModeToggleButton = Create-Button "Switch to Prod Mode" 400 640 150 30 "#d0caca" {
    if ($global:TestMode) {
        $global:TestMode = $false
        $ModeLabel.Text = "Mode: PRODUCTION MODE"
        $ModeLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FF0000") # Red for Production Mode
        $ModeToggleButton.Text = "Switch to Test Mode"
        Log-Message -Message "Application switched to PRODUCTION MODE." -Type "INFO"
        Show-AutoDismissMessage -Message "Application is now running in PRODUCTION MODE." -Title "Mode Switch" -Timeout 5
    } else {
        $global:TestMode = $true
        $ModeLabel.Text = "Mode: TEST MODE"
        $ModeLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#00FF00") # Green for Test Mode
        $ModeToggleButton.Text = "Switch to Prod Mode"
        Log-Message -Message "Application switched to TEST MODE." -Type "INFO"
        Show-AutoDismissMessage -Message "Application is now running in TEST MODE." -Title "Mode Switch" -Timeout 5
    }
}

Create-Button "Validate Account" 20 560 150 30 "#d0caca" { Log-Message -Message "Validating Account..." }
Create-Button "Generate Password" 200 560 150 30 "#d0caca" { Log-Message -Message "Generated Password: Pass@123" }
Create-Button "Send Notification" 20 600 150 30 "#d0caca" { Show-AutoDismissMessage -Message "Notification Sent Successfully!" -Title "Success" -Timeout 5 }
Create-Button "Export Log" 200 600 150 30 "#d0caca" { Log-Message -Message "Exporting Logs..." }
Create-Button "View Log File" 400 560 150 30 "#d0caca" {
    if (Test-Path -Path $LogFilePath) {
        Start-Process -FilePath "notepad.exe" -ArgumentList $LogFilePath
        Show-AutoDismissMessage -Message "Log file opened successfully!" -Title "Success" -Timeout 5
    } else {
        Show-AutoDismissMessage -Message "Log file does not exist." -Title "Error" -Timeout 5
    }
}
Create-Button "Clear Log File" 400 600 150 30 "#d0caca" {
    if (Test-Path -Path $LogFilePath) {
        Remove-Item -Path $LogFilePath -Force
        Log-Message -Message "Log file cleared by user." -Type "INFO"
        Show-AutoDismissMessage -Message "Log file cleared successfully!" -Title "Success" -Timeout 5
    } else {
        Log-Message -Message "Log file does not exist. Nothing to clear." -Type "WARN"
    }
}
Create-Button "Clear All" 20 640 150 30 "#d0caca" {
    $Firstname.Clear()
    $Lastname.Clear()
    $ManagerEmail.Clear()
    $HREmail.Clear()
    $Office365Checkbox.Checked = $false
    $OutputLogs.Clear()
    Log-Message -Message "All fields cleared by user." -Type "INFO"
}

# Office 365 Checkbox Logic
Create-Button "Apply Settings" 200 640 150 30 "#d0caca" {
    if ($Office365Checkbox.Checked) {
        if (-not $ManagerEmail.Text) {
            Show-AutoDismissMessage -Message "Manager Email is required for Office 365 accounts." -Title "Error" -Timeout 5
            Log-Message -Message "Office 365 setup failed: Manager Email is missing." -Type "ERROR"
        } else {
            Log-Message -Message "Office 365 account setup initiated for $ManagerEmail.Text." -Type "INFO"
        }
    } else {
        Log-Message -Message "No Office 365 account needed for this user." -Type "INFO"
    }
}

# Run the Form
[void]$Form.ShowDialog()