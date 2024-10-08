Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    MFA AdminTool
.SYNOPSIS
    Tool to Check MFA Status and make changes -working logic
.DESCRIPTION
    This tool will help you do basic MFA checks in O365 if you are not a global admin but have Azure P1 license and are assigned a PIM role for MFA Authentication Administrator
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MFA_Admin_Tool                  = New-Object system.Windows.Forms.Form
$MFA_Admin_Tool.ClientSize       = New-Object System.Drawing.Point(575,662)
$MFA_Admin_Tool.text             = "MFA Admin Tool"
$MFA_Admin_Tool.TopMost          = $false
$MFA_Admin_Tool.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

$btnConnect                      = New-Object system.Windows.Forms.Button
$btnConnect.text                 = "Connect MSOnline"
$btnConnect.width                = 130
$btnConnect.height               = 42
$btnConnect.location             = New-Object System.Drawing.Point(5,40)
$btnConnect.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnConnect.ForeColor            = [System.Drawing.ColorTranslator]::FromHtml("#000000")
$btnConnect.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$txtOut                          = New-Object system.Windows.Forms.TextBox
$txtOut.multiline                = $true
$txtOut.width                    = 380
$txtOut.height                   = 160
$txtOut.Anchor                   = 'top,right'
$txtOut.location                 = New-Object System.Drawing.Point(143,17)
$txtOut.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lblConnect                      = New-Object system.Windows.Forms.Label
$lblConnect.text                 = "Disconnected"
$lblConnect.AutoSize             = $false
$lblConnect.width                = 87
$lblConnect.height               = 27
$lblConnect.Anchor               = 'right'
$lblConnect.location             = New-Object System.Drawing.Point(428,181)
$lblConnect.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))
$lblConnect.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#d0021b")

$lblEmail                        = New-Object system.Windows.Forms.Label
$lblEmail.text                   = "Enter Email Address"
$lblEmail.AutoSize               = $true
$lblEmail.width                  = 25
$lblEmail.height                 = 10
$lblEmail.location               = New-Object System.Drawing.Point(9,207)
$lblEmail.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$txtUser                         = New-Object system.Windows.Forms.TextBox
$txtUser.multiline               = $false
$txtUser.width                   = 167
$txtUser.height                  = 20
$txtUser.Anchor                  = 'left'
$txtUser.location                = New-Object System.Drawing.Point(7,227)
$txtUser.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#$ToolTip1                        = New-Object system.Windows.Forms.ToolTip

$txtStatus                       = New-Object system.Windows.Forms.TextBox
$txtStatus.multiline             = $true
$txtStatus.width                 = 380
$txtStatus.height                = 160
$txtStatus.Anchor                = 'left'
$txtStatus.location              = New-Object System.Drawing.Point(4,273)
$txtStatus.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnStatus                       = New-Object system.Windows.Forms.Button
$btnStatus.text                  = "Check Status"
$btnStatus.width                 = 130
$btnStatus.height                = 30
$btnStatus.Anchor                = 'right'
$btnStatus.location              = New-Object System.Drawing.Point(396,234)
$btnStatus.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnStatus.BackColor             = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnEnable                       = New-Object system.Windows.Forms.Button
$btnEnable.text                  = "Enable MFA"
$btnEnable.width                 = 130
$btnEnable.height                = 30
$btnEnable.location              = New-Object System.Drawing.Point(395,270)
$btnEnable.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnEnable.BackColor             = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnDisable                      = New-Object system.Windows.Forms.Button
$btnDisable.text                 = "Disable MFA"
$btnDisable.width                = 130
$btnDisable.height               = 30
$btnDisable.Anchor               = 'right'
$btnDisable.location             = New-Object System.Drawing.Point(396,353)
$btnDisable.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnDisable.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnReset                        = New-Object system.Windows.Forms.Button
$btnReset.text                   = "Reset MFA"
$btnReset.width                  = 130
$btnReset.height                 = 30
$btnReset.Anchor                 = 'right'
$btnReset.location               = New-Object System.Drawing.Point(395,392)
$btnReset.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnReset.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnRevoke                       = New-Object system.Windows.Forms.Button
$btnRevoke.text                  = "Revoke MFA"
$btnRevoke.width                 = 130
$btnRevoke.height                = 30
$btnRevoke.Anchor                = 'right'
$btnRevoke.location              = New-Object System.Drawing.Point(395,433)
$btnRevoke.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnRevoke.BackColor             = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnClear                        = New-Object system.Windows.Forms.Button
$btnClear.text                   = "Clear Fields"
$btnClear.width                  = 140
$btnClear.height                 = 31
$btnClear.location               = New-Object System.Drawing.Point(390,537)
$btnClear.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnClear.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnClose                        = New-Object system.Windows.Forms.Button
$btnClose.text                   = "Close"
$btnClose.width                  = 60
$btnClose.height                 = 30
$btnClose.location               = New-Object System.Drawing.Point(15,535)
$btnClose.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnClose.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$btnEnforce                      = New-Object system.Windows.Forms.Button
$btnEnforce.text                 = "Enforce MFA"
$btnEnforce.width                = 130
$btnEnforce.height               = 30
$btnEnforce.location             = New-Object System.Drawing.Point(395,311)
$btnEnforce.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$btnEnforce.BackColor            = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

#$None.SetToolTip($MFA_Admin_Tool,'Enter Email Address')
#$None.SetToolTip($txtUser,'Enter Email Address')
$MFA_Admin_Tool.controls.AddRange(@($btnConnect,$txtOut,$lblConnect,$lblEmail,$txtUser,$txtStatus,$btnStatus,$btnEnable,$btnDisable,$btnReset,$btnRevoke,$btnClear,$btnClose,$btnEnforce))

$btnConnect.Add_Click({ Connect-Msol })
$btnStatus.Add_Click({ Check_MFA })
$btnEnable.Add_Click({ Enable_MFA })
$btnDisable.Add_Click({ Disable_MFA })
$btnReset.Add_Click({ Reset_MFA })
$btnRevoke.Add_Click({ Revoke_MFA })
$btnClear.Add_Click({ Clear_Form })
$btnClose.Add_Click({ CloseForm })
$btnEnforce.Add_Click({ Enforce_MFA })

function closeForm
{
    $MFA_Admin_Tool.Close();
}
function Clear_Form { 
    $txtStatus.Clear(); 
    $txtUser.Clear();
}

function Connect-Msol { 
    
    $msonline = Get-InstalledModule -Name MSOnline -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

    $txtOut.text = "Connecting to all Services `r`n" 

    # Import Modules or Install if not found
    if (-not $msonline) {
        $txtOut.text += "Installing Required Module MSOnline `r`n"
        Set-PSRepository PSGallery -InstallationPolicy Trusted
        Install-Module MSOnline -Repository PSGallery -Confirm:$False -Force -Scope CurrentUser
        $txtOut.text += "The following Module have been installed `r`n $($msonline)"
    } else {
        $txtOut.text += "Loaded Module `r`n"
        Import-Module MSOnline
    }
}

    
    Connect-MsolService  
    

    $txtOut.text = "Connected to: Service & Required Module Installed"
    $lblConnect.BackColor = "#417505"
    $lblConnect.text = "Connected"


function Check_MFA {
# This will first check the MFA status defined by method type SMS, Voice, PhoneApp, the State returns Enforced, or Enabled, if it's blank its disabled    
$txtStatus.Clear();
$MFA_state = get-msoluser -UserPrincipalName $txtUser.Text -erroraction silentlycontinue # search for a user continue on if error
        if($MFA_state.count -LT 1) { # If statement to check if what is entered is less than 1 character
                $txtStatus.Text = "User $($txtUser.text) does not exist" # Display message if user doesn't exist
                $txtUser.Clear(); # Clear Username Field

} else {

    $CheckEnabled = Get-MsolUser -UserPrincipalName $txtUser.Text | Select-Object DisplayName,UserPrincipalName,@{N="MFA Status"; E={ if( $_.StrongAuthenticationMethods.IsDefault -eq $true) {($_.StrongAuthenticationMethods | Where-object IsDefault -eq $True).MethodType} else { "Disabled"}}}
                            
                $txtStatus.Text = $txtStatus.text + "MFA Status: $($CheckEnabled.'MFA Status')`r`n" # Display MFA Status of device used
                $txtStatus.text = $txtStatus.text + "MFA State: $($Mfa_State.StrongAuthenticationRequirements.State)`r`n" # Display MFA State ie: Enabled, Enforced, or blank for Disabled
                $txtStatus.Text = $txtStatus.text + "If MFA State is blank it means disabled" #If MFAT State is blank mention its disabled


}

}

function Enable_MFA { 
$txtStatus.Clear();
    $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $st.RelyingParty = "*"
    $st.State = "Enabled"
    $sta = @($st)
                
                Set-MsolUser -UserPrincipalName $txtUser.text -StrongAuthenticationRequirements $sta # Sets the State flag to enabled
                $Check = Get-MsolUser -UserPrincipalName $txtUser.Text
                
                $txtStatus.Text = $txtStatus.Text + "Enabling MFA For: $($txtUser.Text)`r`n" 
                $txtStatus.Text = $txtStatus.Text + "Enabled Flag: $($Check.StrongAuthenticationRequirements.State)`r`n" # Return the flag Enabled
}


function Disable_MFA { 
$txtStatus.Clear();
        Set-MsolUser -UserPrincipalName $txtUser.text -StrongAuthenticationRequirements @() # Disables MFA by removing the state

        $txtStatus.Text = "User: $($txtUser.text) MFA has been DISABLED"

}

function Reset_MFA { 
$txtStatus.Clear();
        Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName $txtUser.text # Re-quires the user to re-register in MFA

        $txtStatus.Text = "User: $($txtUser.text) MFA has been RESET"
}

function Revoke_MFA {
$txtStatus.Clear(); 
        
        Revoke-AzureADUserAllRefreshToken -ObjectId $txtUser.text # Revoke all tokens and session for user
        
        $txtStatus.Text = "Revoking all tokens for user $($txtUser.text)"
}
function Enforce_MFA { 
    $txtStatus.Clear();
    $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $st.RelyingParty = "*"
    $st.State = "Enforced"
    $sta = @($st)
                
                Set-MsolUser -UserPrincipalName $txtUser.text -StrongAuthenticationRequirements $sta # Sets the State flag to enabled
                $Check = Get-MsolUser -UserPrincipalName $txtUser.Text
                
                $txtStatus.Text = $txtStatus.Text + "Enforcing MFA For: $($txtUser.Text)`r`n" 
                $txtStatus.Text = $txtStatus.Text + "Enforced Flag: $($Check.StrongAuthenticationRequirements.State)`r`n" # Return the flag Enabled
}
[void]$MFA_Admin_Tool.ShowDialog()
