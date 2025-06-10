
# Cleanup-GraphModules-NoExit.ps1
# This script disconnects from Microsoft Graph, removes all Microsoft.Graph modules, and reinstalls the necessary ones.
# It does NOT close the PowerShell window or stop any processes.

# Step 1: Disconnect from Microsoft Graph
Write-Host "🔐 Disconnecting from Microsoft Graph..."
Disconnect-MgGraph -ErrorAction SilentlyContinue

# Step 2: Uninstall all Microsoft.Graph modules
Write-Host "📦 Uninstalling all Microsoft.Graph modules..."
Get-InstalledModule Microsoft.Graph* | Uninstall-Module -Force -AllVersions

# Step 3: Clean up residual module folders
Write-Host "🧹 Cleaning up residual module folders..."
Remove-Item "$env:USERPROFILE\Documents\PowerShell\Modules\Microsoft.Graph*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "C:\Program Files\WindowsPowerShell\Modules\Microsoft.Graph*" -Recurse -Force -ErrorAction SilentlyContinue

# Step 4: Reinstall the latest stable Microsoft.Graph module
Write-Host "📥 Reinstalling the latest stable Microsoft.Graph module..."
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Step 5: Reinstall the Microsoft.Graph.Beta.Users.Actions module
Write-Host "📥 Reinstalling the Microsoft.Graph.Beta.Users.Actions module..."
Install-Module Microsoft.Graph.Beta.Users.Actions -Scope CurrentUser -Force

# Step 6: Display success message
Write-Host "✅ Cleanup and reinstallation complete. You can now run your scripts."
