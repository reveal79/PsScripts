# TeleVantage-Forwarder-User-Mapper.ps1
#===============================================================================
# Script Name: TeleVantage-Forwarder-User-Mapper.ps1
# Created On: April 1, 2025
#
# Description:
#   This script analyzes TeleVantage extensions to identify forwarding
#   relationships and maps them to actual users for migration planning.
#
# Usage:
#   .\TeleVantage-Forwarder-User-Mapper.ps1
#===============================================================================

# Database connection settings for TeleVantage
$server = "srv-phone-103.usgroup.loc"
$database = "TVDB"
$outputFile = "C:\Temp\TeleVantage_Forwarding_User_Map.xlsx"

try {
    # Load required assemblies
    Add-Type -AssemblyName System.Data

    # Build connection string with TrustServerCertificate
    $connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True"
    
    Write-Host "Connecting to SQL Server..."
    
    # Create connection
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()
    
    # First, find what columns contain user information
    $schemaQuery = "SELECT TOP 1 * FROM ExtensionSettings"
    $schemaCommand = New-Object System.Data.SqlClient.SqlCommand($schemaQuery, $connection)
    $schemaAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($schemaCommand)
    $schemaDataset = New-Object System.Data.DataSet
    $schemaAdapter.Fill($schemaDataset) | Out-Null
    
    # Look for potential name/user fields
    $nameFields = @()
    foreach ($column in $schemaDataset.Tables[0].Columns) {
        if ($column.ColumnName -like "*Name*" -or 
            $column.ColumnName -like "*User*" -or
            $column.ColumnName -like "*Owner*") {
            $nameFields += $column.ColumnName
        }
    }
    
    Write-Host "Found $($nameFields.Count) potential name fields: $($nameFields -join ', ')"
    
    # Choose at least one name field to include
    $nameField = if ($nameFields -contains "FromFirstName") { 
        "FromFirstName + ' ' + FromLastName" 
    } elseif ($nameFields -contains "ToFirstName") { 
        "ToFirstName + ' ' + ToLastName" 
    } elseif ($nameFields.Count -gt 0) { 
        $nameFields[0] 
    } else { 
        "ID" # Fallback
    }
    
    Write-Host "Using $nameField for user names"
    
    # Get all extensions with user info
    $query = "SELECT ID, Number, $nameField AS UserName FROM ExtensionSettings WHERE Number IS NOT NULL ORDER BY Number"
    $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    
    $extensions = $dataset.Tables[0]
    Write-Host "Found $($extensions.Rows.Count) extensions in total"
    
    # Convert to PS objects for easier manipulation
    $extensionList = @()
    foreach ($row in $extensions.Rows) {
        $extensionList += [PSCustomObject]@{
            'ID' = $row['ID']
            'Number' = $row['Number'].ToString().Trim()
            'Length' = $row['Number'].ToString().Trim().Length
            'UserName' = $row['UserName']
        }
    }
    
    # Identify potential forwarding relationships
    $forwardingMap = @()
    
    # Look for 3-digit to 4-digit patterns (e.g., 200 to 1200)
    $threeDigitExtensions = $extensionList | Where-Object { $_.Length -eq 3 }
    
    foreach ($source in $threeDigitExtensions) {
        # Try to find potential matches using common patterns
        $potentialMatches = @()
        
        # Pattern 1: Add 1000 (e.g., 200 to 1200)
        $pattern1 = "1" + $source.Number
        $match1 = $extensionList | Where-Object { $_.Number -eq $pattern1 }
        if ($match1) { $potentialMatches += @{Pattern = "Add 1000"; Match = $match1} }
        
        # Pattern 2: Add 5 prefix (e.g., 407 to 5407)
        $pattern2 = "5" + $source.Number
        $match2 = $extensionList | Where-Object { $_.Number -eq $pattern2 }
        if ($match2) { $potentialMatches += @{Pattern = "Add 5 prefix"; Match = $match2} }
        
        # Pattern 3: Add 9 prefix (e.g., 407 to 9407)
        $pattern3 = "9" + $source.Number
        $match3 = $extensionList | Where-Object { $_.Number -eq $pattern3 }
        if ($match3) { $potentialMatches += @{Pattern = "Add 9 prefix"; Match = $match3} }
        
        # Add each potential match to the mapping
        foreach ($potential in $potentialMatches) {
            $forwardingMap += [PSCustomObject]@{
                'SourceID' = $source.ID
                'SourceExtension' = $source.Number
                'SourceUserName' = $source.UserName
                'MappingPattern' = $potential.Pattern
                'TargetID' = $potential.Match.ID
                'TargetExtension' = $potential.Match.Number
                'TargetUserName' = $potential.Match.UserName
                'ForwardingExists' = "Potential"
                'IsForwarder' = $source.UserName -like "*Forward*"
                'RequiresAction' = if ($source.UserName -like "*Forward*") {
                    "Already configured as forwarder"
                } else {
                    "Create forwarding: $($source.Number) to $($potential.Match.Number)"
                }
                'UserHint' = if ($potential.Match.UserName -and $potential.Match.UserName -ne $potential.Match.ID) {
                    "For user: $($potential.Match.UserName)"
                } else {
                    "User info not available"
                }
            }
        }
        
        # If no matches found, note it
        if ($potentialMatches.Count -eq 0) {
            $forwardingMap += [PSCustomObject]@{
                'SourceID' = $source.ID
                'SourceExtension' = $source.Number
                'SourceUserName' = $source.UserName
                'MappingPattern' = "None"
                'TargetID' = $null
                'TargetExtension' = $null
                'TargetUserName' = $null
                'ForwardingExists' = "No"
                'IsForwarder' = $source.UserName -like "*Forward*"
                'RequiresAction' = "No target found"
                'UserHint' = ""
            }
        }
    }
    
    # Close connection
    $connection.Close()
    
    # Export to Excel
    Write-Host "Exporting forwarding map to Excel..."
    Import-Module ImportExcel
    $forwardingMap | Export-Excel -Path $outputFile -AutoSize -TableName "ForwardingMap" -WorksheetName "Extension Forwarding Map"
    
    # Summary
    Write-Host "Summary of Extension Forwarding Map:"
    Write-Host "Total 3-digit extensions analyzed: $($threeDigitExtensions.Count)"
    Write-Host "Potential forwarding relationships found: $(($forwardingMap | Where-Object { $_.ForwardingExists -eq 'Potential' }).Count)"
    Write-Host "Extensions with no forwarding target: $(($forwardingMap | Where-Object { $_.ForwardingExists -eq 'No' }).Count)"
    Write-Host "Existing forwarders identified: $(($forwardingMap | Where-Object { $_.IsForwarder -eq $true }).Count)"
    Write-Host "Mapping exported to: $outputFile"
}
catch {
    Write-Host "❌ Error: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}