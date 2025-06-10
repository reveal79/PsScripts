# TeleVantage-DID-Migration-Report-Generator.ps1
#===============================================================================
# Script Name: TeleVantage-DID-Migration-Report-Generator.ps1
# Created By: Don Cook
# Created On: April 1, 2025
# Last Modified: April 1, 2025
#
# Description:
#   This script generates a comprehensive report on TeleVantage phone system usage
#   to assist with migration planning to Microsoft Teams or other platforms.
#   It queries the TeleVantage database to identify:
#   - Active/inactive DIDs and extensions
#   - Call activity patterns (total calls, recent calls, outbound calls, etc.)
#   - Forwarding configurations
#
# The script categorizes extensions based on activity levels:
#   - Active: Used within the last 3 months
#   - Low Usage: Used within the last year
#   - Very Low Usage: Used, but not in the last year
#   - Inactive: No recorded usage
#
# Dependencies:
#   - System.Data assembly
#   - ImportExcel PowerShell module
#
# Parameters:
#   $server - TeleVantage database server
#   $database - TeleVantage database name
#   $outputFile - Path for the Excel report output
#
# Usage:
#   .\TeleVantage-DID-Migration-Report-Generator.ps1
#===============================================================================

# Define SQL query - extended lookback period and focused on actual usage
$query = @"
SELECT 
    es.DIDNumber,
    es.Number AS Extension,
    MAX(cl.StartTime) AS LastAnyCall,
    COUNT(cl.ID) AS TotalCalls,
    SUM(CASE WHEN cl.StartTime >= DATEADD(MONTH, -3, GETDATE()) THEN 1 ELSE 0 END) AS CallsLast3Months,
    SUM(CASE WHEN cl.Direction = 1 THEN 1 ELSE 0 END) AS OutboundCalls,
    SUM(CASE WHEN DATEDIFF(second, cl.StartTime, cl.StopTime) > 30 THEN 1 ELSE 0 END) AS CallsOver30Sec,
    es.ForwardAddressID,
    es.DefaultForwardingID,
    CASE WHEN es.AllowExternalForward = 1 THEN 'Yes' ELSE 'No' END AS AllowExternalForward,
    CASE 
        WHEN es.ForwardAddressID IS NOT NULL OR es.DefaultForwardingID IS NOT NULL THEN 'Yes'
        ELSE 'No'
    END AS HasForwarding
FROM ExtensionSettings es
LEFT JOIN CallLog cl 
    ON es.DIDNumber = cl.DIDNumber
    AND cl.StartTime >= DATEADD(YEAR, -3, GETDATE())
WHERE es.DIDNumber IS NOT NULL OR es.Number IS NOT NULL
GROUP BY 
    es.DIDNumber, 
    es.Number, 
    es.ForwardAddressID, 
    es.DefaultForwardingID, 
    es.AllowExternalForward
ORDER BY CASE WHEN MAX(cl.StartTime) IS NULL THEN 1 ELSE 0 END, MAX(cl.StartTime) DESC;
"@

# Setup export path and SQL connection settings
$server = "srv-phone-103.usgroup.loc"
$database = "TVDB"
$outputFile = "C:\Temp\DID_Migration_Report.xlsx"

# Create directory if it doesn't exist
$outputDir = Split-Path -Path $outputFile -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    Write-Host "Created directory: $outputDir"
}

try {
    # Load required assemblies
    Add-Type -AssemblyName System.Data

    # Build connection string with TrustServerCertificate
    $connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True"
    
    Write-Host "Connecting to SQL Server..."
    
    # Create connection and command
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
    
    # Create adapter and dataset
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    
    # Get data
    Write-Host "Executing query..."
    $connection.Open()
    $adapter.Fill($dataset) | Out-Null
    $connection.Close()
    
    # Get results table
    $results = $dataset.Tables[0]
    
    Write-Host "✅ Retrieved $($results.Rows.Count) records from SQL"
    
    # Clean up the data to avoid System.Object[] issues
    $cleanResults = @()
    foreach ($row in $results.Rows) {
        $cleanResults += [PSCustomObject]@{
            'DIDNumber' = $row['DIDNumber']
            'Extension' = $row['Extension']
            'LastUsed' = $row['LastAnyCall']
            'TotalCalls' = $row['TotalCalls']
            'CallsLast3Months' = $row['CallsLast3Months']
            'OutboundCalls' = $row['OutboundCalls']
            'CallsOver30Sec' = $row['CallsOver30Sec']
            'HasForwarding' = $row['HasForwarding']
            'ForwardAddressID' = $row['ForwardAddressID']
            'DefaultForwardingID' = $row['DefaultForwardingID'] 
            'AllowExternalForward' = $row['AllowExternalForward']
            'Status' = if ([string]::IsNullOrEmpty($row['LastAnyCall'])) {
                'Inactive'
            } elseif ($row['CallsLast3Months'] -gt 0) {
                'Active'
            } elseif ($row['LastAnyCall'] -ge [DateTime]::Now.AddYears(-1)) {
                'Low Usage'
            } else {
                'Very Low Usage'
            }
        }
    }
    
    # Export to Excel using ImportExcel module
    Write-Host "Exporting to Excel..."
    Import-Module ImportExcel
    
    # Export without conditional formatting first
    $cleanResults | Export-Excel -Path $outputFile -AutoSize -TableName "MigrationPlanning" -WorksheetName "Migration Status" 
    
    Write-Host "✅ Export complete: $outputFile"
}
catch {
    Write-Host "❌ Error: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}