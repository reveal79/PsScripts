# TeleVantage-Basic-Analysis.ps1
#===============================================================================
# Script Name: TeleVantage-Basic-Analysis.ps1
# Created On: April 7, 2025
#
# Description:
#   Simple extraction of TeleVantage call data for analysis.
#===============================================================================

# Database connection settings
$server = "srv-phone-103.usgroup.loc"
$database = "TVDB"
$outputFile = "C:\Temp\TeleVantage_Basic_Usage.csv"

try {
    # Load required assemblies
    Add-Type -AssemblyName System.Data

    # Build connection string with TrustServerCertificate
    $connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True"
    
    Write-Host "Connecting to SQL Server..."
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()
    
    # Simplified query to get basic call activity
    $query = @"
    SELECT 
        es.Number AS Extension,
        es.DIDNumber,
        COUNT(cl.ID) AS TotalCalls,
        MAX(cl.StartTime) AS LastCallDate,
        SUM(CASE WHEN cl.StartTime >= DATEADD(MONTH, -3, GETDATE()) THEN 1 ELSE 0 END) AS RecentCalls
    FROM ExtensionSettings es
    LEFT JOIN CallLog cl ON es.DIDNumber = cl.DIDNumber
    WHERE es.Number IS NOT NULL
    GROUP BY 
        es.Number,
        es.DIDNumber
    ORDER BY COUNT(cl.ID) DESC
"@
    
    $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    
    $usageData = $dataset.Tables[0]
    Write-Host "Found call activity data for $($usageData.Rows.Count) extensions"
    
    # Close connection
    $connection.Close()
    
    # Export raw data to CSV (no processing)
    Write-Host "Exporting to CSV..."
    $dataTable = $dataset.Tables[0]
    
    # Create a list to hold the results
    $results = @()
    
    # Convert DataTable to custom objects
    foreach ($row in $dataTable.Rows) {
        $obj = New-Object PSObject
        foreach ($column in $dataTable.Columns) {
            $obj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $row[$column.ColumnName]
        }
        $results += $obj
    }
    
    # Export to CSV
    $results | Export-Csv -Path $outputFile -NoTypeInformation
    
    Write-Host "✅ Export complete: $outputFile"
    
    # Display summary
    Write-Host "`nTeleVantage Usage Summary:"
    Write-Host "Total extensions analyzed: $($results.Count)"
    $activeCount = ($results | Where-Object { $_.TotalCalls -gt 0 }).Count
    Write-Host "Extensions with call activity: $activeCount"
    Write-Host "Extensions with recent calls: $(($results | Where-Object { $_.RecentCalls -gt 0 }).Count)"
}
catch {
    Write-Host "❌ Error: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}