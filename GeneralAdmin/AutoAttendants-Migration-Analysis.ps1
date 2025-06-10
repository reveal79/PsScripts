# AutoAttendants-Migration-Analysis.ps1
# Created: April 8, 2025

# Configuration
$server = "srv-phone-103.usgroup.loc"
$database = "TVDB"
$outputPath = "C:\Temp\TeleVantage-Migration\AutoAttendants"

# Ensure output directory exists
if (-not (Test-Path -Path $outputPath)) {
    New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
}

# Function to deduplicate DIDs
function Remove-DIDDuplicates {
    param([string]$DIDString)
    
    # Split the DID string
    $didArray = $DIDString -split '\|'
    
    # Remove duplicates and remove numbers with '1' prefix if a matching number exists without '1'
    $cleanDids = $didArray | ForEach-Object {
        $currentDid = $_
        # Remove '1' prefix if another matching number exists
        if ($currentDid -match '^1(\d+)$') {
            $baseNumber = $matches[1]
            if ($didArray -contains $baseNumber) {
                return $null
            }
        }
        $currentDid
    } | Select-Object -Unique
    
    # Join the unique DIDs back together
    return ($cleanDids -join '|')
}

# Main Auto Attendants Migration Analysis Function
function Get-AutoAttendantsMigrationAnalysis {
    try {
        # Load SQL client
        Add-Type -AssemblyName System.Data

        # Connection string
        $connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True"
        
        # Open database connection
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()

        # Comprehensive Auto Attendants query
        $autoAttendantsQuery = @"
        SELECT 
            aa.ID AS AutoAttendantID,
            aa.Name AS AutoAttendantName,
            ISNULL(ISNULL(aa.DIDNumber, ex.DIDNumber), '') AS DIDNumber,
            pbe.FirstName + ' ' + pbe.PrimaryName AS FullName,
            ex.Number AS Extension,
            pbe.ID AS PhoneBookID
        FROM AutoAttendant aa
        LEFT JOIN PhoneBookEntry pbe ON aa.PhoneBookID = pbe.ID
        LEFT JOIN ExtensionSettings ex ON pbe.ID = ex.ID
"@

        # Call activity query for Auto Attendants
        $callActivityQuery = @"
        SELECT 
            cl.DIDNumber,
            COUNT(cl.ID) AS TotalCalls,
            SUM(CASE WHEN cl.StartTime >= DATEADD(year, -1, GETDATE()) THEN 1 ELSE 0 END) AS YearCalls,
            MAX(cl.StartTime) AS LastCallDate,
            MIN(cl.StartTime) AS FirstCallDate
        FROM CallLog cl
        INNER JOIN (
            SELECT DISTINCT 
                CASE 
                    WHEN DIDNumber LIKE '1%' THEN SUBSTRING(DIDNumber, 2, LEN(DIDNumber))
                    ELSE DIDNumber 
                END AS CleanDID
            FROM AutoAttendant
        ) aa ON 
            cl.DIDNumber = aa.CleanDID OR 
            cl.DIDNumber = '1' + aa.CleanDID
        GROUP BY cl.DIDNumber
        HAVING COUNT(cl.ID) > 0
"@

        # Execute Auto Attendants query
        $autoAttendantsCommand = New-Object System.Data.SqlClient.SqlCommand($autoAttendantsQuery, $connection)
        $autoAttendantsAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($autoAttendantsCommand)
        $autoAttendantsDataset = New-Object System.Data.DataSet
        $autoAttendantsAdapter.Fill($autoAttendantsDataset) | Out-Null

        # Execute call activity query
        $callActivityCommand = New-Object System.Data.SqlClient.SqlCommand($callActivityQuery, $connection)
        $callActivityAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($callActivityCommand)
        $callActivityDataset = New-Object System.Data.DataSet
        $callActivityAdapter.Fill($callActivityDataset) | Out-Null

        # Get data tables
        $autoAttendantsData = $autoAttendantsDataset.Tables[0]
        $callActivityData = $callActivityDataset.Tables[0]

        # Close database connection
        $connection.Close()

        # Process Auto Attendants
        $autoAttendantsMappings = @()

        foreach ($row in $autoAttendantsData.Rows) {
            $name = $row["AutoAttendantName"].ToString()
            $extension = $row["Extension"].ToString()
            $originalDid = $row["DIDNumber"].ToString()
            $fullName = $row["FullName"].ToString()

            # Deduplicate DIDs
            $cleanDid = Remove-DIDDuplicates -DIDString $originalDid

            # Find call activity for each DID
            $callDetails = @()
            $didCallDetails = @()
            $totalYearCalls = 0
            $cleanDid -split '\|' | ForEach-Object {
                $did = $_
                $activityRow = $callActivityData.Select("DIDNumber = '$did'")
                
                if ($activityRow.Count -gt 0) {
                    $yearCalls = [int]$activityRow[0]["YearCalls"]
                    $totalYearCalls += $yearCalls

                    $didCallDetail = [PSCustomObject]@{
                        DIDNumber = $did
                        YearCalls = $yearCalls
                        LastCallDate = $activityRow[0]["LastCallDate"]
                        FirstCallDate = $activityRow[0]["FirstCallDate"]
                    }
                    $callDetails += $didCallDetail
                    $didCallDetails += "$did($yearCalls)"
                }
            }

            # Determine usage level and migration priority
            $usageLevel = "Inactive"
            if ($totalYearCalls -gt 500) { $usageLevel = "High" }
            elseif ($totalYearCalls -gt 100) { $usageLevel = "Medium" }
            elseif ($totalYearCalls -gt 0) { $usageLevel = "Low" }

            $migrationPriority = "Low"
            if ($usageLevel -in @("High", "Medium")) {
                $migrationPriority = "High"
            }

            # Create mapping entry
            $mappingEntry = [PSCustomObject]@{
                'Name' = $name
                'FullName' = $fullName
                'Extension' = $extension
                'OriginalDIDNumber' = $originalDid
                'CleanDIDNumber' = $cleanDid
                'YearCalls' = $totalYearCalls
                'LastCallDate' = ($callDetails | Measure-Object -Property LastCallDate -Maximum).Maximum
                'FirstCallDate' = ($callDetails | Measure-Object -Property FirstCallDate -Minimum).Minimum
                'UsageLevel' = $usageLevel
                'MigrationPriority' = $migrationPriority
                'CallDetailsSummary' = ($didCallDetails -join '; ')
                'AutoAttendantID' = $row["AutoAttendantID"]
            }

            $autoAttendantsMappings += $mappingEntry
        }

        # Export to CSV
        $csvPath = Join-Path $outputPath "AutoAttendants_Migration_Mapping.csv"
        $autoAttendantsMappings | Export-Csv -LiteralPath $csvPath -NoTypeInformation

        # Generate report
        $reportPath = Join-Path $outputPath "AutoAttendants_Migration_Report.txt"
        
        # Create report content
        $reportContent = @"
Auto Attendants Migration Analysis Report
=========================================
Generated: $(Get-Date)

Total Auto Attendants: $($autoAttendantsMappings.Count)
Active Auto Attendants (Yearly Calls > 0): $(($autoAttendantsMappings | Where-Object { $_.YearCalls -gt 0 }).Count)

Usage and Migration Priority:
----------------------------
- High Priority Auto Attendants: $(($autoAttendantsMappings | Where-Object { $_.MigrationPriority -eq 'High' }).Count)
- Medium Usage Auto Attendants: $(($autoAttendantsMappings | Where-Object { $_.UsageLevel -eq 'Medium' }).Count)
- Low Usage Auto Attendants: $(($autoAttendantsMappings | Where-Object { $_.UsageLevel -eq 'Low' }).Count)
- Inactive Auto Attendants: $(($autoAttendantsMappings | Where-Object { $_.UsageLevel -eq 'Inactive' }).Count)

Top 10 Active Auto Attendants:
-----------------------------
$( 
    ($autoAttendantsMappings | 
        Where-Object { $_.YearCalls -gt 0 } | 
        Sort-Object -Property YearCalls -Descending | 
        Select-Object -First 10 | 
        ForEach-Object { 
            "- $($_.Name) (Ext: $($_.Extension)) | Yearly Calls: $($_.YearCalls) | DID: $($_.CleanDIDNumber)" 
        }) -join "`n"
)
"@

        # Write report to file
        [System.IO.File]::WriteAllText($reportPath, $reportContent)

        # Output final summary
        Write-Host "`nAuto Attendants Migration Analysis Complete" -ForegroundColor Green
        Write-Host "Detailed Auto Attendants Mapping: $csvPath"
        Write-Host "Auto Attendants Migration Report: $reportPath"

        return $autoAttendantsMappings
    }
    catch {
        Write-Host "Error during Auto Attendants Migration Analysis: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        return $null
    }
}

# Execute the Auto Attendants Migration Analysis
Get-AutoAttendantsMigrationAnalysis