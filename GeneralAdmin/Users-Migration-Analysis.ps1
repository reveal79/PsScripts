# Users-Migration-Analysis.ps1
# Created: April 8, 2025

# Configuration
$server = "srv-phone-103.usgroup.loc"
$database = "TVDB"
$outputPath = "C:\Temp\TeleVantage-Migration\Users"

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

# Function to get forwarding details
function Get-ForwardingDetails {
    param($forwardAddressId, $extensionsData)
    
    if ($null -eq $forwardAddressId) {
        return $null
    }
    
    $forwardRow = $extensionsData.Select("PhoneBookID = '$forwardAddressId'")
    
    if ($forwardRow.Count -gt 0) {
        return [PSCustomObject]@{
            ForwardedToExtension = $forwardRow[0]["Extension"]
            ForwardedToName = $forwardRow[0]["FullName"]
            ForwardedToDID = $forwardRow[0]["DIDNumber"]
        }
    }
    
    return $null
}

# Main Users Migration Analysis Function
function Get-UsersMigrationAnalysis {
    try {
        # Load SQL client
        Add-Type -AssemblyName System.Data

        # Connection string
        $connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True"
        
        # Open database connection
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()

        # Comprehensive users query
        $usersQuery = @"
        SELECT 
            pbe.ID as PhoneBookID, 
            pbe.IsDeleted,
            pbe.FirstName + ' ' + pbe.PrimaryName AS FullName,
            pbe.FirstName, 
            pbe.PrimaryName, 
            ex.Number as Extension, 
            ex.DIDNumber,
            ex.HoldMusicStation,
            ter.DeviceID,
            dev.name AS DeviceName,
            ter.terminaltype,
            ex.UserType, 
            ex.status, 
            ex.AllowExternalCallback,
            ex.CallBackEnabled,
            ex.AllowExternalForward,
            ex.ForwardAddressID, 
            ex.NotifyEnabled,
            ex.NotifyAddress,
            ex.IsOperator,
            ex.OperatorID,
            ex.OperatorType
        FROM PhoneBookEntry pbe 
        LEFT JOIN ExtensionSettings ex ON pbe.id = ex.id 
        LEFT JOIN Terminal ter ON ter.MyPermPbeid = pbe.id 
        LEFT JOIN Device dev ON dev.deviceID = ter.deviceID 
        WHERE ex.Number IS NOT NULL AND ex.UserType = '0'
"@

        # Call activity query for Users
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
            FROM ExtensionSettings
            WHERE UserType = '0'
        ) aa ON 
            cl.DIDNumber = aa.CleanDID OR 
            cl.DIDNumber = '1' + aa.CleanDID
        GROUP BY cl.DIDNumber
"@

        # Execute users query
        $usersCommand = New-Object System.Data.SqlClient.SqlCommand($usersQuery, $connection)
        $usersAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($usersCommand)
        $usersDataset = New-Object System.Data.DataSet
        $usersAdapter.Fill($usersDataset) | Out-Null

        # Execute call activity query
        $callActivityCommand = New-Object System.Data.SqlClient.SqlCommand($callActivityQuery, $connection)
        $callActivityAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($callActivityCommand)
        $callActivityDataset = New-Object System.Data.DataSet
        $callActivityAdapter.Fill($callActivityDataset) | Out-Null

        # Get data tables
        $usersData = $usersDataset.Tables[0]
        $callActivityData = $callActivityDataset.Tables[0]

        # Close database connection
        $connection.Close()

        # Process users
        $usersMappings = @()

        foreach ($row in $usersData.Rows) {
            $extension = $row["Extension"].ToString()
            $fullName = $row["FullName"].ToString()
            $originalDid = $row["DIDNumber"].ToString()

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

            # Get forwarding details
            $forwardingInfo = $null
            $forwardedToExtension = $null
            $forwardedToName = $null
            $forwardedToDID = $null

            if ($row["ForwardAddressID"] -ne [System.DBNull]::Value) {
                $forwardingInfo = Get-ForwardingDetails -forwardAddressId $row["ForwardAddressID"] -extensionsData $usersData
                
                if ($forwardingInfo) {
                    $forwardedToExtension = $forwardingInfo.ForwardedToExtension
                    $forwardedToName = $forwardingInfo.ForwardedToName
                    $forwardedToDID = $forwardingInfo.ForwardedToDID
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
                'Name' = $fullName
                'Extension' = $extension
                'OriginalDIDNumber' = $originalDid
                'CleanDIDNumber' = $cleanDid
                'YearCalls' = $totalYearCalls
                'LastCallDate' = ($callDetails | Measure-Object -Property LastCallDate -Maximum).Maximum
                'FirstCallDate' = ($callDetails | Measure-Object -Property FirstCallDate -Minimum).Minimum
                'UsageLevel' = $usageLevel
                'MigrationPriority' = $migrationPriority
                'CallDetailsSummary' = ($didCallDetails -join '; ')
                
                # Additional Details
                'DeviceID' = $row["DeviceID"]
                'DeviceName' = $row["DeviceName"]
                'TerminalType' = $row["terminaltype"]
                'Status' = $row["status"]
                'AllowExternalCallback' = $row["AllowExternalCallback"]
                'CallBackEnabled' = $row["CallBackEnabled"]
                'AllowExternalForward' = $row["AllowExternalForward"]
                'NotifyEnabled' = $row["NotifyEnabled"]
                'NotifyAddress' = $row["NotifyAddress"]
                'IsOperator' = $row["IsOperator"]
                'OperatorType' = $row["OperatorType"]
                'IsDeleted' = $row["IsDeleted"]

                # Forwarding Information
                'ForwardedToExtension' = $forwardedToExtension
                'ForwardedToName' = $forwardedToName
                'ForwardedToDID' = $forwardedToDID
            }

            $usersMappings += $mappingEntry
        }

        # Export to CSV
        $csvPath = Join-Path $outputPath "Users_Migration_Mapping.csv"
        $usersMappings | Export-Csv -LiteralPath $csvPath -NoTypeInformation

        # Generate report
        $reportPath = Join-Path $outputPath "Users_Migration_Report.txt"
        
        # Create report content
        $reportContent = @"
Users Migration Analysis Report
===============================
Generated: $(Get-Date)

Total Users: $($usersMappings.Count)
Active Users (Yearly Calls > 0): $(($usersMappings | Where-Object { $_.YearCalls -gt 0 }).Count)

Usage and Migration Priority:
----------------------------
- High Priority Users: $(($usersMappings | Where-Object { $_.MigrationPriority -eq 'High' }).Count)
- Medium Usage Users: $(($usersMappings | Where-Object { $_.UsageLevel -eq 'Medium' }).Count)
- Low Usage Users: $(($usersMappings | Where-Object { $_.UsageLevel -eq 'Low' }).Count)
- Inactive Users: $(($usersMappings | Where-Object { $_.UsageLevel -eq 'Inactive' }).Count)

Top 10 Active Users:
-------------------
$( 
    ($usersMappings | 
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
        Write-Host "`nUsers Migration Analysis Complete" -ForegroundColor Green
        Write-Host "Detailed Users Mapping: $csvPath"
        Write-Host "Users Migration Report: $reportPath"

        return $usersMappings
    }
    catch {
        Write-Host "Error during Users Migration Analysis: $_" -ForegroundColor Red
        Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        return $null
    }
}

# Execute the Users Migration Analysis
Get-UsersMigrationAnalysis