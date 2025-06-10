# PowerShell script to process the OpticsPlant email forwarders data
# This script transforms the forwarders data into a more structured format

param (
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "opticsplanet_forwarders_processed.xlsx"
)

# Check if Excel is available (for better output formatting)
$excelAvailable = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excelAvailable = $true
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel
}
catch {
    Write-Host "Excel not available. Will output to CSV format instead." -ForegroundColor Yellow
}

# Import the data
Write-Host "Importing data from $InputFile..."
$rawData = Import-Csv -Path $InputFile -Header "ForwarderRule" -Encoding UTF8

# Parse the forwarder data into structured format
$forwarders = @()
foreach ($row in $rawData) {
    $rule = $row.ForwarderRule
    
    # Skip empty rows
    if ([string]::IsNullOrWhiteSpace($rule)) { continue }
    
    # Parse the rule - format is typically "source: destination"
    if ($rule -match "^(.+?):(.+)$") {
        $source = $matches[1].Trim()
        $destination = $matches[2].Trim()
        
        # Determine if destination is internal (Office.OpticsPlant.com) or external
        $destinationType = "External"
        if ($destination -match "@office\.opticsplanet\.com") {
            $destinationType = "Internal Office"
        }
        elseif ($destination -match "@ecentria\.com") {
            $destinationType = "Ecentria"
        }
        elseif ($destination -match "@opticsplanet\.com") {
            $destinationType = "OpticsPlant"
        }
        
        # Create a custom object for this forwarder
        $forwarderObject = [PSCustomObject]@{
            SourceEmail = $source
            DestinationEmail = $destination
            DestinationType = $destinationType
            
            # Extract domains for analysis
            SourceDomain = if ($source -match "@(.+)$") { $matches[1] } else { "" }
            DestinationDomain = if ($destination -match "@(.+)$") { $matches[1] } else { "" }
            
            # Extract username parts for analysis
            SourceUsername = if ($source -match "^(.+?)@") { $matches[1] } else { $source }
            DestinationUsername = if ($destination -match "^(.+?)@") { $matches[1] } else { $destination }
        }
        
        # Add to our collection
        $forwarders += $forwarderObject
    }
}

# Generate analysis and statistics
$totalForwarders = $forwarders.Count
$internalForwards = ($forwarders | Where-Object { $_.DestinationType -eq "Internal Office" }).Count
$externalForwards = $totalForwarders - $internalForwards
$uniqueDestinations = ($forwarders | Select-Object -Property DestinationEmail -Unique).Count
$uniqueDestinationDomains = ($forwarders | Select-Object -Property DestinationDomain -Unique).Count

# Get top destination domains
$topDestinationDomains = $forwarders | Group-Object -Property DestinationDomain | 
                         Sort-Object -Property Count -Descending | 
                         Select-Object -First 10 -Property Name, Count

# Display summary
Write-Host "`n===== OpticsPlant Email Forwarders Analysis =====" -ForegroundColor Green
Write-Host "Total forwarders: $totalForwarders" -ForegroundColor Cyan
Write-Host "Forwarded to internal addresses: $internalForwards" -ForegroundColor Cyan
Write-Host "Forwarded to external addresses: $externalForwards" -ForegroundColor Cyan
Write-Host "Unique destination addresses: $uniqueDestinations" -ForegroundColor Cyan
Write-Host "Unique destination domains: $uniqueDestinationDomains" -ForegroundColor Cyan

Write-Host "`n===== Top Destination Domains =====" -ForegroundColor Green
foreach ($domain in $topDestinationDomains) {
    Write-Host "$($domain.Name): $($domain.Count) forwarders" -ForegroundColor Cyan
}

# Export the processed data
$csvOutputPath = $OutputFile -replace '\.xlsx$', '.csv'
$forwarders | Export-Csv -Path $csvOutputPath -NoTypeInformation
Write-Host "`nData exported to CSV: $csvOutputPath" -ForegroundColor Green

# If Excel is available, create a more detailed Excel report
if ($excelAvailable) {
    try {
        Write-Host "Creating Excel report..." -ForegroundColor Yellow
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        
        $workbook = $excel.Workbooks.Add()
        
        # Main forwarders sheet
        $mainSheet = $workbook.Worksheets.Item(1)
        $mainSheet.Name = "Forwarders"
        
        # Add headers
        $mainSheet.Cells.Item(1, 1) = "Source Email"
        $mainSheet.Cells.Item(1, 2) = "Destination Email" 
        $mainSheet.Cells.Item(1, 3) = "Destination Type"
        $mainSheet.Cells.Item(1, 4) = "Source Username"
        $mainSheet.Cells.Item(1, 5) = "Source Domain"
        
        # Format headers
        $headerRange = $mainSheet.Range("A1:E1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Add data
        for ($i = 0; $i -lt $forwarders.Count; $i++) {
            $row = $i + 2
            $mainSheet.Cells.Item($row, 1) = $forwarders[$i].SourceEmail
            $mainSheet.Cells.Item($row, 2) = $forwarders[$i].DestinationEmail
            $mainSheet.Cells.Item($row, 3) = $forwarders[$i].DestinationType
            $mainSheet.Cells.Item($row, 4) = $forwarders[$i].SourceUsername
            $mainSheet.Cells.Item($row, 5) = $forwarders[$i].SourceDomain
            
            # Color code by destination type
            if ($forwarders[$i].DestinationType -eq "External") {
                $mainSheet.Range("C$row").Interior.ColorIndex = 6 # Yellow
            }
        }
        
        # Add summary sheet
        $summarySheet = $workbook.Worksheets.Add()
        $summarySheet.Name = "Summary"
        
        $summarySheet.Cells.Item(1, 1) = "OpticsPlant Email Forwarders - Summary"
        $summarySheet.Range("A1").Font.Bold = $true
        $summarySheet.Range("A1").Font.Size = 14
        
        $summarySheet.Cells.Item(3, 1) = "Total forwarders:"
        $summarySheet.Cells.Item(3, 2) = $totalForwarders
        
        $summarySheet.Cells.Item(4, 1) = "Internal forwards:"
        $summarySheet.Cells.Item(4, 2) = $internalForwards
        
        $summarySheet.Cells.Item(5, 1) = "External forwards:"
        $summarySheet.Cells.Item(5, 2) = $externalForwards
        
        $summarySheet.Cells.Item(6, 1) = "Unique destinations:"
        $summarySheet.Cells.Item(6, 2) = $uniqueDestinations
        
        $summarySheet.Cells.Item(8, 1) = "Top Destination Domains"
        $summarySheet.Range("A8").Font.Bold = $true
        
        $summarySheet.Cells.Item(9, 1) = "Domain"
        $summarySheet.Cells.Item(9, 2) = "Count"
        $summarySheet.Range("A9:B9").Font.Bold = $true
        
        for ($i = 0; $i -lt $topDestinationDomains.Count; $i++) {
            $row = $i + 10
            $summarySheet.Cells.Item($row, 1) = $topDestinationDomains[$i].Name
            $summarySheet.Cells.Item($row, 2) = $topDestinationDomains[$i].Count
        }
        
        # Auto-fit columns
        $mainSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        $summarySheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        
        # Save the workbook
        $workbook.SaveAs($OutputFile)
        $workbook.Close()
        $excel.Quit()
        
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mainSheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($summarySheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "Excel report created successfully: $OutputFile" -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating Excel report: $_" -ForegroundColor Red
        Write-Host "Falling back to CSV output only." -ForegroundColor Yellow
    }
}

Write-Host "`nProcessing complete!" -ForegroundColor Green