# PowerShell script to compare email forwarders with Active Directory
# This script adds a Status column based on AD account existence and status

param (
    [Parameter(Mandatory=$true)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "forwarders_with_ad_status.xlsx",
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController = "",  # Leave empty to use current domain
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDisabled = $false  # Set to true to include detailed disabled status
)

# Ensure Active Directory module is available
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    Write-Host "The Active Directory module is required. Attempting to install..." -ForegroundColor Yellow
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
    catch {
        Write-Error "Cannot import the Active Directory module. Please install RSAT tools."
        Write-Host "Run 'Add-WindowsCapability -Name Rsat.ActiveDirectory* -Online' to install." -ForegroundColor Yellow
        exit 1
    }
}

# Import the Active Directory module
Import-Module ActiveDirectory

# Check if Excel is available for better output formatting
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
    $OutputFile = $OutputFile -replace '\.xlsx$', '.csv'
}

# Function to check if a user exists in Active Directory and is enabled
function Get-ADUserStatus {
    param (
        [string]$Username,
        [string]$DomainController
    )
    
    $adParams = @{
        Filter = "SamAccountName -eq '$Username'"
        Properties = @("Enabled")
        ErrorAction = "SilentlyContinue"
    }
    
    if ($DomainController) {
        $adParams.Server = $DomainController
    }
    
    try {
        $user = Get-ADUser @adParams
        
        if ($user) {
            if ($user.Enabled) {
                return "Active"
            }
            else {
                if ($IncludeDisabled) {
                    return "Disabled"
                }
                else {
                    return "Not Active"
                }
            }
        }
        else {
            return "Not Found"
        }
    }
    catch {
        # Fixed error handling
        Write-Warning "Error checking AD status for $Username`: $($_.Exception.Message)"
        return "Error"
    }
}

# Check if input file exists
if (-not (Test-Path $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    exit 1
}

# Determine input file type and import data
$fileExtension = [System.IO.Path]::GetExtension($InputFile)
$forwarders = @()

if ($fileExtension -eq ".csv") {
    Write-Host "Importing data from CSV file: $InputFile" -ForegroundColor Cyan
    $forwarders = Import-Csv -Path $InputFile
}
elseif ($fileExtension -eq ".xlsx") {
    Write-Host "Importing data from Excel file: $InputFile" -ForegroundColor Cyan
    
    if ($excelAvailable) {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Open($InputFile)
            $sheet = $workbook.Sheets.Item(1)
            
            # Find the last row with data
            $lastRow = $sheet.UsedRange.Rows.Count
            
            # Get headers
            $headers = @()
            $colIndex = 1
            while ($sheet.Cells.Item(1, $colIndex).Value2 -ne $null) {
                $headers += $sheet.Cells.Item(1, $colIndex).Value2
                $colIndex++
            }
            
            # Read data rows
            for ($row = 2; $row -le $lastRow; $row++) {
                $rowData = [ordered]@{}
                for ($col = 1; $col -le $headers.Count; $col++) {
                    $rowData[$headers[$col-1]] = $sheet.Cells.Item($row, $col).Value2
                }
                $forwarders += [PSCustomObject]$rowData
            }
            
            # Clean up Excel objects
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        catch {
            Write-Error "Failed to read Excel file: $_"
            exit 1
        }
    }
    else {
        Write-Error "Excel is required to read .xlsx files but is not available."
        exit 1
    }
}
else {
    Write-Error "Unsupported file format: $fileExtension. Use .csv or .xlsx."
    exit 1
}

# Verify that the required fields exist
if (-not ($forwarders | Get-Member -MemberType NoteProperty -Name "SourceUsername")) {
    Write-Error "The input file must contain a 'SourceUsername' column."
    exit 1
}

# Process each forwarder and check AD status
Write-Host "Checking Active Directory status for $(($forwarders | Measure-Object).Count) accounts..." -ForegroundColor Cyan
$progressCounter = 0
$totalCount = ($forwarders | Measure-Object).Count

foreach ($forwarder in $forwarders) {
    $progressCounter++
    $percent = [math]::Round(($progressCounter / $totalCount) * 100, 0)
    
    Write-Progress -Activity "Checking AD Status" -Status "$progressCounter of $totalCount" -PercentComplete $percent
    
    # Clean up username (remove any potential domain prefix)
    $username = $forwarder.SourceUsername
    if ($username -match "\\(.+)") {
        $username = $matches[1]
    }
    
    # Check AD status
    $status = Get-ADUserStatus -Username $username -DomainController $DomainController
    
    # Add the status as a property
    $forwarder | Add-Member -MemberType NoteProperty -Name "ADStatus" -Value $status -Force
    
    # For detailed reporting, we'll categorize the overall status
    if ($status -eq "Active") {
        $forwarder | Add-Member -MemberType NoteProperty -Name "Status" -Value "Active" -Force
    }
    else {
        $forwarder | Add-Member -MemberType NoteProperty -Name "Status" -Value "Not Active" -Force
    }
}

Write-Progress -Activity "Checking AD Status" -Completed

# Generate summary statistics
$totalAccounts = ($forwarders | Measure-Object).Count
$activeAccounts = ($forwarders | Where-Object { $_.Status -eq "Active" } | Measure-Object).Count
$inactiveAccounts = $totalAccounts - $activeAccounts

$notFoundAccounts = ($forwarders | Where-Object { $_.ADStatus -eq "Not Found" } | Measure-Object).Count
$disabledAccounts = ($forwarders | Where-Object { $_.ADStatus -eq "Disabled" } | Measure-Object).Count
$errorAccounts = ($forwarders | Where-Object { $_.ADStatus -eq "Error" } | Measure-Object).Count

# Display summary
Write-Host "`n===== Active Directory Status Summary =====" -ForegroundColor Green
Write-Host "Total forwarders analyzed: $totalAccounts" -ForegroundColor Cyan
Write-Host "Active accounts: $activeAccounts ($([math]::Round(($activeAccounts / $totalAccounts) * 100, 1))%)" -ForegroundColor Cyan
Write-Host "Inactive accounts: $inactiveAccounts ($([math]::Round(($inactiveAccounts / $totalAccounts) * 100, 1))%)" -ForegroundColor Cyan

if ($IncludeDisabled) {
    Write-Host "   - Not found in AD: $notFoundAccounts" -ForegroundColor Cyan
    Write-Host "   - Disabled in AD: $disabledAccounts" -ForegroundColor Cyan
    Write-Host "   - Error checking: $errorAccounts" -ForegroundColor Cyan
}

# Export the updated data
$outputExtension = [System.IO.Path]::GetExtension($OutputFile)

if ($outputExtension -eq ".csv" -or -not $excelAvailable) {
    $csvPath = $OutputFile -replace '\.xlsx$', '.csv'
    $forwarders | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nUpdated data exported to CSV: $csvPath" -ForegroundColor Green
}
elseif ($excelAvailable) {
    try {
        Write-Host "`nCreating Excel report..." -ForegroundColor Yellow
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        
        $workbook = $excel.Workbooks.Add()
        $mainSheet = $workbook.Worksheets.Item(1)
        $mainSheet.Name = "Forwarders with AD Status"
        
        # Get all property names
        $properties = ($forwarders | Get-Member -MemberType NoteProperty).Name
        
        # Add headers
        for ($i = 0; $i -lt $properties.Count; $i++) {
            $mainSheet.Cells.Item(1, $i+1) = $properties[$i]
        }
        
        # Format headers
        $headerRange = $mainSheet.Range($mainSheet.Cells.Item(1,1), $mainSheet.Cells.Item(1, $properties.Count))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Add data
        for ($i = 0; $i -lt $forwarders.Count; $i++) {
            $row = $i + 2
            
            for ($j = 0; $j -lt $properties.Count; $j++) {
                $mainSheet.Cells.Item($row, $j+1) = $forwarders[$i].$($properties[$j])
            }
            
            # Color code by status
            $statusCol = $properties.IndexOf("Status") + 1
            if ($statusCol -gt 0) {
                if ($forwarders[$i].Status -eq "Not Active") {
                    $mainSheet.Cells.Item($row, $statusCol).Interior.ColorIndex = 3  # Red
                }
                else {
                    $mainSheet.Cells.Item($row, $statusCol).Interior.ColorIndex = 4  # Green
                }
            }
        }
        
        # Add summary sheet
        $summarySheet = $workbook.Worksheets.Add()
        $summarySheet.Name = "Summary"
        
        $summarySheet.Cells.Item(1, 1) = "Email Forwarders AD Status - Summary"
        $summarySheet.Range("A1").Font.Bold = $true
        $summarySheet.Range("A1").Font.Size = 14
        
        $summarySheet.Cells.Item(3, 1) = "Total forwarders:"
        $summarySheet.Cells.Item(3, 2) = $totalAccounts
        
        $summarySheet.Cells.Item(4, 1) = "Active accounts:"
        $summarySheet.Cells.Item(4, 2) = $activeAccounts
        $summarySheet.Cells.Item(4, 3) = "($([math]::Round(($activeAccounts / $totalAccounts) * 100, 1))%)"
        
        $summarySheet.Cells.Item(5, 1) = "Inactive accounts:"
        $summarySheet.Cells.Item(5, 2) = $inactiveAccounts
        $summarySheet.Cells.Item(5, 3) = "($([math]::Round(($inactiveAccounts / $totalAccounts) * 100, 1))%)"
        
        if ($IncludeDisabled) {
            $summarySheet.Cells.Item(7, 1) = "Detailed Breakdown:"
            $summarySheet.Range("A7").Font.Bold = $true
            
            $summarySheet.Cells.Item(8, 1) = "Not found in AD:"
            $summarySheet.Cells.Item(8, 2) = $notFoundAccounts
            
            $summarySheet.Cells.Item(9, 1) = "Disabled in AD:"
            $summarySheet.Cells.Item(9, 2) = $disabledAccounts
            
            $summarySheet.Cells.Item(10, 1) = "Error checking:"
            $summarySheet.Cells.Item(10, 2) = $errorAccounts
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
        
        # Fall back to CSV if Excel export fails
        $csvPath = $OutputFile -replace '\.xlsx$', '.csv'
        $forwarders | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Falling back to CSV output: $csvPath" -ForegroundColor Yellow
    }
}

Write-Host "`nProcessing complete!" -ForegroundColor Green