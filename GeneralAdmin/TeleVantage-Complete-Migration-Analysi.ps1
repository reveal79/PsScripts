# TeleVantage-Complete-Migration-Analysis.ps1
# Created: April 8, 2025

# Configuration
$outputBasePath = "C:\Temp\TeleVantage-Migration"
$scriptsPath = $PSScriptRoot  # Uses the directory where the script is located

# Ensure output base directory exists
if (-not (Test-Path -Path $outputBasePath)) {
    New-Item -ItemType Directory -Force -Path $outputBasePath | Out-Null
}

# Function to run a migration analysis script
function Invoke-MigrationAnalysisScript {
    param(
        [string]$ScriptName,
        [string]$ScriptPath
    )

    try {
        Write-Host "Running $ScriptName Migration Analysis..." -ForegroundColor Cyan
        
        # Execute the script and capture its output
        $result = & $ScriptPath

        return $result
    }
    catch {
        Write-Host "Error running $ScriptName Migration Analysis: $_" -ForegroundColor Red
        return $null
    }
}

# Main Migration Analysis Function
function Start-TeleVantageMigrationAnalysis {
    # List of migration analysis scripts
    $migrationScripts = @(
        @{
            Name = "Users"
            FileName = "Users-Migration-Analysis.ps1"
        },
        @{
            Name = "Auto Attendants"
            FileName = "AutoAttendants-Migration-Analysis.ps1"
        },
        @{
            Name = "Groups/Workgroups"
            FileName = "Groups-Migration-Analysis.ps1"
        }
    )

    # Results storage
    $analysisResults = @{}

    # Track start time
    $startTime = Get-Date

    # Run each migration analysis script
    foreach ($script in $migrationScripts) {
        $scriptPath = Join-Path -Path $scriptsPath -ChildPath $script.FileName
        
        if (Test-Path $scriptPath) {
            $result = Invoke-MigrationAnalysisScript -ScriptName $script.Name -ScriptPath $scriptPath
            $analysisResults[$script.Name] = $result
        }
        else {
            Write-Host "Script not found: $($script.FileName)" -ForegroundColor Yellow
        }
    }

    # Generate comprehensive summary
    $summaryPath = Join-Path -Path $outputBasePath -ChildPath "Migration_Analysis_Summary.txt"
    
    $summaryContent = @"
TeleVantage Migration Analysis - Comprehensive Summary
======================================================
Generated: $startTime
Total Analysis Time: $((Get-Date) - $startTime)

Summary of Migration Analysis:
------------------------------
$(
    $migrationScripts | ForEach-Object {
        $name = $_.Name
        $result = $analysisResults[$name]
        
        if ($result) {
            $totalEntities = $result.Count
            $activeEntities = ($result | Where-Object { $_.YearCalls -gt 0 }).Count
            $highPriorityEntities = ($result | Where-Object { $_.MigrationPriority -eq 'High' }).Count
            
            @"
- $name
  Total Entries: $totalEntities
  Active Entries: $activeEntities
  High Priority Migrations: $highPriorityEntities
"@
        }
        else {
            @"
- $name
  Analysis Failed or No Data
"@
        }
    }
) -join "`n"

Detailed Reports:
----------------
$(
    $migrationScripts | ForEach-Object {
        $name = $_.Name
        $outputPath = Join-Path -Path $outputBasePath -ChildPath $name
        @"
- $name Migration Report: $outputPath\*_Migration_Report.txt
- $name Mapping CSV: $outputPath\*_Migration_Mapping.csv
"@
    }
) -join "`n"
"@

    # Write summary to file
    $summaryContent | Out-File -FilePath $summaryPath -Encoding UTF8

    # Display final summary
    Write-Host "`nMigration Analysis Complete" -ForegroundColor Green
    Write-Host "Comprehensive Summary: $summaryPath"

    return $analysisResults
}

# Execute the complete migration analysis
Start-TeleVantageMigrationAnalysis