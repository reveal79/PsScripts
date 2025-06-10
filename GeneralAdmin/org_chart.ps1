# Define President and CEO user accounts
$president = "Pavel.Shvartsman"
$ceo = "Mark.Levitin"
$outputDir = "OrgCharts_Domains"  # Directory to store individual domain charts

# Create output directory if it doesn't exist
if (!(Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Retrieve enabled AD users and filter out service/admin accounts afterward
$allUsers = Get-ADUser -Filter { enabled -eq $true } -Properties Name, Manager, Department, Title |
    Where-Object { -not ($_.samaccountname -like "sa_*" -or $_.samaccountname -like "a_*") }

# Define mappings for high-level domains
$domainMappings = @{
    "Engineering" = @("*SWD", "*Software Development", "*DevOps", "*Engineering")
    "Operations" = @("Warehouse", "Customer Service", "Operations", "Logistics")
    "Finance" = @("Finance", "Accounting", "Budgeting")
    "Sales and Marketing" = @("Sales", "Marketing", "Business Development")
    "Human Resources" = @("Human Resources", "HR", "Talent Acquisition")
}

# Helper function to determine domain based on title or department
function Get-Domain {
    param ($user)
    foreach ($domain in $domainMappings.Keys) {
        foreach ($keyword in $domainMappings[$domain]) {
            if ($user.Title -like "*$keyword*" -or $user.Department -like "*$keyword*") {
                return $domain
            }
        }
    }
    return "Uncategorized"  # Default if no match found
}

# Process each user and categorize them into domains, generating a .puml file for each
foreach ($domain in $domainMappings.Keys) {
    $outputPath = "$outputDir\OrgChart_$domain.puml"
    $output = "@startuml`n"
    $output += "' Organizational Chart for $domain`n"
    
    # Retrieve top-level manager under the President or CEO for each domain
    $topManagers = $allUsers | Where-Object {
        $_.Manager -eq $president -or $_.Manager -eq $ceo -and (Get-Domain -user $_) -eq $domain
    }
    
    # Process each top manager in the domain
    foreach ($topManager in $topManagers) {
        $output += "`"$domain`" --> `"$topManager.Name`" : leads`n"
        
        # Add all direct reports to the top manager
        $reports = $allUsers | Where-Object { $_.Manager -eq $topManager.DistinguishedName }
        foreach ($report in $reports) {
            $output += "`"$topManager.Name`" --> `"$report.Name`" : reports to`n"
        }
    }

    $output += "@enduml"
    $output | Out-File -FilePath $outputPath -Encoding UTF8

    Write-Host "Organizational chart generated for $domain at $outputPath"
}