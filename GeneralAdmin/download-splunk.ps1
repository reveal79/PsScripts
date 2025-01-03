function Download-Splunk {
    function Get-URLsFromPage {
        param (
            [string]$Url,
            [string]$Pattern
        )
        $content = (Invoke-WebRequest -Uri $Url -UseBasicParsing).Content
        $matches = Select-String -InputObject $content -Pattern $Pattern -AllMatches
        $urls = $matches.Matches | ForEach-Object { $_.Groups[1].Value.Trim() }
        return $urls
    }

    Write-Host "⏳ Fetching the list of Splunk Enterprise URLs..."
    $splunkEnterpriseURLs = Get-URLsFromPage -Url "https://www.splunk.com/en_us/download/splunk-enterprise.html" -Pattern 'data-link=\"([^\"]+)\"'

    Write-Host "⏳ Fetching the list of Splunk Universal Forwarder URLs..."
    $splunkUFURLs = Get-URLsFromPage -Url "https://www.splunk.com/en_us/download/universal-forwarder.html" -Pattern 'data-link=\"([^\"]+)\"'

    $allURLs = $splunkEnterpriseURLs + $splunkUFURLs

    if ($allURLs.Count -eq 0) {
        Write-Host "❌ No URLs were fetched. Please check the URLs or your internet connection." -ForegroundColor Red
        return
    }

    Write-Host "❓ Please choose a value from the following list:" -ForegroundColor Yellow
    $allURLs | ForEach-Object {
        $index = [array]::IndexOf($allURLs, $_)
        Write-Host "$(($index + 1)). $_"
    }

    while ($true) {
        $choice = Read-Host "Enter the number of your choice (1-$($allURLs.Count))"

        if ($choice -as [int] -and $choice -ge 1 -and $choice -le $allURLs.Count) {
            $selectedURL = $allURLs[$choice - 1]
            $filename = [System.IO.Path]::GetFileName($selectedURL)

            # Check for existing file and its size
            if (Test-Path $filename) {
                $existingSize = (Get-Item $filename).Length
                Write-Host "⏳ Resuming download. File already exists: $filename ($existingSize bytes)" -ForegroundColor Cyan
                $headers = @{ Range = "bytes=$existingSize-" }
                Invoke-WebRequest -Uri $selectedURL -OutFile $filename -Headers $headers -UseBasicParsing
            } else {
                Write-Host "⤵️ Downloading to current directory: $filename" -ForegroundColor Cyan
                Invoke-WebRequest -Uri $selectedURL -OutFile $filename -UseBasicParsing
            }

            Write-Host "Downloaded `\"$filename`\"" -ForegroundColor Green
            Write-Host "🎉 Done, have a great day!" -ForegroundColor Green
            break
        } else {
            Write-Host "❌ Invalid selection. Please enter a number between 1 and $($allURLs.Count)." -ForegroundColor Red
        }
    }
}

Download-Splunk