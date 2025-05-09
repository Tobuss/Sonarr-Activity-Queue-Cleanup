# ==== CONFIGURATION ====
$SonarrURL = "http://localhost:8989"     # Change if Sonarr runs elsewhere
$APIKey = "CHANGE"           # Replace with your actual API key
$Headers = @{ "X-Api-Key" = $APIKey }
$MaxAgeDays = 7

# Pull the queue
$queueResponse = Invoke-RestMethod -Uri "$SonarrURL/api/v3/queue?page=1&pageSize=500" -Headers $Headers

$now = Get-Date
$removedCount = 0

foreach ($item in $queueResponse.records) {
    $removeDueToAge = $false
    $removeDueToLnk = $false
    $ageDays = "Unknown"

    # Age check based on "added" field from the API
    if ($item.added) {
        $addedDate = Get-Date $item.added
        $ageDays = ($now - $addedDate).TotalDays
        if ($ageDays -ge $MaxAgeDays) {
            $removeDueToAge = $true
        }
    }

    # .lnk check in outputPath
    $outputPath = $item.outputPath
    if ($outputPath -and $outputPath -like '*.lnk') {
        $removeDueToLnk = $true
    }

    # Remove if either condition is met
    if ($removeDueToAge -or $removeDueToLnk) {
        Write-Host " Removing item: $($item.title) - Age: $([math]::Round($ageDays,2)) days - OutputPath: $outputPath"
        try {
            Invoke-RestMethod -Uri "$SonarrURL/api/v3/queue/$($item.id)?removeFromClient=true" -Method Delete -Headers $Headers
            $removedCount++
        } catch {
            Write-Host "Failed to remove $($item.title): $_"
        }
    }
}

Write-Host "Completed. Total removed: $removedCount"