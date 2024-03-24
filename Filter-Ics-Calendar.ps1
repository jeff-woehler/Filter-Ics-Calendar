param (
    [string]$SourceFile = $(throw "-SourceFile is required."),
    [string]$DestFile = $(throw "-DestFile is required."),
    [string]$FilterSummary = $(throw "-FilterSummary is required."),
    [string]$StartDate,
    [string]$NewTitle = ""
)

$header = $true
$collect = $true
$saveEvent = @()
$saveFile = @()

$startDateProvided = [datetime]::MinValue
if ($StartDate) {
    $startDateOK = [datetime]::TryParseExact($StartDate, "yyyy-MM-dd", 
                   [System.Globalization.CultureInfo]::InvariantCulture, 
                   [System.Globalization.DateTimeStyles]::None, 
                   [ref]$startDateProvided)
}


$lines = Get-Content -Path $SourceFile
$eventCount = 0
$foundCount = 0
foreach ($line in $lines) {
    if ($line -eq 'BEGIN:VEVENT') {
        $header = $false

        if ($collect) {
            # Collection only started at header of file
            # Write all this to the file.
            $saveFile += $saveEvent
        }

        # Start new event collection
        $saveEvent = @()
        $collect = $true
        $eventCount++
    }
    elseif ($line -like 'X-WR-CALNAME:*') {
        Write-Host 'Found old event name: ' + $line
        if ($NewTitle -ne "") {
            $line = 'X-WR-CALNAME:' + $NewTitle
            Write-Host 'Calendar name changed to: ' + $NewTitle
        }
    }
    elseif ($line -like 'DTSTART:*') {
        # Check if start date filter is provided
        if ($header -ne $true) {
            $startDateFromEvent = [datetime]::MinValue            
            $eventDateOK = [datetime]::TryParseExact(($line -replace 'DTSTART:', ''), "yyyyMMddTHHmmssZ", 
                            [System.Globalization.CultureInfo]::InvariantCulture, 
                            [System.Globalization.DateTimeStyles]::None, 
                            [ref]$startDateFromEvent)
            if ($eventDateOK) {
                # If event's start date is before the provided start date, stop collecting this event
                if ($startDateFromEvent -lt $startDateProvided) {
                    $collect = $false
                    $saveEvent = @()
                }
            }
            else {
                Write-Host "Failed to parse date: $($line -replace 'DTSTART:', '')"
            }
        }
    }
    elseif ($line -like 'SUMMARY:*') {
        # Found a Summary line. Test filter.
        if ($line -like ('*' + $FilterSummary + '*')) {
            Write-Host 'Found event: ' $line
        }
        else {
            # Filter failed. Stop collecting this event.
            $collect = $false
            $saveEvent = @()
        }
    }

    if ($collect) {
        $saveEvent += $line
    }

    if ($line -eq 'END:VEVENT') {
        if ($collect) {
            # Finished event, passed filter. Save to file
            $foundCount++
            $saveFile += $saveEvent
        }
        $collect = $false
        $saveEvent = @()
    }

    if ($line -like 'END:VCALENDAR') {
        $saveFile += $line
    }

}

Write-Host 'Checked ' + $eventCount + ' events, selected ' + $foundCount + ' from filter "' + $FilterSummary + '".'
$saveFile | Out-File $DestFile
