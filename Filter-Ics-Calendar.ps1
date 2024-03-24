param (
    [string]$SourceFile = $(throw "-SourceFile is required."),
    [string]$DestFile = $(throw "-DestFile is required."),
    [string]$FilterSummary = $(throw "-FilterSummary is required."),    
    [string]$NewTitle = ""
)

$collect = $true
$saveEvent = @()
$saveFile = @()

$lines = Get-Content -Path $SourceFile
$eventCount = 0
$foundCount = 0;
foreach ($line in $lines) {
    if ($line -eq 'BEGIN:VEVENT') {
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
    elseif ($line -like 'SUMMARY:*') {
        # Found a Summary line.  Test filter.
        if ($line -like ('*' + $FilterSummary + '*')) {
            Write-Host 'Found event.  ' + $line
        }
        else {
            # Filter failed.  Stop collecting this event.
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
            $foundCount++;
            $saveFile += $saveEvent
        }
        $collect = $false
        $saveEvent = @()
    }
}

Write-Host 'Checked ' + $eventCount + ' events, selected ' + $foundCount + ' from filter "' + $FilterSummary + '".'
$saveFile | Out-File $DestFile
