$date = Get-Date -Format "yyyyMMdd"
$displayDate = Get-Date -Format "yyyy-MM-dd"
$vaultPath = "C:\Users\kevin.crump\OneDrive - Candid\Documents\DevDocs\Dev Docs"
$devlogDir = Join-Path $vaultPath "devlog"
$fileName = "devlog $date.md"
$filePath = Join-Path $devlogDir $fileName

if (-not (Test-Path $devlogDir)) {
    New-Item -ItemType Directory -Path $devlogDir | Out-Null
}

function Get-TodayMeetings {
    $wasRunning = $null -ne (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)

    try {
        $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $ns = $ol.GetNamespace("MAPI")

        if (-not $wasRunning) {
            Start-Sleep -Seconds 10
        }

        $calendar = $ns.GetDefaultFolder(9)
        $items = $calendar.Items
        $items.IncludeRecurrences = $true
        $items.Sort("[Start]")

        $today = [DateTime]::Today
        $tomorrow = $today.AddDays(1)
        $startStr = $today.ToString("MM/dd/yyyy") + " 12:00 AM"
        $endStr = $tomorrow.ToString("MM/dd/yyyy") + " 12:00 AM"
        $filter = "[Start] >= '" + $startStr + "' AND [Start] < '" + $endStr + "' AND [AllDayEvent] = False"
        $filtered = $items.Restrict($filter)

        $lines = @()
        foreach ($m in $filtered) {
            $start = $m.Start.ToString("h:mm tt")
            $end = $m.End.ToString("h:mm tt")
            $lines += "- " + $start + " - " + $end + "  " + $m.Subject
        }
        return $lines
    }
    catch {
        return $null
    }
}

function Build-MeetingsBlock {
    param ($meetings)
    $block = "## Meetings`n`n"
    if ($meetings.Count -gt 0) {
        $block += ($meetings -join "`n") + "`n"
    }
    else {
        $block += "_No meetings scheduled_`n"
    }
    return $block
}

if (-not (Test-Path $filePath)) {
    $meetings = Get-TodayMeetings
    if ($null -ne $meetings) {
        $content = "# Devlog - $displayDate`n`n" + (Build-MeetingsBlock $meetings) + "`n"
    }
    else {
        $content = "# Devlog - $displayDate`n`n"
    }
    Set-Content -Path $filePath -Value $content -Encoding UTF8
}
else {
    $existing = Get-Content -Path $filePath -Raw -Encoding UTF8
    if ($existing -notmatch "## Meetings") {
        $meetings = Get-TodayMeetings
        if ($null -ne $meetings) {
            $lines = Get-Content -Path $filePath -Encoding UTF8
            $newContent = @($lines[0], "", (Build-MeetingsBlock $meetings)) + $lines[1..($lines.Length - 1)]
            Set-Content -Path $filePath -Value $newContent -Encoding UTF8
        }
    }
}

$encodedFile = [Uri]::EscapeDataString("devlog/devlog $date")
$uri = 'obsidian://open?vault=Dev%20Docs&file=' + $encodedFile
Start-Process $uri
