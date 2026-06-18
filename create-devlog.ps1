# Combined daily devlog generator for Obsidian.
# Builds a "## Meetings" table (from the published Outlook calendar feed, via
# get-meetings.js) and a "## Tickets" table grouped by sprint (current + next,
# from Jira). Both sections are created if missing and refreshed in place if
# present; any other sections in the note are preserved.

$date = Get-Date -Format "yyyyMMdd"
$monthFolder = Get-Date -Format "yyyyMM"
$titleText = "# " + (Get-Date).ToString("dddd, MMMM d, yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
$vaultPath = "C:\Users\kevin.crump\OneDrive - Candid\Documents\DevDocs\Dev Docs"
$vaultName = "Dev Docs"
$devlogRoot = Join-Path $vaultPath "devlog"
$monthDir = Join-Path $devlogRoot $monthFolder   # per-month folder, e.g. devlog\202606
$fileName = "devlog $date.md"
$filePath = Join-Path $monthDir $fileName

$jiraBase = "https://candidprojects.atlassian.net"
$jiraEmail = "kevin.crump@candid.org"

if (-not (Test-Path $monthDir)) {
    New-Item -ItemType Directory -Path $monthDir -Force | Out-Null
}

function Ensure-AppsRunning {
    param ($vaultName)

    # Obsidian: open the vault so the app is up before we write/show the note.
    # (Outlook is no longer needed; meetings come from the published calendar feed.)
    if (Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue) { return }
    try {
        Start-Process ('obsidian://open?vault=' + [Uri]::EscapeDataString($vaultName))
    }
    catch { return }

    # Wait for the process to appear (up to ~40s), then let the vault finish loading.
    $deadline = (Get-Date).AddSeconds(40)
    while ((Get-Date) -lt $deadline) {
        if (Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue) { break }
        Start-Sleep -Seconds 2
    }
    Start-Sleep -Seconds 10
}

function Finalize-Cell {
    param ($s)
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $s = ($s -replace "[`r`n]", " ") -replace "\s+", " "
    $s = $s.Trim()
    if ($s.Length -gt 100) { $s = $s.Substring(0, 100).Trim() + "..." }
    return ($s -replace "\|", "\|")
}

function Shorten-Location {
    param ($loc)
    if ([string]::IsNullOrWhiteSpace($loc)) { return "" }
    $s = $loc -replace "[`r`n]", " "
    # Replace long Zoom/Teams URLs with a short clickable markdown link, keeping
    # the real URL as the link target. No length truncation here so the link
    # target stays intact.
    $s = [regex]::Replace($s, '(?i)https?://[^\s;]*zoom\.us[^\s;]*', '[Zoom](${0})')
    $s = [regex]::Replace($s, '(?i)https?://[^\s;]*teams\.(microsoft|live)\.com[^\s;]*', '[Teams](${0})')
    $s = $s -replace '\s*;\s*', '; '
    $s = ($s -replace '\s+', ' ').Trim().Trim(';').Trim()
    return ($s -replace '\|', '\|')
}

function Clean-Summary {
    param ($text, $location)

    $body = ""
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $t = $text -replace "[`r`n]", " "

        # Cut the body at the earliest Zoom/Teams join boilerplate marker; the
        # meaningful agenda (if any) sits before it.
        $markers = @(
            'is inviting you to a scheduled zoom meeting',
            'join zoom meeting',
            'microsoft teams meeting',
            'join on your computer',
            'click here to join',
            'join the meeting now',
            'dial by your location',
            'one tap mobile',
            'meeting id:',
            'passcode:',
            '________',
            '----------',
            '=========='
        )
        $lower = $t.ToLower()
        $cut = $t.Length
        foreach ($m in $markers) {
            $idx = $lower.IndexOf($m)
            if ($idx -ge 0 -and $idx -lt $cut) { $cut = $idx }
        }
        # Box-drawing separator runs (Zoom invites start with these). Build the
        # char range from codepoints so this script stays plain ASCII.
        $boxPattern = '[' + [char]0x2500 + '-' + [char]0x257F + ']{2,}'
        $box = [regex]::Match($t, $boxPattern)
        if ($box.Success -and $box.Index -lt $cut) { $cut = $box.Index }

        if ($cut -lt $t.Length) { $t = $t.Substring(0, $cut) }
        $t = $t -replace '[_\-=]{3,}', ' '
        $body = Finalize-Cell $t

        # A Zoom personal invite ("<Name> is inviting you...") leaves only the
        # organizer name before the marker; that just duplicates the Owner column.
        if ($lower -match 'is inviting you to a scheduled zoom meeting' -and $body.Length -le 40) { $body = "" }
        $body = $body -replace '^[\*\-]\s+', ''   # drop a leading bullet marker
    }

    # Show the location (Zoom/Teams URLs become short links); if a genuine agenda
    # survived the boilerplate stripping, append it labelled "Agenda:".
    $locCell = Shorten-Location $location
    $parts = @()
    if ($locCell -ne "") { $parts += $locCell }
    if ($body.Length -ge 3) { $parts += ("Agenda: " + $body) }
    return ($parts -join " - ")
}

function Resolve-NodeExe {
    # Node is provided per-user by fnm (kevin.crump is a standard, non-admin user,
    # so there is no system-wide Node on PATH). In an interactive dev shell fnm's
    # shell-integration puts node on PATH; in this non-interactive scheduled task
    # it does NOT, so fall back to fnm's default-version junction (a stable path
    # that always tracks `fnm default`).
    $cmd = Get-Command node -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $fnmDefault = Join-Path $env:APPDATA "fnm\aliases\default\node.exe"
    if (Test-Path $fnmDefault) { return $fnmDefault }

    # Last resort: the newest installed fnm Node version.
    $versions = Join-Path $env:APPDATA "fnm\node-versions"
    if (Test-Path $versions) {
        $node = Get-ChildItem $versions -Directory -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending |
            ForEach-Object { Join-Path $_.FullName "installation\node.exe" } |
            Where-Object { Test-Path $_ } |
            Select-Object -First 1
        if ($node) { return $node }
    }
    return $null
}

function Get-TodayMeetings {
    # Cloud fetch: get-meetings.js pulls the published Outlook ICS feed (URL in
    # the OUTLOOK_ICS_URL user env var), expands recurring events, and prints
    # today's non-all-day meetings as JSON. No Outlook desktop dependency.
    $icsUrl = $env:OUTLOOK_ICS_URL
    if ([string]::IsNullOrWhiteSpace($icsUrl)) {
        $icsUrl = [System.Environment]::GetEnvironmentVariable("OUTLOOK_ICS_URL", "User")
    }
    if ([string]::IsNullOrWhiteSpace($icsUrl)) {
        return [PSCustomObject]@{ Ok = $false; Meetings = @() }
    }

    try {
        $env:OUTLOOK_ICS_URL = $icsUrl
        $helper = Join-Path $PSScriptRoot "get-meetings.js"
        $nodeExe = Resolve-NodeExe
        if (-not $nodeExe) {
            return [PSCustomObject]@{ Ok = $false; Meetings = @() }
        }
        $json = & $nodeExe $helper
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace("$json")) {
            return [PSCustomObject]@{ Ok = $false; Meetings = @() }
        }

        $meetings = @()
        foreach ($e in ("$json" | ConvertFrom-Json)) {
            $start = [DateTimeOffset]::Parse($e.start).LocalDateTime
            $end = [DateTimeOffset]::Parse($e.end).LocalDateTime
            $summary = Clean-Summary $e.description $e.location
            $meetings += [PSCustomObject]@{
                Time    = $start.ToString("h:mm tt") + " - " + $end.ToString("h:mm tt")
                Subject = $e.subject
                Owner   = $e.organizer
                Summary = $summary
            }
        }
        return [PSCustomObject]@{ Ok = $true; Meetings = $meetings }
    }
    catch {
        return [PSCustomObject]@{ Ok = $false; Meetings = @() }
    }
}

function Get-JiraTickets {
    # Prefer the process env var, but fall back to the persisted User value so a
    # freshly spawned task process is not bitten by a stale inherited environment.
    $token = $env:JIRA_API_TOKEN
    if ([string]::IsNullOrWhiteSpace($token)) {
        $token = [System.Environment]::GetEnvironmentVariable("JIRA_API_TOKEN", "User")
    }
    if ([string]::IsNullOrWhiteSpace($token)) {
        return [PSCustomObject]@{ Ok = $false; Tickets = @() }
    }

    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($jiraEmail + ":" + $token)
        $headers = @{ Authorization = "Basic " + [System.Convert]::ToBase64String($bytes); Accept = "application/json" }

        # Current (open) + next (future) sprints assigned to me.
        $jql = 'project=UMT AND assignee=currentUser() AND (sprint in openSprints() OR sprint in futureSprints()) AND issuetype not in subTaskIssueTypes() ORDER BY rank ASC'
        # Note: /search returns 410 Gone on this tenant; use /search/jql.
        # customfield_12108 = "Story point estimate"; customfield_10006 = "Sprint".
        $url = $jiraBase + '/rest/api/3/search/jql?jql=' + [Uri]::EscapeDataString($jql) + '&fields=summary,status,customfield_12108,customfield_10006&maxResults=50'

        $resp = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ContentType "application/json" -ErrorAction Stop

        $out = @()
        foreach ($i in $resp.issues) {
            $statusName = ""
            if ($i.fields.status) { $statusName = $i.fields.status.name }

            # An issue lists every sprint it has touched; pick the relevant one:
            # the active sprint if present, otherwise the (earliest) future sprint.
            $sprints = $i.fields.customfield_10006
            $rel = $null
            if ($sprints) {
                $rel = $sprints | Where-Object { $_.state -eq 'active' } | Select-Object -First 1
                if ($null -eq $rel) {
                    $rel = $sprints | Where-Object { $_.state -eq 'future' } |
                        Sort-Object @{ Expression = { if ($_.startDate) { [DateTime]::Parse($_.startDate) } else { [DateTime]::MaxValue } } } |
                        Select-Object -First 1
                }
            }

            $sprintId = 0; $sprintName = "(no sprint)"; $sprintState = "none"; $sprintOrder = 9; $sprintStart = [DateTime]::MaxValue
            if ($rel) {
                $sprintId = $rel.id
                $sprintName = $rel.name
                $sprintState = $rel.state
                if ($rel.startDate) { $sprintStart = [DateTime]::Parse($rel.startDate) }
                if ($rel.state -eq 'active') { $sprintOrder = 0 } elseif ($rel.state -eq 'future') { $sprintOrder = 1 } else { $sprintOrder = 2 }
            }

            $out += [PSCustomObject]@{
                Key         = $i.key
                Summary     = $i.fields.summary
                Status      = $statusName
                Points      = $i.fields.customfield_12108
                Link        = $jiraBase + "/browse/" + $i.key
                SprintId    = $sprintId
                SprintName  = $sprintName
                SprintState = $sprintState
                SprintOrder = $sprintOrder
                SprintStart = $sprintStart
            }
        }
        return [PSCustomObject]@{ Ok = $true; Tickets = $out }
    }
    catch {
        return [PSCustomObject]@{ Ok = $false; Tickets = @() }
    }
}

function Format-TicketRow {
    param ($ticket)
    $summary = $ticket.Summary
    if ($null -eq $summary) { $summary = "" }
    $summary = $summary.Trim() -replace '\|', '\|'
    $points = "-"
    if ($null -ne $ticket.Points -and "$($ticket.Points)" -ne "") {
        $points = ("{0:0.##}" -f [double]$ticket.Points)
    }
    return "| [" + $ticket.Key + "](" + $ticket.Link + ") | " + $summary + " | " + $ticket.Status + " | " + $points + " |"
}

function Format-MeetingRow {
    param ($mtg)
    $subject = (("" + $mtg.Subject).Trim()) -replace '\|', '\|'
    $owner = (("" + $mtg.Owner).Trim()) -replace '\|', '\|'
    $summary = "" + $mtg.Summary
    if ($summary -eq "") { $summary = "-" }
    return "| " + $mtg.Time + " | " + $subject + " | " + $owner + " | " + $summary + " |"
}

function Parse-Note {
    param ($lines)
    $preamble = New-Object System.Collections.ArrayList
    $sections = New-Object System.Collections.ArrayList
    $current = $null
    foreach ($line in $lines) {
        if ($line -match '^##\s') {
            $current = @{ Heading = $line; Body = (New-Object System.Collections.ArrayList) }
            [void]$sections.Add($current)
        }
        elseif ($null -eq $current) {
            [void]$preamble.Add($line)
        }
        else {
            [void]$current.Body.Add($line)
        }
    }
    return @{ Preamble = $preamble; Sections = $sections }
}

function Rebuild-Note {
    param ($note)
    $out = New-Object System.Collections.ArrayList
    foreach ($line in $note.Preamble) { [void]$out.Add($line) }
    foreach ($section in $note.Sections) {
        [void]$out.Add($section.Heading)
        foreach ($bodyLine in $section.Body) { [void]$out.Add($bodyLine) }
    }
    return $out.ToArray()
}

function Find-Section {
    param ($note, $pattern)
    foreach ($section in $note.Sections) {
        if ($section.Heading -match $pattern) { return $section }
    }
    return $null
}

function Ensure-Header {
    param ($note, $title)
    # Deterministically rebuild the header: drop any existing title (single #),
    # legacy day-of-week line, and leading blanks, then prepend the canonical
    # title. Self-heals duplicates or a malformed title from earlier runs.
    $kept = New-Object System.Collections.ArrayList
    foreach ($l in $note.Preamble) {
        if ($l -match '^#(?!#)') { continue }          # old title line
        if ($l -match '^_[A-Za-z]+_\s*$') { continue }  # legacy day-of-week line
        [void]$kept.Add($l)
    }
    while ($kept.Count -gt 0 -and [string]::IsNullOrWhiteSpace($kept[0])) { $kept.RemoveAt(0) }

    $new = New-Object System.Collections.ArrayList
    [void]$new.Add($title)
    [void]$new.Add("")
    foreach ($l in $kept) { [void]$new.Add($l) }
    $note.Preamble = $new
}

function Get-HandWrittenRemainder {
    param ($bodyLines)
    # A managed section (Meetings/Tickets) holds an auto-generated block at the
    # top: optional blanks, a markdown pipe table (or a "_No ..._" placeholder),
    # then blanks. Strip exactly that leading block and return whatever the user
    # hand-wrote after it, so a refresh rewrites the table but keeps their notes.
    $list = @($bodyLines)
    $i = 0
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    while ($i -lt $list.Count -and ($list[$i] -match '^\s*\|' -or $list[$i] -match '^\s*_No .*_\s*$')) { $i++ }
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    if ($i -ge $list.Count) { return @() }
    return @($list[$i..($list.Count - 1)])
}

function Append-HandNotes {
    param ($body, $handNotes)
    if ($handNotes.Count -gt 0) {
        foreach ($l in $handNotes) { [void]$body.Add($l) }
        while ($body.Count -gt 0 -and [string]::IsNullOrWhiteSpace($body[$body.Count - 1])) { $body.RemoveAt($body.Count - 1) }
        [void]$body.Add("")
    }
}

function Set-MeetingsBody {
    param ($section, $meetings)
    $hand = Get-HandWrittenRemainder $section.Body
    $body = New-Object System.Collections.ArrayList
    [void]$body.Add("")
    if ($meetings.Count -gt 0) {
        [void]$body.Add("| Time | Meeting | Owner | Summary |")
        [void]$body.Add("| --- | --- | --- | --- |")
        foreach ($m in $meetings) { [void]$body.Add((Format-MeetingRow $m)) }
    }
    else {
        [void]$body.Add("_No meetings scheduled_")
    }
    [void]$body.Add("")
    Append-HandNotes $body $hand
    $section.Body = $body
}

function Set-TicketsBody {
    param ($section, $tickets)
    $hand = Get-HandWrittenRemainder $section.Body
    $body = New-Object System.Collections.ArrayList
    [void]$body.Add("")
    if ($tickets.Count -eq 0) {
        [void]$body.Add("_No tickets assigned in current or next sprint_")
        [void]$body.Add("")
        Append-HandNotes $body $hand
        $section.Body = $body
        return
    }

    [void]$body.Add("| Ticket | Summary | Status | Points |")
    [void]$body.Add("| --- | --- | --- | --- |")

    # Group by sprint; active sprint(s) first, then future, then by start date.
    $groups = $tickets | Group-Object -Property SprintId
    $ordered = $groups | Sort-Object `
        @{ Expression = { $_.Group[0].SprintOrder } }, `
        @{ Expression = { $_.Group[0].SprintStart } }

    foreach ($g in $ordered) {
        $first = $g.Group[0]
        $label = "" + $first.SprintName
        if ($first.SprintState -eq 'active') { $label += " (active)" }
        elseif ($first.SprintState -eq 'future') { $label += " (next)" }
        $label = $label -replace '\|', '\|'
        [void]$body.Add("| **" + $label + "** |  |  |  |")
        foreach ($t in $g.Group) { [void]$body.Add((Format-TicketRow $t)) }
    }

    [void]$body.Add("")
    Append-HandNotes $body $hand
    $section.Body = $body
}

# ---------------- Main ----------------

# Open Obsidian (if needed) and wait for it before doing the work.
Ensure-AppsRunning $vaultName

if (Test-Path $filePath) {
    $lines = Get-Content -Path $filePath -Encoding UTF8
    if ($null -eq $lines) { $lines = @() }
}
else {
    $lines = @($titleText, "")
}

$note = Parse-Note $lines
Ensure-Header $note $titleText

# Meetings: create or refresh
$meetings = Get-TodayMeetings
if ($meetings.Ok) {
    $meetingsSection = Find-Section $note '^##\s+Meetings'
    if ($null -eq $meetingsSection) {
        $meetingsSection = @{ Heading = "## Meetings"; Body = (New-Object System.Collections.ArrayList) }
        Set-MeetingsBody $meetingsSection $meetings.Meetings
        $note.Sections.Insert(0, $meetingsSection)
    }
    else {
        Set-MeetingsBody $meetingsSection $meetings.Meetings
    }
}

# Tickets: create or refresh (only when the Jira call succeeded)
$jira = Get-JiraTickets
if ($jira.Ok) {
    $ticketsSection = Find-Section $note '^##\s+(Current\s+)?Tickets'
    if ($null -eq $ticketsSection) {
        $ticketsSection = @{ Heading = "## Tickets"; Body = (New-Object System.Collections.ArrayList) }
        Set-TicketsBody $ticketsSection $jira.Tickets
        $insertAt = $note.Sections.Count
        $meetingsSection = Find-Section $note '^##\s+Meetings'
        if ($null -ne $meetingsSection) {
            $insertAt = $note.Sections.IndexOf($meetingsSection) + 1
        }
        $note.Sections.Insert($insertAt, $ticketsSection)
    }
    else {
        $ticketsSection.Heading = "## Tickets"   # migrate legacy "## Current Tickets"
        Set-TicketsBody $ticketsSection $jira.Tickets
    }
}

$content = Rebuild-Note $note
Set-Content -Path $filePath -Value $content -Encoding UTF8

# Open today's note in Obsidian
$encodedFile = [Uri]::EscapeDataString("devlog/$monthFolder/devlog $date")
$uri = 'obsidian://open?vault=Dev%20Docs&file=' + $encodedFile
Start-Process $uri
