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

# Per-user log location (no admin needed). A rolling devlog.log holds one line per
# event; a per-run transcript captures full output. Both are pruned after 30 days.
$logDir = Join-Path $env:LOCALAPPDATA "obsidian-devlog\logs"
$script:logFile = $null

if (-not (Test-Path $monthDir)) {
    New-Item -ItemType Directory -Path $monthDir -Force | Out-Null
}

function Write-Log {
    param([string] $Message, [string] $Level = "INFO")
    if (-not $script:logFile) { return }
    try {
        $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Add-Content -Path $script:logFile -Value "[$ts] [$Level] $Message" -Encoding UTF8
    }
    catch { }
}

function New-Banner {
    # Obsidian callout shown in the note when a section genuinely failed to load
    # (as opposed to a real empty day). Kept plain ASCII for PS 5.1.
    param([string] $What, [string] $Reason)
    $r = $Reason
    if ([string]::IsNullOrWhiteSpace($r)) { $r = "unknown error" }
    return "> [!warning] $What unavailable - $r (see log)"
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

function Wait-ForNetwork {
    # Logon/startup runs can fire before DNS/network is ready, so the fetches
    # below fail silently and the note is written with no Meetings/Tickets. Poll
    # until every host we depend on resolves, up to a bounded timeout; if it
    # never comes up we fall through and the fetches degrade as before.
    param(
        [string[]] $Hosts,
        [int] $TimeoutSeconds = 180,
        [int] $IntervalSeconds = 10
    )
    $targets = @($Hosts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($targets.Count -eq 0) { return $true }
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ($true) {
        $allOk = $true
        foreach ($h in $targets) {
            try { [void][System.Net.Dns]::GetHostAddresses($h) }
            catch { $allOk = $false; break }
        }
        if ($allOk) { return $true }
        if ((Get-Date) -ge $deadline) { return $false }
        Start-Sleep -Seconds $IntervalSeconds
    }
}

function Invoke-WithRetry {
    # Retry a flaky network operation a few times before giving up. The action
    # must throw on failure (use -ErrorAction Stop / throw); its return value is
    # passed straight back on success.
    param(
        [scriptblock] $Action,
        [int] $MaxAttempts = 3,
        [int] $DelaySeconds = 5
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try { return (& $Action) }
        catch {
            if ($attempt -ge $MaxAttempts) { throw }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
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
        Write-Log "Meetings skipped: OUTLOOK_ICS_URL not set" "WARN"
        return [PSCustomObject]@{ Ok = $false; Meetings = @(); Reason = "OUTLOOK_ICS_URL not set" }
    }

    try {
        $env:OUTLOOK_ICS_URL = $icsUrl
        $helper = Join-Path $PSScriptRoot "get-meetings.js"
        $nodeExe = Resolve-NodeExe
        if (-not $nodeExe) {
            Write-Log "Meetings failed: Node not found (fnm default junction missing)" "ERROR"
            return [PSCustomObject]@{ Ok = $false; Meetings = @(); Reason = "Node not found" }
        }
        Write-Log "Using node: $nodeExe"
        # Keep stdout (the JSON) clean; send stderr to a file so Node's fetch
        # ExperimentalWarning can't corrupt the JSON, while still capturing the
        # real error text for the log on failure.
        $errPath = Join-Path $logDir "get-meetings.err.txt"
        $json = Invoke-WithRetry -Action {
            $out = & $nodeExe $helper 2>$errPath
            if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace("$out")) {
                $errText = ""
                if (Test-Path $errPath) { $errText = (Get-Content $errPath -Raw -ErrorAction SilentlyContinue) }
                throw "get-meetings.js failed (exit $LASTEXITCODE): $errText"
            }
            return $out
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
        Write-Log "Meetings loaded: $($meetings.Count) event(s)"
        return [PSCustomObject]@{ Ok = $true; Meetings = $meetings; Reason = "" }
    }
    catch {
        Write-Log "Meetings failed: $($_.Exception.Message)" "ERROR"
        return [PSCustomObject]@{ Ok = $false; Meetings = @(); Reason = "fetch error" }
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
        Write-Log "Tickets skipped: JIRA_API_TOKEN not set" "WARN"
        return [PSCustomObject]@{ Ok = $false; Tickets = @(); Reason = "JIRA_API_TOKEN not set" }
    }

    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($jiraEmail + ":" + $token)
        $headers = @{ Authorization = "Basic " + [System.Convert]::ToBase64String($bytes); Accept = "application/json" }

        # Current (open) + next (future) sprints assigned to me.
        $jql = 'project=UMT AND assignee=currentUser() AND (sprint in openSprints() OR sprint in futureSprints()) AND issuetype not in subTaskIssueTypes() ORDER BY rank ASC'
        # Note: /search returns 410 Gone on this tenant; use /search/jql.
        # customfield_12108 = "Story point estimate"; customfield_10006 = "Sprint".
        $url = $jiraBase + '/rest/api/3/search/jql?jql=' + [Uri]::EscapeDataString($jql) + '&fields=summary,status,customfield_12108,customfield_10006&maxResults=50'

        # Retry + bounded timeout: the call previously had neither, so a slow or
        # transient Jira response could hang past the task limit or fail outright.
        $resp = Invoke-WithRetry -Action {
            Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ContentType "application/json" -TimeoutSec 30 -ErrorAction Stop
        }

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
        Write-Log "Tickets loaded: $($out.Count) issue(s)"
        return [PSCustomObject]@{ Ok = $true; Tickets = $out; Reason = "" }
    }
    catch {
        Write-Log "Tickets failed: $($_.Exception.Message)" "ERROR"
        return [PSCustomObject]@{ Ok = $false; Tickets = @(); Reason = "Jira request failed" }
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
    # top: optional blanks, an optional "> [!warning] ... unavailable" banner, a
    # markdown pipe table (or a "_No ..._" placeholder), then blanks. Strip exactly
    # that leading block and return whatever the user hand-wrote after it, so a
    # refresh rewrites the table but keeps their notes. The banner match is narrow
    # ("[!warning] ... unavailable") so it never eats a user's own blockquote.
    $list = @($bodyLines)
    $i = 0
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    while ($i -lt $list.Count -and $list[$i] -match '^\s*>\s*\[!(warning|failure|error)\].*unavailable') { $i++ }
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    while ($i -lt $list.Count -and ($list[$i] -match '^\s*\|' -or $list[$i] -match '^\s*_No .*_\s*$')) { $i++ }
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    if ($i -ge $list.Count) { return @() }
    return @($list[$i..($list.Count - 1)])
}

function Get-LastTable {
    param ($bodyLines)
    # Pull the existing pipe-table rows out of a managed section's auto block, so a
    # later failed refresh can preserve the last good table instead of blanking it.
    # Only real "| ... |" lines are kept (a "_No ..._" placeholder is not a table).
    $list = @($bodyLines)
    $i = 0
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    while ($i -lt $list.Count -and $list[$i] -match '^\s*>\s*\[!(warning|failure|error)\].*unavailable') { $i++ }
    while ($i -lt $list.Count -and [string]::IsNullOrWhiteSpace($list[$i])) { $i++ }
    $rows = New-Object System.Collections.ArrayList
    while ($i -lt $list.Count -and $list[$i] -match '^\s*\|') { [void]$rows.Add($list[$i]); $i++ }
    return @($rows.ToArray())
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
    param ($section, $result)
    $hand = Get-HandWrittenRemainder $section.Body
    $body = New-Object System.Collections.ArrayList
    [void]$body.Add("")
    if (-not $result.Ok) {
        # Genuine failure: show a visible banner instead of silence, and keep the
        # last good table if a prior run in this note had one.
        [void]$body.Add((New-Banner "Meetings" $result.Reason))
        $prior = Get-LastTable $section.Body
        if ($prior.Count -gt 0) {
            [void]$body.Add("")
            foreach ($l in $prior) { [void]$body.Add($l) }
        }
    }
    elseif ($result.Meetings.Count -gt 0) {
        [void]$body.Add("| Time | Meeting | Owner | Summary |")
        [void]$body.Add("| --- | --- | --- | --- |")
        foreach ($m in $result.Meetings) { [void]$body.Add((Format-MeetingRow $m)) }
    }
    else {
        [void]$body.Add("_No meetings scheduled_")
    }
    [void]$body.Add("")
    Append-HandNotes $body $hand
    $section.Body = $body
}

function Set-TicketsBody {
    param ($section, $result)
    $hand = Get-HandWrittenRemainder $section.Body
    $body = New-Object System.Collections.ArrayList
    [void]$body.Add("")
    if (-not $result.Ok) {
        [void]$body.Add((New-Banner "Tickets" $result.Reason))
        $prior = Get-LastTable $section.Body
        if ($prior.Count -gt 0) {
            [void]$body.Add("")
            foreach ($l in $prior) { [void]$body.Add($l) }
        }
        [void]$body.Add("")
        Append-HandNotes $body $hand
        $section.Body = $body
        return
    }

    $tickets = $result.Tickets
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

function Normalize-OngoingLines {
    param ($lines)
    $arr = @($lines)
    # Demote any top-level (# / ##) headings in the source to h3+ so the copied
    # content never introduces new sections into the daily note (that would break
    # the in-place refresh and leave orphan sections behind).
    $arr = @(foreach ($l in $arr) {
        if ($l -match '^(#{1,2})(\s.*)$') { '###' + $matches[2] } else { $l }
    })
    # Trim leading/trailing blank lines.
    $start = 0
    while ($start -lt $arr.Count -and [string]::IsNullOrWhiteSpace($arr[$start])) { $start++ }
    $end = $arr.Count - 1
    while ($end -ge $start -and [string]::IsNullOrWhiteSpace($arr[$end])) { $end-- }
    if ($start -gt $end) { return @() }
    return @($arr[$start..$end])
}

function Get-OngoingBody {
    # Carry the "## Ongoing" section forward from the most recent PRIOR devlog note
    # (previous day, or the latest note before today). This lets the user maintain
    # Ongoing directly inside the daily notes; combined with the seed-once rule in
    # Main, deletions stick because we read the last note's final state and never
    # re-copy a frozen snapshot. Falls back to the hand-edited vault-root
    # ONGOING.md only when no prior devlog note exists (e.g. the very first note).
    try {
        $prior = $null
        if (Test-Path $devlogRoot) {
            $prior = Get-ChildItem -Path $devlogRoot -Recurse -Filter "devlog *.md" -ErrorAction SilentlyContinue |
                ForEach-Object {
                    if ($_.BaseName -match '^devlog\s+(\d{8})$') {
                        [PSCustomObject]@{ File = $_; Date = $matches[1] }
                    }
                } |
                Where-Object { $_.Date -lt $date } |   # exclude today's note (and any future)
                Sort-Object Date -Descending |
                Select-Object -First 1
        }

        if ($prior) {
            $raw = Get-Content -Path $prior.File.FullName -Encoding UTF8
            if ($null -eq $raw) { $raw = @() }
            $pnote = Parse-Note (@($raw))
            $psec = Find-Section $pnote '^##\s+Ongoing'
            if ($null -ne $psec) {
                Write-Log "Ongoing source: prior note $($prior.File.Name)"
                return [PSCustomObject]@{ Ok = $true; Lines = (Normalize-OngoingLines $psec.Body) }
            }
            # Prior note exists but has no Ongoing section: nothing to carry.
            Write-Log "Ongoing source: prior note $($prior.File.Name) has no Ongoing section"
            return [PSCustomObject]@{ Ok = $true; Lines = @() }
        }

        # Fallback: hand-edited root ONGOING.md (only when there is no prior note).
        $ongoingPath = Join-Path $vaultPath "ONGOING.md"
        if (-not (Test-Path $ongoingPath)) {
            return [PSCustomObject]@{ Ok = $false; Lines = @() }
        }
        $raw = Get-Content -Path $ongoingPath -Encoding UTF8
        if ($null -eq $raw) { $raw = @() }
        Write-Log "Ongoing source: fallback ONGOING.md (no prior devlog note)"
        return [PSCustomObject]@{ Ok = $true; Lines = (Normalize-OngoingLines @($raw)) }
    }
    catch {
        return [PSCustomObject]@{ Ok = $false; Lines = @() }
    }
}

function Set-OngoingBody {
    param ($section, $ongoingLines)
    # Fill a NEWLY created "## Ongoing" section from the carried-forward source.
    # Only called when the section did not already exist in today's note (see the
    # seed-once rule in Main), so it never clobbers edits made during the day.
    $body = New-Object System.Collections.ArrayList
    [void]$body.Add("")
    if ($ongoingLines.Count -gt 0) {
        foreach ($l in $ongoingLines) { [void]$body.Add($l) }
    }
    else {
        [void]$body.Add("_ONGOING note is empty_")
    }
    [void]$body.Add("")
    $section.Body = $body
}

# ---------------- Main ----------------

# --- Logging: per-run transcript + a rolling one-line-per-event devlog.log, both
# under %LOCALAPPDATA% (no admin). Prune transcripts after 30 days; cap devlog.log
# at ~1 MB (keep one .1 backup). Best-effort: logging must never break the run.
try {
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:logFile = Join-Path $logDir "devlog.log"
    if ((Test-Path $script:logFile) -and ((Get-Item $script:logFile).Length -gt 1MB)) {
        Move-Item -Path $script:logFile -Destination (Join-Path $logDir "devlog.log.1") -Force -ErrorAction SilentlyContinue
    }
    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    Start-Transcript -Path (Join-Path $logDir "transcript-$stamp.log") -Force | Out-Null
    Get-ChildItem -Path $logDir -Filter "transcript-*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}
catch { }

Write-Log "=== Run start (user=$env:USERNAME, note=$fileName) ==="

try {
    # Wait for DNS to come up before the network fetches. Logon/wake runs can fire
    # before the network is ready; without this the fetches throw and the sections
    # come up empty. Safe now that the task ExecutionTimeLimit is 10 min.
    $netHosts = New-Object System.Collections.ArrayList
    try { [void]$netHosts.Add(([Uri]$jiraBase).Host) } catch { }
    $icsForHost = $env:OUTLOOK_ICS_URL
    if ([string]::IsNullOrWhiteSpace($icsForHost)) {
        $icsForHost = [System.Environment]::GetEnvironmentVariable("OUTLOOK_ICS_URL", "User")
    }
    if (-not [string]::IsNullOrWhiteSpace($icsForHost)) {
        try { [void]$netHosts.Add(([Uri]$icsForHost).Host) } catch { }
    }
    if ((Wait-ForNetwork -Hosts $netHosts.ToArray() -TimeoutSeconds 180)) {
        Write-Log "Network ready (hosts: $($netHosts -join ', '))"
    }
    else {
        Write-Log "Network wait timed out; proceeding (fetches may degrade)" "WARN"
    }

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

    # Meetings: always ensure the section exists, then render the table on success
    # or a "> [!warning] ... unavailable" banner on failure (never silently absent).
    $meetings = Get-TodayMeetings
    $meetingsSection = Find-Section $note '^##\s+Meetings'
    if ($null -eq $meetingsSection) {
        $meetingsSection = @{ Heading = "## Meetings"; Body = (New-Object System.Collections.ArrayList) }
        $note.Sections.Insert(0, $meetingsSection)
    }
    Set-MeetingsBody $meetingsSection $meetings

    # Tickets: same contract as Meetings, placed right after it.
    $jira = Get-JiraTickets
    $ticketsSection = Find-Section $note '^##\s+(Current\s+)?Tickets'
    if ($null -eq $ticketsSection) {
        $ticketsSection = @{ Heading = "## Tickets"; Body = (New-Object System.Collections.ArrayList) }
        $insertAt = $note.Sections.Count
        $ms = Find-Section $note '^##\s+Meetings'
        if ($null -ne $ms) { $insertAt = $note.Sections.IndexOf($ms) + 1 }
        $note.Sections.Insert($insertAt, $ticketsSection)
    }
    else {
        $ticketsSection.Heading = "## Tickets"   # migrate legacy "## Current Tickets"
    }
    Set-TicketsBody $ticketsSection $jira

    # Ongoing: carried forward from the most recent prior devlog note. SEED ONCE -
    # only create+fill it when today's note has no Ongoing section yet; if one is
    # already present, leave it completely untouched so any edits or deletions made
    # during the day are never clobbered by a later run. Placed beneath Meetings
    # and Tickets; never touches other sections.
    $ongoingSection = Find-Section $note '^##\s+Ongoing'
    if ($null -eq $ongoingSection) {
        $ongoing = Get-OngoingBody
        if ($ongoing.Ok) {
            $ongoingSection = @{ Heading = "## Ongoing"; Body = (New-Object System.Collections.ArrayList) }
            Set-OngoingBody $ongoingSection $ongoing.Lines
            $insertAt = $note.Sections.Count
            $ts = Find-Section $note '^##\s+(Current\s+)?Tickets'
            $ms = Find-Section $note '^##\s+Meetings'
            if ($null -ne $ts) {
                $insertAt = $note.Sections.IndexOf($ts) + 1
            }
            elseif ($null -ne $ms) {
                $insertAt = $note.Sections.IndexOf($ms) + 1
            }
            $note.Sections.Insert($insertAt, $ongoingSection)
            Write-Log "Ongoing seeded ($($ongoing.Lines.Count) line(s))"
        }
        else {
            Write-Log "Ongoing skipped (no prior note and ONGOING.md missing/unreadable)"
        }
    }
    else {
        Write-Log "Ongoing preserved (already present in today's note)"
    }

    # Write the note as UTF-8 WITHOUT a BOM (PS 5.1's Set-Content -Encoding UTF8
    # emits a BOM). WriteAllLines uses the platform newline (CRLF), which the read
    # path (Get-Content) handles fine.
    $content = Rebuild-Note $note
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($filePath, [string[]]$content, $utf8NoBom)
    Write-Log "Note written: $filePath"

    # Open today's note in Obsidian
    $encodedFile = [Uri]::EscapeDataString("devlog/$monthFolder/devlog $date")
    $uri = 'obsidian://open?vault=Dev%20Docs&file=' + $encodedFile
    Start-Process $uri
    Write-Log "=== Run OK ==="
}
catch {
    Write-Log "Run failed: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
}
