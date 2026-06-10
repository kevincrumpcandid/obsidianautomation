$date = Get-Date -Format "yyyyMMdd"
$displayDate = Get-Date -Format "yyyy-MM-dd"
$vaultPath = "C:\Users\kevin.crump\OneDrive - Candid\Documents\DevDocs\Dev Docs"
$devlogDir = Join-Path $vaultPath "devlog"
$fileName = "devlog $date.md"
$filePath = Join-Path $devlogDir $fileName

$jiraBase = "https://candidprojects.atlassian.net"
$jiraEmail = "kevin.crump@candid.org"
$jiraProject = "UMT"
$jiraBoardId = "1273"

if (-not (Test-Path $devlogDir)) {
    New-Item -ItemType Directory -Path $devlogDir | Out-Null
}

function Get-AuthHeaders {
    $token = $env:JIRA_API_TOKEN
    if (-not $token) { return $null }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($jiraEmail + ":" + $token)
    $auth = "Basic " + [System.Convert]::ToBase64String($bytes)
    return @{ "Authorization" = $auth; "Accept" = "application/json" }
}

function Get-SprintInfo {
    $headers = Get-AuthHeaders
    if (-not $headers) { return $null }
    try {
        $url = $jiraBase + "/rest/agile/1.0/board/" + $jiraBoardId + "/sprint?state=active"
        $r = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop
        if ($r.values.Count -eq 0) { return $null }
        $sprint = $r.values[0]
        $name = $sprint.name
        $start = $null
        $end = $null
        if ($sprint.startDate) { $start = [DateTime]::Parse($sprint.startDate).ToString("MMM d") }
        if ($sprint.endDate)   { $end   = [DateTime]::Parse($sprint.endDate).ToString("MMM d, yyyy") }
        if ($start -and $end) { return $name + " (" + $start + " - " + $end + ")" }
        return $name
    }
    catch { return $null }
}

function Get-JiraTickets {
    $headers = Get-AuthHeaders
    if (-not $headers) { return $null }

    $fileExists = Test-Path $filePath

    try {
        if (-not $fileExists) {
            Start-Sleep -Seconds 30
        }

        $jql = "project=" + $jiraProject + " AND sprint in openSprints() AND assignee=currentUser() AND issuetype not in subTaskIssueTypes() ORDER BY rank ASC"
        $encoded = [Uri]::EscapeDataString($jql)
        $fields = "summary,customfield_10016,priority,labels"
        $url = $jiraBase + "/rest/api/3/search/jql?jql=" + $encoded + "&fields=" + $fields + "&maxResults=50"

        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction Stop

        $lines = @()
        foreach ($issue in $response.issues) {
            $key      = $issue.key
            $title    = $issue.fields.summary
            $pts      = $issue.fields.customfield_10016
            $priority = if ($issue.fields.priority) { $issue.fields.priority.name } else { "None" }
            $labels   = if ($issue.fields.labels -and $issue.fields.labels.Count -gt 0) { $issue.fields.labels -join ", " } else { "none" }
            $link     = $jiraBase + "/browse/" + $key
            $ptsTxt   = if ($null -ne $pts) { [string][int]$pts + " pts" } else { "--" }

            $lines += "- [" + $key + "](" + $link + ") " + $title
            $lines += "  *" + $priority + " | " + $ptsTxt + " | " + $labels + "*"
        }
        return ,$lines
    }
    catch {
        return $null
    }
}

function Build-TicketsBlock {
    param ($tickets, $sprintInfo)
    $block = "## Current Tickets`n"
    if ($sprintInfo) {
        $block += "_" + $sprintInfo + "_`n"
    }
    $block += "`n"
    if ($null -ne $tickets -and $tickets.Count -gt 0) {
        $block += ($tickets -join "`n") + "`n"
    }
    else {
        $block += "_No tickets assigned in current sprint_`n"
    }
    return $block
}

if (-not (Test-Path $filePath)) {
    $sprintInfo = Get-SprintInfo
    $tickets = Get-JiraTickets
    if ($null -ne $tickets) {
        $content = "# Devlog - $displayDate`n`n" + (Build-TicketsBlock $tickets $sprintInfo) + "`n"
        Set-Content -Path $filePath -Value $content -Encoding UTF8
    }
}
else {
    $existing = Get-Content -Path $filePath -Raw -Encoding UTF8
    if ($existing -notmatch "## Current Tickets") {
        $sprintInfo = Get-SprintInfo
        $tickets = Get-JiraTickets
        if ($null -ne $tickets) {
            $lines = Get-Content -Path $filePath -Encoding UTF8

            $inMeetings = $false
            $insertIdx = $lines.Length

            for ($i = 0; $i -lt $lines.Length; $i++) {
                if ($lines[$i] -match "^## Meetings") {
                    $inMeetings = $true
                    continue
                }
                if ($inMeetings -and $lines[$i] -match "^#+") {
                    $insertIdx = $i
                    break
                }
            }

            $blockLines = (Build-TicketsBlock $tickets $sprintInfo).TrimEnd() -split "`n"

            if ($insertIdx -gt 0 -and $insertIdx -lt $lines.Length) {
                $newContent = $lines[0..($insertIdx - 1)] + @("") + $blockLines + @("") + $lines[$insertIdx..($lines.Length - 1)]
            }
            else {
                $newContent = $lines + @("") + $blockLines
            }

            Set-Content -Path $filePath -Value $newContent -Encoding UTF8
        }
    }
    else {
        # Section already present: append any Jira tickets not already in the file
        $sprintInfo = Get-SprintInfo
        $tickets = Get-JiraTickets
        if ($null -ne $tickets -and $tickets.Count -gt 0) {
            $lines = Get-Content -Path $filePath -Encoding UTF8

            $sectionStart = -1
            for ($i = 0; $i -lt $lines.Length; $i++) {
                if ($lines[$i] -match "^## Current Tickets") { $sectionStart = $i; break }
            }

            $sectionEnd = $lines.Length
            for ($i = $sectionStart + 1; $i -lt $lines.Length; $i++) {
                if ($lines[$i] -match "^## ") { $sectionEnd = $i; break }
            }

            # Keep only tickets whose key isn't already somewhere in the file (idempotent re-runs)
            $newLines = @()
            for ($j = 0; $j -lt $tickets.Count; $j += 2) {
                $ticketLine = $tickets[$j]
                $subLine = if ($j + 1 -lt $tickets.Count) { $tickets[$j + 1] } else { $null }
                $key = $null
                if ($ticketLine -match "\[([A-Z][A-Z0-9]+-\d+)\]") { $key = $matches[1] }
                if ($key -and ($existing -match [regex]::Escape($key))) { continue }
                $newLines += $ticketLine
                if ($null -ne $subLine) { $newLines += $subLine }
            }

            if ($newLines.Count -gt 0) {
                $before = @($lines[0..($sectionEnd - 1)])
                $bi = $before.Length - 1
                while ($bi -ge 0 -and $before[$bi].Trim() -eq "") { $bi-- }
                $before = @($before[0..$bi])

                $after = @()
                if ($sectionEnd -lt $lines.Length) { $after = @($lines[$sectionEnd..($lines.Length - 1)]) }

                $newContent = @($before + $newLines + @("") + $after)
                Set-Content -Path $filePath -Value $newContent -Encoding UTF8
            }
        }
    }
}
