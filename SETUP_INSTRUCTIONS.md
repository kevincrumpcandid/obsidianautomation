# Obsidian Daily Devlog Automation — Setup Instructions

Automatically creates and opens a dated devlog note in Obsidian every morning,
with today's Outlook calendar meetings pre-populated.

---

## What it does

The script runs at 7 AM daily and again on each logon:
1. Creates `devlog YYYYMMDD.md` in a `devlog` subfolder of your vault (skips if it already exists)
2. Fetches today's meetings from Outlook and appends a `## Meetings` section (skipped if already present, so running twice is safe)
3. Opens the note in Obsidian

---

## Prerequisites

- Windows 10 or 11
- Obsidian installed and configured with a vault
- Microsoft Outlook desktop installed (used for calendar data via COM)
- Obsidian set to **open on startup** (Settings -> About -> "Open Obsidian on system startup")

Outlook does not need to be open when the task fires — the script will launch it silently
and wait 10 seconds for it to sync before querying the calendar. For best results, Outlook
should already be running when you log in each morning.

---

## Option A: Automated setup with Claude Code

If you have [Claude Code](https://claude.ai/code) installed, clone this repo and run:

```
/setup-obsidian-devlog
```

Claude will find your vault, ask a couple of questions, and handle everything below automatically.

---

## Option B: Manual setup

### Step 1 — Find your vault path

Open `%APPDATA%\obsidian\obsidian.json` in a text editor. You will see something like:

```json
{"vaults":{"abc123":{"path":"C:\\Users\\you\\Documents\\My Vault"}}}
```

Note the `path` value and your vault's display name (the last folder in the path).

### Step 2 — Create the devlog script

Create `create-devlog.ps1` anywhere permanent (e.g. `C:\Users\you\scripts\`):

```powershell
$date = Get-Date -Format "yyyyMMdd"
$displayDate = Get-Date -Format "yyyy-MM-dd"
$vaultPath = "C:\Users\you\Documents\My Vault"
$devlogDir = Join-Path $vaultPath "devlog"
$fileName = "devlog $date.md"
$filePath = Join-Path $devlogDir $fileName

if (-not (Test-Path $devlogDir)) {
    New-Item -ItemType Directory -Path $devlogDir | Out-Null
}

if (-not (Test-Path $filePath)) {
    Set-Content -Path $filePath -Value "# Devlog - $displayDate`n`n" -Encoding UTF8
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

$existing = Get-Content -Path $filePath -Raw -Encoding UTF8
if ($existing -notmatch "## Meetings") {
    $meetings = Get-TodayMeetings
    if ($null -ne $meetings) {
        $section = "`n## Meetings`n`n"
        if ($meetings.Count -gt 0) {
            $section += ($meetings -join "`n") + "`n"
        }
        else {
            $section += "_No meetings scheduled_`n"
        }
        Add-Content -Path $filePath -Value $section -Encoding UTF8
    }
}

$encodedFile = [Uri]::EscapeDataString("devlog/devlog $date")
$uri = 'obsidian://open?vault=My%20Vault&file=' + $encodedFile
Start-Process $uri
```

Replace:
- `C:\Users\you\Documents\My Vault` with your actual vault path
- `My%20Vault` in the URI with your vault name (spaces as `%20`)
- `devlog` with your preferred subfolder name (same value in both places)

> **Important:** Use only plain ASCII characters in `.ps1` files. Special characters
> like em dashes cause PowerShell 5.1 encoding errors on unrelated lines.

### Step 3 — Register the scheduled task

Create `setup-scheduled-task.ps1` in the same folder:

```powershell
$scriptPath = "C:\Users\you\scripts\create-devlog.ps1"
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -File `"$scriptPath`""

$triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "7:00AM"),
    (New-ScheduledTaskTrigger -AtLogOn)
)

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 2)

Register-ScheduledTask `
    -TaskName "Obsidian Daily Devlog" `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description "Creates daily devlog note in Obsidian with Outlook meetings." `
    -Force
```

Replace the script path and adjust `7:00AM` to your preferred time. Run it once —
`-Force` means re-running it safely updates the task in place.

### Step 4 — Verify

Open Task Scheduler, find "Obsidian Daily Devlog" under the root folder, confirm it
shows as Ready. Right-click -> Run to test immediately.

---

## Adjusting the schedule

Re-run `setup-scheduled-task.ps1` with the new time, or edit the trigger directly in Task Scheduler.

## Uninstalling

```powershell
Unregister-ScheduledTask -TaskName "Obsidian Daily Devlog" -Confirm:$false
```

---

## Roadmap

See `PHASE2_GRAPH_API.md` for planned migration from Outlook COM to Microsoft Graph API,
which will remove the Outlook desktop dependency.
