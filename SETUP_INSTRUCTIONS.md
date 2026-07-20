# Obsidian Daily Devlog Automation — Setup Instructions

Automatically creates and opens a dated devlog note in Obsidian every morning,
with today's Outlook calendar meetings pre-populated.

---

## What it does

The script runs at 7 AM daily and again on each logon:
1. Creates `devlog YYYYMMDD.md` in a `devlog` subfolder of your vault (skips if it already exists)
2. Fetches today's meetings from your published Outlook calendar feed (cloud — no Outlook desktop needed) and appends a `## Meetings` section (refreshed in place, so running twice is safe)
3. Opens the note in Obsidian

---

## Prerequisites

- Windows 10 or 11
- Obsidian installed and configured with a vault
- Node.js (used by `get-meetings.js` to fetch and expand the calendar feed)
- A published Outlook calendar ICS URL (see below)
- Obsidian set to **open on startup** (Settings -> About -> "Open Obsidian on system startup")

### Publishing your calendar

1. Go to [outlook.office.com](https://outlook.office.com) -> Settings -> **Calendar** -> **Shared calendars**
2. Under **Publish a calendar**, select your calendar, set the permission to
   **"Can view all details"** (lesser levels only show "Busy" with no subjects), and click **Publish**
3. Copy the **ICS** link and store it as a user environment variable (it is a secret —
   anyone with the URL can read your calendar; never commit it):

```powershell
[Environment]::SetEnvironmentVariable("OUTLOOK_ICS_URL", "https://outlook.office365.com/owa/calendar/.../calendar.ics", "User")
```

Then install the Node dependency once, from the repo folder:

```powershell
npm install
```

Note: the published feed can lag calendar changes by a few minutes, which is fine for a
morning snapshot. If your tenant has calendar publishing disabled, see `PHASE2_GRAPH_API.md`
for the Graph API alternative.

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

### Step 2 — Get the devlog script

Clone this repo (or copy `create-devlog.ps1`, `get-meetings.js`, and `package.json`
to a permanent folder — they must stay together), then run `npm install` in that folder.

In `create-devlog.ps1`, replace:
- the `$vaultPath` value with your actual vault path
- `Dev%20Docs` in the Obsidian URI at the bottom with your vault name (spaces as `%20`)
- `devlog` with your preferred subfolder name

Set the `OUTLOOK_ICS_URL` environment variable as described in Prerequisites.

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

## History

Meetings originally came from Outlook desktop via COM; as of 2026-06-10 they come from
the published calendar ICS feed instead, removing the Outlook desktop dependency.
`PHASE2_GRAPH_API.md` documents the Graph API alternative if calendar publishing is
ever disabled by the tenant.
