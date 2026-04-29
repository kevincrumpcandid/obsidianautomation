# Obsidian Daily Devlog Automation — Setup Instructions

Automatically creates and opens a dated devlog note in Obsidian every morning.

---

## What it does

At a scheduled time each day, a PowerShell script:
1. Creates `devlog YYYYMMDD.md` in a folder of your choice inside your vault (skips if it already exists)
2. Opens it directly in Obsidian

---

## Prerequisites

- Windows 10 or 11
- Obsidian installed and configured with a vault
- Obsidian set to **open on startup** (Settings -> About -> "Open Obsidian on system startup"), otherwise the URI won't open the note if Obsidian isn't already running

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

Open `%APPDATA%\obsidian\obsidian.json` in a text editor. You'll see something like:

```json
{"vaults":{"abc123":{"path":"C:\\Users\\you\\Documents\\My Vault"}}}
```

Note down the `path` value and your vault's display name (the last folder in the path).

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

$encodedFile = [Uri]::EscapeDataString("devlog/devlog $date")
$uri = 'obsidian://open?vault=My%20Vault&file=' + $encodedFile
Start-Process $uri
```

Replace:
- `C:\Users\you\Documents\My Vault` with your actual vault path
- `My%20Vault` with your vault name (spaces encoded as `%20`)
- `devlog` with your preferred subfolder name (use the same value in both places)

> **Important:** Use only plain ASCII characters in this file. Special characters like em dashes cause PowerShell 5.1 to misread the file encoding and produce parser errors on unrelated lines.

### Step 3 — Register the scheduled task

Create `setup-scheduled-task.ps1` in the same folder:

```powershell
$scriptPath = "C:\Users\you\scripts\create-devlog.ps1"
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At "7:00AM"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 1)

Register-ScheduledTask `
    -TaskName "Obsidian Daily Devlog" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Creates daily devlog note in Obsidian and opens it" `
    -Force
```

Replace `C:\Users\you\scripts\create-devlog.ps1` with the actual path to your script and adjust `7:00AM` to your preferred time.

Run it once in PowerShell (right-click -> Run with PowerShell, or open a PowerShell terminal and call it directly). You only need to run this once to register the task.

### Step 4 — Verify

Open Task Scheduler (search "Task Scheduler" in Start), find "Obsidian Daily Devlog" under the root folder, and confirm it shows as Ready. Right-click -> Run to test it immediately.

---

## Adjusting the schedule

To change the time after initial setup, either re-run `setup-scheduled-task.ps1` with the new time (the `-Force` flag updates in place), or edit the task directly in Task Scheduler.

## Uninstalling

```powershell
Unregister-ScheduledTask -TaskName "Obsidian Daily Devlog" -Confirm:$false
```
