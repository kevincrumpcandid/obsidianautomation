Set up a daily Obsidian devlog automation for this Windows user. Follow these steps in order:

1. Read `C:\Users\$USERNAME\AppData\Roaming\obsidian\obsidian.json` to find available vaults. Use the actual Windows username from the environment.
2. If multiple vaults exist, ask the user which one to use. If only one, confirm it with them.
3. Ask: what subfolder inside the vault should devlogs go in? (default: devlog)
4. Ask: what time should the daily note be created? (default: 7:00AM)

Then create these two files in the current working directory:

**create-devlog.ps1**

Behavior:
- Gets today's date as yyyyMMdd and yyyy-MM-dd
- Builds devlog folder path and creates it if missing
- Fetches today's meetings from Outlook via COM (see below)
- If the note does not exist: creates it with the title on line 1, then the Meetings section immediately below as the first subsection
- If the note already exists but has no Meetings section: inserts the Meetings section after the title line (not appended at the bottom)
- The Meetings section is idempotent — if it already exists the script skips it, so running twice is safe
- Opens the note via Obsidian URI

Outlook COM meeting fetch:
- Check if OUTLOOK.EXE is already running with Get-Process
- Instantiate via `New-Object -ComObject Outlook.Application`
- If Outlook was not already running, `Start-Sleep -Seconds 10` to allow sync
- Get the default Calendar folder (folder type 9)
- Set IncludeRecurrences = $true and Sort("[Start]") before calling Restrict
- Filter: `[Start] >= 'MM/dd/yyyy 12:00 AM' AND [Start] < 'MM/dd/yyyy 12:00 AM' AND [AllDayEvent] = False` using the next day for the upper bound
- Build meeting lines as: `- h:mm tt - h:mm tt  Subject`
- Return $null on any exception (so a COM failure does not block note creation)
- If $null is returned, create the note without a Meetings section

URI rules (PowerShell 5.1):
- NEVER put & inside a double-quoted string — it breaks the parser
- Build the URI by concatenating a single-quoted base string with the encoded file variable:
  `$uri = 'obsidian://open?vault=My%20Vault&file=' + $encodedFile`
- Encode the file path with `[Uri]::EscapeDataString(...)`

IMPORTANT: use only plain ASCII characters in the .ps1 file. Non-ASCII characters (em dashes, curly quotes, etc.) cause PowerShell 5.1 to misread file encoding and produce parser errors on unrelated lines.

**setup-scheduled-task.ps1**
- Task name: "Obsidian Daily Devlog"
- Action: `powershell.exe -WindowStyle Hidden -NonInteractive -File "<absolute_path>"`
- Two triggers: daily at the user's chosen time AND `New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME`
  (use -User to scope to the current user — omitting -User requires admin and will fail)
- Settings: StartWhenAvailable, ExecutionTimeLimit 2 minutes
- Uses -Force so re-running safely updates in place

After writing both files, run setup-scheduled-task.ps1 via PowerShell to register the task.

Remind the user: Obsidian must be open (or set to open on startup) for the URI to open the note. The logon trigger handles cases where the machine was asleep or off at the scheduled time.
