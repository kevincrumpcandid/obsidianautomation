Set up a daily Obsidian devlog automation for this Windows user. Follow these steps in order:

1. Read `C:\Users\$USERNAME\AppData\Roaming\obsidian\obsidian.json` to find available vaults. Use the actual Windows username from the environment.
2. If multiple vaults exist, ask the user which one to use. If only one, confirm it with them.
3. Ask: what subfolder inside the vault should devlogs go in? (default: devlog)
4. Ask: what time should the daily note be created? (default: 7:00AM)

Then create these two files in the current working directory:

**create-devlog.ps1**
- Gets today's date as yyyyMMdd and yyyy-MM-dd
- Builds the devlog folder path from the vault path and chosen subfolder
- Creates the subfolder if it does not exist
- Builds the note file path as: `devlog yyyyMMdd.md`
- If the file does not already exist, creates it with content: `# Devlog - yyyy-MM-dd` followed by two newlines
- Builds the Obsidian URI by concatenating a single-quoted base string with the encoded file variable — never put & inside a double-quoted string in PowerShell 5.1, it breaks the parser
- Opens the note via `obsidian://open?vault=<encoded_name>&file=<encoded_path>` using Start-Process
- IMPORTANT: use only ASCII characters in the script — non-ASCII characters like em dashes cause PowerShell 5.1 to misread encoding and break subsequent lines

**setup-scheduled-task.ps1**
- Creates a scheduled task named "Obsidian Daily Devlog"
- Action: `powershell.exe -WindowStyle Hidden -NonInteractive -File "<absolute_path_to_create-devlog.ps1>"`
- Trigger: daily at the user's chosen time
- Settings: StartWhenAvailable, ExecutionTimeLimit 1 minute
- Uses -Force so re-running it is safe

After writing both files, run setup-scheduled-task.ps1 via PowerShell to register the task.

Finish by reminding the user that Obsidian must be running (or configured to open on startup) for the URI to open the note automatically.
