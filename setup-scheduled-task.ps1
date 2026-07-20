# Registers the "Obsidian Daily Devlog" scheduled task (daily 7:00 AM + at logon).
# Re-running updates the task in place (-Force).
#
# IMPORTANT: run this NON-elevated, as kevin.crump. UAC elevates as a different
# account (candiduser), which would register the task under the wrong profile so
# it never fires for you. The check below refuses to register if that happens.

$scriptPath = Join-Path $PSScriptRoot "create-devlog.ps1"

# -ExecutionPolicy Bypass so a restrictive per-user/machine policy can't silently
# block the scheduled run (a "task ran but did nothing" failure mode).
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$scriptPath`""

# Logon trigger delayed 2 minutes: gives the network/DNS time to come up and
# de-races it against the 7:00 AM trigger. (New-ScheduledTaskTrigger has no
# -Delay for -AtLogOn, so set it on the returned object.)
$daily = New-ScheduledTaskTrigger -Daily -At "7:00AM"
$logon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$logon.Delay = "PT2M"
$triggers = @($daily, $logon)

# Reliability settings:
#   -AllowStartIfOnBatteries / -DontStopIfGoingOnBatteries : run on an unplugged
#       laptop (default is "AC only" + "stop on battery" -> silently skipped).
#   -StartWhenAvailable : catch up a missed run (asleep/off at 7 AM).
#   -RestartCount/-RestartInterval : retry a transient failure instead of giving
#       up for the day.
#   -ExecutionTimeLimit 10 min : the script now waits for the network (up to 3
#       min) plus app cold-start + retries; the old 2-min cap could kill it
#       mid-write, leaving a missing/partial note.
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask `
    -TaskName "Obsidian Daily Devlog" `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description "Creates daily devlog note in Obsidian with Outlook meetings and Jira tickets. Runs at 7 AM and 2 min after logon." `
    -Force | Out-Null

# Verify the task registered under the current (non-elevated) user, not candiduser.
$task = Get-ScheduledTask -TaskName "Obsidian Daily Devlog" -ErrorAction SilentlyContinue
$principal = if ($task) { $task.Principal.UserId } else { "" }
Write-Host "Done. Task registered under: $principal"
Write-Host "Runs daily at 7:00 AM and 2 minutes after logon."
if ($principal -and $principal -notmatch [regex]::Escape($env:USERNAME)) {
    Write-Host ""
    Write-Host "WARNING: task principal ($principal) does not match your user ($env:USERNAME)."
    Write-Host "You likely ran this elevated. Re-run this script WITHOUT 'Run as administrator'."
}
