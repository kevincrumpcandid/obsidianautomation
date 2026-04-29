$scriptPath = Join-Path $PSScriptRoot "create-devlog.ps1"
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -File `"$scriptPath`""

$triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "7:00AM"),
    (New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME)
)

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 2)

Register-ScheduledTask `
    -TaskName "Obsidian Daily Devlog" `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description "Creates daily devlog note in Obsidian with Outlook meetings. Runs at 7 AM and on logon." `
    -Force

Write-Host "Done. Task will run daily at 7:00 AM and on each logon."
