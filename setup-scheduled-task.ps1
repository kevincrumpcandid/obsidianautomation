$scriptPath = Join-Path $PSScriptRoot "create-devlog.ps1"
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

Write-Host "Done. Task 'Obsidian Daily Devlog' will run every day at 7:00 AM."
