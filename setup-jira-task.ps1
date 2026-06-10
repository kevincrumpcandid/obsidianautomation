if (-not $env:JIRA_API_TOKEN) {
    Write-Host "WARNING: JIRA_API_TOKEN user environment variable is not set."
    Write-Host "1. Create an API token at: https://id.atlassian.com/manage-profile/security/api-tokens"
    Write-Host "2. Set it as a User environment variable named JIRA_API_TOKEN"
    Write-Host "   (System Properties -> Advanced -> Environment Variables -> User variables -> New)"
    Write-Host "3. Log off and back on so the task runner inherits it."
    Write-Host ""
    Write-Host "Registering scheduled task anyway. Script will skip silently until the token is set."
    Write-Host ""
}

$scriptPath = Join-Path $PSScriptRoot "create-devlog-jira.ps1"
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -File `"$scriptPath`""

$triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "7:05AM"),
    (New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME)
)

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 2)

Register-ScheduledTask `
    -TaskName "Obsidian Daily Jira Tickets" `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description "Inserts current sprint Jira stories into the daily devlog. Runs at 7:05 AM and on logon." `
    -Force

Write-Host "Done. Task will run daily at 7:05 AM and on each logon."
