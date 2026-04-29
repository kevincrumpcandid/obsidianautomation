# Phase 2: Microsoft Graph API Calendar Integration

Replace the Outlook COM approach in `create-devlog.ps1` with direct Graph API calls.
This removes the dependency on Outlook being installed or running.

## Why we are on COM for now

Graph Explorer returned `ErrorAccessDenied` for `GET /me/calendarView` when tested
with kevin.crump@candid.org on 2026-04-29. Candid's tenant either restricts user
consent for Graph API apps or has blocked the Graph Explorer client ID entirely.

COM works without any tenant permissions and Outlook is reliably open during working hours.

---

## What is needed to unblock Phase 2

1. **Option A — IT grants user consent**
   Ask the Candid IT/Azure team to allow delegated `Calendars.Read` for your account,
   or to pre-authorize a registered app. This is the lowest-friction path.

2. **Option B — Register your own Azure AD app**
   - Go to portal.azure.com -> Azure Active Directory -> App registrations -> New registration
   - Name: "Obsidian Devlog" / Supported account types: single tenant
   - No redirect URI needed (device code flow)
   - Under API Permissions add: Microsoft Graph -> Delegated -> Calendars.Read
   - If admin consent is required, submit the request from that screen

---

## Graph API endpoint

```
GET https://graph.microsoft.com/v1.0/me/calendarView
    ?startDateTime=2026-04-29T00:00:00
    &endDateTime=2026-04-30T00:00:00
    &$select=subject,start,end,isAllDay
    &$orderby=start/dateTime
```

Filter out all-day events (`isAllDay -eq $false`) after fetching.

---

## PowerShell implementation sketch

Requires the `Microsoft.Graph` module or raw REST calls via `Invoke-RestMethod`.

### Device code flow (interactive, run once to cache token)

```powershell
# Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -ClientId "YOUR_APP_CLIENT_ID" -TenantId "YOUR_TENANT_ID" -Scopes "Calendars.Read"

$today = (Get-Date).Date.ToString("yyyy-MM-ddTHH:mm:ss")
$tomorrow = (Get-Date).Date.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")

$events = Get-MgUserCalendarView -UserId "me" `
    -StartDateTime $today `
    -EndDateTime $tomorrow `
    -Property "subject,start,end,isAllDay" |
    Where-Object { -not $_.IsAllDay } |
    Sort-Object { $_.Start.DateTime }

$lines = foreach ($e in $events) {
    $start = [DateTime]$e.Start.DateTime
    $end = [DateTime]$e.End.DateTime
    "- " + $start.ToString("h:mm tt") + " - " + $end.ToString("h:mm tt") + "  " + $e.Subject
}
```

### Token caching

`Connect-MgGraph` caches the token on disk via the Graph SDK. Subsequent script runs
will reuse it silently until expiry (typically 1 hour access token, 90-day refresh token).
No re-authentication needed for the daily scheduled task once the initial login is done.

---

## Migration steps when ready

1. Confirm Graph access works in Graph Explorer
2. Register the Azure AD app (or get client ID from IT)
3. Run `Connect-MgGraph` once interactively to cache credentials
4. Replace the `Get-TodayMeetings` function in `create-devlog.ps1` with the Graph version above
5. Extend `setup-scheduled-task.ps1` execution time limit to 3 minutes (token refresh can be slow)
6. Remove the `Start-Sleep` that was needed for Outlook cold-start

---

## Files to change

- `create-devlog.ps1` — replace `Get-TodayMeetings` function only, nothing else changes
- `setup-scheduled-task.ps1` — bump `-ExecutionTimeLimit` from 2 to 3 minutes
- `SETUP_INSTRUCTIONS.md` — add a note about running `Connect-MgGraph` once after setup
