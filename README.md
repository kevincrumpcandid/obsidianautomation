# Obsidian Daily Devlog Automation

A Windows PowerShell automation that builds a dated **developer log** note in an
[Obsidian](https://obsidian.md/) vault every morning, pre-populated with that day's
Outlook calendar meetings and the current/next Jira sprint stories assigned to me.
Both Claude Code and Obsidian work in Markdown, which makes this a lightweight daily
log compared to OneNote.

---

## What it does

A single scheduled task runs `create-devlog.ps1` daily at **7:00 AM** and again **at
logon**. Each run:

1. **Waits for the network, then ensures Obsidian is up** — polls DNS until the Jira
   and calendar hosts resolve (bounded), then launches Obsidian if it is not already
   running. Meetings come from the cloud calendar feed, so Outlook desktop is not needed.
2. **Locates/creates today's note** at `…\<vault>\devlog\<yyyyMM>\devlog <yyyyMMdd>.md`
   (a per-month folder, e.g. `devlog\202606\devlog 20260610.md`).
3. **Writes a friendly title** — `# Wednesday, June 10, 2026` (rebuilt deterministically
   each run, so it self-heals duplicate/malformed headers).
4. **`## Meetings` table** — `| Time | Meeting | Owner | Summary |` from the published
   Outlook calendar feed (via `get-meetings.js`). The Summary shows the meeting location
   (Zoom/Teams URLs become short clickable links like `[Zoom](…)`) and appends `Agenda: …`
   when the invite body contains a genuine agenda (join-link boilerplate is stripped).
5. **`## Tickets` table** — `| Ticket | Summary | Status | Points |` from Jira, grouped
   under bold **sprint subsection rows** for the active sprint and the next (future)
   sprint, e.g. `**UMT Sprint 5 (active)**` / `**UMT Sprint 6 (next)**`.
6. **Opens/focuses today's note** in Obsidian.

Both managed sections are created if missing and refreshed in place if present; any
other sections you add to the note by hand are preserved. Running multiple times a day
is safe and idempotent.

### Reliability & logging

The run is designed to fail loudly, never silently:

- **Network wait** — logon/wake runs poll DNS until the Jira and calendar hosts resolve
  (up to 3 min) before fetching, so a not-yet-ready network no longer yields empty sections.
- **Bounded, retried fetches** — the Jira call retries with a 30s timeout; the calendar
  fetch is bounded by a 20s timeout in `get-meetings.js`.
- **Failure banners** — if a section genuinely fails to load, the note shows a
  `> [!warning] Meetings unavailable - <reason> (see log)` callout instead of quietly
  omitting the section, and the **last good table is preserved** beneath it. A truly empty
  day still shows `_No meetings scheduled_`.
- **Logs** — every run writes to `%LOCALAPPDATA%\obsidian-devlog\logs`: a rolling
  `devlog.log` (one line per event) plus a per-run `transcript-<timestamp>.log`. Both are
  pruned after 30 days. Check these first when a section comes up empty.
- **Scheduled task** — registered to run on battery, restart up to 3× on failure, and
  with a 10-minute execution limit (the network wait + retries need headroom). The logon
  trigger is delayed 2 minutes to let the network come up.

### Example output

```markdown
# Wednesday, June 10, 2026

## Meetings

| Time | Meeting | Owner | Summary |
| --- | --- | --- | --- |
| 12:00 PM - 1:00 PM | UMT Sprint Planning & Demo | Matthew Grahovac | [Zoom](https://candid.zoom.us/j/84144113303) |
| 2:30 PM - 3:30 PM | UMT Sprint Planning | Virginia Holmstrom | [Zoom](https://candid.zoom.us/j/8839...); Wolf359 Conference Room |

## Tickets

| Ticket | Summary | Status | Points |
| --- | --- | --- | --- |
| **UMT Sprint 5 (active)** |  |  |  |
| [UMT-110](https://candidprojects.atlassian.net/browse/UMT-110) | Clone StripeDev sandbox | Done | 8 |
| **UMT Sprint 6 (next)** |  |  |  |
| [UMT-70](https://candidprojects.atlassian.net/browse/UMT-70) | Store/retrieve Stripe Customer ID | In Progress | 5 |
```

---

## Files

| File | Summary |
| --- | --- |
| `create-devlog.ps1` | **The script.** Combined meetings + Jira generator (everything in "What it does"). Self-contained: app-launch/wait, section-aware Markdown parser (`Parse-Note`/`Rebuild-Note`/`Find-Section`), `Ensure-Header`, `Get-TodayMeetings` (published ICS via `get-meetings.js`), `Get-JiraTickets` (Jira REST), table builders, `Clean-Summary`/`Shorten-Location` for meeting summaries. Edit the config block at the top (`$vaultPath`, `$vaultName`, `$jiraEmail`) for another machine/user. |
| `setup-scheduled-task.ps1` | Registers the **`Obsidian Daily Devlog`** scheduled task (daily 7:00 AM + at logon) pointing at `create-devlog.ps1` in this folder (`$PSScriptRoot`). Uses `-Force`, so re-running updates the task in place. This is the only task you need. |
| `setup-jira-task.ps1` | **Legacy/superseded.** Used to register a separate `Obsidian Daily Jira Tickets` task (7:05 AM) for the standalone Jira script. That task has been unregistered now that `create-devlog.ps1` does both jobs. Kept for reference; do not run it. |
| `create-devlog-jira.ps1` | **Legacy/superseded.** The original standalone Jira-tickets script before it was merged into `create-devlog.ps1`. No longer scheduled; safe to delete. |
| `SETUP_INSTRUCTIONS.md` | Original step-by-step setup guide. **Predates** the Jira integration, monthly folders, friendly title, and tables — describes only the meetings-only version. Useful for the general approach (finding your vault, registering the task) but not literally current. |
| `PHASE2_GRAPH_API.md` | Roadmap doc: planned migration of `Get-TodayMeetings` from Outlook COM to the Microsoft Graph API (removes the Outlook-desktop dependency). Blocked as of 2026-04-29 by tenant Graph permissions; includes the implementation sketch and unblock options. |
| `README.md` | This file. |
| `.claude/` | Claude Code project settings for this repo. |

---

## Configuration

Set in the config block at the top of `create-devlog.ps1`:

- `$vaultPath` — absolute path to the Obsidian vault (currently the Candid OneDrive `Dev Docs` vault).
- `$vaultName` — vault display name used in `obsidian://open` URIs (`Dev Docs`).
- `$jiraEmail` — Atlassian account email for Basic auth (`kevin.crump@candid.org`).
- Notes are written to `…\<vault>\devlog\<yyyyMM>\devlog <yyyyMMdd>.md`.

### Jira integration

- **Auth:** HTTP Basic with `$jiraEmail` + a **personal Atlassian API token** read from the
  **User** environment variable `JIRA_API_TOKEN`. Create one at
  <https://id.atlassian.com/manage-profile/security/api-tokens> — it must be a real API
  token (starts with `ATATT…`), **not** a Marketplace-app JWT (`eyJ…`), which 401s.
  The script reads the token via `[Environment]::GetEnvironmentVariable(...,"User")` to
  avoid a stale inherited process value. If the token is missing, the Tickets section is
  skipped silently (meetings still run).
- **Endpoint:** `/rest/api/3/search/jql` (the plain `/search` endpoint returns 410 on this tenant).
- **JQL:** `project=UMT AND assignee=currentUser() AND (sprint in openSprints() OR sprint in futureSprints()) AND issuetype not in subTaskIssueTypes() ORDER BY rank ASC`.
- **Field ids (Candid tenant, project UMT):** story points = `customfield_12108`
  ("Story point estimate"); sprint = `customfield_10006`. (The Jira defaults
  `customfield_10016`/`10020` do **not** apply here.)

---

## First-time setup

1. Ensure Obsidian and Outlook desktop are installed; the target vault exists.
2. Create a Jira API token and set it as a **User** env var `JIRA_API_TOKEN` (log off/on
   so scheduled tasks inherit it).
3. From this folder, run `setup-scheduled-task.ps1` once (registers the daily task).
4. Optional: run `create-devlog.ps1` directly to generate today's note immediately.

## Maintenance

- **Change the schedule:** edit the time in `setup-scheduled-task.ps1` and re-run it,
  or edit the trigger in Task Scheduler.
- **Uninstall:** `Unregister-ScheduledTask -TaskName "Obsidian Daily Devlog" -Confirm:$false`
- **Rule for `.ps1` edits:** keep the script **plain ASCII** — non-ASCII characters (em
  dashes, curly quotes, box-drawing chars) cause PowerShell 5.1 encoding errors. Build
  any needed Unicode (e.g. box-drawing separators) from `[char]0xNNNN` at runtime.

---

## Roadmap

See `PHASE2_GRAPH_API.md` — migrate calendar reads from Outlook COM to Microsoft Graph
to drop the Outlook desktop dependency (pending tenant Graph permissions).
