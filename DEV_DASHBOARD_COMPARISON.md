# Future: One-Stop Developer Dashboard — Build-Path Comparison

A dev-effort comparison for rebuilding this Obsidian devlog automation into a single
Candid **Developer Dashboard**. Goals driving the comparison:

- Access **Outlook Calendar in the cloud via Microsoft Graph** (not Outlook COM).
- Produce richer output: **OneNote or Word** (instead of Markdown).
- Add modules for **Azure DevOps** and **Slack** alongside the existing **Jira** integration.
- Two candidate builds: a **Power App / Power Platform** solution, or an **npm/Node.js**
  project executable from a PowerShell scheduled task ("cron").

---

## Shared foundation (both paths need this)

| Concern | Reality |
|---|---|
| **Cloud Outlook calendar** | Microsoft Graph `GET /me/calendarView`. Needs an Azure AD app + delegated `Calendars.Read`. **Known blocker** — see `PHASE2_GRAPH_API.md`: tenant denied Graph Explorer consent on 2026-04-29. |
| **OneNote output** | Graph `POST /me/onenote/pages` (HTML body) — clean fit for the existing tables. |
| **Word output** | No native Graph "create Word." Either generate `.docx` in code or fill a Word template via a low-code connector. More work than Markdown/OneNote. |
| **Azure DevOps** | REST API (work items via WIQL) — analogous to the current Jira calls. |
| **Slack** | Slack Web API (`chat.postMessage`, `conversations.history`, etc.). |

---

## Path comparison

| Dimension | **Power Platform** (Power Apps UI + Power Automate flows) | **Node.js/TypeScript CLI** (scheduled via Task Scheduler) |
|---|---|---|
| **Cloud Outlook auth** | Office 365 Outlook connector is **pre-consented** — sidesteps the Graph app-registration/consent blocker entirely | Must get IT to register an AAD app + admin-consent `Calendars.Read` (current blocker); you own MSAL token caching/refresh |
| **OneNote / Word** | Native OneNote connector; "Populate a Word template" connector | Graph for OneNote; `docx`/`docxtemplater` npm for Word (full control, more code) |
| **Azure DevOps / Jira / Slack** | Native connectors exist — but **premium** | `azure-devops-node-api`, existing Jira logic, `@slack/web-api` — all free |
| **Scheduling** | Built-in scheduled cloud flows | PowerShell scheduled task running `node dist/index.js` (real cron on Linux) |
| **Complex formatting** (sprint grouping, summary cleanup, link shortening) | Painful — string/array logic in flows is clunky | Trivial — current PowerShell logic ports ~1:1 to TS |
| **Interactive dashboard** | Power Apps gives a UI out of the box | Needs a web frontend (e.g. React + small Express/API) — extra build |
| **Cost / licensing** | Premium connectors + Power Apps per-user/per-app licensing; governance approval | Free; nothing to host (local task) or a cheap container if a UI is added |
| **Source control / testing / reuse** | Weak (solution ALM, hard to unit test) | Strong (git, Jest, reusable `modules/{outlook,jira,ado,slack,onenote}`) |
| **Maintenance owner** | Low-code; ops-friendly | Real code; dev-friendly |

---

## Rough effort (relative)

- **Power Platform:** low-code build of ~4-5 flows + 1 app; the slog is **licensing/governance
  approval** and fighting flow logic for formatting. Fast to a working passive doc; slower
  for rich/clean output.
- **Node.js:** higher upfront (project scaffold, MSAL auth, 4 source modules, OneNote/Word
  renderer); existing logic transfers, and it scales into a real dashboard. The slog is the
  **Graph AAD consent** (same blocker) + owning secrets/auth.

---

## Recommendation

- **Priority = cloud calendar + multi-source, soon, least auth pain:** Power Platform — its
  pre-consented Outlook connector dodges the exact Graph consent wall currently blocking us,
  and Jira/ADO/Slack/OneNote connectors exist. Accept premium licensing + formatting limits.
- **Priority = control, rich output, testability, free, reuse of existing logic:** Node/TS CLI —
  but **first unblock the AAD app + `Calendars.Read` consent with IT** (Option A/B in
  `PHASE2_GRAPH_API.md`); that single dependency gates the whole cloud-Outlook story.
- **Pragmatic hybrid:** keep the generation engine in Node (reuse logic, emit OneNote via
  Graph), and add a Power App / React surface later only if an interactive dashboard is wanted
  rather than a daily document.

---

## Next step (if pursued)

Turn the chosen path into a concrete architecture plan: repo layout, auth flow, per-integration
API calls (Graph calendar + OneNote, Azure DevOps WIQL, Slack Web API, Jira), and the formal IT
ask for Graph `Calendars.Read` consent.
