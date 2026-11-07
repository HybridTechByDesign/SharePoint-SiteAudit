# SharePoint Site Audit — Project Brief

## What this is
A read-only PowerShell tool that scans one SharePoint Online site and produces a plain-language
HTML report showing where folder structures have become too deep (path length), why, and what to
move. Built by an IT vendor who manages multiple client tenants.

The audience for the report is a **non-technical team** (e.g. an accounting department). The goal is
to show them it is a **behaviour problem** (copying deep folder trees in), not a technical fault.

## The problem it solves
- A team syncs a SharePoint library with the **Sync button** (not shortcuts — shortcuts are ruled out).
- They collect data from another team, save it to their desktop, then move whole folder trees into
  the synced library.
- Result: paths too long → OneDrive sync errors → conflict copies (`<name>-<PCNAME>`) → OneDrive
  freezes → Outlook can't attach files from OneDrive.
- Typical structure: `Shared Documents/<Section>/<Year> Operation/<Client name>/<N. Topic>/...`
  The fixed part already uses ~63 characters before any user folders.

## Hard constraints
- **Read-only.** The scanner must never create, move, rename, delete or change permissions on anything.
- **Never hardcode a client company name** anywhere (code, examples, tests, sample data). Use
  placeholders like `<Client A>`, `<org>`, `<tenant>`. All tenant specifics live in config files.
- **Keep it simple and low-maintenance.** The owner is not a programmer. No web server, no database,
  no build step, no frameworks that need npm. PowerShell + one self-contained HTML file.
- Must handle libraries far above 5,000 items (page through with `-PageSize`, never unfiltered views).
- Respect throttling (HTTP 429): back off and retry.

## Stack
- PowerShell 7 + `PnP.PowerShell` module.
- Auth: per-tenant Entra app registration (`Register-PnPEntraIDAppForInteractiveLogin`), read-only
  SharePoint permissions, interactive sign-in.
- Report: one self-contained `Report.html` (inline CSS/JS; Chart.js may be embedded inline or loaded
  from a CDN — prefer inline so it works offline and can be emailed).
- Data export: `AllFindings.csv` (full per-file detail, IT use only).

## Usage (target)
```powershell
.\Invoke-SiteAudit.ps1 -Config .\config\<client>.json -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>"
# optional: -Libraries "Documents","Archive"   (default: all document libraries)
# optional: -CompareTo .\output\<previous-run-folder>
```
Output goes to `.\output\<site>_<yyyyMMdd-HHmm>\` containing `Report.html`, `AllFindings.csv`, `summary.json`.

## Config file (one per client tenant)
```json
{
  "TenantName": "<tenant>",
  "ClientId": "<entra-app-id>",
  "SyncRootLength": 70,
  "PcNames": ["<PC-01>", "<PC-02>"],
  "Thresholds": { "Warn": 200, "Excel": 218, "Windows": 260, "SharePoint": 400 },
  "MaxRecommendedDepthBelowSection": 3
}
```
- `SyncRootLength` = measured length of the local sync root on a user PC
  (e.g. `C:\Users\<user>\<Org>\<Site> - <Library>`). Used to estimate local path length.
- `PcNames` = used to detect conflict copies.

## Path length limits (thresholds)
Measured on the **estimated local path** = `SyncRootLength` + 1 + path relative to the library root
(backslashes), except the 400 rule which uses the decoded SharePoint server-relative path.

| Level | Length | Meaning (plain language for report) |
|---|---|---|
| Green | < 200 | Fine |
| Amber | 200–218 | Close to the Excel limit |
| Red | > 218 | Excel can't open or save it ("Filename is not valid") |
| Dark red | > 260 | Most Windows apps fail (Explorer, Word, Outlook attach, zip) — LongPathsEnabled does not help because apps must opt in |
| Blocked | > 400 (SharePoint path) | SharePoint/OneDrive refuses it; sync errors |
| Also flag | any single name > 255 | Name too long |

## Checks the scanner performs
Per item (files and folders):
1. Server-relative path, decoded path length, estimated local path length, folder depth, size,
   created/modified dates, modified by.
2. Threshold level (table above).
3. **Repeated folder levels** — same or near-same folder name appearing twice in a path
   (e.g. `X\X\` or `Projects\Projects 2024\Projects`).
4. **Conflict copies** — file name ends with `-<PCNAME>` from config.
5. **Likely duplicates** — same file name + same size in different folders.
6. **Bad names** — leading/trailing spaces, name ending in `.`, reserved names (`CON`, `PRN`, `AUX`,
   `NUL`, `COM1-9`, `LPT1-9`), `_vti_` in name.
7. **Junk files** — `~$*`, `Thumbs.db`, `desktop.ini`, `*.tmp`, `*.pst`.
8. **Large files** — over a configurable size (default 2 GB).
9. **Bulk-dump detection** — folders where many items were created on the same day by the same person
   (signals a copied-in folder tree).

Per folder:
10. Direct child count > 5,000 (list view threshold) and subtree count > 100,000.

Per library:
11. Total item count vs 300,000 (sync recommendation).
12. Settings that break sync: Require Check Out, required columns, content approval, offline sync
    disabled, checked-out files.
13. Count of items with unique permissions.

## Grouping logic (the key feature — do NOT list every file to the team)
1. Take every file at Red or worse.
2. For each, find its **break point**: the highest folder such that every file beneath it that is
   over the limit shares that folder, and the folder sits below the fixed structure
   (Section/Year/Client/Topic). Group files by break point.
3. For each group compute: files affected, deepest path length, number of extra folder levels inside
   the break point, and the **fix**: "move contents of `<break point>` up into `<parent>`" plus the
   characters saved and how many files would drop back under the Excel limit (218).
4. Rank groups by files affected. Report shows **top 5**; all groups and all files go in the CSV.

## Report layout (Report.html)
Written for non-technical readers. Sentence case, no jargon, every card ends with an action.
1. **Headline sentence** — e.g. "300 files can't be reliably opened or synced. They're in 6 places.
   214 of them can't be opened in Excel."
2. **Four tiles** — files scanned · over Excel limit · over sync limit · deepest nesting.
3. **Hanging bar chart** — one bar per client folder (configurable grouping level), bars hang
   downward from the folder label, length = longest local path in that folder, value label on each
   bar, dashed lines at 218 and 400, bar colour by threshold level. Click a bar → show its subfolders
   (drill-down). Tooltip: files over 218 in that folder.
4. **Top 5 problem areas** — one card each: short breadcrumb tree to the break point, files affected,
   deepest path, fix, characters saved.
5. **How to avoid this** — fixed four rules:
   - Bring files, not folder trees.
   - Skip the desktop step — upload in the browser or use a shared link.
   - Don't repeat the client name in sub-folder or file names.
   - No more than 3 folder levels below the section folder.
6. **Other findings (collapsed)** — conflict copies, duplicates, bad names, junk files, library
   settings — counts only, details in CSV.
7. **Footer** — scan date, site, library list, and comparison with previous run if `-CompareTo` given
   ("last scan: 300 → now: 12").
- Hide "modified by" names in the HTML by default (`-ShowNames` switch to include); always present in CSV.

## Build phases
1. **Auth + config** — load config, connect with PnP, list libraries. Test on a small site.
2. **Scanner** — enumerate all items with paging, compute per-item fields, write raw CSV. Handle 429.
3. **Checks** — implement checks 2–13, add columns to CSV.
4. **Grouping** — break-point logic and fix suggestions; write `summary.json`.
5. **Report** — generate `Report.html` from `summary.json` (template file + data injection).
6. **Compare** — `-CompareTo` previous run.
7. **Docs** — short README: one-time app registration per tenant, how to run, how to read the report.

## Acceptance criteria
- Scans a library with 50k+ items without hitting the 5,000 threshold error.
- No write operations anywhere (grep for `Set-`, `Remove-`, `Move-`, `Add-`, `New-PnP` on site objects
  should find nothing in the scanner).
- Report opens offline by double-click, readable by a non-technical user, top 5 areas only.
- Path lengths verified by hand against 3–5 sample files.
- No client names anywhere in the repo.

## Out of scope (for now)
- Any fix/move/rename mode (future: separate script driven by an approved CSV).
- PC-side check script (future, Intune-deployed, read-only).
- OneDrive shortcuts guidance (ruled out).
- Purview audit log lookup of who moved what (possible later add-on).
- Integration into the owner's planned multi-tenant M365 script launcher (later).
