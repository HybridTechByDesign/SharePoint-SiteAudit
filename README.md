# SharePoint site audit

A read-only PowerShell tool that scans one SharePoint Online site and shows, in plain language,
where folder structures have become too deep and what to move. See [CLAUDE.md](CLAUDE.md) for the
full brief.

**Status:** phase 4 of 7. It signs in, reads every file and folder, runs the checks, groups the problem files into a few places to fix, and writes the data for the report. The HTML report for the team comes next.

The audit never creates, moves, renames, deletes or changes permissions on anything in SharePoint.

## Quick start

1. Install **PowerShell 7** once: https://aka.ms/powershell
2. Double-click **`Start-SiteAudit.cmd`**.
3. Pick a client from the menu, or choose **N** to set up a new one. Setup asks five questions:

   | Question | What to do |
   |---|---|
   | Client name | Any label, e.g. `<Client A>` |
   | Site address | Open the site in your browser and paste the address bar. Any page of the site works, and so does the admin center address. |
   | Sign-in tenant | Usually found automatically. If not, type the domain you sign in with, e.g. `<org>.onmicrosoft.com`. This can differ from the SharePoint address, which is common for trial and demo tenants. |
   | Sign-in app | Say **N** the first time for a tenant, then **Y** to set it up. A browser opens: sign in as a tenant admin and click Accept. The app can only read. |
   | Synced folder (optional) | On a team member's PC, copy the address bar of the synced library in File Explorer. Press Enter to skip. |
   | PC names (optional) | The team's PC names, separated by commas. Press Enter to skip. |

4. Press Enter to scan the remembered site, or paste a different one.

If the PnP.PowerShell module is missing, the tool offers to install it for you (no admin needed).

Everything you enter is saved in the `config` folder on your PC. Those files are never committed to
git, so client details stay local.

## What you get

Each run creates a folder `output\<site>_<yyyyMMdd-HHmm>\` with:

- `AllFindings.csv`: one row per file and folder (for IT, opens in Excel). The main columns:

  | Column | Meaning |
  |---|---|
  | `PathInLibrary` | Where the item sits inside the library |
  | `LocalPathLength` | Estimated path length on a synced PC (sync folder + path). Excel fails above 218, most Windows apps above 260 |
  | `ServerPathLength` | Length of the SharePoint path. SharePoint and OneDrive refuse paths above 400 |
  | `FolderDepth` | How many folders deep the item is inside the library |
  | `SizeBytes`, `Created`, `CreatedBy`, `Modified`, `ModifiedBy` | As shown in SharePoint |
  | `DirectChildren` | For folders: how many files and folders sit directly inside |
  | `CheckedOutTo` | Who has the file checked out, if anyone |

  Check columns (blank means fine). Filter `Issues` in Excel to see everything flagged on a row.

  | Column | Meaning |
  |---|---|
  | `Level` | `Green` under 200 · `Amber` 200 to 218 (close to the Excel limit) · `Red` over 218 (Excel can't open or save) · `DarkRed` over 260 (most Windows apps fail) · `Blocked` SharePoint path over 400 (won't sync) |
  | `Issues` | Everything flagged on this row, separated by `;` |
  | `NameTooLong` | A single name over 255 characters |
  | `RepeatedFolder` | The same or nearly the same folder name twice in the path, e.g. `Projects / Projects 2024` |
  | `ConflictCopyOf` | The file name ends in one of the PC names from setup (a OneDrive conflict copy) |
  | `LikelyDuplicates` | How many files have this same name and size across the site |
  | `BadName` | Leading or trailing space, ends with a dot, reserved name (CON, PRN, AUX, NUL, COM1-9, LPT1-9), or `_vti_` |
  | `JunkFile` | `~$` temp files, Thumbs.db, desktop.ini, .tmp, .pst |
  | `LargeFile` | Over the size set in the config (default 2 GB) |
  | `BulkCopyFolder` | The top folder of a tree that was copied in at once: 50 or more items created the same day by the same person who created the folder |
  | `SubtreeItems` | For folders: all items below it |
  | `TooManyInFolder` / `HugeSubtree` | Folder has over 5,000 items directly inside / over 100,000 below it |
  | `UniquePermissions` | The item has its own permissions instead of inheriting them |
  | `ProblemArea` | For files over the Excel limit: the break point folder this file is grouped under (see below) |

- `run-info.json`: when the scan ran, the totals for each check, and per library: item count against
  the 300,000 sync limit, settings that break sync (Require Check Out, content approval, offline sync
  off, required columns), checked-out files, and items with their own permissions.

- `ProblemAreas.csv`: every problem area, biggest first (see below).
- `summary.json`: the numbers the team report is built from: the headline sentence, the four tiles,
  the top 5 problem areas, the bar chart data and the other findings. Names of people are not in it.

At the end of a run, the headline and the top 5 places to fix are shown on screen.

### Problem areas (how files are grouped)

The team doesn't get a list of every long path. Files over the Excel limit are grouped by their
**break point**: the first folder below the fixed structure (Section / Year / Client / Topic), which
is usually where a copied-in folder tree starts. For each area you get:

- how many files are affected, the deepest path, and how many extra folder levels sit inside it
- the fix: *move the contents of the break point up into its parent*, how many characters that
  saves, and how many files drop back under the Excel limit
- if that isn't enough: *bring the files up out of the extra folder levels*, and how many files
  that fixes

Areas are ranked by files affected. The report shows the top 5; all of them are in `ProblemAreas.csv`.

Big libraries are read 2,000 items at a time, so libraries well above 5,000 items work. If SharePoint
asks the tool to slow down, it waits and tries again on its own.

## Running it from a PowerShell window

```powershell
.\Start-SiteAudit.ps1                     # same menu as the double-click

# or skip the menu:
.\Invoke-SiteAudit.ps1 -Config .\config\<client>.json
.\Invoke-SiteAudit.ps1 -Config .\config\<client>.json -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -Libraries "Documents","Archive"
```

## Changing a client's settings

Run setup again with the same client name and say yes to replace it. Or open
`config\<client>.json` in Notepad. The advanced settings there are:

| Setting | Meaning |
|---|---|
| `TenantId` | The sign-in domain or tenant ID, when it differs from the SharePoint name |
| `Thresholds` | Path length limits (Warn 200, Excel 218, Windows 260, SharePoint 400). Leave as is. |
| `MaxRecommendedDepthBelowSection` | Folder levels allowed below the section folder (default 3) |
| `LargeFileSizeGB` | Files above this size are flagged (default 2) |
| `FixedFolderLevels` | How many folder levels make up the fixed structure (default 4: Section / Year / Client / Topic). Break points are looked for below this. If the usual structure sits inside an extra test folder at the top of the library, add 1 (e.g. 5) |
| `ChartFolderLevel` | Which folder level gets one bar in the report chart (default 3: the client folder) |

## One-time app setup (manual way)

Setup does this for you. If you'd rather do it by hand, in PowerShell 7:

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SharePoint Site Audit (read-only)" `
    -Tenant <tenant>.onmicrosoft.com -SharePointDelegatePermissions "AllSites.Read"
```

Always include `-SharePointDelegatePermissions "AllSites.Read"`. Without it, PnP asks for full
control. Afterwards, check in the Entra admin center (App registrations > API permissions) that only
read permissions are listed, then paste the Application (client) ID into setup.

Registering the app is the only step that changes anything in a tenant, and setup asks before
doing it. The audit itself is read-only.

## Checks and tests

- `.\tools\Test-ReadOnly.ps1` checks the audit code for any command that could change SharePoint.
- Config tests use Pester 5 and never connect to SharePoint:
  ```powershell
  Install-Module Pester -MinimumVersion 5.5 -Scope CurrentUser -SkipPublisherCheck
  Invoke-Pester .\tests
  ```
