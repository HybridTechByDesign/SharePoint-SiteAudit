# SharePoint site audit

A read-only PowerShell tool that scans one SharePoint Online site and shows, in plain language,
where folder structures have become too deep and what to move. See [CLAUDE.md](CLAUDE.md) for the
full brief.

**Status:** phase 2 of 7. It signs in, reads every file and folder, and writes a CSV with path lengths. The checks, grouping and the team report come next.

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

- `run-info.json`: when the scan ran, and the totals per library.

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
