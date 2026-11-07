# SharePoint site audit

A read-only PowerShell tool that scans one SharePoint Online site and shows, in plain language,
where folder structures have become too deep and what to move. See [CLAUDE.md](CLAUDE.md) for the
full brief.

**Status:** phase 1 of 7 (sign-in, config, list libraries). Scanning and the report come next.

The tool never creates, moves, renames, deletes or changes permissions on anything in SharePoint.
Run `.\tools\Test-ReadOnly.ps1` at any time to check the code for write commands.

## What you need

- PowerShell 7 (`pwsh`). Get it from https://aka.ms/powershell
- The PnP.PowerShell module, version 2.12 or later:
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```

## One-time setup per client tenant

1. **Register a read-only app** in the client's Entra ID. You need to be a Global Admin (or
   Application Admin) in that tenant to grant consent.
   ```powershell
   Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SharePoint Site Audit (read-only)" `
       -Tenant <tenant>.onmicrosoft.com -SharePointDelegatePermissions "AllSites.Read"
   ```
   Always pass `-SharePointDelegatePermissions "AllSites.Read"`. Without it, the command asks for
   full control. A browser opens; sign in and accept. Afterwards, in the Entra portal under
   App registrations > API permissions, check that the app lists read permissions only.
   Copy the **Application (client) ID** it prints.

2. **Create the config file.** Copy `config\example.json` to `config\<client>.json` and fill in:

   | Setting | What to put |
   |---|---|
   | `TenantName` | The part before `.sharepoint.com`, e.g. `<tenant>` |
   | `ClientId` | The Application (client) ID from step 1 |
   | `SyncRootLength` | Length of the synced folder path on a user PC, e.g. `C:\Users\<user>\<Org>\<Site> - <Library>`. In PowerShell: `'C:\Users\...'.Length` |
   | `PcNames` | Names of the team's PCs (used to spot conflict copies), or `[]` |
   | `Thresholds` | Leave as is unless you have a reason |
   | `MaxRecommendedDepthBelowSection` | Folder levels allowed below the section folder (default 3) |
   | `LargeFileSizeGB` | Files above this size are flagged (default 2) |

   Real config files are ignored by git, so client details never get committed.

## Running it

```powershell
.\Invoke-SiteAudit.ps1 -Config .\config\<client>.json -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>"

# only some libraries (title or URL name)
.\Invoke-SiteAudit.ps1 -Config .\config\<client>.json -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -Libraries "Documents","Archive"
```

In phase 1 this signs you in and lists the document libraries it would scan, with item counts and
the settings that affect sync (Require Check Out, content approval, offline sync off).

The script refuses to run if the site address doesn't belong to the tenant in the config file, so a
config can't be used against the wrong client by mistake.

## Tests

The config tests use Pester 5 and never connect to SharePoint:

```powershell
Install-Module Pester -MinimumVersion 5.5 -Scope CurrentUser -SkipPublisherCheck
Invoke-Pester .\tests
```
