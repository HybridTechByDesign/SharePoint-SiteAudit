<#
.SYNOPSIS
    Read-only audit of one SharePoint Online site: finds folder structures that have become too deep.

.DESCRIPTION
    Phase 1 (current): loads the tenant config, signs in, and lists the document libraries that
    would be scanned. Later phases add the scan, checks, grouping and the HTML report.

    This script never creates, moves, renames, deletes or changes permissions on anything in
    SharePoint. It only writes to the local .\output folder.

.EXAMPLE
    .\Invoke-SiteAudit.ps1 -Config .\config\<client>.json -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>"

.EXAMPLE
    .\Invoke-SiteAudit.ps1 -Config .\config\<client>.json -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -Libraries "Documents","Archive"
#>
[CmdletBinding()]
param(
    # Path to the per-tenant config file, e.g. .\config\<client>.json
    [Parameter(Mandatory)] [string] $Config,

    # Full site address, e.g. https://<tenant>.sharepoint.com/sites/<site>
    [Parameter(Mandatory)] [string] $SiteUrl,

    # Library titles or URL names to scan. Default: all visible document libraries.
    [string[]] $Libraries,

    # Where run folders are created. Default: .\output next to this script.
    [string] $OutputRoot = (Join-Path $PSScriptRoot 'output')
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'lib\Config.ps1')
. (Join-Path $PSScriptRoot 'lib\Connection.ps1')

$cfg = Import-AuditConfig -Path $Config
$SiteUrl = $SiteUrl.TrimEnd('/')
$null = Test-SiteUrlMatchesTenant -SiteUrl $SiteUrl -TenantName $cfg.TenantName

Assert-PnPModule
$web = Connect-AuditSite -SiteUrl $SiteUrl -Config $cfg

try {
    $libs = Get-AuditLibraries -Libraries $Libraries
    if (-not $libs) {
        Write-Warning "No document libraries found on this site."
        return
    }

    Write-Host ""
    Write-Host "Document libraries to scan on '$($web.Title)':"
    $libs | Format-Table Title, UrlName, ItemCount, RequireCheckOut, ContentApproval, OfflineSyncOff -AutoSize | Out-Host

    $total = ($libs | Measure-Object ItemCount -Sum).Sum
    Write-Host ("{0} libraries, {1:N0} items in total." -f $libs.Count, $total)
    Write-Host "Phase 1 complete: sign-in and library listing work. Scanning comes in phase 2."
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
