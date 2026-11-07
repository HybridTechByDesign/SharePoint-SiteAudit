<#
.SYNOPSIS
    Read-only audit of one SharePoint Online site: finds folder structures that have become too deep.

.DESCRIPTION
    Current (phase 2): loads the tenant config, signs in, lists the document libraries, reads every
    file and folder, and writes AllFindings.csv with path lengths, depth, size, dates and people.
    Later phases add the checks, grouping and the HTML report.

    Results go to .\output\<site>_<yyyyMMdd-HHmm>\. The script returns that folder path.

    This script never creates, moves, renames, deletes or changes permissions on anything in
    SharePoint. It only writes to the local .\output folder.

.NOTES
    Most people should use Start-SiteAudit.cmd (double-click) instead, which asks questions and
    builds the config for you. This script is for running directly or from other scripts.

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
    # Optional if the config file has a DefaultSiteUrl.
    [string] $SiteUrl,

    # Library titles or URL names to scan. Default: all visible document libraries.
    [string[]] $Libraries,

    # Where run folders are created. Default: .\output next to this script.
    [string] $OutputRoot = (Join-Path $PSScriptRoot 'output')
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'lib\Config.ps1')
. (Join-Path $PSScriptRoot 'lib\Connection.ps1')
. (Join-Path $PSScriptRoot 'lib\Scanner.ps1')

$cfg = Import-AuditConfig -Path $Config
if (-not $SiteUrl) { $SiteUrl = $cfg.DefaultSiteUrl }
if (-not $SiteUrl) { throw "No site given. Use -SiteUrl, or run .\Start-SiteAudit.ps1 which remembers the site for you." }
$SiteUrl = Get-SiteRootUrl $SiteUrl
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
    $libs | Format-Table Title, UrlName, ItemCount, RequireCheckOut, ContentApproval, OfflineSyncOff -AutoSize | Out-String -Width 200 | Write-Host

    $total = ($libs | Measure-Object ItemCount -Sum).Sum
    Write-Host ("{0} libraries, {1:N0} items in total." -f $libs.Count, $total)

    # Results folder: output\<site>_<yyyyMMdd-HHmm>
    $siteName = ([uri]$SiteUrl).AbsolutePath.TrimEnd('/') -split '/' | Select-Object -Last 1
    if (-not $siteName) { $siteName = 'root' }
    $siteName = $siteName -replace '[^\w\-]', ''
    $started = Get-Date
    $runFolder = Join-Path $OutputRoot ('{0}_{1:yyyyMMdd-HHmm}' -f $siteName, $started)
    $partsFolder = Join-Path $runFolder 'parts'
    $null = New-Item -ItemType Directory -Path $partsFolder -Force

    $results = foreach ($lib in $libs) {
        Write-Host ""
        Write-Host ("Reading '{0}' ({1:N0} items)..." -f $lib.Title, $lib.ItemCount)
        $part = Join-Path $partsFolder ('{0}.csv' -f ($lib.Id))
        $stats = Invoke-LibraryScan -Library $lib -SyncRootLength $cfg.SyncRootLength -OutFile $part
        Write-Host ("  {0:N0} files, {1:N0} folders. Longest local path: {2} characters." -f $stats.Files, $stats.Folders, $stats.LongestLocalPath)
        [pscustomobject]@{ Library = $lib; Stats = $stats; Part = $part }
    }

    $csvPath = Join-Path $runFolder 'AllFindings.csv'
    Join-CsvFiles -Parts $results.Part -OutFile $csvPath
    Get-ChildItem -LiteralPath $partsFolder -Filter *.csv | ForEach-Object { [System.IO.File]::Delete($_.FullName) }
    [System.IO.Directory]::Delete($partsFolder)

    $longest = $results.Stats | Sort-Object LongestLocalPath -Descending | Select-Object -First 1
    $runInfo = [ordered]@{
        Site            = $SiteUrl
        SiteTitle       = $web.Title
        ScanStarted     = $started.ToString('yyyy-MM-dd HH:mm:ss')
        ScanFinished    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        SyncRootLength  = $cfg.SyncRootLength
        Libraries       = @($results | ForEach-Object {
            [ordered]@{
                Title           = $_.Library.Title
                UrlName         = $_.Library.UrlName
                Items           = $_.Stats.Items
                Files           = $_.Stats.Files
                Folders         = $_.Stats.Folders
                LongestLocalPath = $_.Stats.LongestLocalPath
                RequireCheckOut = $_.Library.RequireCheckOut
                ContentApproval = $_.Library.ContentApproval
                OfflineSyncOff  = $_.Library.OfflineSyncOff
            }
        })
    }
    $runInfo | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $runFolder 'run-info.json') -Encoding utf8

    $files = ($results.Stats | Measure-Object Files -Sum).Sum
    $folders = ($results.Stats | Measure-Object Folders -Sum).Sum
    Write-Host ""
    Write-Host ("Scan finished: {0:N0} files and {1:N0} folders read." -f $files, $folders) -ForegroundColor Green
    if ($longest -and $longest.LongestLocalPath -gt 0) {
        Write-Host ("Longest estimated local path: {0} characters, in '{1}':" -f $longest.LongestLocalPath, $longest.Library)
        Write-Host "  $($longest.LongestPath)" -ForegroundColor DarkGray
    }
    Write-Host "Results: $csvPath"
    Write-Host "The checks and the report for the team come in the next phases."
    $runFolder
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
