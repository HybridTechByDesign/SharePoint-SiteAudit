<#
.SYNOPSIS
    Read-only audit of one SharePoint Online site: finds folder structures that have become too deep.

.DESCRIPTION
    Current (phase 3): loads the tenant config, signs in, lists the document libraries, reads every
    file and folder, runs the checks, and writes AllFindings.csv (one row per item, with a column per
    check) and run-info.json (totals and library findings). Later phases add the grouping and the
    HTML report.

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
. (Join-Path $PSScriptRoot 'lib\Checks.ps1')

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
        Write-Host "  Checking library settings and permissions..."
        $required = Get-RequiredColumns -Library $lib
        $unique = Get-UniquePermissionPaths -Library $lib
        [pscustomobject]@{ Library = $lib; Stats = $stats; Part = $part; Required = $required; Unique = $unique }
    }

    Write-Host ""
    Write-Host "Running the checks..."
    $rawCsv = Join-Path $partsFolder 'raw.csv'
    Join-CsvFiles -Parts $results.Part -OutFile $rawCsv
    $uniqueSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $results) { foreach ($p in @($r.Unique)) { if ($p) { $null = $uniqueSet.Add($p) } } }
    $csvPath = Join-Path $runFolder 'AllFindings.csv'
    $findings = Invoke-AuditChecks -RawCsv $rawCsv -OutCsv $csvPath -Config $cfg -UniquePermissionPaths $uniqueSet

    # Checked-out files per library, for the library findings.
    $checkedOut = @{}
    Import-Csv -LiteralPath $csvPath | Where-Object CheckedOutTo | ForEach-Object { $checkedOut[$_.Library] = 1 + [int]$checkedOut[$_.Library] }

    Get-ChildItem -LiteralPath $partsFolder -Filter *.csv | ForEach-Object { [System.IO.File]::Delete($_.FullName) }
    [System.IO.Directory]::Delete($partsFolder)

    $longest = $results.Stats | Sort-Object LongestLocalPath -Descending | Select-Object -First 1
    $runInfo = [ordered]@{
        Site            = $SiteUrl
        SiteTitle       = $web.Title
        ScanStarted     = $started.ToString('yyyy-MM-dd HH:mm:ss')
        ScanFinished    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        SyncRootLength  = $cfg.SyncRootLength
        Thresholds      = $cfg.Thresholds
        Findings        = $findings
        Libraries       = @($results | ForEach-Object {
            $uniqueCount = if ($null -eq $_.Unique) { 'unknown' } else { @($_.Unique).Count }
            $lf = Get-LibraryFindings -Library $_.Library -Stats $_.Stats -RequiredColumns $_.Required -UniqueCount $uniqueCount -CheckedOut ([int]$checkedOut[$_.Library.Title])
            [ordered]@{
                Title            = $_.Library.Title
                UrlName          = $_.Library.UrlName
                Files            = $_.Stats.Files
                Folders          = $_.Stats.Folders
                LongestLocalPath = $_.Stats.LongestLocalPath
            } + $lf
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
    Write-Host ""
    Write-Host "What the checks found:" -ForegroundColor Cyan
    $th = $cfg.Thresholds
    $lines = @(
        ,@($findings.OverExcel, "files are over the Excel limit ($($th.Excel) characters)")
        ,@(($findings.DarkRed + $findings.Blocked), "of those are also over the Windows limit ($($th.Windows))")
        ,@($findings.Blocked, "files have a SharePoint path over $($th.SharePoint) and can't sync")
        ,@($findings.Amber, "files are close to the Excel limit")
        ,@($findings.RepeatedFolder, "items sit under a repeated folder name (e.g. X\X)")
        ,@($findings.BulkCopyFolders, "folders look like whole folder trees copied in at once")
        ,@($findings.ConflictCopy, "conflict copies (name ending in a PC name)")
        ,@($findings.LikelyDuplicate, "files look like duplicates (same name and size elsewhere)")
        ,@($findings.BadName, "items have names Windows or OneDrive dislike")
        ,@($findings.Junk, "junk files (temp files, Thumbs.db, desktop.ini, .pst)")
        ,@($findings.LargeFile, "files over $($cfg.LargeFileSizeGB) GB")
        ,@($findings.NameTooLong, "names over 255 characters")
        ,@($findings.TooManyChildren, "folders with over 5,000 items directly inside")
        ,@($findings.CheckedOut, "files are checked out")
        ,@($findings.UniquePermissions, "items have their own permissions")
    )
    foreach ($l in $lines) { if ($l[0] -gt 0) { Write-Host ("  {0,7:N0}  {1}" -f $l[0], $l[1]) } }
    if (-not ($lines | Where-Object { $_[0] -gt 0 })) { Write-Host "  Nothing to report." -ForegroundColor Green }
    Write-Host ("  Deepest folder nesting: {0} levels." -f $findings.DeepestFolderDepth)
    foreach ($lib in $runInfo.Libraries) {
        foreach ($n in $lib.Notes) { Write-Host "  Library '$($lib.Title)': $n" -ForegroundColor Yellow }
    }
    Write-Host ""
    Write-Host "Results: $csvPath"
    Write-Host "The report for the team comes in the next phases."
    $runFolder
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
