<#
.SYNOPSIS
    The easy way to run the SharePoint site audit. Double-click Start-SiteAudit.cmd to open this.

.DESCRIPTION
    Shows a menu of your clients, sets up new ones by asking questions (no files to edit),
    remembers each client's site, then runs the read-only audit.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'lib\Config.ps1')
. (Join-Path $PSScriptRoot 'lib\Connection.ps1')
. (Join-Path $PSScriptRoot 'setup\Setup.ps1')

$configFolder = Join-Path $PSScriptRoot 'config'

function Get-ClientList {
    Get-ChildItem -LiteralPath $configFolder -Filter *.json |
        Where-Object Name -ne 'example.json' |
        ForEach-Object {
            $name = $_.BaseName; $site = ''
            try {
                $c = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
                if ($c.PSObject.Properties['ClientName'] -and $c.ClientName) { $name = $c.ClientName }
                if ($c.PSObject.Properties['DefaultSiteUrl'] -and $c.DefaultSiteUrl) { $site = $c.DefaultSiteUrl }
            }
            catch { $site = '(file has a problem)' }
            [pscustomobject]@{ Name = $name; Site = $site; Path = $_.FullName }
        } | Sort-Object Name
}

function Select-Client {
    while ($true) {
        $clients = @(Get-ClientList)
        Write-Host ""
        Write-Host "Which client?" -ForegroundColor Cyan
        for ($i = 0; $i -lt $clients.Count; $i++) {
            Write-Host ("  {0}. {1}" -f ($i + 1), $clients[$i].Name) -NoNewline
            Write-Host ("   {0}" -f $clients[$i].Site) -ForegroundColor DarkGray
        }
        Write-Host "  N. Set up a new client"
        Write-Host "  Q. Quit"
        $choice = (Read-Host "Choose").Trim()

        if ($choice -match '^[nN]$') { return Invoke-ClientSetup -ConfigFolder $configFolder }
        if ($choice -match '^[qQ]$') { return $null }
        $n = 0
        if ([int]::TryParse($choice, [ref]$n) -and $n -ge 1 -and $n -le $clients.Count) { return $clients[$n - 1].Path }
        Write-Host "Please type a number from the list, N or Q." -ForegroundColor Yellow
    }
}

try {
    Write-Host ""
    Write-Host "SharePoint site audit (read-only)" -ForegroundColor White
    Write-Host "This only reads from SharePoint. It never moves, renames or deletes anything." -ForegroundColor DarkGray

    Install-PnPIfMissing

    $configPath = Select-Client
    if (-not $configPath) { return }
    $cfg = Import-AuditConfig -Path $configPath

    Write-Host ""
    $site = Read-Answer "Site to scan (press Enter to keep)" $cfg.DefaultSiteUrl
    while (-not (Get-TenantFromSiteUrl $site) -or (Get-TenantFromSiteUrl $site) -ne $cfg.TenantName) {
        Write-Host "  That site isn't in this client's tenant ($($cfg.TenantName).sharepoint.com)." -ForegroundColor Yellow
        $site = Read-Answer "Site to scan" $cfg.DefaultSiteUrl
    }
    $site = Get-SiteRootUrl $site
    if ($site -ne $cfg.DefaultSiteUrl -and (Read-YesNo "Remember this site for next time?")) {
        $cfg.DefaultSiteUrl = $site
        Save-AuditConfig -Config $cfg -Path $configPath
    }

    $libs = Read-Answer "Libraries to scan (press Enter for all, or type names separated by commas)"
    $params = @{ Config = $configPath; SiteUrl = $site }
    if ($libs) { $params.Libraries = @($libs -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

    $runFolder = & (Join-Path $PSScriptRoot 'Invoke-SiteAudit.ps1') @params | Select-Object -Last 1
    if ($runFolder -and (Test-Path -LiteralPath $runFolder) -and $IsWindows) {
        Write-Host ""
        if (Read-YesNo "Open the results folder?") { Invoke-Item -LiteralPath $runFolder }
    }
}
catch {
    Write-Host ""
    Write-Host "Something went wrong:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host ""
    Read-Host "Press Enter to close" | Out-Null
}
