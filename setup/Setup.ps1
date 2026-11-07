# Guided setup: asks plain questions and writes the client config file for you.
#
# This is the only part of the tool that can change anything in a tenant, and only if you say yes:
# it can register the read-only "SharePoint Site Audit" app in Entra ID (one time per tenant).
# The scanner itself (Invoke-SiteAudit.ps1 and lib\) is read-only.

$script:AppName = 'SharePoint Site Audit (read-only)'

function Write-Step([string] $Text) { Write-Host ""; Write-Host $Text -ForegroundColor Cyan }
function Write-Hint([string] $Text) { Write-Host "  $Text" -ForegroundColor DarkGray }

function Read-Answer {
    # Asks a question. Pressing Enter returns the default (if there is one).
    param([string] $Prompt, [string] $Default)
    $suffix = if ($Default) { " [$Default]" } else { '' }
    $answer = Read-Host "$Prompt$suffix"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    $answer.Trim().Trim('"')
}

function Read-YesNo {
    param([string] $Prompt, [bool] $Default = $true)
    $hint = if ($Default) { 'Y/n' } else { 'y/N' }
    while ($true) {
        $a = Read-Host "$Prompt ($hint)"
        if ([string]::IsNullOrWhiteSpace($a)) { return $Default }
        switch -Regex ($a.Trim()) {
            '^(y|yes)$' { return $true }
            '^(n|no)$' { return $false }
        }
        Write-Host "  Please type Y or N." -ForegroundColor Yellow
    }
}

function Install-PnPIfMissing {
    # Checks PowerShell 7 and the PnP module, and offers to install PnP for the current user.
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw "This needs PowerShell 7. Install it from https://aka.ms/powershell, then double-click Start-SiteAudit.cmd again."
    }
    $module = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
    if ($module -and $module.Version -ge $script:MinPnPVersion) { return }

    if ($module) {
        Write-Host "The PnP.PowerShell module on this PC is too old ($($module.Version))." -ForegroundColor Yellow
    }
    else {
        Write-Host "The PnP.PowerShell module is not installed on this PC yet." -ForegroundColor Yellow
    }
    if (-not (Read-YesNo "Install it now for your user account? (takes a minute, no admin needed)")) {
        throw "PnP.PowerShell is needed. Install it later with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    Write-Host "Installed." -ForegroundColor Green
}

function Get-SafeFileName([string] $Name) {
    $clean = ($Name -replace '[^\w\- ]', '').Trim() -replace '\s+', '-'
    if (-not $clean) { $clean = 'client' }
    $clean.ToLowerInvariant()
}

function Register-AuditApp {
    # Registers the read-only app in the tenant and returns its client ID, or $null if it didn't work.
    param([Parameter(Mandatory)] [string] $TenantName)

    Write-Hint "A browser window will open. Sign in as a Global Admin (or Application Admin) of this tenant"
    Write-Hint "and click Accept. The app only gets permission to READ SharePoint sites."
    try {
        $output = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $script:AppName `
            -Tenant "$TenantName.onmicrosoft.com" -SharePointDelegatePermissions 'AllSites.Read' -ErrorAction Stop
        $record = $output | Where-Object { $_ -isnot [string] -and $_.PSObject.Properties['AzureAppId/ClientId'] } | Select-Object -First 1
        if ($record) { return [string]$record.'AzureAppId/ClientId' }
        # Fall back to the text line "... -ClientId <guid>"
        $line = $output | Where-Object { $_ -is [string] -and $_ -match '-ClientId\s+([0-9a-fA-F-]{36})' } | Select-Object -First 1
        if ($line -and $line -match '-ClientId\s+([0-9a-fA-F-]{36})') { return $Matches[1] }
    }
    catch {
        Write-Host "  App registration did not finish: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    $null
}

function Read-ClientId {
    while ($true) {
        $id = Read-Answer "  Paste the Application (client) ID"
        $g = [guid]::Empty
        if ([guid]::TryParse([string]$id, [ref]$g)) { return $g.ToString() }
        Write-Host "  That doesn't look like an ID. It looks like 00000000-0000-0000-0000-000000000000." -ForegroundColor Yellow
    }
}

function Invoke-ClientSetup {
    # Walks through the questions and saves config\<client>.json. Returns the file path.
    param([Parameter(Mandatory)] [string] $ConfigFolder)

    Write-Host ""
    Write-Host "=== Set up a new client ===" -ForegroundColor Cyan
    Write-Hint "Answers are saved in the 'config' folder on this PC only."

    # 1. Name
    Write-Step "1 of 5  Client name"
    Write-Hint "Just a label so you can pick it from the menu next time."
    do { $clientName = Read-Answer "  Name" } while (-not $clientName)
    $path = Join-Path $ConfigFolder ("{0}.json" -f (Get-SafeFileName $clientName))
    if ((Test-Path -LiteralPath $path) -and -not (Read-YesNo "  A client with this name already exists. Replace it?" $false)) {
        return $path
    }

    # 2. Site
    Write-Step "2 of 5  SharePoint site"
    Write-Hint "Open the site in your browser and copy the address bar. Any page of the site is fine."
    while ($true) {
        $pasted = Read-Answer "  Site address"
        $tenant = if ($pasted) { Get-TenantFromSiteUrl $pasted } else { $null }
        if ($tenant) { break }
        Write-Host "  That doesn't look like a SharePoint Online address (https://<tenant>.sharepoint.com/sites/<site>)." -ForegroundColor Yellow
    }
    $siteUrl = Get-SiteRootUrl $pasted
    Write-Hint "Site: $siteUrl   Tenant: $tenant"

    # 3. App registration
    Write-Step "3 of 5  Sign-in app for this tenant"
    Write-Hint "Each tenant needs a small read-only app in Entra ID so the tool can sign in. This is done once per tenant."
    $clientId = $null
    if (Read-YesNo "  Has the '$script:AppName' app already been set up in this tenant?" $false) {
        Write-Hint "Find it in the Entra admin center > App registrations > '$script:AppName' > Application (client) ID."
        $clientId = Read-ClientId
    }
    elseif (Read-YesNo "  Set it up now? (needs a tenant admin account)") {
        $clientId = Register-AuditApp -TenantName $tenant
        if ($clientId) {
            Write-Host "  App created. ID: $clientId" -ForegroundColor Green
        }
        else {
            Write-Hint "If the app was created anyway, copy its ID from the Entra admin center."
            $clientId = Read-ClientId
        }
    }
    else {
        Write-Hint "See 'One-time app setup' in README.md, then paste the ID here."
        $clientId = Read-ClientId
    }

    # 4. Sync folder length
    Write-Step "4 of 5  Synced folder on a user's PC (optional)"
    Write-Hint "On a team member's PC, open the synced library in File Explorer, click the address bar and copy it."
    Write-Hint "It looks like C:\Users\<user>\<Org>\<Site> - <Library>. This makes path lengths more accurate."
    $syncRoot = Read-Answer "  Paste it, or press Enter to use a typical estimate"
    $syncLength = if ($syncRoot) { $syncRoot.TrimEnd('\').Length } else { 70 }
    Write-Hint "Using $syncLength characters for the local folder."

    # 5. PC names
    Write-Step "5 of 5  Team PC names (optional)"
    Write-Hint "Used to spot conflict copies like 'Budget-PC01.xlsx'. Separate with commas."
    $pcs = Read-Answer "  PC names, or press Enter to skip"
    $pcNames = @(if ($pcs) { $pcs -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } })

    $config = [ordered]@{
        ClientName                      = $clientName
        DefaultSiteUrl                  = $siteUrl
        TenantName                      = $tenant
        ClientId                        = $clientId
        SyncRootLength                  = $syncLength
        PcNames                         = $pcNames
        Thresholds                      = [ordered]@{ Warn = 200; Excel = 218; Windows = 260; SharePoint = 400 }
        MaxRecommendedDepthBelowSection = 3
        LargeFileSizeGB                 = 2
    }
    $problems = Test-AuditConfig -Config $config
    if ($problems) { throw "Setup produced an invalid config:`n - " + ($problems -join "`n - ") }

    Save-AuditConfig -Config $config -Path $path
    Write-Host ""
    Write-Host "Saved. Next time just pick '$clientName' from the menu." -ForegroundColor Green
    $path
}
