# Loads and checks a per-tenant config file.
# Read-only: this file only reads a local JSON file. It never touches SharePoint.

function Import-AuditConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Config file not found: $Path"
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Config file '$Path' is not valid JSON. $($_.Exception.Message)"
    }

    # Defaults for anything optional. Values in the file win.
    $config = [ordered]@{
        ClientName                      = $null
        DefaultSiteUrl                  = $null
        TenantName                      = $null
        TenantId                        = $null
        ClientId                        = $null
        SyncRootLength                  = 70
        PcNames                         = @()
        Thresholds                      = [ordered]@{ Warn = 200; Excel = 218; Windows = 260; SharePoint = 400 }
        MaxRecommendedDepthBelowSection = 3
        LargeFileSizeGB                 = 2
        FixedFolderLevels               = 4
        ChartFolderLevel                = 3
    }

    foreach ($prop in $raw.PSObject.Properties) {
        if ($prop.Name -eq 'Thresholds' -and $prop.Value) {
            foreach ($t in $prop.Value.PSObject.Properties) { $config.Thresholds[$t.Name] = $t.Value }
        }
        elseif ($config.Contains($prop.Name)) {
            $config[$prop.Name] = $prop.Value
        }
        else {
            Write-Warning "Config: unknown setting '$($prop.Name)' is ignored."
        }
    }
    $config.PcNames = @($config.PcNames | Where-Object { $_ })

    $problems = Test-AuditConfig -Config $config
    if ($problems) {
        throw "Config file '$Path' has problems:`n - " + ($problems -join "`n - ")
    }

    [pscustomobject]$config
}

function Test-AuditConfig {
    # Returns a list of plain-language problems. Empty list means the config is fine.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Config
    )

    $problems = [System.Collections.Generic.List[string]]::new()

    if ([string]::IsNullOrWhiteSpace($Config.TenantName) -or $Config.TenantName -match '[<>]') {
        $problems.Add("TenantName is missing. Use the part before .sharepoint.com, e.g. 'https://<tenant>.sharepoint.com' -> '<tenant>'.")
    }
    elseif ($Config.TenantName -notmatch '^[a-zA-Z0-9-]+$') {
        $problems.Add("TenantName '$($Config.TenantName)' should be just the tenant prefix (letters, numbers, dashes), not a URL.")
    }

    if ($Config.TenantId -and -not (Test-TenantIdFormat $Config.TenantId)) {
        $problems.Add("TenantId '$($Config.TenantId)' should be the sign-in domain (e.g. <org>.onmicrosoft.com) or the tenant ID (a GUID).")
    }

    $guid = [guid]::Empty
    if (-not [guid]::TryParse([string]$Config.ClientId, [ref]$guid)) {
        $problems.Add("ClientId must be the Application (client) ID of the Entra app, a GUID like 00000000-0000-0000-0000-000000000000.")
    }

    if (-not ($Config.SyncRootLength -is [ValueType]) -or [int]$Config.SyncRootLength -lt 3 -or [int]$Config.SyncRootLength -gt 250) {
        $problems.Add("SyncRootLength must be a number between 3 and 250 (length of the local sync folder path).")
    }

    $t = $Config.Thresholds
    foreach ($name in 'Warn', 'Excel', 'Windows', 'SharePoint') {
        if (-not ($t[$name] -is [ValueType]) -or [int]$t[$name] -le 0) {
            $problems.Add("Thresholds.$name must be a positive number.")
        }
    }
    if ($problems.Count -eq 0 -and -not ($t.Warn -lt $t.Excel -and $t.Excel -lt $t.Windows -and $t.Windows -lt $t.SharePoint)) {
        $problems.Add("Thresholds must go up in order: Warn < Excel < Windows < SharePoint.")
    }

    if (-not ($Config.MaxRecommendedDepthBelowSection -is [ValueType]) -or [int]$Config.MaxRecommendedDepthBelowSection -lt 1) {
        $problems.Add("MaxRecommendedDepthBelowSection must be a number of 1 or more.")
    }

    foreach ($name in 'FixedFolderLevels', 'ChartFolderLevel') {
        if (-not ($Config.$name -is [ValueType]) -or [int]$Config.$name -lt 0 -or [int]$Config.$name -gt 20) {
            $problems.Add("$name must be a number from 0 to 20.")
        }
    }

    if (-not ($Config.LargeFileSizeGB -is [ValueType]) -or [double]$Config.LargeFileSizeGB -le 0) {
        $problems.Add("LargeFileSizeGB must be a positive number.")
    }

    if ($Config.DefaultSiteUrl) {
        try { $null = Test-SiteUrlMatchesTenant -SiteUrl $Config.DefaultSiteUrl -TenantName $Config.TenantName }
        catch { $problems.Add("DefaultSiteUrl: $($_.Exception.Message)") }
    }

    foreach ($pc in $Config.PcNames) {
        if ($pc -match '[<>]') { $problems.Add("PcNames still contains the placeholder '$pc'. Replace it with real PC names, or use [].") }
    }

    , $problems.ToArray()
}

function Test-SiteUrlMatchesTenant {
    # Makes sure the site URL belongs to the tenant in the config, so a config is never used against the wrong client.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SiteUrl,
        [Parameter(Mandatory)] [string] $TenantName
    )

    $uri = $null
    if (-not [uri]::TryCreate($SiteUrl, [urikind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') {
        throw "SiteUrl '$SiteUrl' is not a valid https address."
    }
    $expected = "$TenantName.sharepoint.com"
    if ($uri.Host -ne $expected) {
        throw "SiteUrl host '$($uri.Host)' does not match the config tenant ('$expected'). Check you picked the right config file."
    }
    $true
}

function Get-TenantFromSiteUrl {
    # 'https://<tenant>.sharepoint.com/sites/<site>' -> '<tenant>'. Returns $null if it isn't a SharePoint Online address.
    # Admin center and OneDrive addresses (<tenant>-admin / <tenant>-my) give the same '<tenant>'.
    param([string] $SiteUrl)
    $uri = $null
    if ([uri]::TryCreate($SiteUrl.Trim(), [urikind]::Absolute, [ref]$uri) -and $uri.Host -match '^([a-zA-Z0-9-]+?)(-admin|-my)?\.sharepoint\.com$') {
        return $Matches[1].ToLowerInvariant()
    }
    $null
}

function Test-TenantIdFormat {
    # True for a tenant GUID or a domain like <org>.onmicrosoft.com / <org>.com.
    param([string] $Value)
    $g = [guid]::Empty
    [guid]::TryParse($Value, [ref]$g) -or $Value -match '^[a-zA-Z0-9-]+(\.[a-zA-Z0-9-]+)+$'
}

function Get-SignInTenant {
    # What to sign in to: TenantId from the config if set, otherwise the usual '<tenant>.onmicrosoft.com'.
    # They differ when a tenant was renamed or set up with another domain (common for trial and demo tenants).
    param([Parameter(Mandatory)] $Config)
    if ($Config.TenantId) { return [string]$Config.TenantId }
    "$($Config.TenantName).onmicrosoft.com"
}

function Get-SiteRootUrl {
    # Trims a pasted browser address (e.g. ending in /Shared Documents/Forms/AllItems.aspx) back to the site.
    param([string] $SiteUrl)
    $uri = $null
    if (-not [uri]::TryCreate($SiteUrl.Trim(), [urikind]::Absolute, [ref]$uri)) { return $SiteUrl.Trim() }
    # An admin center or OneDrive address is turned into the tenant's main SharePoint address.
    $hostName = $uri.Host -replace '-(admin|my)\.sharepoint\.com$', '.sharepoint.com'
    if ($uri.AbsolutePath -match '^/(sites|teams)/([^/]+)') {
        return "https://$hostName/$($Matches[1])/$($Matches[2])"
    }
    "$($uri.Scheme)://$hostName"
}

function Save-AuditConfig {
    # Writes a config object to a local JSON file. Local only, never touches SharePoint.
    param(
        [Parameter(Mandatory)] $Config,
        [Parameter(Mandatory)] [string] $Path
    )
    $Config | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding utf8
}
