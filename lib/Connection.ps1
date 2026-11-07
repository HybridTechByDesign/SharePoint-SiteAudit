# Signs in to SharePoint and reads library information.
# Read-only: only Connect-, Get- and Disconnect- PnP cmdlets are used here.

$script:MinPnPVersion = [version]'2.12.0'

function Assert-PnPModule {
    [CmdletBinding()]
    param()

    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw "PowerShell 7 or later is needed. You are running $($PSVersionTable.PSVersion). Install it from https://aka.ms/powershell and run this from 'pwsh'."
    }

    $module = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) {
        throw "The PnP.PowerShell module is not installed. Run once: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    if ($module.Version -lt $script:MinPnPVersion) {
        throw "PnP.PowerShell $($module.Version) is too old (need $script:MinPnPVersion or later). Run: Update-Module PnP.PowerShell"
    }
    Import-Module PnP.PowerShell -MinimumVersion $script:MinPnPVersion -ErrorAction Stop
}

function Connect-AuditSite {
    # Opens a browser sign-in using the tenant's own read-only Entra app.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $SiteUrl,
        [Parameter(Mandatory)] $Config
    )

    Write-Host "Signing in to $SiteUrl (a browser window will open)..."
    Connect-PnPOnline -Url $SiteUrl -ClientId $Config.ClientId -Tenant (Get-SignInTenant $Config) -Interactive -ForceAuthentication -ErrorAction Stop
    try {
        $web = Get-PnPWeb -Includes Title, Url, ServerRelativeUrl -ErrorAction Stop
    }
    catch {
        throw (Get-AccessHelp -Step 'open this site' -ErrorRecord $_)
    }
    Write-Host "Connected to '$($web.Title)'."
    $web
}

$script:UnreadableSettings = @{}

function Format-Setting($Value) {
    if ($null -eq $Value) { return 'unknown' }
    if ($Value) { 'Yes' } else { 'No' }
}

function Get-AccessHelp {
    # Turns an "access denied" from SharePoint into plain steps to check.
    param([string] $Step, $ErrorRecord)
    $message = $ErrorRecord.Exception.Message
    if ($message -notmatch 'denied|unauthori[sz]ed|forbidden|401|403') {
        return "Couldn't $Step. $message"
    }
    @"
Signed in, but SharePoint refused to $Step ("$message").
Check these, in order:
  1. Your account can open this site in the browser. Being a tenant or SharePoint admin is NOT enough:
     you must be a member or owner of the site (or a site collection admin).
  2. The sign-in app has admin consent: Entra admin center > App registrations >
     'SharePoint Site Audit (read-only)' > API permissions. 'SharePoint - AllSites.Read' must show
     'Granted for <org>'. If not, click 'Grant admin consent'.
  3. You signed in with the right account. To force a fresh sign-in, close this window and run it again.
"@
}

function Get-AuditLibraries {
    # Returns the visible document libraries on the site, optionally filtered by name.
    # -Libraries matches the library title ("Documents") or its URL name ("Shared Documents").
    [CmdletBinding()]
    param(
        [string[]] $Libraries
    )

    # Step 1: the basic list of lists. If even this is denied, it's a permissions problem.
    try {
        $lists = Get-PnPList -Includes ItemCount -ErrorAction Stop
    }
    catch {
        throw (Get-AccessHelp -Step 'read the list of libraries on this site' -ErrorRecord $_)
    }

    # BaseTemplate 101 = document library. Hidden libraries are skipped.
    $docLibs = @($lists | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden })

    $result = foreach ($lib in $docLibs) {
        # Step 2: settings that affect sync. Some of these can be refused with read-only access,
        # so each one is read on its own and shown as unknown if refused, instead of stopping the run.
        $settings = @{}
        foreach ($prop in 'IsCatalog', 'IsSystemList', 'ForceCheckout', 'EnableModeration', 'ExcludeFromOfflineClient') {
            try { $settings[$prop] = Get-PnPProperty -ClientObject $lib -Property $prop -ErrorAction Stop }
            catch {
                $settings[$prop] = $null
                Write-Verbose "Couldn't read '$prop' on library '$($lib.Title)': $($_.Exception.Message)"
                $script:UnreadableSettings[$prop] = $true
            }
        }
        if ($settings.IsCatalog -or $settings.IsSystemList) { continue }

        $rootUrl = $lib.RootFolder.ServerRelativeUrl
        [pscustomobject]@{
            Title             = $lib.Title
            UrlName           = ($rootUrl -split '/')[-1]
            ServerRelativeUrl = $rootUrl
            ItemCount         = $lib.ItemCount
            RequireCheckOut   = Format-Setting $settings.ForceCheckout
            ContentApproval   = Format-Setting $settings.EnableModeration
            OfflineSyncOff    = Format-Setting $settings.ExcludeFromOfflineClient
            Id                = $lib.Id
        }
    }
    if ($script:UnreadableSettings.Count) {
        Write-Warning ("Some library settings couldn't be read with read-only access and show as 'unknown': {0}." -f (($script:UnreadableSettings.Keys | Sort-Object) -join ', '))
    }

    if ($Libraries) {
        $missing = $Libraries | Where-Object { $name = $_; -not ($result | Where-Object { $_.Title -eq $name -or $_.UrlName -eq $name }) }
        if ($missing) {
            $available = ($result.Title | Sort-Object) -join ', '
            throw "Library not found: $($missing -join ', '). Available: $available"
        }
        $result = $result | Where-Object { $_.Title -in $Libraries -or $_.UrlName -in $Libraries }
    }

    @($result | Sort-Object Title)
}
