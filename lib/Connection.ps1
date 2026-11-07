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
    Connect-PnPOnline -Url $SiteUrl -ClientId $Config.ClientId -Tenant "$($Config.TenantName).onmicrosoft.com" -Interactive -ErrorAction Stop
    $web = Get-PnPWeb -Includes Title, Url, ServerRelativeUrl -ErrorAction Stop
    Write-Host "Connected to '$($web.Title)'."
    $web
}

function Get-AuditLibraries {
    # Returns the visible document libraries on the site, optionally filtered by name.
    # -Libraries matches the library title ("Documents") or its URL name ("Shared Documents").
    [CmdletBinding()]
    param(
        [string[]] $Libraries
    )

    $lists = Get-PnPList -Includes RootFolder, ItemCount, Hidden, BaseTemplate, IsCatalog, IsSystemList, ForceCheckout, EnableModeration, ExcludeFromOfflineClient -ErrorAction Stop
    # BaseTemplate 101 = document library. Hidden and system libraries are skipped.
    $docLibs = $lists | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden -and -not $_.IsCatalog -and -not $_.IsSystemList }

    $result = foreach ($lib in $docLibs) {
        $rootUrl = $lib.RootFolder.ServerRelativeUrl
        [pscustomobject]@{
            Title             = $lib.Title
            UrlName           = ($rootUrl -split '/')[-1]
            ServerRelativeUrl = $rootUrl
            ItemCount         = $lib.ItemCount
            RequireCheckOut   = [bool]$lib.ForceCheckout
            ContentApproval   = [bool]$lib.EnableModeration
            OfflineSyncOff    = [bool]$lib.ExcludeFromOfflineClient
            Id                = $lib.Id
        }
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
