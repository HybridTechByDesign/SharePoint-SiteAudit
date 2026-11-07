# Pester 5 tests for config loading. Run: Invoke-Pester .\tests
# These tests never connect to SharePoint.

BeforeAll {
    . (Join-Path $PSScriptRoot '..\lib\Config.ps1')

    function New-TestConfig([hashtable] $Values) {
        $base = @{
            TenantName = 'exampletenant'
            ClientId   = '11111111-2222-3333-4444-555555555555'
        }
        foreach ($k in $Values.Keys) { $base[$k] = $Values[$k] }
        $path = Join-Path $TestDrive ("cfg-{0}.json" -f [guid]::NewGuid())
        $base | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path
        $path
    }
}

Describe 'Import-AuditConfig' {
    It 'loads a minimal config and fills in defaults' {
        $cfg = Import-AuditConfig -Path (New-TestConfig @{})
        $cfg.SyncRootLength | Should -Be 70
        $cfg.Thresholds.Excel | Should -Be 218
        $cfg.Thresholds.SharePoint | Should -Be 400
        $cfg.MaxRecommendedDepthBelowSection | Should -Be 3
        $cfg.LargeFileSizeGB | Should -Be 2
        @($cfg.PcNames).Count | Should -Be 0
    }

    It 'lets the file override one threshold and keeps the others' {
        $cfg = Import-AuditConfig -Path (New-TestConfig @{ Thresholds = @{ Excel = 210 } })
        $cfg.Thresholds.Excel | Should -Be 210
        $cfg.Thresholds.Windows | Should -Be 260
    }

    It 'rejects the example file because it still has placeholders' {
        { Import-AuditConfig -Path (Join-Path $PSScriptRoot '..\config\example.json') } | Should -Throw '*TenantName is missing*'
    }

    It 'rejects a ClientId that is not a GUID' {
        { Import-AuditConfig -Path (New-TestConfig @{ ClientId = 'abc' }) } | Should -Throw '*ClientId*'
    }

    It 'rejects thresholds out of order' {
        { Import-AuditConfig -Path (New-TestConfig @{ Thresholds = @{ Excel = 300 } }) } | Should -Throw '*go up in order*'
    }

    It 'rejects a missing file' {
        { Import-AuditConfig -Path (Join-Path $TestDrive 'nope.json') } | Should -Throw '*not found*'
    }

    It 'rejects broken JSON' {
        $p = Join-Path $TestDrive 'bad.json'
        Set-Content -LiteralPath $p -Value '{ "TenantName": '
        { Import-AuditConfig -Path $p } | Should -Throw '*not valid JSON*'
    }
}

Describe 'Test-SiteUrlMatchesTenant' {
    It 'accepts a site in the same tenant' {
        Test-SiteUrlMatchesTenant -SiteUrl 'https://exampletenant.sharepoint.com/sites/finance' -TenantName 'exampletenant' | Should -BeTrue
    }

    It 'refuses a site in another tenant' {
        { Test-SiteUrlMatchesTenant -SiteUrl 'https://othertenant.sharepoint.com/sites/finance' -TenantName 'exampletenant' } | Should -Throw '*does not match*'
    }

    It 'refuses a non-https address' {
        { Test-SiteUrlMatchesTenant -SiteUrl 'http://exampletenant.sharepoint.com/sites/finance' -TenantName 'exampletenant' } | Should -Throw '*https*'
    }
}

Describe 'Site address helpers' {
    It 'gets the tenant from a pasted browser address' {
        Get-TenantFromSiteUrl 'https://exampletenant.sharepoint.com/sites/finance/Shared%20Documents/Forms/AllItems.aspx' | Should -Be 'exampletenant'
    }

    It 'returns nothing for a non-SharePoint address' {
        Get-TenantFromSiteUrl 'https://example.com/sites/finance' | Should -BeNullOrEmpty
        Get-TenantFromSiteUrl 'not a url' | Should -BeNullOrEmpty
    }

    It 'trims a pasted address back to the site' {
        Get-SiteRootUrl 'https://exampletenant.sharepoint.com/sites/finance/Shared%20Documents/Forms/AllItems.aspx' | Should -Be 'https://exampletenant.sharepoint.com/sites/finance'
        Get-SiteRootUrl 'https://exampletenant.sharepoint.com/teams/hr/' | Should -Be 'https://exampletenant.sharepoint.com/teams/hr'
    }
}

Describe 'Save-AuditConfig' {
    It 'writes a file that loads back the same' {
        $cfg = Import-AuditConfig -Path (New-TestConfig @{ ClientName = '<Client A>'; DefaultSiteUrl = 'https://exampletenant.sharepoint.com/sites/finance'; PcNames = @('PC-01') })
        $path = Join-Path $TestDrive 'saved.json'
        Save-AuditConfig -Config $cfg -Path $path
        $again = Import-AuditConfig -Path $path
        $again.DefaultSiteUrl | Should -Be 'https://exampletenant.sharepoint.com/sites/finance'
        @($again.PcNames) | Should -Be @('PC-01')
        $again.Thresholds.Excel | Should -Be 218
    }

    It 'rejects a remembered site from another tenant' {
        { Import-AuditConfig -Path (New-TestConfig @{ DefaultSiteUrl = 'https://othertenant.sharepoint.com/sites/x' }) } | Should -Throw '*DefaultSiteUrl*'
    }
}
