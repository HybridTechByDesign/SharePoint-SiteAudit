# Pester 5 tests for building Report.html. Run: Invoke-Pester .\tests

BeforeAll {
    . (Join-Path $PSScriptRoot '..\lib\Report.ps1')
    $summary = [ordered]@{
        SiteTitle = 'Site'
        Headline  = 'Test headline'
        Bars      = @([ordered]@{ Label = 'x</script><b>bold</b>'; Longest = 230 })
        TopAreas  = @()
    }
}

Describe 'Export-AuditReport' {
    It 'puts the scan results into the page' {
        $out = Join-Path $TestDrive 'Report.html'
        Export-AuditReport -Summary $summary -OutFile $out | Should -Be $out
        $html = Get-Content -LiteralPath $out -Raw
        $html | Should -Match 'const DATA = \{'
        $html | Should -Match 'Test headline'
        $html.Contains('/*__SUMMARY_JSON__*/null') | Should -BeFalse
    }

    It 'cannot be broken by a folder name that looks like HTML' {
        $out = Join-Path $TestDrive 'Report.html'
        Export-AuditReport -Summary $summary -OutFile $out | Out-Null
        $html = Get-Content -LiteralPath $out -Raw
        # Only the template's own closing tag; the folder name's one is escaped.
        ([regex]::Matches($html, '</script>')).Count | Should -Be 1
        $html | Should -Match ([regex]::Escape('x<\/script><b>bold<\/b>'))
    }

    It 'refuses a template without the data marker' {
        $bad = Join-Path $TestDrive 'bad.html'
        Set-Content -LiteralPath $bad -Value '<html></html>'
        { Export-AuditReport -Summary $summary -OutFile (Join-Path $TestDrive 'r.html') -TemplatePath $bad } | Should -Throw '*marker*'
    }
}

Describe 'Hide-AreaNames' {
    It 'removes who added the files and keeps the rest' {
        $a = [pscustomobject]@{ BreakPoint = 'A/B'; AddedOn = '2024-05-14'; AddedBy = 'Person One' }
        $r = @(Hide-AreaNames @($a))
        $r[0].AddedBy | Should -Be ''
        $r[0].AddedOn | Should -Be '2024-05-14'
        $a.AddedBy | Should -Be 'Person One'
    }
}
