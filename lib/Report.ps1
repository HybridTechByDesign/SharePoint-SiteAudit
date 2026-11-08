# Builds Report.html: the one page for the team. Local only: nothing here talks to SharePoint.
#
# The page itself lives in report\Report.template.html (plain HTML, CSS and JavaScript). This
# just puts the scan results into it, so the finished file works offline and can be emailed.

function Export-AuditReport {
    param(
        # The same object that is written to summary.json.
        [Parameter(Mandatory)] $Summary,
        [Parameter(Mandatory)] [string] $OutFile,
        [string] $TemplatePath = (Join-Path $PSScriptRoot '..\report\Report.template.html')
    )
    $placeholder = '/*__SUMMARY_JSON__*/null'
    $template = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $TemplatePath).Path)
    if (-not $template.Contains($placeholder)) { throw "The report template is missing the $placeholder marker: $TemplatePath" }

    # "</" would end the <script> block early if a folder name contained "</script>". "<\/" means
    # the same in JavaScript. U+2028/2029 are valid in JSON but break older JavaScript engines.
    $json = $Summary | ConvertTo-Json -Depth 8 -Compress
    $json = $json.Replace('</', '<\/').Replace([string][char]0x2028, ' ').Replace([string][char]0x2029, ' ')

    $html = $template.Replace($placeholder, $json)
    [System.IO.File]::WriteAllText($OutFile, $html, [System.Text.UTF8Encoding]::new($false))
    $OutFile
}

function Hide-AreaNames {
    # Removes who added the files from each problem area, unless names were asked for.
    param([AllowEmptyCollection()] $Areas)
    foreach ($a in $Areas) {
        $copy = $a | Select-Object -Property *
        $copy.AddedBy = ''
        $copy
    }
}
