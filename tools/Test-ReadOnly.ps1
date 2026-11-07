<#
.SYNOPSIS
    Checks the scanner code for any cmdlet that could change SharePoint.

.DESCRIPTION
    Searches Invoke-SiteAudit.ps1 and lib\ for write verbs (Set-, Remove-, Move-, Add-, New-,
    Rename-, Copy-, Clear-, Grant-, Revoke-, Publish-, Submit-, Restore-, Update-, Invoke-PnPSPRestMethod
    with a non-GET method). Exits with code 1 if anything is found.
    Local-only commands that are known to be safe are allowed.
#>
[CmdletBinding()]
param(
    [string] $Root = (Split-Path $PSScriptRoot -Parent)
)

$files = @(Get-Item (Join-Path $Root 'Invoke-SiteAudit.ps1')) + @(Get-ChildItem (Join-Path $Root 'lib') -Filter *.ps1 -Recurse)

# Write verbs. Matched as whole cmdlet names, e.g. "Set-PnPListItem".
$pattern = '(Set|Remove|Move|Add|New|Rename|Copy|Clear|Grant|Revoke|Publish|Submit|Restore|Update|Unpublish|Approve|Deny)-[A-Za-z]+'

# Local-only cmdlets that never touch SharePoint.
$allowed = @(
    'Set-StrictMode', 'Set-Content', 'Add-Content', 'New-Item', 'New-Object', 'Add-Member',
    'Set-Location', 'Copy-Item', 'Set-Variable', 'New-Variable', 'Remove-Variable', 'Clear-Host'
)

# Use the PowerShell parser so only real commands are checked (not text in messages or comments).
$hits = foreach ($f in $files) {
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errors)
    $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    foreach ($cmd in $commands) {
        $name = $cmd.GetCommandName()
        if (-not $name) { continue }
        $line = $cmd.Extent.StartLineNumber
        if ($name -match "^$pattern$" -and $name -notin $allowed) {
            [pscustomobject]@{ File = $f.Name; Line = $line; Command = $name }
        }
        if ($name -eq 'Invoke-PnPSPRestMethod' -and $cmd.Extent.Text -match '-Method\s+(Post|Put|Patch|Delete|Merge)') {
            [pscustomobject]@{ File = $f.Name; Line = $line; Command = 'Invoke-PnPSPRestMethod (write method)' }
        }
    }
}

if ($hits) {
    Write-Host "Possible write operations found:" -ForegroundColor Red
    foreach ($h in $hits) { Write-Host ("  {0} line {1}: {2}" -f $h.File, $h.Line, $h.Command) }
    exit 1
}
Write-Host "OK: no write operations found in $($files.Count) files." -ForegroundColor Green
exit 0
