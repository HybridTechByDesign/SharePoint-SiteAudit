# Pester 5 tests for the checks. Run: Invoke-Pester .\tests
# Local only: no SharePoint connection.

BeforeAll {
    . (Join-Path $PSScriptRoot '..\lib\Checks.ps1')
    . (Join-Path $PSScriptRoot '..\lib\Grouping.ps1')
    $th = [ordered]@{ Warn = 200; Excel = 218; Windows = 260; SharePoint = 400 }
}

Describe 'Get-PathLevel' {
    It 'uses the brief''s thresholds' -ForEach @(
        @{ Local = 143; Server = 100; Expected = 'Green' }
        @{ Local = 199; Server = 100; Expected = 'Green' }
        @{ Local = 200; Server = 100; Expected = 'Amber' }
        @{ Local = 218; Server = 100; Expected = 'Amber' }
        @{ Local = 219; Server = 100; Expected = 'Red' }
        @{ Local = 260; Server = 100; Expected = 'Red' }
        @{ Local = 261; Server = 100; Expected = 'DarkRed' }
        @{ Local = 416; Server = 377; Expected = 'DarkRed' }
        @{ Local = 150; Server = 401; Expected = 'Blocked' }
    ) {
        Get-PathLevel $Local $Server $th | Should -Be $Expected
    }
}

Describe 'Find-RepeatedFolder' {
    It 'finds X\X' { Find-RepeatedFolder 'A/Client docs/Client docs/file.pdf' $false | Should -Be 'Client docs / Client docs' }
    It 'finds near-same names' { Find-RepeatedFolder 'Projects/Projects 2024/Projects/file.pdf' $false | Should -Be 'Projects / Projects 2024' }
    It 'finds Final / Final v2' { Find-RepeatedFolder 'X/Final/Final v2' $true | Should -Be 'Final / Final v2' }
    It 'ignores different names' { Find-RepeatedFolder 'Section A/2024 Operation/<Client A>/1. Tax/file.pdf' $false | Should -Be '' }
    It 'does not count the file name' { Find-RepeatedFolder 'Budget/Budget.xlsx' $false | Should -Be '' }
}

Describe 'Get-ConflictCopyPc' {
    It 'spots name-PC.ext' { Get-ConflictCopyPc 'Budget-PC-01.xlsx' @('PC-01') | Should -Be 'PC-01' }
    It 'spots name-PC-2.ext' { Get-ConflictCopyPc 'Budget-PC-01-2.xlsx' @('PC-01') | Should -Be 'PC-01' }
    It 'ignores other files' { Get-ConflictCopyPc 'Budget.xlsx' @('PC-01') | Should -Be '' }
    It 'does nothing without PC names' { Get-ConflictCopyPc 'Budget-PC-01.xlsx' @() | Should -Be '' }
}

Describe 'Get-BadNameReason' {
    It 'flags trailing space' { Get-BadNameReason 'notes ' | Should -Match 'ends with a space' }
    It 'flags trailing dot' { Get-BadNameReason 'notes.' | Should -Match 'ends with a dot' }
    It 'flags reserved names' { Get-BadNameReason 'CON.txt' | Should -Match 'reserved' }
    It 'flags _vti_' { Get-BadNameReason 'a_vti_b.doc' | Should -Match '_vti_' }
    It 'passes normal names' { Get-BadNameReason 'Console log.txt' | Should -Be '' }
}

Describe 'Get-JunkReason' {
    It 'flags junk' -ForEach @('~$budget.xlsx', 'Thumbs.db', 'desktop.ini', 'x.tmp', 'old.pst') { Get-JunkReason $_ | Should -Not -Be '' }
    It 'passes normal files' { Get-JunkReason 'budget.xlsx' | Should -Be '' }
}

Describe 'Invoke-AuditChecks' {
    BeforeAll {
        $raw = Join-Path $TestDrive 'raw.csv'; $out = Join-Path $TestDrive 'out.csv'
        function New-Row($Type, $Path, $Size = '', $Created = '2025-01-10 09:05:00', $By = 'Person One', $Kids = '') {
            $name = ($Path -split '/')[-1]
            [pscustomobject][ordered]@{
                Library = 'Documents'; Type = $Type; Name = $name; PathInLibrary = $Path; ServerPath = "/sites/x/Shared Documents/$Path"
                LocalPathLength = 71 + $Path.Length; ServerPathLength = 26 + $Path.Length; NameLength = $name.Length
                FolderDepth = ($Path -split '/').Count - $(if ($Type -eq 'File') { 1 } else { 0 }); SizeBytes = $Size; DirectChildren = $Kids
                Created = $Created; CreatedBy = $By; Modified = $Created; ModifiedBy = $By; CheckedOutTo = ''
            }
        }
        $rows = @(
            New-Row 'Folder' 'Copy' -Created '2025-02-02 14:00:00' -By 'Person Three'
            New-Row 'Folder' 'Copy/Sub' -Created '2025-02-02 14:00:00' -By 'Person Three'
            1..55 | ForEach-Object { New-Row 'File' "Copy/Sub/doc$_.pdf" -Size (100 + $_) -Created '2025-02-02 14:01:00' -By 'Person Three' }
            New-Row 'File' 'A/budget.xlsx' -Size 4242
            New-Row 'File' 'B/budget.xlsx' -Size 4242
        )
        $rows | Export-Csv $raw -NoTypeInformation -Encoding utf8BOM
        $config = [pscustomobject]@{ Thresholds = $th; PcNames = @(); LargeFileSizeGB = 2; FixedFolderLevels = 4; ChartFolderLevel = 3 }
        $set = [System.Collections.Generic.HashSet[string]]::new([string[]]@('/sites/x/Shared Documents/A/budget.xlsx'))
        $result = (Invoke-AuditChecks -RawCsv $raw -OutCsv $out -Config $config -UniquePermissionPaths $set).Counts
        $csv = Import-Csv $out
    }

    It 'finds one bulk copy-in at the top folder' {
        $result.BulkCopyFolders | Should -Be 1
        $result.BulkCopyFolderList | Should -Be @('Copy')
        ($csv | Where-Object Name -eq 'doc1.pdf').BulkCopyFolder | Should -Be 'Copy'
    }

    It 'marks likely duplicates' {
        ($csv | Where-Object Name -eq 'budget.xlsx').LikelyDuplicates | Should -Be @('2', '2')
    }

    It 'counts items below a folder' {
        ($csv | Where-Object PathInLibrary -eq 'Copy').SubtreeItems | Should -Be '56'
    }

    It 'marks unique permissions' {
        ($csv | Where-Object PathInLibrary -eq 'A/budget.xlsx').UniquePermissions | Should -Be 'Yes'
    }
}
