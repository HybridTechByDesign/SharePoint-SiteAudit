# Pester 5 tests for turning SharePoint items into CSV rows. Run: Invoke-Pester .\tests
# These tests never connect to SharePoint: items are plain hashtables shaped like SharePoint's.

BeforeAll {
    . (Join-Path $PSScriptRoot '..\lib\Scanner.ps1')
    $root = '/sites/finance/Shared Documents'
    $person = [pscustomobject]@{ LookupValue = 'Person One' }

    function New-FakeItem([string] $Path, [switch] $Folder, $Size = $null) {
        @{
            FileRef = "$root/$Path"; FileLeafRef = ($Path -split '/')[-1]; FSObjType = $(if ($Folder) { '1' } else { '0' })
            File_x0020_Size = $Size; Created = [datetime]'2024-03-01 09:05'; Modified = [datetime]'2024-04-01 10:00'
            Author = $person; Editor = $person
            ItemChildCount = [pscustomobject]@{ LookupValue = '3' }; FolderChildCount = [pscustomobject]@{ LookupValue = '2' }
        }
    }
}

Describe 'ConvertTo-AuditRow' {
    It 'works out the local path length, depth and size of a file' {
        $path = 'Section A/2024 Operation/<Client A>/1. Tax/report.xlsx'
        $row = ConvertTo-AuditRow -Item (New-FakeItem $path -Size '2048') -LibraryTitle 'Documents' -LibraryRoot $root -SyncRootLength 70
        $row.Type | Should -Be 'File'
        $row.PathInLibrary | Should -Be $path
        $row.LocalPathLength | Should -Be (70 + 1 + $path.Length)
        $row.ServerPathLength | Should -Be ("$root/$path").Length
        $row.FolderDepth | Should -Be 4
        $row.SizeBytes | Should -Be 2048
        $row.CreatedBy | Should -Be 'Person One'
        $row.Created | Should -Be '2024-03-01 09:05:00'
    }

    It 'counts direct children and depth for a folder' {
        $row = ConvertTo-AuditRow -Item (New-FakeItem 'Section A/2024 Operation' -Folder) -LibraryTitle 'Documents' -LibraryRoot $root -SyncRootLength 70
        $row.Type | Should -Be 'Folder'
        $row.FolderDepth | Should -Be 2
        $row.DirectChildren | Should -Be 5
        $row.SizeBytes | Should -Be ''
    }

    It 'gives depth 0 for a file at the library root' {
        $row = ConvertTo-AuditRow -Item (New-FakeItem 'readme.docx' -Size '10') -LibraryTitle 'Documents' -LibraryRoot $root -SyncRootLength 70
        $row.FolderDepth | Should -Be 0
        $row.LocalPathLength | Should -Be (70 + 1 + 'readme.docx'.Length)
    }
}

Describe 'Join-CsvFiles' {
    It 'keeps one header row and all data rows' {
        $a = Join-Path $TestDrive 'a.csv'; $b = Join-Path $TestDrive 'b.csv'; $out = Join-Path $TestDrive 'all.csv'
        @([pscustomobject]@{ X = 1 }, [pscustomobject]@{ X = 2 }) | Export-Csv $a -NoTypeInformation -Encoding utf8BOM
        @([pscustomobject]@{ X = 3 }) | Export-Csv $b -NoTypeInformation -Encoding utf8BOM
        Join-CsvFiles -Parts $a, $b, (Join-Path $TestDrive 'missing.csv') -OutFile $out
        (Import-Csv $out).X | Should -Be @('1', '2', '3')
    }
}

Describe 'Test-IsThrottled' {
    It 'recognises throttling errors' {
        $err = [System.Management.Automation.ErrorRecord]::new([Exception]::new('The remote server returned an error: (429) Too Many Requests.'), 'x', 'NotSpecified', $null)
        Test-IsThrottled $err | Should -BeTrue
        $other = [System.Management.Automation.ErrorRecord]::new([Exception]::new('Access denied.'), 'x', 'NotSpecified', $null)
        Test-IsThrottled $other | Should -BeFalse
    }
}
