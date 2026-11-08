# Pester 5 tests for grouping over-limit files into problem areas. Run: Invoke-Pester .\tests

BeforeAll {
    . (Join-Path $PSScriptRoot '..\lib\Grouping.ps1')
    function New-Red([string] $Path, [int] $Length, [string] $Level = 'Red') {
        [pscustomobject]@{ Library = 'Documents'; PathInLibrary = $Path; LocalPathLength = $Length; Level = $Level }
    }
    $fixed = 'Section B/2024 Operation/<Client A>/1. Tax'
}

Describe 'Get-ProblemAreas' {
    BeforeAll {
        $red = @(
            New-Red "$fixed/Copy of desktop/A/B/C/file1.xlsx" 240
            New-Red "$fixed/Copy of desktop/A/B/file2.xlsx" 230
            New-Red "$fixed/Copy of desktop/X/file3.xlsx" 225 'DarkRed'
            New-Red 'Section B/2024 Operation/<Client B>/2. Payroll/From payroll/Y/file4.csv' 250
        )
        $areas = Get-ProblemAreas -RedFiles $red -FixedLevels 4 -ExcelLimit 218
    }

    It 'groups by the first folder below Section/Year/Client/Topic' {
        $areas.Count | Should -Be 2
        $areas[0].BreakPoint | Should -Be "$fixed/Copy of desktop"
        $areas[0].FilesAffected | Should -Be 3
        $areas[1].BreakPoint | Should -Be 'Section B/2024 Operation/<Client B>/2. Payroll/From payroll'
    }

    It 'describes the fix and what it saves' {
        $a = $areas[0]
        $a.Fix | Should -Be "Move the contents of 'Copy of desktop' up into '1. Tax'"
        $a.CharactersSaved | Should -Be ('Copy of desktop'.Length + 1)
        # 240-16=224 still over, 230-16=214 fixed, 225-16=209 fixed
        $a.FixedByMove | Should -Be 2
        $a.StillTooLong | Should -Be 1
        $a.DeepestLocalPath | Should -Be 240
        $a.ExtraLevels | Should -Be 3
        $a.OverWindowsLimit | Should -Be 1
    }

    It 'offers flattening when one move is not enough' {
        $areas[0].FlattenFix | Should -Match 'Bring the files up out of the 3 folder level'
        # Inner folders here are short (A/B/C), so flattening alone doesn't bring these under 218.
        $areas[0].FixedByFlatten | Should -Be 0
    }

    It 'uses the file''s own folder when it sits above the fixed levels' {
        $a = Get-ProblemAreas -RedFiles @(New-Red 'AuditTest/416/long/f.xlsx' 416) -FixedLevels 4
        $a[0].BreakPoint | Should -Be 'AuditTest/416/long'
    }

    It 'works with one folder level above the break point, under strict mode like the real run' {
        # The audit runs with Set-StrictMode -Version 3, where a one-item list that PowerShell
        # unwrapped into a plain string has no .Count. FixedLevels 1 or a file in a top-level folder hit this.
        Set-StrictMode -Version 3
        try {
            $a = Get-ProblemAreas -RedFiles @(New-Red 'AuditTest/Copied/A/f.xlsx' 300) -FixedLevels 1
            $a[0].BreakPoint | Should -Be 'AuditTest/Copied'
            $a[0].Fix | Should -Be "Move the contents of 'Copied' up into 'AuditTest'"
            $b = Get-ProblemAreas -RedFiles @(New-Red 'Top/f.xlsx' 300) -FixedLevels 4
            $b[0].Fix | Should -Be "Move the contents of 'Top' up into the top of the 'Documents' library"
        }
        finally { Set-StrictMode -Off }
    }

    It 'handles no over-limit files' {
        @(Get-ProblemAreas -RedFiles @() -FixedLevels 4).Count | Should -Be 0
    }
}

Describe 'Get-Headline' {
    It 'reads naturally' {
        Get-Headline -OverExcel 300 -Places 6 -OverWindows 214 | Should -Be "300 files can't be reliably opened or synced. They're in 6 places. 214 of them are too long even for File Explorer and Outlook."
        Get-Headline -OverExcel 1 -Places 1 -OverWindows 0 | Should -Be "1 file can't be reliably opened or synced. They're all in 1 place."
        Get-Headline -OverExcel 0 -Places 0 -OverWindows 0 | Should -Match '^No files'
    }
}

Describe 'Chart totals' {
    It 'builds one bar per client folder with subfolders' {
        $bars = @{}; $children = @{}
        Measure-BarStat $bars $children 'Documents' "$fixed/Copy/f1.xlsx" 240 $true 3
        Measure-BarStat $bars $children 'Documents' 'Section B/2024 Operation/<Client A>/2. Payroll/f2.xlsx' 150 $false 3
        $result = Get-FolderBars -BarStats $bars -ChildStats $children -ShowLibrary $false
        $result.Count | Should -Be 1
        $result[0].Label | Should -Be 'Section B/2024 Operation/<Client A>'
        $result[0].Longest | Should -Be 240
        $result[0].OverExcel | Should -Be 1
        $result[0].Children.Count | Should -Be 2
        $result[0].Children[0].Name | Should -Be '1. Tax'
    }
}
