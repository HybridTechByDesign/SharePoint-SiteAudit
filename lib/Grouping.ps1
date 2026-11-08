# Groups the files that are over the Excel limit into a few "problem areas", each with one fix.
# Local only: nothing here talks to SharePoint.
#
# The idea: the team shouldn't get a list of every long path. Long paths almost always come from
# one folder tree that was copied in. So for each over-limit file we find its break point: the
# highest folder below the fixed structure (Section/Year/Client/Topic = 4 levels by default), which
# is where the copied-in tree starts. Files are grouped by break point. If a file sits higher than
# that, its own folder is used.
# The fix is "move the contents of <break point> up into <parent>", which shortens every path
# under it by the length of that folder's name plus one.

function Get-Headline {
    # The one sentence at the top of the report, in plain language.
    param([int] $OverExcel, [int] $Places, [int] $OverWindows)
    if ($OverExcel -eq 0) { return "No files are too deep to open. Nothing needs moving right now." }
    $s = if ($OverExcel -eq 1) { "1 file can't be reliably opened or synced." } else { "$('{0:N0}' -f $OverExcel) files can't be reliably opened or synced." }
    $s += if ($Places -eq 1) { " They're all in 1 place." } else { " They're in $Places places." }
    if ($OverWindows -gt 0) {
        $s += if ($OverWindows -eq $OverExcel) { " All of them are too long even for File Explorer and Outlook." }
        else { " $('{0:N0}' -f $OverWindows) of them are too long even for File Explorer and Outlook." }
    }
    $s
}

function ConvertTo-PublicArea {
    # A problem area without the internal member list, for summary.json and the CSV.
    param($Area)
    $Area | Select-Object -Property * -ExcludeProperty Members
}

function Get-FolderSegments {
    # Folders of a file's path inside the library: "A/B/c.xlsx" -> A, B
    param([string] $PathInLibrary)
    $s = @($PathInLibrary -split '/' | Where-Object { $_ })
    if ($s.Count -le 1) { return , @() }
    , @($s[0..($s.Count - 2)])
}

function Get-ProblemAreas {
    # RedFiles: objects with Library, PathInLibrary, LocalPathLength, Level.
    # Returns one object per problem area, biggest first.
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] $RedFiles,
        [int] $FixedLevels = 4,
        [int] $ExcelLimit = 218
    )

    # 1. Group by the folder just below the fixed structure (or the file's own folder if it's shallower).
    $groups = [ordered]@{}
    foreach ($f in $RedFiles) {
        $folders = Get-FolderSegments $f.PathInLibrary
        $depth = [Math]::Min($folders.Count, $FixedLevels + 1)
        $top = if ($depth -gt 0) { $folders[0..($depth - 1)] -join '/' } else { '' }
        $key = "$($f.Library)|$top"
        if (-not $groups.Contains($key)) { $groups[$key] = [System.Collections.Generic.List[object]]::new() }
        $groups[$key].Add([pscustomobject]@{ File = $f; Folders = $folders; MinDepth = $depth })
    }

    # 2. Describe each group and its fix.
    $areas = foreach ($key in $groups.Keys) {
        $members = $groups[$key]
        $library = $key.Substring(0, $key.IndexOf('|'))
        $depth = $members[0].MinDepth
        $segments = @(if ($depth -gt 0) { $members[0].Folders[0..($depth - 1)] })
        $breakPoint = $segments -join '/'
        $name = if ($segments.Count) { $segments[-1] } else { '' }
        $parentSegments = @(if ($segments.Count -gt 1) { $segments[0..($segments.Count - 2)] })
        $parentName = if ($parentSegments.Count) { "'$($parentSegments[-1])'" } else { "the top of the '$library' library" }

        $files = @($members | ForEach-Object { $_.File })
        $saved = if ($name) { $name.Length + 1 } else { 0 }
        $deepest = $files | Sort-Object { [int]$_.LocalPathLength } -Descending | Select-Object -First 1
        $maxDepth = ($members | ForEach-Object { $_.Folders.Count } | Measure-Object -Maximum).Maximum
        $fixedAfter = @($files | Where-Object { ([int]$_.LocalPathLength - $saved) -le $ExcelLimit }).Count

        # Second option, for when moving one level isn't enough: bring the files up out of the
        # extra folders inside the break point, so they sit directly in it.
        $fixedFlat = 0
        foreach ($m in $members) {
            $inner = if ($m.Folders.Count -gt $depth) { ($m.Folders[$depth..($m.Folders.Count - 1)] -join '/').Length + 1 } else { 0 }
            if (([int]$m.File.LocalPathLength - $inner) -le $ExcelLimit) { $fixedFlat++ }
        }

        # When the files arrived, and who added most of them (names are only shown with -ShowNames).
        $whenGroup = $files | Where-Object { $_.PSObject.Properties['CreatedOn'] -and $_.CreatedOn } | Group-Object CreatedOn | Sort-Object Count -Descending | Select-Object -First 1
        $whoGroup = $files | Where-Object { $_.PSObject.Properties['CreatedBy'] -and $_.CreatedBy } | Group-Object CreatedBy | Sort-Object Count -Descending | Select-Object -First 1

        $fix = if ($name) {
            "Move the contents of '$name' up into $parentName"
        }
        else {
            "Shorten these file names"
        }

        [pscustomobject][ordered]@{
            Library            = $library
            BreakPoint         = $breakPoint
            BreakPointName     = $name
            Parent             = $parentSegments -join '/'
            Segments           = $segments
            FilesAffected      = $files.Count
            OverWindowsLimit   = @($files | Where-Object { $_.Level -in 'DarkRed', 'Blocked' }).Count
            BlockedBySharePoint = @($files | Where-Object { $_.Level -eq 'Blocked' }).Count
            DeepestLocalPath   = [int]$deepest.LocalPathLength
            DeepestFile        = $deepest.PathInLibrary
            ExtraLevels        = [Math]::Max(0, $maxDepth - $segments.Count)
            Fix                = $fix
            CharactersSaved    = $saved
            FixedByMove        = $fixedAfter
            StillTooLong       = $files.Count - $fixedAfter
            FlattenFix         = $(if ($maxDepth -gt $segments.Count) { $n = $maxDepth - $segments.Count; "Bring the files up out of the $n folder $(if ($n -eq 1) { 'level' } else { 'levels' }) inside '$name'" } else { '' })
            FixedByFlatten     = $fixedFlat
            AddedOn            = $(if ($whenGroup) { $whenGroup.Name } else { '' })
            AddedBy            = $(if ($whoGroup) { $whoGroup.Name } else { '' })
            # "Library|PathInLibrary" of each file in the area. Used to fill the CSV, not shown in the report.
            Members            = @($files | ForEach-Object { "$($_.Library)|$($_.PathInLibrary)" })
        }
    }

    @($areas | Sort-Object @{ Expression = 'FilesAffected'; Descending = $true }, @{ Expression = 'DeepestLocalPath'; Descending = $true })
}

function Get-FolderBars {
    # Data for the report's bar chart: one bar per folder at the chosen level (default 3 = Section/Year/Client),
    # with its subfolders for drill-down. Shapes the running totals collected by Measure-BarStat.
    param([Parameter(Mandatory)] $BarStats, [Parameter(Mandatory)] $ChildStats, [bool] $ShowLibrary)

    $childrenOf = @{}
    foreach ($ck in $ChildStats.Keys) {
        $c = $ChildStats[$ck]
        if (-not $childrenOf.ContainsKey($c.Parent)) { $childrenOf[$c.Parent] = [System.Collections.Generic.List[object]]::new() }
        $childrenOf[$c.Parent].Add([pscustomobject][ordered]@{ Name = $c.Name; Longest = $c.Longest; OverExcel = $c.OverExcel; Files = $c.Files })
    }

    $bars = foreach ($key in $BarStats.Keys) {
        $s = $BarStats[$key]
        $lib, $path = $key -split '\|', 2
        $label = if ($path) { $path } else { '(top level)' }
        if ($ShowLibrary) { $label = "$lib / $label" }
        $children = @(if ($childrenOf.ContainsKey($key)) { $childrenOf[$key] | Sort-Object Longest -Descending })
        [pscustomobject][ordered]@{ Label = $label; Library = $lib; Path = $path; Longest = $s.Longest; OverExcel = $s.OverExcel; Files = $s.Files; Children = $children }
    }
    @($bars | Sort-Object Longest -Descending)
}

function Measure-BarStat {
    # Updates the running totals for the chart for one file.
    param($BarStats, $ChildStats, [string] $Library, [string] $PathInLibrary, [int] $LocalLength, [bool] $OverExcel, [int] $Level)
    $folders = Get-FolderSegments $PathInLibrary
    $depth = [Math]::Min($folders.Count, $Level)
    $barPath = if ($depth -gt 0) { $folders[0..($depth - 1)] -join '/' } else { '' }
    $barKey = "$Library|$barPath"
    Measure-FolderStat $BarStats $barKey $LocalLength $OverExcel $null $null
    if ($folders.Count -gt $depth) {
        $childName = $folders[$depth]
        $childKey = if ($barPath) { "$barKey/$childName" } else { "$barKey$childName" }
        Measure-FolderStat $ChildStats $childKey $LocalLength $OverExcel $barKey $childName
    }
}

function Measure-FolderStat($Table, [string] $Key, [int] $Length, [bool] $OverExcel, $Parent, $Name) {
    if (-not $Table.ContainsKey($Key)) {
        $Table[$Key] = @{ Longest = 0; OverExcel = 0; Files = 0; Parent = $Parent; Name = $Name }
    }
    $s = $Table[$Key]
    $s.Files++
    if ($OverExcel) { $s.OverExcel++ }
    if ($Length -gt $s.Longest) { $s.Longest = $Length }
}
