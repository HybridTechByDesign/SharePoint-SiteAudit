# Checks run on the scan results (AllFindings.csv). Local only: nothing here talks to SharePoint.
#
# The raw scan is read twice, row by row, so memory stays low on big sites:
#   pass 1 counts things that need the whole site (duplicates, bulk copy-ins, folder sizes)
#   pass 2 adds the check columns to each row and writes the final CSV.

# A folder counts as a bulk copy-in when at least this many items inside it were created on the
# same day, by the same person, who also created the folder itself that day.
$script:BulkCopyMinItems = 50

$script:FolderChildLimit = 5000      # list view threshold for direct children
$script:SubtreeLimit = 100000        # items below one folder
$script:LibrarySyncLimit = 300000    # OneDrive sync recommendation per library

$script:ReservedNames = '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$'
$script:JunkPatterns = @(
    @{ Pattern = '^~\$'; Reason = 'Office temp file (~$)' }
    @{ Pattern = '^thumbs\.db$'; Reason = 'Thumbs.db' }
    @{ Pattern = '^desktop\.ini$'; Reason = 'desktop.ini' }
    @{ Pattern = '\.tmp$'; Reason = 'Temp file (.tmp)' }
    @{ Pattern = '\.pst$'; Reason = 'Outlook data file (.pst)' }
)

function Get-PathLevel {
    # Green < Warn <= Amber <= Excel < Red <= Windows < DarkRed. Blocked when the SharePoint path is over the limit.
    param([int] $LocalLength, [int] $ServerLength, $Thresholds)
    if ($ServerLength -gt $Thresholds.SharePoint) { return 'Blocked' }
    if ($LocalLength -gt $Thresholds.Windows) { return 'DarkRed' }
    if ($LocalLength -gt $Thresholds.Excel) { return 'Red' }
    if ($LocalLength -ge $Thresholds.Warn) { return 'Amber' }
    'Green'
}

function Get-NormalizedFolderName([string] $Name) {
    # "Projects 2024" -> "projects", "Final v2" -> "final v", "2024 Operation" -> "operation"
    $n = $Name.ToLowerInvariant() -replace '\d+', ' ' -replace '[^\p{L}\s]', ' ' -replace '\s+', ' '
    $n.Trim()
}

function Find-RepeatedFolder {
    # Returns "A / B" when the same or nearly the same folder name appears twice in the path, else ''.
    param([string] $PathInLibrary, [bool] $IsFolder)
    $segments = @($PathInLibrary -split '/' | Where-Object { $_ })
    if (-not $IsFolder -and $segments.Count) { $segments = @($segments | Select-Object -SkipLast 1) }
    if ($segments.Count -lt 2) { return '' }

    $norm = @($segments | ForEach-Object { Get-NormalizedFolderName $_ })
    for ($i = 0; $i -lt $segments.Count; $i++) {
        if ($norm[$i].Length -lt 3) { continue }
        for ($j = $i + 1; $j -lt $segments.Count; $j++) {
            if ($norm[$j].Length -lt 3) { continue }
            $a = $norm[$i]; $b = $norm[$j]
            if ($a -eq $b -or $b.StartsWith("$a ") -or $a.StartsWith("$b ")) {
                return "$($segments[$i]) / $($segments[$j])"
            }
        }
    }
    ''
}

function Get-ConflictCopyPc {
    # OneDrive names conflict copies "<name>-<PCNAME>.<ext>" (sometimes with "-1", "-2" after it).
    param([string] $Name, [string[]] $PcNames)
    if (-not $PcNames) { return '' }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    foreach ($pc in $PcNames) {
        if ($base -match ('-' + [regex]::Escape($pc) + '(-\d+)?$')) { return $pc }
    }
    ''
}

function Get-BadNameReason {
    param([string] $Name)
    $reasons = [System.Collections.Generic.List[string]]::new()
    if ($Name -ne $Name.TrimStart()) { $reasons.Add('starts with a space') }
    if ($Name -ne $Name.TrimEnd()) { $reasons.Add('ends with a space') }
    if ($Name.EndsWith('.')) { $reasons.Add('ends with a dot') }
    $stem = ($Name -split '\.')[0].Trim()
    if ($stem -match $script:ReservedNames) { $reasons.Add("reserved name ($stem)") }
    if ($Name -match '_vti_') { $reasons.Add('contains _vti_') }
    $reasons -join '; '
}

function Get-JunkReason {
    param([string] $Name)
    foreach ($j in $script:JunkPatterns) {
        if ($Name -match $j.Pattern) { return $j.Reason }
    }
    ''
}

function Get-AncestorKeys {
    # "Lib", "A/B/C/file.txt" -> "Lib|A", "Lib|A/B", "Lib|A/B/C" (plus "Lib|A/B/C/file.txt"'s own key if IncludeSelf).
    param([string] $Library, [string] $PathInLibrary, [switch] $IncludeSelf)
    $segments = @($PathInLibrary -split '/' | Where-Object { $_ })
    $last = if ($IncludeSelf) { $segments.Count } else { $segments.Count - 1 }
    $keys = [System.Collections.Generic.List[string]]::new()
    for ($k = 1; $k -le $last; $k++) { $keys.Add($Library + '|' + ($segments[0..($k - 1)] -join '/')) }
    $keys
}

function Invoke-AuditChecks {
    # Reads the raw scan CSV, adds the check columns, writes the final CSV. Returns counts for the summary.
    param(
        [Parameter(Mandatory)] [string] $RawCsv,
        [Parameter(Mandatory)] [string] $OutCsv,
        [Parameter(Mandatory)] $Config,
        # Item paths (ServerPath) that have their own permissions, from Get-UniquePermissionPaths.
        [System.Collections.Generic.HashSet[string]] $UniquePermissionPaths
    )

    $largeBytes = [long]([double]$Config.LargeFileSizeGB * 1GB)
    $dupCount = [System.Collections.Generic.Dictionary[string, int]]::new()
    $folderInfo = [System.Collections.Generic.Dictionary[string, string]]::new()   # folder key -> "day|person"
    $tally = [System.Collections.Generic.Dictionary[string, int]]::new()           # "folder key|day|person" -> items
    $subtree = [System.Collections.Generic.Dictionary[string, int]]::new()         # folder key -> items below

    # ---- Pass 1: counts that need the whole site ----
    Import-Csv -LiteralPath $RawCsv | ForEach-Object {
        $isFolder = $_.Type -eq 'Folder'
        $day = ($_.Created -split ' ')[0]
        $stamp = "$day|$($_.CreatedBy)"
        if ($isFolder) {
            $folderInfo[$_.Library + '|' + $_.PathInLibrary] = $stamp
        }
        elseif ([long]$_.SizeBytes -gt 0) {
            $key = $_.Name.ToLowerInvariant() + '|' + $_.SizeBytes
            $dupCount[$key] = 1 + $(if ($dupCount.ContainsKey($key)) { $dupCount[$key] } else { 0 })
        }
        foreach ($a in (Get-AncestorKeys $_.Library $_.PathInLibrary)) {
            $subtree[$a] = 1 + $(if ($subtree.ContainsKey($a)) { $subtree[$a] } else { 0 })
            $t = "$a|$stamp"
            $tally[$t] = 1 + $(if ($tally.ContainsKey($t)) { $tally[$t] } else { 0 })
        }
    }

    # Bulk copy-in folders: enough same-day, same-person items, and the folder itself created that day by that person.
    $bulk = [System.Collections.Generic.Dictionary[string, string]]::new()
    foreach ($f in $folderInfo.Keys) {
        $t = "$f|$($folderInfo[$f])"
        if ($tally.ContainsKey($t) -and $tally[$t] -ge $script:BulkCopyMinItems) { $bulk[$f] = $folderInfo[$f] }
    }
    # Keep only the top folder of each copied-in tree.
    $topBulk = [System.Collections.Generic.Dictionary[string, string]]::new()
    foreach ($f in $bulk.Keys) {
        $parent = $f.Substring(0, [Math]::Max($f.LastIndexOf('/'), $f.IndexOf('|')))
        if (-not ($bulk.ContainsKey($parent) -and $bulk[$parent] -eq $bulk[$f])) { $topBulk[$f] = $bulk[$f] }
    }

    # ---- Pass 2: add columns and write ----
    $counts = [ordered]@{
        Items = 0; Files = 0; Folders = 0
        Green = 0; Amber = 0; Red = 0; DarkRed = 0; Blocked = 0
        OverExcel = 0; NameTooLong = 0; RepeatedFolder = 0; ConflictCopy = 0; LikelyDuplicate = 0
        BadName = 0; Junk = 0; LargeFile = 0; BulkCopyFolders = $topBulk.Count; TooManyChildren = 0
        HugeSubtree = 0; CheckedOut = 0; UniquePermissions = 0; DeepestFolderDepth = 0; LongestLocalPath = 0
    }
    $th = $Config.Thresholds

    Import-Csv -LiteralPath $RawCsv | ForEach-Object {
        $row = $_
        $isFolder = $row.Type -eq 'Folder'
        $selfKey = $row.Library + '|' + $row.PathInLibrary
        $issues = [System.Collections.Generic.List[string]]::new()

        $level = Get-PathLevel ([int]$row.LocalPathLength) ([int]$row.ServerPathLength) $th
        if ($level -ne 'Green') { $issues.Add("Path $level") }
        $nameTooLong = if ([int]$row.NameLength -gt 255) { 'Yes' } else { '' }
        if ($nameTooLong) { $issues.Add('Name too long') }
        $repeated = Find-RepeatedFolder $row.PathInLibrary $isFolder
        if ($repeated) { $issues.Add('Repeated folder') }
        $badName = Get-BadNameReason $row.Name
        if ($badName) { $issues.Add('Bad name') }

        $conflict = ''; $dup = ''; $junk = ''; $large = ''; $tooMany = ''; $huge = ''; $subCount = ''
        if ($isFolder) {
            $subCount = if ($subtree.ContainsKey($selfKey)) { $subtree[$selfKey] } else { 0 }
            if ([int]("0$($row.DirectChildren)") -gt $script:FolderChildLimit) { $tooMany = 'Yes'; $issues.Add('Over 5,000 items in folder') }
            if ($subCount -gt $script:SubtreeLimit) { $huge = 'Yes'; $issues.Add('Over 100,000 items below folder') }
        }
        else {
            $conflict = Get-ConflictCopyPc $row.Name $Config.PcNames
            if ($conflict) { $issues.Add('Conflict copy') }
            if ([long]$row.SizeBytes -gt 0) {
                $n = $dupCount[$row.Name.ToLowerInvariant() + '|' + $row.SizeBytes]
                if ($n -gt 1) { $dup = $n; $issues.Add('Likely duplicate') }
            }
            $junk = Get-JunkReason $row.Name
            if ($junk) { $issues.Add('Junk file') }
            if ([long]$row.SizeBytes -gt $largeBytes) { $large = 'Yes'; $issues.Add('Large file') }
        }

        # Top copied-in folder this item belongs to (same day and person as the copy).
        $bulkFolder = ''
        if ($topBulk.Count) {
            $stamp = "$(($row.Created -split ' ')[0])|$($row.CreatedBy)"
            foreach ($a in (Get-AncestorKeys $row.Library $row.PathInLibrary -IncludeSelf:$isFolder)) {
                if ($topBulk.ContainsKey($a) -and $topBulk[$a] -eq $stamp) { $bulkFolder = $a.Substring($a.IndexOf('|') + 1); break }
            }
        }
        if ($bulkFolder) { $issues.Add('Bulk copy-in') }

        $unique = if ($UniquePermissionPaths -and $UniquePermissionPaths.Contains($row.ServerPath)) { 'Yes' } else { '' }
        if ($unique) { $issues.Add('Unique permissions') }
        if ($row.CheckedOutTo) { $issues.Add('Checked out') }

        # Tally for the summary.
        $counts.Items++
        if ($isFolder) { $counts.Folders++ } else { $counts.Files++ }
        if (-not $isFolder) {
            $counts[$level]++
            if ($level -in 'Red', 'DarkRed', 'Blocked') { $counts.OverExcel++ }
            if ($conflict) { $counts.ConflictCopy++ }
            if ($dup) { $counts.LikelyDuplicate++ }
            if ($junk) { $counts.Junk++ }
            if ($large) { $counts.LargeFile++ }
            if ($row.CheckedOutTo) { $counts.CheckedOut++ }
        }
        if ($nameTooLong) { $counts.NameTooLong++ }
        if ($repeated) { $counts.RepeatedFolder++ }
        if ($badName) { $counts.BadName++ }
        if ($tooMany) { $counts.TooManyChildren++ }
        if ($huge) { $counts.HugeSubtree++ }
        if ($unique) { $counts.UniquePermissions++ }
        if ([int]$row.FolderDepth -gt $counts.DeepestFolderDepth) { $counts.DeepestFolderDepth = [int]$row.FolderDepth }
        if ([int]$row.LocalPathLength -gt $counts.LongestLocalPath) { $counts.LongestLocalPath = [int]$row.LocalPathLength }

        $row | Add-Member -NotePropertyMembers ([ordered]@{
                Level             = $level
                Issues            = $issues -join '; '
                NameTooLong       = $nameTooLong
                RepeatedFolder    = $repeated
                ConflictCopyOf    = $conflict
                LikelyDuplicates  = $dup
                BadName           = $badName
                JunkFile          = $junk
                LargeFile         = $large
                BulkCopyFolder    = $bulkFolder
                SubtreeItems      = $subCount
                TooManyInFolder   = $tooMany
                HugeSubtree       = $huge
                UniquePermissions = $unique
            }) -PassThru
    } | Export-Csv -LiteralPath $OutCsv -NoTypeInformation -Encoding utf8BOM

    $counts['BulkCopyFolderList'] = @($topBulk.Keys | ForEach-Object { $_.Substring($_.IndexOf('|') + 1) } | Sort-Object)
    [pscustomobject]$counts
}

function Get-LibraryFindings {
    # Library-level checks for the summary (check 11-13). Settings come from Get-AuditLibraries,
    # required columns and unique permissions from the scanner.
    param($Library, $Stats, [string[]] $RequiredColumns, $UniqueCount, [int] $CheckedOut)
    $notes = [System.Collections.Generic.List[string]]::new()
    if ($Stats.Items -gt $script:LibrarySyncLimit) { $notes.Add(("Over 300,000 items ({0:N0}): too big to sync reliably" -f $Stats.Items)) }
    if ($Library.RequireCheckOut -eq 'Yes') { $notes.Add('Require Check Out is on: synced files open read-only') }
    if ($Library.ContentApproval -eq 'Yes') { $notes.Add('Content approval is on: can stop files from syncing') }
    if ($Library.OfflineSyncOff -eq 'Yes') { $notes.Add('Offline sync is turned off for this library') }
    if ($RequiredColumns) { $notes.Add("Required columns ($($RequiredColumns -join ', ')): new files can't sync until filled in") }
    if ($CheckedOut -gt 0) { $notes.Add("$CheckedOut file(s) checked out") }
    [ordered]@{
        Items                 = $Stats.Items
        OverSyncLimit         = $Stats.Items -gt $script:LibrarySyncLimit
        RequireCheckOut       = $Library.RequireCheckOut
        ContentApproval       = $Library.ContentApproval
        OfflineSyncOff        = $Library.OfflineSyncOff
        RequiredColumns       = @($RequiredColumns)
        CheckedOutFiles       = $CheckedOut
        UniquePermissionItems = $UniqueCount
        Notes                 = @($notes)
    }
}
