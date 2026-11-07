# Reads every file and folder in a library and works out path lengths.
# Read-only: the only SharePoint command used here is Get-PnPListItem.

# Columns asked for from SharePoint. Asking only for these keeps each page small and fast.
$script:ScanFields = @(
    'FileRef', 'FileLeafRef', 'FSObjType', 'File_x0020_Size',
    'Created', 'Modified', 'Author', 'Editor', 'ItemChildCount', 'FolderChildCount'
)

# Items per page. Paging is what lets libraries far above 5,000 items be read without
# hitting the list view threshold.
$script:ScanPageSize = 2000

function Get-FieldText($Value) {
    # SharePoint returns some columns as objects (people, counts). This turns them into plain text.
    if ($null -eq $Value) { return '' }
    if ($Value.PSObject.Properties['LookupValue']) { return [string]$Value.LookupValue }
    [string]$Value
}

function ConvertTo-AuditRow {
    # Turns one SharePoint item into one row of the CSV.
    param(
        [Parameter(Mandatory)] $Item,
        [Parameter(Mandatory)] [string] $LibraryTitle,
        [Parameter(Mandatory)] [string] $LibraryRoot,
        [Parameter(Mandatory)] [int] $SyncRootLength
    )

    $fileRef = [string]$Item['FileRef']
    $isFolder = ([string]$Item['FSObjType']) -eq '1'

    # Path inside the library, e.g. "Section/2024 Operation/<Client A>/file.xlsx"
    $relative = $fileRef
    if ($fileRef.StartsWith($LibraryRoot + '/', [StringComparison]::OrdinalIgnoreCase)) {
        $relative = $fileRef.Substring($LibraryRoot.Length + 1)
    }
    $segments = @($relative -split '/' | Where-Object { $_ })
    $folderDepth = if ($isFolder) { $segments.Count } else { [Math]::Max(0, $segments.Count - 1) }

    $sizeText = Get-FieldText $Item['File_x0020_Size']
    $size = 0L
    $null = [long]::TryParse($sizeText, [ref]$size)

    $children = ''
    if ($isFolder) {
        $files = 0; $folders = 0
        $null = [int]::TryParse((Get-FieldText $Item['ItemChildCount']), [ref]$files)
        $null = [int]::TryParse((Get-FieldText $Item['FolderChildCount']), [ref]$folders)
        $children = $files + $folders
    }

    [pscustomobject][ordered]@{
        Library          = $LibraryTitle
        Type             = if ($isFolder) { 'Folder' } else { 'File' }
        Name             = [string]$Item['FileLeafRef']
        PathInLibrary    = $relative
        ServerPath       = $fileRef
        # Estimated local path = sync folder + "\" + path inside the library.
        LocalPathLength  = $SyncRootLength + 1 + $relative.Length
        ServerPathLength = $fileRef.Length
        NameLength       = ([string]$Item['FileLeafRef']).Length
        FolderDepth      = $folderDepth
        SizeBytes        = if ($isFolder) { '' } else { $size }
        DirectChildren   = $children
        Created          = Format-ScanDate $Item['Created']
        CreatedBy        = Get-FieldText $Item['Author']
        Modified         = Format-ScanDate $Item['Modified']
        ModifiedBy       = Get-FieldText $Item['Editor']
    }
}

function Format-ScanDate($Value) {
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    [string]$Value
}

function Test-IsThrottled($ErrorRecord) {
    $ErrorRecord.Exception.Message -match '429|503|throttl|too many requests|server is busy|temporarily unavailable'
}

function Invoke-LibraryScan {
    # Reads one library page by page and writes its rows to a CSV file.
    # PnP already retries single requests when SharePoint throttles (HTTP 429). If it still gives up,
    # this waits and starts the library again, up to 3 times. Returns totals for the library.
    param(
        [Parameter(Mandatory)] $Library,
        [Parameter(Mandatory)] [int] $SyncRootLength,
        [Parameter(Mandatory)] [string] $OutFile,
        [int[]] $RetryWaitSeconds = @(30, 120, 300)
    )

    $attempt = 0
    while ($true) {
        try {
            return Read-LibraryToCsv -Library $Library -SyncRootLength $SyncRootLength -OutFile $OutFile
        }
        catch {
            if (-not (Test-IsThrottled $_) -or $attempt -ge $RetryWaitSeconds.Count) { throw }
            $wait = $RetryWaitSeconds[$attempt]
            $attempt++
            Write-Warning "SharePoint asked us to slow down. Waiting $wait seconds, then reading '$($Library.Title)' again (try $attempt of $($RetryWaitSeconds.Count))."
            Start-Sleep -Seconds $wait
        }
    }
}

function Read-LibraryToCsv {
    param($Library, [int] $SyncRootLength, [string] $OutFile)

    $stats = [ordered]@{ Library = $Library.Title; Items = 0; Files = 0; Folders = 0; LongestLocalPath = 0; LongestPath = '' }
    $buffer = [System.Collections.Generic.List[object]]::new()
    $expected = [Math]::Max(1, [int]$Library.ItemCount)

    # The first write replaces any file left by an earlier try; later writes add to it.
    $state = @{ Append = $false }
    $flush = {
        if ($buffer.Count) {
            $buffer | Export-Csv -LiteralPath $OutFile -Append:$state.Append -NoTypeInformation -Encoding utf8BOM
            $state.Append = $true
            $buffer.Clear()
        }
    }

    Get-PnPListItem -List $Library.Id -PageSize $script:ScanPageSize -Fields $script:ScanFields -ErrorAction Stop |
        ForEach-Object {
            $row = ConvertTo-AuditRow -Item $_ -LibraryTitle $Library.Title -LibraryRoot $Library.ServerRelativeUrl -SyncRootLength $SyncRootLength
            $buffer.Add($row)
            $stats.Items++
            if ($row.Type -eq 'Folder') { $stats.Folders++ } else { $stats.Files++ }
            if ($row.LocalPathLength -gt $stats.LongestLocalPath) {
                $stats.LongestLocalPath = $row.LocalPathLength
                $stats.LongestPath = $row.PathInLibrary
            }
            if ($buffer.Count -ge $script:ScanPageSize) {
                & $flush
                $pct = [Math]::Min(100, [int](100 * $stats.Items / $expected))
                Write-Progress -Activity "Reading '$($Library.Title)'" -Status ("{0:N0} of about {1:N0} items" -f $stats.Items, $expected) -PercentComplete $pct
            }
        }
    & $flush
    Write-Progress -Activity "Reading '$($Library.Title)'" -Completed
    [pscustomobject]$stats
}

function Join-CsvFiles {
    # Combines the per-library CSV files into one, keeping a single header row.
    param([string[]] $Parts, [string] $OutFile)
    $first = $true
    $writer = [System.IO.StreamWriter]::new($OutFile, $false, [System.Text.UTF8Encoding]::new($true))
    try {
        foreach ($part in $Parts) {
            if (-not (Test-Path -LiteralPath $part)) { continue }
            $reader = [System.IO.StreamReader]::new($part, [System.Text.Encoding]::UTF8)
            try {
                $header = $reader.ReadLine()
                if ($first -and $null -ne $header) { $writer.WriteLine($header); $first = $false }
                while ($null -ne ($line = $reader.ReadLine())) { $writer.WriteLine($line) }
            }
            finally { $reader.Dispose() }
        }
    }
    finally { $writer.Dispose() }
}
