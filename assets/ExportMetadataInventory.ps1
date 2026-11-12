<#
.SYNOPSIS
    Generates a detailed CSV inventory of files and folders listed in Excel,
    WITHOUT uploading anything to SharePoint.

.DESCRIPTION
    - Reads the "MetaData" worksheet by default.
    - Processes only rows where "To Be Uploaded?" = "Yes".
    - Uses "Source Folder", with fallback to "UNC Source Folder" and "UNC Folder".
    - Recursively scans folders and files, including long paths and OneDrive-synced folders.
    - Includes the root folder from each Excel row.
    - Copies ALL columns from the source Excel row to every inventory item,
      using the "Excel_" prefix.
    - Collects file-system metadata, timestamps, attributes, extension,
      size, depth, relative path, and other inventory information.
    - Nao conecta no SharePoint e nao executa Add-PnPFile.

.EXAMPLE
    .\ExportMetadataInventory.ps1

.EXAMPLE
.NOTES
#>

param (
    [string]$SheetName = "MetaData"
)

$ErrorActionPreference = "Stop"

$ScriptBase = $PSScriptRoot
$RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFolder = Join-Path $ScriptBase "MetadataReports"
$OutputCsv = Join-Path $OutputFolder ("MetadataInventory_{0}.csv" -f $RunTimestamp)
$ErrorCsv = Join-Path $OutputFolder ("MetadataInventory_Errors_{0}.csv" -f $RunTimestamp)

function Normalize-LocalPath {
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return "" }
    try {
        return ([System.IO.Path]::GetFullPath($PathValue.Trim().Trim('"'))).TrimEnd('\')
    }
    catch {
        return $PathValue.Trim().Trim('"').TrimEnd('\')
    }
}

function Resolve-SourceFolderFromRow {
    param([object]$Row)
    $candidates = @(
        $Row."Source Folder",
        $Row."UNC Source Folder",
        $Row."UNC Folder"
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
            return ([string]$candidate).Trim()
        }
    }
    return ""
}

function Resolve-MappedDriveToUncPath {
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return "" }

    $cleanPath = $PathValue.Trim().Trim('"')
    if ($cleanPath.StartsWith("\\") -or $cleanPath.StartsWith("\\?\")) {
        return $cleanPath
    }

    if ($cleanPath -match "^[A-Za-z]:\\") {
        $driveName = $cleanPath.Substring(0, 1)
        try {
            $drive = Get-PSDrive -Name $driveName -PSProvider FileSystem -ErrorAction Stop
            if ($null -ne $drive -and
                -not [string]::IsNullOrWhiteSpace([string]$drive.DisplayRoot) -and
                ([string]$drive.DisplayRoot).StartsWith("\\")) {
                $relative = $cleanPath.Substring(2).TrimStart("\")
                $uncRoot = ([string]$drive.DisplayRoot).TrimEnd("\")
                if ([string]::IsNullOrWhiteSpace($relative)) { return $uncRoot }
                return $uncRoot + "\" + $relative
            }
        }
        catch {}
    }
    return $cleanPath
}

function Convert-ToExtendedLengthPath {
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return "" }

    $cleanPath = Resolve-MappedDriveToUncPath -PathValue $PathValue
    if ($cleanPath.StartsWith("\\?\")) { return $cleanPath }
    if ($cleanPath.StartsWith("\\")) { return "\\?\UNC\" + $cleanPath.Substring(2) }
    if ($cleanPath -match "^[A-Za-z]:\\") { return "\\?\" + $cleanPath }
    return $cleanPath
}

function Test-IsOneDriveSystemItemName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    return ($Name -match '^\.[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$')
}

function Test-IsSameOrChildPath {
    param([string]$CandidatePath, [string]$RootPath)
    $candidate = Normalize-LocalPath $CandidatePath
    $root = Normalize-LocalPath $RootPath
    if ([string]::IsNullOrWhiteSpace($candidate) -or [string]::IsNullOrWhiteSpace($root)) { return $false }
    if ($candidate.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $candidate.StartsWith($root + "\", [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-RelativePathSafe {
    param([string]$Root, [string]$FullPath)
    $rootNorm = (Normalize-LocalPath $Root).TrimEnd('\')
    $fullNorm = Normalize-LocalPath $FullPath
    if ($fullNorm.Equals($rootNorm, [System.StringComparison]::OrdinalIgnoreCase)) { return "." }
    if ($fullNorm.StartsWith($rootNorm + "\", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullNorm.Substring($rootNorm.Length).TrimStart('\')
    }
    return $fullNorm
}

function Get-PathDepth {
    param([string]$RelativePath)
    if ([string]::IsNullOrWhiteSpace($RelativePath) -or $RelativePath -eq ".") { return 0 }
    return ($RelativePath -split '\\').Count
}


function Convert-ToEnglishValue {
    param([object]$Value)

    # Excel automatically localizes Boolean values when a CSV is opened.
    # Prefixing with an invisible zero-width non-joiner keeps the value as text,
    # so Excel displays "True" / "False" instead of VERDADEIRO / FALSO.
    $textPrefix = [char]0x200C

    if ($Value -is [bool]) {
        if ($Value) { return "$textPrefix" + "True" }
        return "$textPrefix" + "False"
    }

    $stringValue = [string]$Value
    if ($stringValue -match '^(?i:VERDADEIRO|TRUE)$') {
        return "$textPrefix" + "True"
    }
    if ($stringValue -match '^(?i:FALSO|FALSE)$') {
        return "$textPrefix" + "False"
    }

    return $Value
}

function Get-ItemMetadata {
    param(
        [string]$LogicalPath,
        [string]$AccessPath,
        [string]$RootPath,
        [int]$ExcelRow,
        [bool]$IsRoot
    )

    $attributes = [System.IO.File]::GetAttributes($AccessPath)
    $isDirectory = (($attributes -band [System.IO.FileAttributes]::Directory) -ne 0)
    $name = [System.IO.Path]::GetFileName($LogicalPath.TrimEnd('\'))
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $LogicalPath }

    $relative = Get-RelativePathSafe -Root $RootPath -FullPath $LogicalPath
    $parent = [System.IO.Path]::GetDirectoryName($LogicalPath.TrimEnd('\'))
    $extension = if ($isDirectory) { "" } else { [System.IO.Path]::GetExtension($name) }
    $baseName = if ($isDirectory) { $name } else { [System.IO.Path]::GetFileNameWithoutExtension($name) }

    $creationUtc = [System.IO.File]::GetCreationTimeUtc($AccessPath)
    $creationLocal = [System.IO.File]::GetCreationTime($AccessPath)
    $writeUtc = [System.IO.File]::GetLastWriteTimeUtc($AccessPath)
    $writeLocal = [System.IO.File]::GetLastWriteTime($AccessPath)
    $accessUtc = [System.IO.File]::GetLastAccessTimeUtc($AccessPath)
    $accessLocal = [System.IO.File]::GetLastAccessTime($AccessPath)

    [int64]$length = 0
    if (-not $isDirectory) {
        $length = [System.IO.FileInfo]::new($AccessPath).Length
    }
    # Column order optimized for migration analysis:
    # identification -> path -> size -> timestamps -> technical details -> source Excel metadata.
    $obj = [ordered]@{
        ItemType = if ($isDirectory) { "Folder" } else { "File" }
        Name = $name
        BaseName = $baseName
        Extension = $extension

        FullPath = $LogicalPath
        ParentPath = $parent
        SourceRoot = $RootPath
        RelativePath = $relative
        DepthFromSourceRoot = Get-PathDepth $relative
        SizeMB = if ($isDirectory) { "" } else { [math]::Round($length / 1MB, 3) }

        CreationTime = $creationLocal.ToString("yyyy-MM-dd HH:mm:ss.fff")
        LastWriteTime = $writeLocal.ToString("yyyy-MM-dd HH:mm:ss.fff")
        LastAccessTime = $accessLocal.ToString("yyyy-MM-dd HH:mm:ss.fff")
        CreationTimeUtc = $creationUtc.ToString("yyyy-MM-dd HH:mm:ss.fff")
        LastWriteTimeUtc = $writeUtc.ToString("yyyy-MM-dd HH:mm:ss.fff")
        LastAccessTimeUtc = $accessUtc.ToString("yyyy-MM-dd HH:mm:ss.fff")

        PathLength = $LogicalPath.Length
        NameLength = $name.Length
        Attributes = $attributes.ToString()

        ExcelRow = $ExcelRow
        IsRootFromExcel = Convert-ToEnglishValue -Value $IsRoot
    }

    return [PSCustomObject]$obj
}

function Get-LongPathTree {
    param([string]$SourceFolderPath)

    $logicalRoot = (Normalize-LocalPath $SourceFolderPath).TrimEnd('\')
    $accessRoot = (Convert-ToExtendedLengthPath $logicalRoot).TrimEnd('\')

    if ([string]::IsNullOrWhiteSpace($logicalRoot)) {
        throw "Source folder path is empty."
    }

    if (-not [System.IO.Directory]::Exists($accessRoot) -and -not [System.IO.File]::Exists($accessRoot)) {
        throw "Source path not found: $logicalRoot"
    }

    # If the Excel row points directly to a file.
    if ([System.IO.File]::Exists($accessRoot)) {
        [PSCustomObject]@{ LogicalPath=$logicalRoot; AccessPath=$accessRoot; IsRoot=$true }
        return
    }

    # Include the root folder itself.
    [PSCustomObject]@{ LogicalPath=$logicalRoot; AccessPath=$accessRoot; IsRoot=$true }

    $stack = New-Object 'System.Collections.Generic.Stack[object]'
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $stack.Push([PSCustomObject]@{ Logical=$logicalRoot; Access=$accessRoot })

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        if (-not $visited.Add($current.Access)) { continue }

        $entries = [System.IO.Directory]::EnumerateFileSystemEntries($current.Access)
        foreach ($entryAccess in $entries) {
            $entryName = [System.IO.Path]::GetFileName($entryAccess)
            if (Test-IsOneDriveSystemItemName $entryName) { continue }

            $attrs = [System.IO.File]::GetAttributes($entryAccess)
            $isDir = (($attrs -band [System.IO.FileAttributes]::Directory) -ne 0)

            $relativeAccess = $entryAccess.Substring($accessRoot.Length).TrimStart('\')
            $logical = $logicalRoot + "\" + $relativeAccess

            [PSCustomObject]@{ LogicalPath=$logical; AccessPath=$entryAccess; IsRoot=$false }

            if ($isDir) {
                $stack.Push([PSCustomObject]@{ Logical=$logical; Access=$entryAccess })
            }
        }
    }
}

# -------------------------
# Preparation
# -------------------------
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# Use the bundled ImportExcel module when available.
$bundledImportExcel = Join-Path $ScriptBase "Modules\ImportExcel\7.8.10\ImportExcel.psd1"
if (Test-Path $bundledImportExcel) {
    Import-Module $bundledImportExcel -Force
}
elseif (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "ImportExcel module was not found. Keep this script inside the Migration folder used by the original project."
}
else {
    Import-Module ImportExcel -Force
}

Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*"
$dialog.Title = "Select the migration Excel file"

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Excel selection cancelled." -ForegroundColor Yellow
    exit
}

$ExcelFilePath = $dialog.FileName
Write-Host ""
Write-Host "Reading Excel: $ExcelFilePath" -ForegroundColor Cyan
$migrationData = Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName

$yesRows = @($migrationData | Where-Object { ([string]$_."To Be Uploaded?").Trim() -eq "Yes" })
Write-Host "Rows with 'To Be Uploaded?' = Yes: $($yesRows.Count)" -ForegroundColor Cyan

Write-Host ""
Write-Host "Preparing source list..." -ForegroundColor Cyan

# Build a list of valid source rows first.
$sourceRows = New-Object System.Collections.Generic.List[object]
$sourceIndex = 0

foreach ($candidateRow in $migrationData) {
    $sourceIndex++

    if (([string]$candidateRow."To Be Uploaded?").Trim() -ne "Yes") {
        continue
    }

    $candidateSourceRaw = Resolve-SourceFolderFromRow $candidateRow
    if ([string]::IsNullOrWhiteSpace($candidateSourceRaw)) {
        $errors.Add([PSCustomObject]@{
            ExcelRow=$sourceIndex + 1
            SourcePath=""
            ItemPath=""
            Error="No Source Folder / UNC fallback path found."
        })
        continue
    }

    $candidateSource = Normalize-LocalPath $candidateSourceRaw

    $sourceRows.Add([PSCustomObject]@{
        ExcelRow = $sourceIndex + 1
        Row = $candidateRow
        SourceRoot = $candidateSource
    })
}

# Keep only top-level source roots.
# If a source folder is already inside another selected source folder,
# scanning the parent already covers it, so the child source is skipped.
$optimizedSources = New-Object System.Collections.Generic.List[object]

foreach ($candidate in ($sourceRows | Sort-Object { $_.SourceRoot.Length })) {
    $coveredByExistingRoot = $false

    foreach ($existing in $optimizedSources) {
        if (Test-IsSameOrChildPath -CandidatePath $candidate.SourceRoot -RootPath $existing.SourceRoot) {
            $coveredByExistingRoot = $true
            break
        }
    }

    if (-not $coveredByExistingRoot) {
        $optimizedSources.Add($candidate)
    }
}

Write-Host "Selected Excel rows: $($sourceRows.Count)" -ForegroundColor Cyan
Write-Host "Unique top-level source roots: $($optimizedSources.Count)" -ForegroundColor Cyan
Write-Host "Skipped overlapping source rows: $($sourceRows.Count - $optimizedSources.Count)" -ForegroundColor Cyan


# Determine which source-Excel columns contain at least one useful value in the
# optimized rows. Completely empty Excel_* columns are not generated.
$script:ExcelColumnsToInclude = New-Object System.Collections.Generic.List[string]

if ($optimizedSources.Count -gt 0) {
    $allExcelPropertyNames = @(
        $optimizedSources |
        ForEach-Object { $_.Row.PSObject.Properties.Name } |
        Select-Object -Unique
    )

    foreach ($propertyName in $allExcelPropertyNames) {
        $hasValue = $false

        foreach ($sourceInfo in $optimizedSources) {
            $property = $sourceInfo.Row.PSObject.Properties[$propertyName]
            if ($null -ne $property) {
                $value = $property.Value
                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $hasValue = $true
                    break
                }
            }
        }

        if ($hasValue) {
            $script:ExcelColumnsToInclude.Add($propertyName)
        }
    }
}


# Count the actual unique items that will be exported.
# Because overlapping child roots were removed above, this estimate now matches
# the real inventory much more closely and avoids rescanning the same trees.
Write-Host ""
Write-Host "Preparing progress tracking..." -ForegroundColor Cyan

$estimatedTotalItems = 0
$preScanSource = 0

foreach ($sourceInfo in $optimizedSources) {
    $preScanSource++
    $prePercent = if ($optimizedSources.Count -gt 0) {
        [math]::Round(($preScanSource / $optimizedSources.Count) * 100, 0)
    } else {
        100
    }

    Write-Progress -Id 1 `
        -Activity "Preparing inventory" `
        -Status "Counting source $preScanSource of $($optimizedSources.Count)" `
        -CurrentOperation $sourceInfo.SourceRoot `
        -PercentComplete $prePercent

    try {
        foreach ($null in Get-LongPathTree -SourceFolderPath $sourceInfo.SourceRoot) {
            $estimatedTotalItems++
        }
    }
    catch {
        # The main scan records the detailed error.
    }
}

Write-Progress -Id 1 -Activity "Preparing inventory" -Completed
Write-Host "Unique items to process: $estimatedTotalItems" -ForegroundColor Cyan
Write-Host "Output: $OutputCsv" -ForegroundColor Cyan

$errors = New-Object System.Collections.Generic.List[object]
$seenItems = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$firstWrite = $true
$totalExported = 0
foreach ($sourceInfo in $optimizedSources) {
    $row = $sourceInfo.Row
    $excelRow = $sourceInfo.ExcelRow
    $sourceRoot = $sourceInfo.SourceRoot

    Write-Host ""
    Write-Host "[$excelRow] Scanning: $sourceRoot" -ForegroundColor Green

    try {
        foreach ($entry in Get-LongPathTree -SourceFolderPath $sourceRoot) {
            # This remains as an additional safeguard against duplicate paths.
            if (-not $seenItems.Add($entry.LogicalPath)) {
                continue
            }

            try {
                $metadata = Get-ItemMetadata `
                    -LogicalPath $entry.LogicalPath `
                    -AccessPath $entry.AccessPath `
                    -RootPath $sourceRoot `
                    -ExcelRow $excelRow `
                    -IsRoot ([bool]$entry.IsRoot)

                if ($firstWrite) {
                    $metadata | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                    $firstWrite = $false
                }
                else {
                    $metadata | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Append
                }

                $totalExported++

                if ($estimatedTotalItems -gt 0) {
                    $progressPercent = [math]::Min(
                        100,
                        [math]::Round(($totalExported / $estimatedTotalItems) * 100, 1)
                    )

                    Write-Progress -Id 2 `
                        -Activity "Creating metadata inventory" `
                        -Status "$totalExported of $estimatedTotalItems items processed - $progressPercent%" `
                        -CurrentOperation $entry.LogicalPath `
                        -PercentComplete $progressPercent
                }

                if (($totalExported % 500) -eq 0) {
                    Write-Host "Processed: $totalExported of $estimatedTotalItems items..." -ForegroundColor DarkCyan
                }
            }
            catch {
                $errors.Add([PSCustomObject]@{
                    ExcelRow=$excelRow
                    SourcePath=$sourceRoot
                    ItemPath=$entry.LogicalPath
                    Error=$_.Exception.Message
                })
            }
        }
    }
    catch {
        $errors.Add([PSCustomObject]@{
            ExcelRow=$excelRow
            SourcePath=$sourceRoot
            ItemPath=$sourceRoot
            Error=$_.Exception.Message
        })
    }
}

Write-Progress -Id 2 -Activity "Creating metadata inventory" -Completed

if ($errors.Count -gt 0) {
    $errors | Export-Csv -Path $ErrorCsv -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}

Write-Progress -Id 2 -Activity "Creating metadata inventory" -Completed

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "METADATA INVENTORY COMPLETED" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Items exported: $totalExported" -ForegroundColor Green
Write-Host "Errors: $($errors.Count)" -ForegroundColor $(if ($errors.Count -gt 0) { "Yellow" } else { "Green" })
if (-not $firstWrite) {
    Write-Host "CSV: $OutputCsv" -ForegroundColor Green
}
if ($errors.Count -gt 0) {
    Write-Host "Error CSV: $ErrorCsv" -ForegroundColor Yellow
}
Write-Host ""
