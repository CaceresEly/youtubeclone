<#
.SYNOPSIS
    PORTABLE version v16 of the SharePoint migration script. Includes optional SharePoint subfolder routing before /Forms URLs.
    Uses -UseWebLogin for authentication — no Azure App ID required.
    Includes Pause/Resume/Stop controls and immediate persistent state matrix.
#>
param (
    [string]$SheetName = "MetaData"
)
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
# =========================
# Global variables
# =========================
$Global:CurrentSPSiteUrl = ""
$Global:LogEntries = @()
$Global:ProcessingMatrix = @()
$Global:PersistentMatrix = @()
$Global:ScriptBase = $PSScriptRoot
$Global:RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Global:LogsFolder = Join-Path -Path $Global:ScriptBase -ChildPath "Logs"
$Global:PersistentMatrixPath = Join-Path -Path $Global:LogsFolder -ChildPath "StateMatrix.csv"
$Global:LegacyPersistentMatrixPath = Join-Path -Path $Global:LogsFolder -ChildPath "MatrizEstado.csv"
$Global:ValidSubFolders = @("BR", "CR", "EPDM", "EVM", "HNBR", "IIR", "NBR", "SBR")
$Global:DestinationLibraryMap = @{}
$Global:IsPaused = $false
$Global:StopRequested = $false
$Global:CurrentExcelRow = 0
$Global:CurrentDescription = ""

# =========================
# Upload progress
# =========================
$Global:ProgressTotalFiles = 0
$Global:ProgressCompletedFiles = 0
$Global:ProgressItemKeys = @{}
$Global:ProgressCompletedKeys = @{}
$Global:ProgressItemBytes = @{}
$Global:ProgressTotalBytes = [int64]0
$Global:ProgressCompletedBytes = [int64]0

# =========================
# Helper functions
# =========================
function Normalize-LocalPath {
    param (
        [string]$PathValue
    )
    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return ""
    }
    try {
        $full = [System.IO.Path]::GetFullPath($PathValue.Trim().Trim('"'))
        return $full.TrimEnd('\')
    }
    catch {
        return $PathValue.Trim().Trim('"').TrimEnd('\')
    }
}
function Normalize-DestinationKey {
    param (
        [string]$DestinationValue
    )
    if ([string]::IsNullOrWhiteSpace($DestinationValue)) {
        return ""
    }
    return $DestinationValue.Trim().TrimEnd('/').ToLowerInvariant()
}
function Normalize-SubFolderValue {
    param (
        [string]$SubFolderValue
    )
 
    if ([string]::IsNullOrWhiteSpace($SubFolderValue)) {
        return ""
    }
 
    return $SubFolderValue.Trim().Trim("/").Trim("\\").ToUpperInvariant()
}
 
function Get-FinalDestinationUrl {
    param (
        [string]$DestinationFolder,
        [string]$SubFolder
    )
 
    if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
        return ""
    }
 
    $cleanDestination = $DestinationFolder.Trim()
 
    if ([string]::IsNullOrWhiteSpace($SubFolder)) {
        return $cleanDestination
    }
 
    $cleanSubFolder = $SubFolder.Trim().Trim("/").Trim("\\")
 
    # Important: Excel often contains SharePoint view URLs such as:
    # https://tenant/sites/site/Library/Forms/AllItems.aspx
    # The subfolder must be inserted before /Forms/, otherwise Resolve-TargetInfo
    # removes everything after /Forms/ and the upload falls back to the library root.
    $formsIndex = $cleanDestination.IndexOf("/Forms/", [System.StringComparison]::OrdinalIgnoreCase)
 
    if ($formsIndex -ge 0) {
        $libraryUrl = $cleanDestination.Substring(0, $formsIndex).TrimEnd("/")
        return $libraryUrl + "/" + $cleanSubFolder
    }
 
    return $cleanDestination.TrimEnd("/") + "/" + $cleanSubFolder
}
 
function Get-DisplayName {
    param (
        [string]$Description,
        [string]$SourcePath
    )
    if (-not [string]::IsNullOrWhiteSpace($Description)) {
        return $Description.Trim()
    }
    $leaf = Split-Path -Path $SourcePath -Leaf
    if (-not [string]::IsNullOrWhiteSpace($leaf)) {
        return $leaf
    }
    return $SourcePath
}

function Normalize-LibraryNameKey {
    param (
        [string]$LibraryName
    )
    if ([string]::IsNullOrWhiteSpace($LibraryName)) {
        return ""
    }
    return $LibraryName.Trim().ToLowerInvariant()
}

function Initialize-DestinationLibraryMap {
    param (
        [string]$ExcelFilePath
    )

    $Global:DestinationLibraryMap = @{}

    try {
        $listsData = Import-Excel -Path $ExcelFilePath -WorksheetName "Lists" -ErrorAction Stop
        foreach ($libraryRow in $listsData) {
            $libraryName = $libraryRow."Library"
            $libraryUrl  = $libraryRow."Library URL"

            if (-not [string]::IsNullOrWhiteSpace($libraryName) -and -not [string]::IsNullOrWhiteSpace($libraryUrl)) {
                $Global:DestinationLibraryMap[(Normalize-LibraryNameKey -LibraryName $libraryName)] = $libraryUrl.Trim()
            }
        }

        Write-Host "Destination library map loaded: $($Global:DestinationLibraryMap.Count) libraries." -ForegroundColor Cyan
    }
    catch {
        Write-Host "Warning: could not load worksheet 'Lists'. If the MetaData sheet uses library names instead of URLs, rows will fail validation." -ForegroundColor Yellow
        Write-Host $_.Exception.Message -ForegroundColor DarkYellow
    }
}

function Resolve-DestinationLibraryUrlFromRow {
    param (
        [object]$Row
    )

    # Supported formats:
    # 1) Column "Destination Library URL" already containing the full SharePoint URL.
    # 2) Column "Destination Library" containing either the URL itself or a library name that exists in worksheet "Lists".
    $directUrl = $Row."Destination Library URL"
    if (-not [string]::IsNullOrWhiteSpace($directUrl)) {
        return $directUrl.Trim()
    }

    $libraryValue = $Row."Destination Library"
    if ([string]::IsNullOrWhiteSpace($libraryValue)) {
        return ""
    }

    $libraryValue = $libraryValue.Trim()

    if ($libraryValue -match "^https?://") {
        return $libraryValue
    }

    $key = Normalize-LibraryNameKey -LibraryName $libraryValue
    if ($Global:DestinationLibraryMap.ContainsKey($key)) {
        return $Global:DestinationLibraryMap[$key]
    }

    return ""
}
function Resolve-SourceFolderFromRow {
    param (
        [object]$Row
    )

    # Workbook layout support:
    # Priority 1: "Source Folder" - main source path column.
    # Priority 2: "UNC Source Folder" - fallback only when Source Folder is empty.
    # Priority 3: "UNC Folder" - fallback for alternative workbook naming.
    # Important: UNC columns are NOT preferred. They are used only as a last option.
    $sourceCandidates = @(
        $Row."Source Folder",
        $Row."UNC Source Folder",
        $Row."UNC Folder"
    )

    foreach ($candidate in $sourceCandidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
            return ([string]$candidate).Trim()
        }
    }

    return ""
}

function Test-IsSameOrChildPath {
    param (
        [string]$CandidatePath,
        [string]$RootPath
    )
    $candidate = Normalize-LocalPath -PathValue $CandidatePath
    $root = Normalize-LocalPath -PathValue $RootPath
    if ([string]::IsNullOrWhiteSpace($candidate) -or [string]::IsNullOrWhiteSpace($root)) {
        return $false
    }
    if ($candidate.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    return $candidate.StartsWith($root + "\", [System.StringComparison]::OrdinalIgnoreCase)
}

# =========================
# Upload progress functions
# =========================
function Get-ProgressItemKey {
    param (
        [string]$SourcePath,
        [string]$DestinationUrl
    )

    $normalizedSource = Normalize-LocalPath -PathValue $SourcePath
    $destinationKey = Normalize-DestinationKey -DestinationValue $DestinationUrl

    return ($normalizedSource.ToLowerInvariant() + "|" + $destinationKey)
}

function Format-ByteSize {
    param (
        [Int64]$Bytes
    )

    if ($Bytes -lt 1KB) {
        return ("{0:N0} B" -f $Bytes)
    }
    elseif ($Bytes -lt 1MB) {
        return ("{0:N2} KB" -f ($Bytes / 1KB))
    }
    elseif ($Bytes -lt 1GB) {
        return ("{0:N2} MB" -f ($Bytes / 1MB))
    }
    elseif ($Bytes -lt 1TB) {
        return ("{0:N2} GB" -f ($Bytes / 1GB))
    }
    else {
        return ("{0:N2} TB" -f ($Bytes / 1TB))
    }
}

function Show-UploadProgress {
    param (
        [string]$CurrentItem = ""
    )

    $total = [int]$Global:ProgressTotalFiles
    $completed = [int]$Global:ProgressCompletedFiles
    $totalBytes = [int64]$Global:ProgressTotalBytes
    $completedBytes = [int64]$Global:ProgressCompletedBytes

    if ($total -le 0) {
        Write-Progress -Id 1 -Activity "SharePoint upload progress" -Status "No files pending for rows with To Be Uploaded? = Yes" -PercentComplete 100
        return
    }

    if ($completed -lt 0) { $completed = 0 }
    if ($completed -gt $total) { $completed = $total }
    if ($completedBytes -lt 0) { $completedBytes = 0 }
    if ($completedBytes -gt $totalBytes) { $completedBytes = $totalBytes }

    $remaining = $total - $completed
    $remainingBytes = $totalBytes - $completedBytes
    $percent = [math]::Round(($completed / $total) * 100, 1)

    if ($totalBytes -gt 0) {
        $bytePercent = [math]::Round(($completedBytes / $totalBytes) * 100, 1)
    }
    else {
        $bytePercent = 100
    }

    $completedBytesText = Format-ByteSize -Bytes $completedBytes
    $totalBytesText = Format-ByteSize -Bytes $totalBytes
    $remainingBytesText = Format-ByteSize -Bytes $remainingBytes

    $status = "Files: ${percent}% ($completed/$total, remaining $remaining) | Data: ${bytePercent}% ($completedBytesText / $totalBytesText, remaining $remainingBytesText)"
    if (-not [string]::IsNullOrWhiteSpace($CurrentItem)) {
        $status += " | Current: $CurrentItem"
    }

    Write-Progress -Id 1 -Activity "SharePoint upload progress" -Status $status -PercentComplete $percent
}

function Mark-ProgressItemComplete {
    param (
        [string]$SourcePath,
        [string]$DestinationUrl,
        [string]$CurrentItem = ""
    )

    $key = Get-ProgressItemKey -SourcePath $SourcePath -DestinationUrl $DestinationUrl

    if ($Global:ProgressItemKeys.ContainsKey($key) -and -not $Global:ProgressCompletedKeys.ContainsKey($key)) {
        $Global:ProgressCompletedKeys[$key] = $true
        $Global:ProgressCompletedFiles++

        if ($Global:ProgressItemBytes.ContainsKey($key)) {
            $Global:ProgressCompletedBytes += [int64]$Global:ProgressItemBytes[$key]
        }
    }

    Show-UploadProgress -CurrentItem $CurrentItem
}

function Initialize-UploadProgress {
    param (
        [object[]]$MigrationData
    )

    $Global:ProgressTotalFiles = 0
    $Global:ProgressCompletedFiles = 0
    $Global:ProgressItemKeys = @{}
    $Global:ProgressCompletedKeys = @{}
    $Global:ProgressItemBytes = @{}
    $Global:ProgressTotalBytes = [int64]0
    $Global:ProgressCompletedBytes = [int64]0

    Write-Host ""
    Write-Host "Calculating upload progress..." -ForegroundColor Cyan

    foreach ($row in $MigrationData) {
        $toBeUpload = [string]$row."To Be Uploaded?"

        if ($toBeUpload.Trim() -ne "Yes") {
            continue
        }

        $sourceFolderRaw = Resolve-SourceFolderFromRow -Row $row
        $destinationLibraryUrl = Resolve-DestinationLibraryUrlFromRow -Row $row
        $destinationFolderRaw = $row."Destination Folder"

        if ([string]::IsNullOrWhiteSpace($sourceFolderRaw) -or
            [string]::IsNullOrWhiteSpace($destinationLibraryUrl)) {
            continue
        }

        $sourceFolder = Normalize-LocalPath -PathValue $sourceFolderRaw
        $subFolder = Normalize-SubFolderValue -SubFolderValue $destinationFolderRaw

        if (-not [string]::IsNullOrWhiteSpace($subFolder) -and
            $Global:ValidSubFolders -notcontains $subFolder) {
            continue
        }

        $finalDestinationFolder = Get-FinalDestinationUrl -DestinationFolder $destinationLibraryUrl -SubFolder $subFolder
        $destinationKey = Normalize-DestinationKey -DestinationValue $finalDestinationFolder

        if (-not (Test-Path -LiteralPath $sourceFolder)) {
            continue
        }

        $sourceItem = Get-Item -LiteralPath $sourceFolder -ErrorAction SilentlyContinue
        if ($null -eq $sourceItem) {
            continue
        }

        if ($sourceItem.PSIsContainer) {
            $files = Get-ChildItem -LiteralPath $sourceFolder -Recurse -File -ErrorAction SilentlyContinue

            foreach ($file in $files) {
                $key = Get-ProgressItemKey -SourcePath $file.FullName -DestinationUrl $finalDestinationFolder

                # Prevent double counting when the same file/destination is covered by multiple rows.
                if (-not $Global:ProgressItemKeys.ContainsKey($key)) {
                    $fileBytes = [int64]$file.Length
                    $Global:ProgressItemKeys[$key] = $true
                    $Global:ProgressItemBytes[$key] = $fileBytes
                    $Global:ProgressTotalFiles++
                    $Global:ProgressTotalBytes += $fileBytes

                    if (Test-ItemPreviouslyProcessed -ItemPath $file.FullName -DestinationKey $destinationKey) {
                        $Global:ProgressCompletedKeys[$key] = $true
                        $Global:ProgressCompletedFiles++
                        $Global:ProgressCompletedBytes += $fileBytes
                    }
                }
            }
        }
        else {
            $key = Get-ProgressItemKey -SourcePath $sourceFolder -DestinationUrl $finalDestinationFolder

            if (-not $Global:ProgressItemKeys.ContainsKey($key)) {
                $fileBytes = [int64]$sourceItem.Length
                $Global:ProgressItemKeys[$key] = $true
                $Global:ProgressItemBytes[$key] = $fileBytes
                $Global:ProgressTotalFiles++
                $Global:ProgressTotalBytes += $fileBytes

                if (Test-PathPreviouslyProcessed -SourcePath $sourceFolder -DestinationKey $destinationKey) {
                    $Global:ProgressCompletedKeys[$key] = $true
                    $Global:ProgressCompletedFiles++
                    $Global:ProgressCompletedBytes += $fileBytes
                }
            }
        }
    }

    $remaining = $Global:ProgressTotalFiles - $Global:ProgressCompletedFiles

    $remainingBytes = $Global:ProgressTotalBytes - $Global:ProgressCompletedBytes
    $totalBytesText = Format-ByteSize -Bytes $Global:ProgressTotalBytes
    $completedBytesText = Format-ByteSize -Bytes $Global:ProgressCompletedBytes
    $remainingBytesText = Format-ByteSize -Bytes $remainingBytes

    Write-Host "Files considered (To Be Uploaded? = Yes): $($Global:ProgressTotalFiles)" -ForegroundColor Cyan
    Write-Host "Already uploaded according to StateMatrix.csv: $($Global:ProgressCompletedFiles)" -ForegroundColor DarkGray
    Write-Host "Files remaining: $remaining" -ForegroundColor Yellow
    Write-Host "Total upload size: $totalBytesText ($($Global:ProgressTotalBytes) bytes)" -ForegroundColor Cyan
    Write-Host "Already uploaded size: $completedBytesText ($($Global:ProgressCompletedBytes) bytes)" -ForegroundColor DarkGray
    Write-Host "Remaining upload size: $remainingBytesText ($remainingBytes bytes)" -ForegroundColor Yellow
    Write-Host ""

    Show-UploadProgress
}

# =========================
# Pause / Resume / Stop controls
# =========================
function Show-ControlHelp {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "SCRIPT CONTROLS" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "P = Pause  | Pauses after the current file and waits for a command." -ForegroundColor Yellow
    Write-Host "R = Resume | Resumes processing after Pause." -ForegroundColor Green
    Write-Host "S = Stop   | Saves logs/state and exits so you can continue another day." -ForegroundColor Red
    Write-Host ""
    Write-Host "Note: commands are checked between files. The upload currently in progress is not interrupted midway." -ForegroundColor Gray
    Write-Host "State matrix path: $($Global:PersistentMatrixPath)" -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
}
function Read-ControlKeyIfAvailable {
    try {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            return $key.Key.ToString().ToUpperInvariant()
        }
    }
    catch {
        # Some hosts do not support KeyAvailable. In that case, controls are ignored.
    }
    return ""
}

function Test-ControlCommand {
    param (
        [switch]$AllowResumeOnly
    )

    $key = Read-ControlKeyIfAvailable

    if ([string]::IsNullOrWhiteSpace($key)) {
        return
    }

    switch ($key) {
        "P" {
            if (-not $AllowResumeOnly -and -not $Global:IsPaused) {
                $Global:IsPaused = $true
                Write-Host ""
                Write-Host "PAUSE requested. The script will pause before the next file." -ForegroundColor Yellow
            }
        }
        "R" {
            if ($Global:IsPaused) {
                $Global:IsPaused = $false
                Write-Host ""
                Write-Host "RESUME requested. Processing resumed." -ForegroundColor Green
            }
        }
        "S" {
            $Global:StopRequested = $true
            Write-Host ""
            Write-Host "STOP requested. Saving logs and state before exiting..." -ForegroundColor Red
        }
    }
}

function Wait-IfPausedOrStopped {
    Test-ControlCommand

    if ($Global:IsPaused -and -not $Global:StopRequested) {
        Write-Host "PAUSED: Press R to resume or S to stop." -ForegroundColor Yellow
    }

    while ($Global:IsPaused -and -not $Global:StopRequested) {
        Start-Sleep -Milliseconds 500
        Test-ControlCommand -AllowResumeOnly
    }

    if ($Global:StopRequested) {
        Save-Logs
        Write-Host "Process safely interrupted. Run the script again to continue from the next unprocessed file." -ForegroundColor Yellow
        Read-Host "Press Enter to close"
        exit
    }
}
function Add-ProcessedStateImmediately {
    param (
        [int]$ExcelRow,
        [string]$Description,
        [string]$SourcePath,
        [string]$DestinationUrl,
        [string]$Message = "Item processed successfully."
    )

    Add-MatrixEntry -ExcelRow $ExcelRow -Description $Description -SourcePath $SourcePath -DestinationUrl $DestinationUrl -Status "PROCESSED" -Message $Message
    Add-PersistentMatrixEntry -Description $Description -SourcePath $SourcePath -DestinationUrl $DestinationUrl -Status "PROCESSED" -Message $Message
    Save-PersistentMatrix
}

function Add-ErrorStateImmediately {
    param (
        [int]$ExcelRow,
        [string]$Description,
        [string]$SourcePath,
        [string]$DestinationUrl,
        [string]$Message
    )

    Add-MatrixEntry -ExcelRow $ExcelRow -Description $Description -SourcePath $SourcePath -DestinationUrl $DestinationUrl -Status "ERROR" -Message $Message
    Save-PersistentMatrix
}

# =========================
# Matrix functions
# =========================
function Add-MatrixEntry {
    param (
        [int]$ExcelRow,
        [string]$Description,
        [string]$SourcePath,
        [string]$DestinationUrl,
        [string]$Status,
        [string]$Message
    )
    $Global:ProcessingMatrix += [PSCustomObject]@{
        Timestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ExcelRow       = $ExcelRow
        Description    = $Description
        SourcePath     = (Normalize-LocalPath -PathValue $SourcePath)
        DestinationUrl = $DestinationUrl
        DestinationKey = (Normalize-DestinationKey -DestinationValue $DestinationUrl)
        Status         = $Status
        Message        = $Message
    }
}
function Initialize-PersistentMatrix {
    $matrixPathToLoad = $Global:PersistentMatrixPath

    if (-not (Test-Path -LiteralPath $matrixPathToLoad)) {
        if (Test-Path -LiteralPath $Global:LegacyPersistentMatrixPath) {
            $matrixPathToLoad = $Global:LegacyPersistentMatrixPath
            Write-Host "Legacy state matrix found and loaded: $matrixPathToLoad" -ForegroundColor DarkGray
        }
        else {
            $Global:PersistentMatrix = @()
            return
        }
    }
    try {
        $loaded = Import-Csv -Path $matrixPathToLoad -Encoding UTF8
        if ($null -eq $loaded) {
            $Global:PersistentMatrix = @()
        }
        else {
            $Global:PersistentMatrix = @($loaded)
        }
    }
    catch {
        $Global:PersistentMatrix = @()
    }
}
function Save-PersistentMatrix {
    $Global:PersistentMatrix | Export-Csv -Path $Global:PersistentMatrixPath -NoTypeInformation -Encoding UTF8
}
function Add-PersistentMatrixEntry {
    param (
        [string]$Description,
        [string]$SourcePath,
        [string]$DestinationUrl,
        [string]$Status,
        [string]$Message
    )

    $normalizedSource = Normalize-LocalPath -PathValue $SourcePath
    $destinationKey = Normalize-DestinationKey -DestinationValue $DestinationUrl

    $existing = $Global:PersistentMatrix | Where-Object {
        $_.SourcePath -eq $normalizedSource -and
        $_.DestinationKey -eq $destinationKey
    } | Select-Object -First 1

    if ($null -ne $existing) {
        $existing.Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        $existing.Description = $Description
        $existing.DestinationUrl = $DestinationUrl
        $existing.Status = $Status
        $existing.Message = $Message
    }
    else {
        $Global:PersistentMatrix += [PSCustomObject]@{
            Timestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            Description    = $Description
            SourcePath     = $normalizedSource
            DestinationUrl = $DestinationUrl
            DestinationKey = $destinationKey
            Status         = $Status
            Message        = $Message
        }
    }
}
function Get-ProcessedRootsForDestination {
    param (
        [string]$DestinationKey
    )
    return $Global:ProcessingMatrix | Where-Object {
        $_.DestinationKey -eq $DestinationKey -and $_.Status -in @("PROCESSED", "PROCESSADO")
    }
}
function Test-RowAlreadyCoveredByCurrentExecution {
    param (
        [string]$SourcePath,
        [string]$DestinationKey
    )
    $processedEntries = Get-ProcessedRootsForDestination -DestinationKey $DestinationKey
    foreach ($entry in $processedEntries) {
        if (Test-IsSameOrChildPath -CandidatePath $SourcePath -RootPath $entry.SourcePath) {
            return $true
        }
    }
    return $false
}
function Test-ItemCoveredByCurrentExecution {
    param (
        [string]$ItemPath,
        [string]$DestinationKey
    )
    $processedEntries = Get-ProcessedRootsForDestination -DestinationKey $DestinationKey
    foreach ($entry in $processedEntries) {
        if (Test-IsSameOrChildPath -CandidatePath $ItemPath -RootPath $entry.SourcePath) {
            return $true
        }
    }
    return $false
}
function Test-PathPreviouslyProcessed {
    param (
        [string]$SourcePath,
        [string]$DestinationKey
    )
    $normalizedSource = Normalize-LocalPath -PathValue $SourcePath
    $match = $Global:PersistentMatrix | Where-Object {
        $_.DestinationKey -eq $DestinationKey -and
        $_.Status -in @("PROCESSED", "PROCESSADO") -and
        $_.SourcePath -eq $normalizedSource
    } | Select-Object -First 1
    return ($null -ne $match)
}

function Test-ItemPreviouslyProcessed {
    param (
        [string]$ItemPath,
        [string]$DestinationKey
    )
    $normalizedItem = Normalize-LocalPath -PathValue $ItemPath
    $match = $Global:PersistentMatrix | Where-Object {
        $_.DestinationKey -eq $DestinationKey -and
        $_.Status -in @("PROCESSED", "PROCESSADO") -and
        $_.SourcePath -eq $normalizedItem
    } | Select-Object -First 1
    return ($null -ne $match)
}
# =========================
# Log functions
# =========================
function Initialize-LogSystem {
    if (-not (Test-Path -LiteralPath $Global:LogsFolder)) {
        New-Item -ItemType Directory -Path $Global:LogsFolder | Out-Null
    }
}
function Add-LogEntry {
    param (
        [string]$Status,
        [string]$Type,
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$Message,
        [int]$ExcelRow = 0,
        [string]$Description = ""
    )
    $Global:LogEntries += [PSCustomObject]@{
        Timestamp       = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ExcelRow        = $ExcelRow
        Description     = $Description
        Status          = $Status
        Type            = $Type
        SourcePath      = $SourcePath
        DestinationPath = $DestinationPath
        Message         = $Message
    }
}
function Save-Logs {
    $csvPath = Join-Path -Path $Global:LogsFolder -ChildPath ("MigrationLog_" + $Global:RunTimestamp + ".csv")
    $txtPath = Join-Path -Path $Global:LogsFolder -ChildPath ("MigrationSummary_" + $Global:RunTimestamp + ".txt")
    $matrixPath = Join-Path -Path $Global:LogsFolder -ChildPath ("ProcessingMatrix_" + $Global:RunTimestamp + ".csv")
    $Global:LogEntries | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    $Global:ProcessingMatrix | Export-Csv -Path $matrixPath -NoTypeInformation -Encoding UTF8
    Save-PersistentMatrix
    $successCount = ($Global:LogEntries | Where-Object { $_.Status -in @("SUCCESS", "SUCESSO") }).Count
    $errorCount   = ($Global:LogEntries | Where-Object { $_.Status -in @("ERROR", "ERRO") }).Count
    $skipCount    = ($Global:LogEntries | Where-Object { $_.Status -in @("SKIPPED", "IGNORADO") }).Count
    $summary = @()
    $summary += "============================================================"
    $summary += "MIGRATION SUMMARY"
    $summary += "============================================================"
    $summary += "Date/Time : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $summary += "CSV Log   : $csvPath"
    $summary += "CSV Matrix: $matrixPath"
    $summary += "State     : $($Global:PersistentMatrixPath)"
    $summary += ""
    $summary += "Totals:"
    $summary += "  Successes: $successCount"
    $summary += "  Skipped  : $skipCount"
    $summary += "  Errors   : $errorCount"
    $summary += ""
    $summary += "Details:"
    $summary += ""
    foreach ($item in $Global:LogEntries) {
        $summary += "[$($item.Timestamp)] [Row $($item.ExcelRow)] [$($item.Status)] [$($item.Type)]"
        $summary += "Description: $($item.Description)"
        $summary += "Source     : $($item.SourcePath)"
        $summary += "Destination: $($item.DestinationPath)"
        $summary += "Message    : $($item.Message)"
        $summary += "------------------------------------------------------------"
    }
    $summary | Out-File -FilePath $txtPath -Encoding UTF8
    Write-Host "Logs saved at:" -ForegroundColor Green
    Write-Host "  CSV   : $csvPath" -ForegroundColor Gray
    Write-Host "  TXT   : $txtPath" -ForegroundColor Gray
    Write-Host "  Matrix: $matrixPath" -ForegroundColor Gray
}
# =========================
# Module loading
# =========================
function Load-LocalModules {
    $scriptBase = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptBase)) {
        $scriptBase = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $localModulesPath = Join-Path -Path $scriptBase -ChildPath "Modules"
    $modules = @("PnP.PowerShell", "ImportExcel")
    if (-not (Test-Path -LiteralPath $localModulesPath -PathType Container)) {
        Write-Host "ERROR: 'Modules' folder not found at ${localModulesPath}" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        exit
    }
    Get-ChildItem -LiteralPath $localModulesPath -Recurse -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue
    $env:PSModulePath = $localModulesPath + ";" + $env:PSModulePath
    foreach ($moduleName in $modules) {
        try {
            Import-Module $moduleName -ErrorAction Stop -Force
        }
        catch {
            $moduleManifest = Get-ChildItem -LiteralPath $localModulesPath -Recurse -Filter ($moduleName + ".psd1") -ErrorAction SilentlyContinue | Sort-Object FullName | Select-Object -First 1
            if ($null -ne $moduleManifest) {
                try {
                    Import-Module $moduleManifest.FullName -ErrorAction Stop -Force
                }
                catch {
                    Write-Host "CRITICAL ERROR loading module ${moduleName}: $($_.Exception.Message)" -ForegroundColor Red
                    Read-Host "Press Enter to continue"
                    exit
                }
            }
            else {
                Write-Host "ERROR: Manifest file for ${moduleName} not found in ${localModulesPath}" -ForegroundColor Red
                Read-Host "Press Enter to continue"
                exit
            }
        }
    }
    if (-not (Get-Command "Connect-PnPOnline" -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: The command 'Connect-PnPOnline' was not recognized after loading the modules." -ForegroundColor Red
        Read-Host "Press Enter to continue"
        exit
    }
}
# =========================
# SharePoint helpers
# =========================
function Ensure-SPOFolder {
    param (
        [string]$SiteRelativeUrl
    )
    $decodedPath = [System.Uri]::UnescapeDataString($SiteRelativeUrl)
    $parts = $decodedPath.Split("/")
    $currentPath = ""
    foreach ($part in $parts) {
        if (-not [string]::IsNullOrEmpty($part)) {
            if ($currentPath -eq "") {
                $currentPath = $part
            }
            else {
                $currentPath = "$currentPath/$part"
            }
            Resolve-PnPFolder -SiteRelativePath $currentPath | Out-Null
        }
    }
}
function Connect-ToSPSite {
    param (
        [string]$SiteUrl
    )
    if ($Global:CurrentSPSiteUrl -ne $SiteUrl) {
        try {
            Connect-PnPOnline -Url $SiteUrl -UseWebLogin
            $Global:CurrentSPSiteUrl = $SiteUrl
        }
        catch {
            Write-Host "Error connecting to SharePoint: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
}
function Resolve-TargetInfo {
    param (
        [string]$TargetLibraryUrl
    )
    if ($TargetLibraryUrl -notlike "*/sites/*") {
        throw "Invalid destination URL: $TargetLibraryUrl"
    }
    $decodedUrl = [System.Uri]::UnescapeDataString($TargetLibraryUrl)
    $siteParts = $decodedUrl -split "/sites/"
    $siteBase = $siteParts[0]
    $siteName = $siteParts[1].Split("/")[0]
    $siteUrl = "$siteBase/sites/$siteName"
    $targetLibraryPath = $decodedUrl.Replace($siteUrl, "").Trim("/")
    if ($targetLibraryPath -match "/Forms/") {
        $targetLibraryPath = $targetLibraryPath.Substring(0, $targetLibraryPath.IndexOf("/Forms/"))
    }
    return [PSCustomObject]@{
        SiteUrl           = $siteUrl
        TargetLibraryPath = $targetLibraryPath
    }
}
function Get-SPOFolderFilesMap {
    param (
        [string]$FolderSiteRelativeUrl
    )
    $map = @{}
    try {
        $decodedFolder = [System.Uri]::UnescapeDataString($FolderSiteRelativeUrl)
        $existingFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $decodedFolder -ItemType File -ErrorAction Stop
        foreach ($file in $existingFiles) {
            if ($null -ne $file.Name -and -not [string]::IsNullOrWhiteSpace($file.Name)) {
                $map[$file.Name.ToLowerInvariant()] = $true
            }
        }
    }
    catch {
    }
    return $map
}
function Get-UniqueSPOFileName {
    param (
        [string]$FolderSiteRelativeUrl,
        [string]$OriginalFileName
    )
    if ([string]::IsNullOrWhiteSpace($OriginalFileName)) {
        return ""
    }
    $decodedFolder = [System.Uri]::UnescapeDataString($FolderSiteRelativeUrl)
    $folderFilesMap = Get-SPOFolderFilesMap -FolderSiteRelativeUrl $decodedFolder
    if (-not $folderFilesMap.ContainsKey($OriginalFileName.ToLowerInvariant())) {
        return $OriginalFileName
    }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalFileName)
    $extension = [System.IO.Path]::GetExtension($OriginalFileName)
    $counter = 1
    do {
        $candidate = "${baseName}_${counter}${extension}"
        $exists = $folderFilesMap.ContainsKey($candidate.ToLowerInvariant())
        $counter++
    } while ($exists)
    return $candidate
}
# =========================
# File upload
# =========================
function Upload-FileToSharePoint {
    param (
        [string]$SourceFilePath,
        [string]$TargetLibraryUrl,
        [int]$ExcelRow,
        [string]$Description
    )
    $fileName = Split-Path -Path $SourceFilePath -Leaf
    try {
        Show-UploadProgress -CurrentItem $fileName
        $targetInfo = Resolve-TargetInfo -TargetLibraryUrl $TargetLibraryUrl
        $siteUrl = $targetInfo.SiteUrl
        $targetLibraryPath = $targetInfo.TargetLibraryPath
        Connect-ToSPSite -SiteUrl $siteUrl
        Ensure-SPOFolder -SiteRelativeUrl $targetLibraryPath
        $finalFileName = Get-UniqueSPOFileName -FolderSiteRelativeUrl $targetLibraryPath -OriginalFileName $fileName
        if ([string]::IsNullOrWhiteSpace($finalFileName)) {
            $finalFileName = $fileName
        }
        if ($finalFileName -eq $fileName) {
            Add-PnPFile -Path $SourceFilePath -Folder $targetLibraryPath -ErrorAction Stop
            Write-Host "Uploaded: $fileName" -ForegroundColor Green
            $msg = "File uploaded successfully."
        }
        else {
            Add-PnPFile -Path $SourceFilePath -Folder $targetLibraryPath -NewFileName $finalFileName -ErrorAction Stop
            Write-Host "Renamed: $fileName -> $finalFileName" -ForegroundColor Cyan
            $msg = "File uploaded with duplicate-based renamed filename: $finalFileName"
        }
        Add-LogEntry -Status "SUCCESS" -Type "FILE" -SourcePath $SourceFilePath -DestinationPath ($targetLibraryPath + "/" + $finalFileName) -Message $msg -ExcelRow $ExcelRow -Description $Description
        Add-ProcessedStateImmediately -ExcelRow $ExcelRow -Description $Description -SourcePath $SourceFilePath -DestinationUrl $TargetLibraryUrl -Message $msg
        Mark-ProgressItemComplete -SourcePath $SourceFilePath -DestinationUrl $TargetLibraryUrl -CurrentItem $fileName
        return $true
    }
    catch {
        $msg = $_.Exception.Message
        Write-Host "Error: $fileName" -ForegroundColor Red
        Add-LogEntry -Status "ERROR" -Type "FILE" -SourcePath $SourceFilePath -DestinationPath $TargetLibraryUrl -Message $msg -ExcelRow $ExcelRow -Description $Description
        return $false
    }
}
# =========================
# Folder upload
# =========================
function Upload-FolderContentsToSharePoint {
    param (
        [string]$SourceFolderPath,
        [string]$TargetLibraryUrl,
        [int]$ExcelRow,
        [string]$Description
    )
    $hadAnyAction = $false
    try {
        $targetInfo = Resolve-TargetInfo -TargetLibraryUrl $TargetLibraryUrl
        $siteUrl = $targetInfo.SiteUrl
        $targetLibraryPath = $targetInfo.TargetLibraryPath
        $destinationKey = Normalize-DestinationKey -DestinationValue $TargetLibraryUrl
        Connect-ToSPSite -SiteUrl $siteUrl
        $sourceRootFolderName = Split-Path -Path $SourceFolderPath -Leaf
        $baseTargetFolder = "${targetLibraryPath}/${sourceRootFolderName}"
        Ensure-SPOFolder -SiteRelativeUrl $baseTargetFolder
        Write-Host "Importing folder: $sourceRootFolderName" -ForegroundColor Cyan
        Add-LogEntry -Status "SUCCESS" -Type "FOLDER" -SourcePath $SourceFolderPath -DestinationPath $baseTargetFolder -Message "Root folder created/validated in SharePoint." -ExcelRow $ExcelRow -Description $Description
        $hadAnyAction = $true
        $items = Get-ChildItem -Path $SourceFolderPath -Recurse
        foreach ($item in $items) {
            Wait-IfPausedOrStopped
            $fullPath = $item.FullName

            if ((-not $item.PSIsContainer) -and (Test-ItemPreviouslyProcessed -ItemPath $fullPath -DestinationKey $destinationKey)) {
                Write-Host "Skipped previously processed: $($item.Name)" -ForegroundColor DarkGray
                continue
            }

            if (Test-ItemCoveredByCurrentExecution -ItemPath $fullPath -DestinationKey $destinationKey) {
                if (-not $item.PSIsContainer) {
                    Write-Host "Skipped: $($item.Name)" -ForegroundColor Yellow
                }
                Add-LogEntry -Status "SKIPPED" -Type $(if ($item.PSIsContainer) { "FOLDER" } else { "FILE" }) -SourcePath $fullPath -DestinationPath $TargetLibraryUrl -Message "Skipped because it was already covered by a previous item in this same execution." -ExcelRow $ExcelRow -Description $Description
                continue
            }
            $relPath = $fullPath.Substring($SourceFolderPath.Length).TrimStart("\")
            $relPathWeb = $relPath.Replace("\", "/")
            if ($item.PSIsContainer) {
                $targetFolder = "${baseTargetFolder}/${relPathWeb}"
                Ensure-SPOFolder -SiteRelativeUrl $targetFolder
                Add-LogEntry -Status "SUCCESS" -Type "FOLDER" -SourcePath $fullPath -DestinationPath $targetFolder -Message "Folder validated/created in SharePoint." -ExcelRow $ExcelRow -Description $Description
                $hadAnyAction = $true
            }
            else {
                $relativeParent = Split-Path -Path $relPathWeb -Parent
                if ([string]::IsNullOrWhiteSpace($relativeParent) -or $relativeParent -eq ".") {
                    $folderOfFile = $baseTargetFolder
                }
                else {
                    $folderOfFile = "${baseTargetFolder}/${relativeParent}"
                }
                Ensure-SPOFolder -SiteRelativeUrl $folderOfFile
                $finalFileName = Get-UniqueSPOFileName -FolderSiteRelativeUrl $folderOfFile -OriginalFileName $item.Name
                if ([string]::IsNullOrWhiteSpace($finalFileName)) {
                    $finalFileName = $item.Name
                }
                try {
                    Show-UploadProgress -CurrentItem $item.Name
                    if ($finalFileName -eq $item.Name) {
                        Add-PnPFile -Path $fullPath -Folder $folderOfFile -ErrorAction Stop
                        Write-Host "Uploaded: $($item.Name)" -ForegroundColor Green
                        $msg = "File uploaded successfully."
                    }
                    else {
                        Add-PnPFile -Path $fullPath -Folder $folderOfFile -NewFileName $finalFileName -ErrorAction Stop
                        Write-Host "Renamed: $($item.Name) -> $finalFileName" -ForegroundColor Cyan
                        $msg = "File uploaded with duplicate-based renamed filename: $finalFileName"
                    }
                    Add-LogEntry -Status "SUCCESS" -Type "FILE" -SourcePath $fullPath -DestinationPath ($folderOfFile + "/" + $finalFileName) -Message $msg -ExcelRow $ExcelRow -Description $Description
                    Add-ProcessedStateImmediately -ExcelRow $ExcelRow -Description $Description -SourcePath $fullPath -DestinationUrl $TargetLibraryUrl -Message $msg
                    Mark-ProgressItemComplete -SourcePath $fullPath -DestinationUrl $TargetLibraryUrl -CurrentItem $item.Name
                    $hadAnyAction = $true
                }
                catch {
                    $msg = $_.Exception.Message
                    Write-Host "Error: $($item.Name)" -ForegroundColor Red
                    Add-LogEntry -Status "ERROR" -Type "FILE" -SourcePath $fullPath -DestinationPath ($folderOfFile + "/" + $finalFileName) -Message $msg -ExcelRow $ExcelRow -Description $Description
                }
            }
        }
        return $hadAnyAction
    }
    catch {
        $msg = $_.Exception.Message
        Write-Host "Error while processing '${Description}': $msg" -ForegroundColor Red
        Add-LogEntry -Status "ERROR" -Type "FOLDER" -SourcePath $SourceFolderPath -DestinationPath $TargetLibraryUrl -Message $msg -ExcelRow $ExcelRow -Description $Description
        return $false
    }
}
# =========================
# Script start
# =========================
Initialize-LogSystem
Initialize-PersistentMatrix
Show-ControlHelp
Load-LocalModules
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$OpenFileDialog.Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*"
$OpenFileDialog.Title = "Select the Excel file"
$OpenFileDialog.Multiselect = $false
if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $ExcelFilePath = $OpenFileDialog.FileName
}
else {
    Write-Host "Excel file selection was cancelled." -ForegroundColor Red
    Add-LogEntry -Status "ERROR" -Type "PROCESS" -SourcePath "" -DestinationPath "" -Message "Excel file selection was cancelled by the user." -ExcelRow 0 -Description ""
    Save-Logs
    exit
}
Write-Host "Reading Excel file..." -ForegroundColor Cyan
Initialize-DestinationLibraryMap -ExcelFilePath $ExcelFilePath
$migrationData = Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName
Write-Host "Rows loaded from worksheet '$SheetName': $($migrationData.Count)" -ForegroundColor Cyan

Initialize-UploadProgress -MigrationData $migrationData

$excelRowNumber = 2
foreach ($row in $migrationData) {
    Wait-IfPausedOrStopped
    $description = $row."Description"
    $sourceFolderRaw = Resolve-SourceFolderFromRow -Row $row
    $destinationLibraryUrl = Resolve-DestinationLibraryUrlFromRow -Row $row
    $destinationLibraryValue = $row."Destination Library"
    $destinationFolderRaw = $row."Destination Folder"
    $toBeUpload = $row."To Be Uploaded?"
    # Skip rows where "To Be Uploaded?" is not "Yes"
    if ($toBeUpload -ne "Yes") {
        Write-Host "Skipped (To Be Uploaded? = '$toBeUpload'): Row $excelRowNumber" -ForegroundColor DarkGray
        Add-LogEntry -Status "SKIPPED" -Type "ROW" -SourcePath $sourceFolderRaw -DestinationPath $destinationLibraryUrl -Message "Row skipped because 'To Be Uploaded?' is not 'Yes' (value: '$toBeUpload')." -ExcelRow $excelRowNumber -Description $description
        $excelRowNumber++
        continue
    }
    if ([string]::IsNullOrWhiteSpace($sourceFolderRaw)) {
        Write-Host "Skipped empty source path in Source Folder / UNC fallback columns: Row $excelRowNumber" -ForegroundColor DarkGray
        Add-LogEntry -Status "SKIPPED" -Type "ROW" -SourcePath "" -DestinationPath $destinationLibraryUrl -Message "Row skipped because no source path was found in 'Source Folder' or UNC fallback columns." -ExcelRow $excelRowNumber -Description $description
        $excelRowNumber++
        continue
    }

    if ([string]::IsNullOrWhiteSpace($destinationLibraryUrl)) {
        $msg = "Destination library could not be resolved. Fill 'Destination Library URL' with the full URL, or fill 'Destination Library' with a name that exists in worksheet 'Lists'. Current value: '$destinationLibraryValue'."
        Write-Host "Error in row ${excelRowNumber}: $msg" -ForegroundColor Red
        Add-LogEntry -Status "ERROR" -Type "VALIDATION" -SourcePath $sourceFolderRaw -DestinationPath $destinationLibraryValue -Message $msg -ExcelRow $excelRowNumber -Description $description
        Add-MatrixEntry -ExcelRow $excelRowNumber -Description $description -SourcePath $sourceFolderRaw -DestinationUrl $destinationLibraryValue -Status "ERROR" -Message $msg
        $excelRowNumber++
        continue
    }
    $sourceFolder = Normalize-LocalPath -PathValue $sourceFolderRaw
    $subFolder = Normalize-SubFolderValue -SubFolderValue $destinationFolderRaw
    $finalDestinationFolder = Get-FinalDestinationUrl -DestinationFolder $destinationLibraryUrl -SubFolder $subFolder
    Write-Host "Destination library URL: $destinationLibraryUrl" -ForegroundColor DarkCyan
    Write-Host "Destination folder/subfolder: $subFolder" -ForegroundColor DarkCyan
    Write-Host "Final destination: $finalDestinationFolder" -ForegroundColor DarkCyan
    $destinationKey = Normalize-DestinationKey -DestinationValue $finalDestinationFolder
    $displayName = Get-DisplayName -Description $description -SourcePath $sourceFolder
 
    # Validate optional Destination Folder. Empty is allowed. If filled, it must match one of the known SharePoint subfolders.
    if (-not [string]::IsNullOrWhiteSpace($subFolder)) {
        if ($Global:ValidSubFolders -notcontains $subFolder) {
            $msg = "Invalid Destination Folder value: '$destinationFolderRaw'. Allowed values: $($Global:ValidSubFolders -join ', ')"
            Write-Host "Error in row ${excelRowNumber}: $msg" -ForegroundColor Red
            Add-LogEntry -Status "ERROR" -Type "VALIDATION" -SourcePath $sourceFolder -DestinationPath $destinationLibraryUrl -Message $msg -ExcelRow $excelRowNumber -Description $displayName
            Add-MatrixEntry -ExcelRow $excelRowNumber -Description $displayName -SourcePath $sourceFolder -DestinationUrl $destinationLibraryUrl -Status "ERROR" -Message $msg
            $excelRowNumber++
            continue
        }
    }
    if (-not (Test-Path -Path $sourceFolder)) {
        $msg = "Source path not found."
        Write-Host "The source path for '${displayName}' was not found." -ForegroundColor Yellow
        Add-LogEntry -Status "ERROR" -Type "PATH" -SourcePath $sourceFolder -DestinationPath $finalDestinationFolder -Message $msg -ExcelRow $excelRowNumber -Description $displayName
        Add-MatrixEntry -ExcelRow $excelRowNumber -Description $displayName -SourcePath $sourceFolder -DestinationUrl $finalDestinationFolder -Status "ERROR" -Message $msg
        $excelRowNumber++
        continue
    }
    if (Test-RowAlreadyCoveredByCurrentExecution -SourcePath $sourceFolder -DestinationKey $destinationKey) {
        $msg = "This item is already included in a previous upload to the same destination in this execution."
        Write-Host "Skipped: $displayName (already uploaded previously)" -ForegroundColor Yellow
        Add-LogEntry -Status "SKIPPED" -Type "ROW" -SourcePath $sourceFolder -DestinationPath $finalDestinationFolder -Message $msg -ExcelRow $excelRowNumber -Description $displayName
        Add-MatrixEntry -ExcelRow $excelRowNumber -Description $displayName -SourcePath $sourceFolder -DestinationUrl $finalDestinationFolder -Status "SKIPPED" -Message $msg
        $excelRowNumber++
        continue
    }
    if ((Get-Item $sourceFolder).PSIsContainer) {
        Write-Host "Processing folder: ${displayName}..." -ForegroundColor Cyan
        Write-Host "Previously processed files inside this folder will be skipped automatically." -ForegroundColor DarkGray
    }
    else {
        if (Test-PathPreviouslyProcessed -SourcePath $sourceFolder -DestinationKey $destinationKey) {
            Write-Host "Skipped previously processed file: ${displayName}" -ForegroundColor DarkGray
            $excelRowNumber++
            continue
        }
        else {
            Write-Host "Processing file: ${displayName}..." -ForegroundColor Cyan
        }
    }
    $processed = $false
    if ((Get-Item $sourceFolder).PSIsContainer) {
        $processed = Upload-FolderContentsToSharePoint -SourceFolderPath $sourceFolder -TargetLibraryUrl $finalDestinationFolder -ExcelRow $excelRowNumber -Description $displayName
    }
    else {
        $processed = Upload-FileToSharePoint -SourceFilePath $sourceFolder -TargetLibraryUrl $finalDestinationFolder -ExcelRow $excelRowNumber -Description $displayName
    }
    if ($processed) {
        if ((Get-Item $sourceFolder).PSIsContainer) {
            Add-ProcessedStateImmediately -ExcelRow $excelRowNumber -Description $displayName -SourcePath $sourceFolder -DestinationUrl $finalDestinationFolder -Message "Folder processed. Files inside were saved individually in the state matrix."
        }
    }
    else {
        Add-ErrorStateImmediately -ExcelRow $excelRowNumber -Description $displayName -SourcePath $sourceFolder -DestinationUrl $finalDestinationFolder -Message "Item did not complete processing successfully."
    }
    $excelRowNumber++
}
 
Show-UploadProgress
Write-Progress -Id 1 -Activity "SharePoint upload progress" -Completed
Save-Logs
Write-Host "Process completed." -ForegroundColor Green
Read-Host "Press Enter to continue"
