$Path = "COPY_PATH_HERE"

$NormalizedPath = [System.IO.Path]::GetFullPath($Path.Trim().Trim('"')).TrimEnd('\')

Write-Host "Original path: [$Path]"
Write-Host "Normalized path: [$NormalizedPath]"

Write-Host "Original path exists:"
Test-Path -LiteralPath $Path

Write-Host "Normalized path exists:"
Test-Path -LiteralPath $NormalizedPath
