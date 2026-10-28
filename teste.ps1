$Caminho = "COPY_PATH_HERE"

Write-Host "Caminho informado: [$Caminho]"

Write-Host "Existe?"
Test-Path -LiteralPath $Caminho

Write-Host "Item encontrado:"
Get-Item -LiteralPath $Caminho -ErrorAction Continue
