$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
& (Join-Path $PSScriptRoot 'tools\test.ps1') -Profile Gui
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
