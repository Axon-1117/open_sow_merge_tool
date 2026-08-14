param([switch]$Force)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$venv = Join-Path (Get-Location).Path '.venv'
if ($Force -and (Test-Path -LiteralPath $venv)) {
  Remove-Item -LiteralPath $venv -Recurse -Force
}
if (-not (Test-Path -LiteralPath $venv)) {
  python -m venv $venv
}
$python = Join-Path $venv 'Scripts\python.exe'
& $python -m pip install -U pip
& $python -m pip install -e '.[dev]'
if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed." }
Write-Host "Environment ready: $python" -ForegroundColor Green
