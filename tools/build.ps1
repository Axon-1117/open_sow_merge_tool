param(
  [switch]$Clean,
  [string]$OutputDir = ''
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
$python = Join-Path $repo '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) { throw 'Run tools\bootstrap.ps1 first.' }

$artifactRoot = if ($OutputDir) { [IO.Path]::GetFullPath($OutputDir) } else { Join-Path $repo 'artifacts\build' }
$dist = Join-Path $artifactRoot 'dist'
$work = Join-Path $artifactRoot 'work'
New-Item -ItemType Directory -Force -Path $artifactRoot | Out-Null
if ($Clean) {
  foreach ($path in @($dist, $work)) {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
  }
}

& $python -m PyInstaller --noconfirm --clean --distpath $dist --workpath $work sow_merge_tool.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }
$exe = Join-Path $dist 'sow_merge_tool.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw "Build output missing: $exe" }
$hash = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
Set-Content -LiteralPath (Join-Path $artifactRoot 'sow_merge_tool.exe.sha256') -Value "$hash  sow_merge_tool.exe" -Encoding ASCII
Write-Host "Built $exe" -ForegroundColor Green
Write-Host "SHA256 $hash" -ForegroundColor Green
