param(
  [string]$BuildDir = '',
  [string]$Version = '2026-08-14.update77'
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
$buildRoot = if ($BuildDir) { [IO.Path]::GetFullPath($BuildDir) } else { Join-Path $repo 'artifacts\build' }
$exe = Join-Path $buildRoot 'dist\sow_merge_tool.exe'
if (-not (Test-Path -LiteralPath $exe)) { throw "Build output missing: $exe" }

$slug = $Version -replace '[^A-Za-z0-9._-]', '_'
$release = Join-Path $repo "artifacts\release\$slug"
if (Test-Path -LiteralPath $release) { Remove-Item -LiteralPath $release -Recurse -Force }
New-Item -ItemType Directory -Force -Path $release | Out-Null
Copy-Item -LiteralPath $exe -Destination (Join-Path $release 'sow_merge_tool.exe')
foreach ($file in @('install_context_menu.bat','uninstall_context_menu.bat','register_tortoisesvn_sow_tool.bat')) {
  Copy-Item -LiteralPath (Join-Path $repo $file) -Destination (Join-Path $release $file)
}
Copy-Item -LiteralPath (Join-Path $repo 'README.md') -Destination (Join-Path $release 'README.md')
$hash = (Get-FileHash -LiteralPath (Join-Path $release 'sow_merge_tool.exe') -Algorithm SHA256).Hash
Set-Content -LiteralPath (Join-Path $release 'SHA256SUMS.txt') -Value "$hash  sow_merge_tool.exe" -Encoding ASCII
$zip = Join-Path $repo "artifacts\release\sow_merge_tool_$slug.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $release '*') -DestinationPath $zip -Force
Write-Host "Package: $zip" -ForegroundColor Green
Write-Host "SHA256: $hash" -ForegroundColor Green
