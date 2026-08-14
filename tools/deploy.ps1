param(
  [string]$PackageDir = '',
  [string]$DeployPath = 'C:\sow_main\excel\excel_merge_tool'
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
$source = if ($PackageDir) { [IO.Path]::GetFullPath($PackageDir) } else { Join-Path $repo 'artifacts\release\2026-08-14.update77' }
$target = [IO.Path]::GetFullPath($DeployPath)
$sourceExe = Join-Path $source 'sow_merge_tool.exe'
if (-not (Test-Path -LiteralPath $sourceExe)) { throw "Package EXE missing: $sourceExe" }

$running = Get-Process -Name 'sow_merge_tool' -ErrorAction SilentlyContinue
if ($running) { throw 'sow_merge_tool.exe is running; close it before deployment.' }

New-Item -ItemType Directory -Force -Path $target | Out-Null
$backupRoot = Join-Path $repo 'artifacts\backups'
$backup = Join-Path $backupRoot (Get-Date -Format 'yyyyMMdd_HHmmss')
New-Item -ItemType Directory -Force -Path $backup | Out-Null
Get-ChildItem -LiteralPath $target -Force -ErrorAction SilentlyContinue | ForEach-Object {
  Copy-Item -LiteralPath $_.FullName -Destination $backup -Recurse -Force
}

$stage = Join-Path $target '.deploy-staging'
if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $stage | Out-Null
Copy-Item -LiteralPath (Join-Path $source '*') -Destination $stage -Recurse -Force
foreach ($item in Get-ChildItem -LiteralPath $stage -Force) {
  Move-Item -LiteralPath $item.FullName -Destination $target -Force
}
Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue

$installedExe = Join-Path $target 'sow_merge_tool.exe'
$expected = (Get-FileHash -LiteralPath $sourceExe -Algorithm SHA256).Hash
$actual = (Get-FileHash -LiteralPath $installedExe -Algorithm SHA256).Hash
if ($expected -ne $actual) { throw "Installed EXE hash mismatch: expected $expected actual $actual" }
$installer = Join-Path $target 'install_context_menu.bat'
if (Test-Path -LiteralPath $installer) {
  & cmd.exe /c $installer /quiet
  if ($LASTEXITCODE -ne 0) { throw "Context-menu installation failed with exit code $LASTEXITCODE" }
}
Write-Host "Deployed to $target" -ForegroundColor Green
Write-Host "Backup: $backup" -ForegroundColor Green
Write-Host "SHA256: $actual" -ForegroundColor Green
