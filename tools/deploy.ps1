param(
  [string]$PackageDir = '',
  [string]$DeployPath = 'C:\sow_main\excel\excel_merge_tool'
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
$releaseRoot = Join-Path $repo 'artifacts\release'
$source = if ($PackageDir) {
  [IO.Path]::GetFullPath($PackageDir)
} else {
  $candidate = Get-ChildItem -LiteralPath $releaseRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'sow_merge_tool.exe') } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if ($null -eq $candidate) { throw "No packaged release found under $releaseRoot" }
  $candidate.FullName
}
$target = [IO.Path]::GetFullPath($DeployPath)
$sourceExe = Join-Path $source 'sow_merge_tool.exe'
if (-not (Test-Path -LiteralPath $sourceExe)) { throw "Package EXE missing: $sourceExe" }

$running = Get-Process -Name 'sow_merge_tool' -ErrorAction SilentlyContinue
if ($running) { throw 'sow_merge_tool.exe is running; close it before deployment.' }

New-Item -ItemType Directory -Force -Path $target | Out-Null
$managedNames = @(
  'sow_merge_tool.exe', '安装.bat', '卸载.bat', '使用说明.md', 'SHA256SUMS.txt',
  'install_context_menu.bat', 'uninstall_context_menu.bat', 'register_tortoisesvn_sow_tool.bat',
  'install.bat', 'uninstall.bat', 'register_tortoisesvn_excel_merge_tool.bat',
  'restore_tortoisesvn_config_latest.bat', 'README.md', '使用说明.txt',
  'excel_merge_tool.zip', 'sow_merge_tool_release.zip'
)
$backupRoot = Join-Path $target 'backups'
$backup = Join-Path $backupRoot (Get-Date -Format 'yyyyMMdd_HHmmss')
New-Item -ItemType Directory -Force -Path $backup | Out-Null
foreach ($name in $managedNames) {
  $existing = Join-Path $target $name
  if (Test-Path -LiteralPath $existing) {
    Copy-Item -LiteralPath $existing -Destination (Join-Path $backup $name) -Recurse -Force
  }
}

foreach ($name in $managedNames) {
  $existing = Join-Path $target $name
  if (Test-Path -LiteralPath $existing) { Remove-Item -LiteralPath $existing -Recurse -Force }
}
foreach ($name in @('sow_merge_tool.exe','安装.bat','卸载.bat','使用说明.md','SHA256SUMS.txt')) {
  $item = Join-Path $source $name
  if (-not (Test-Path -LiteralPath $item)) { throw "Package file missing: $item" }
  Copy-Item -LiteralPath $item -Destination (Join-Path $target $name) -Force
}

$installedExe = Join-Path $target 'sow_merge_tool.exe'
$expected = (Get-FileHash -LiteralPath $sourceExe -Algorithm SHA256).Hash
$actual = (Get-FileHash -LiteralPath $installedExe -Algorithm SHA256).Hash
if ($expected -ne $actual) { throw "Installed EXE hash mismatch: expected $expected actual $actual" }
$installer = Join-Path $target '安装.bat'
if (Test-Path -LiteralPath $installer) {
  & cmd.exe /c "`"$installer`" /quiet"
  if ($LASTEXITCODE -ne 0) { throw "Context-menu installation failed with exit code $LASTEXITCODE" }
}
Write-Host "已部署到：$target" -ForegroundColor Green
Write-Host "备份目录：$backup" -ForegroundColor Green
Write-Host "EXE SHA256：$actual" -ForegroundColor Green
