param(
  [string]$BuildDir = '',
  [string]$Version = '2026-09-14.update113'
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
$runtimeCandidates = @(Get-ChildItem -LiteralPath (Join-Path $repo '.local\tools') -Directory -Filter 'SlikSVN-*' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
$runtimeBin = if ($runtimeCandidates.Count -gt 0) {
  Join-Path $runtimeCandidates[0].FullName 'portable\PFiles\bin'
} else { '' }
$runtimeFiles = @()
if ($runtimeBin -and (Test-Path -LiteralPath $runtimeBin)) {
  $runtimeFiles = @(Get-ChildItem -LiteralPath $runtimeBin -File | Where-Object {
    $_.Name -eq 'svn.exe' -or $_.Extension -ieq '.dll'
  })
  if ($runtimeFiles.Count -eq 0) { throw "Bundled SVN runtime is empty: $runtimeBin" }
  $runtimeRelease = Join-Path $release 'svn_runtime'
  New-Item -ItemType Directory -Force -Path $runtimeRelease | Out-Null
  foreach ($runtimeFile in $runtimeFiles) {
    Copy-Item -LiteralPath $runtimeFile.FullName -Destination (Join-Path $runtimeRelease $runtimeFile.Name)
  }
}
$releaseFiles = @()
# Resolve owned documentation/scripts by extension and size so this remains
# reliable under legacy Windows PowerShell code pages.
$releaseFiles += @(Get-ChildItem -LiteralPath $repo -File -Filter '*.bat')
$usageGuide = Join-Path $repo '使用说明.md'
if (Test-Path -LiteralPath $usageGuide) {
  $releaseFiles += Get-Item -LiteralPath $usageGuide
} else {
  $markdownFiles = @(Get-ChildItem -LiteralPath $repo -File -Filter '*.md' | Sort-Object Length)
  if ($markdownFiles.Count -ge 2) { $releaseFiles += $markdownFiles[1] }
}
foreach ($file in $releaseFiles) {
  Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $release $file.Name)
}
$hashLines = @()
foreach ($file in @(Get-ChildItem -LiteralPath $release -File -Recurse | Sort-Object FullName)) {
  $relative = [IO.Path]::GetRelativePath($release, $file.FullName).Replace('\','/')
  $fileHash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
  $hashLines += "$fileHash  $relative"
}
$hash = (Get-FileHash -LiteralPath (Join-Path $release 'sow_merge_tool.exe') -Algorithm SHA256).Hash
Set-Content -LiteralPath (Join-Path $release 'SHA256SUMS.txt') -Value $hashLines -Encoding ASCII
$manifest = [ordered]@{
  version = $Version
  package = "sow_merge_tool_$slug.zip"
  directory = $slug
  files = @('sow_merge_tool.exe') + @($releaseFiles.Name) + @('svn_runtime/*') + @('SHA256SUMS.txt')
  sha256 = $hash
  generatedAt = (Get-Date).ToUniversalTime().ToString('o')
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $repo 'artifacts\release\latest.json') -Encoding UTF8
$zip = Join-Path $repo "artifacts\release\sow_merge_tool_$slug.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $release '*') -DestinationPath $zip -Force
Write-Host "Package: $zip" -ForegroundColor Green
Write-Host "SHA256: $hash" -ForegroundColor Green
