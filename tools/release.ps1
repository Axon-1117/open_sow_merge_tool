param(
  [string]$DeployPath = 'C:\sow_main\excel\excel_merge_tool',
  [string]$Version = '2026-08-17.update88',
  [switch]$SkipDeploy
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
& (Join-Path $repo 'tools\test.ps1') -Profile Fast
if (-not $?) { throw 'Fast gate failed.' }
& (Join-Path $repo 'tools\build.ps1') -Clean
if (-not $?) { throw 'Build failed.' }
& (Join-Path $repo 'tools\package.ps1') -Version $Version
if (-not $?) { throw 'Package failed.' }
if (-not $SkipDeploy) {
  & (Join-Path $repo 'tools\deploy.ps1') -DeployPath $DeployPath
  if (-not $?) { throw 'Deployment failed.' }
}
Write-Host 'Release workflow passed.' -ForegroundColor Green
