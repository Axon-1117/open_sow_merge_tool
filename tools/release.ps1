param(
  [string]$DeployPath = 'C:\sow_main\excel\excel_merge_tool',
  [switch]$SkipDeploy
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
$repo = (Get-Location).Path
& (Join-Path $repo 'tools\test.ps1') -Profile Fast
if ($LASTEXITCODE -ne 0) { throw 'Fast gate failed.' }
& (Join-Path $repo 'tools\build.ps1') -Clean
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
& (Join-Path $repo 'tools\package.ps1')
if ($LASTEXITCODE -ne 0) { throw 'Package failed.' }
if (-not $SkipDeploy) {
  & (Join-Path $repo 'tools\deploy.ps1') -DeployPath $DeployPath
  if ($LASTEXITCODE -ne 0) { throw 'Deployment failed.' }
}
Write-Host 'Release workflow passed.' -ForegroundColor Green
