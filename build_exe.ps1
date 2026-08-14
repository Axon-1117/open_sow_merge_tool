param(
  [switch]$Clean,
  [switch]$SkipPublish,
  [string]$DeployPath = 'C:\sow_main\excel\excel_merge_tool'
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
& (Join-Path $PSScriptRoot 'tools\build.ps1') -Clean:$Clean
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& (Join-Path $PSScriptRoot 'tools\package.ps1')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
if (-not $SkipPublish) {
  & (Join-Path $PSScriptRoot 'tools\deploy.ps1') -DeployPath $DeployPath
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
