param(
  [switch]$ForceDeps,
  [switch]$SkipPublish,
  [switch]$GitCommit,
  [switch]$GitPush,
  [string]$GitCommitMessage = '',
  [string]$GitRemote = 'origin',
  [string]$GitBranch = '',
  [string[]]$GitInclude = @()
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$venv = Join-Path $PSScriptRoot '.venv'
$py = 'python'
$reqFile = Join-Path $PSScriptRoot 'requirements.txt'
$depsStamp = Join-Path $venv '.requirements.sha256'

if (-not (Test-Path $venv)) {
  & $py -m venv $venv
}

$python = Join-Path $venv 'Scripts\python.exe'

$reqHash = ''
if (Test-Path $reqFile) {
  $reqHash = (Get-FileHash $reqFile -Algorithm SHA256).Hash
}

$needDeps = $true
if (-not $ForceDeps) {
  $needDeps = -not (Test-Path $depsStamp)
  if (-not $needDeps) {
    try {
      $oldHash = (Get-Content $depsStamp -ErrorAction SilentlyContinue | Select-Object -First 1)
      if ($oldHash -ne $reqHash) {
        $needDeps = $true
      }
    } catch {
      $needDeps = $true
    }
  }
  if (-not $needDeps) {
    try {
      & $python -m PyInstaller --version *> $null
      if ($LASTEXITCODE -ne 0) {
        $needDeps = $true
      }
    } catch {
      $needDeps = $true
    }
  }
}

if ($needDeps) {
  & $python -m pip install -U pip wheel
  if (Test-Path $reqFile) {
    & $python -m pip install -r $reqFile
  }
  & $python -m pip install pyinstaller
  Set-Content -Path $depsStamp -Value $reqHash -Encoding ASCII
} else {
  Write-Host 'Dependencies unchanged; skip pip install.' -ForegroundColor Yellow
}

function Get-AppConstant {
  param(
    [string]$Name,
    [string]$Default = 'unknown'
  )

  try {
    $pattern = '^{0}\s*=\s*"([^"]+)"' -f [regex]::Escape($Name)
    $m = Select-String -Path (Join-Path $PSScriptRoot 'sow_merge_tool.py') -Pattern $pattern
    if ($m -and $m.Matches.Count -gt 0) {
      return $m.Matches[0].Groups[1].Value
    }
  } catch {
  }
  return $Default
}

function Invoke-Git {
  param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Args
  )

  & git @Args
  if ($LASTEXITCODE -ne 0) {
    throw ("git {0} failed with exit code {1}" -f ($Args -join ' '), $LASTEXITCODE)
  }
}

function Get-GitText {
  param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Args
  )

  $text = (& git @Args 2>$null) | Out-String
  if ($LASTEXITCODE -ne 0) {
    throw ("git {0} failed with exit code {1}" -f ($Args -join ' '), $LASTEXITCODE)
  }
  return ($text.Trim())
}

$appVersion = Get-AppConstant -Name 'APP_VERSION'
$buildTag = Get-AppConstant -Name 'APP_BUILD_TAG'
$usageStem = (-join @([char]0x4F7F, [char]0x7528, [char]0x8BF4, [char]0x660E))
$commonSoftwareDir = (-join @([char]0x5E38, [char]0x7528, [char]0x8F6F, [char]0x4EF6))

$tmpBuild = Join-Path $PSScriptRoot 'build_tmp'
$tmpDist = Join-Path $PSScriptRoot 'dist_tmp'
$stableDist = Join-Path $PSScriptRoot 'dist'
$archiveDist = Join-Path $stableDist 'archive'
$stableExe = Join-Path $stableDist 'sow_merge_tool.exe'

$releaseDir = Join-Path $PSScriptRoot 'release'
$releaseExe = Join-Path $releaseDir 'sow_merge_tool.exe'
$releaseZip = Join-Path $releaseDir 'sow_merge_tool_release.zip'
$releaseCompatZip = Join-Path $releaseDir 'excel_merge_tool.zip'
$releaseUsageSource = Join-Path $releaseDir 'README.md'
$releaseUsageMd = Join-Path $releaseDir ("{0}.md" -f $usageStem)
$releaseUsageTxt = Join-Path $releaseDir ("{0}.txt" -f $usageStem)
$releaseSha = Join-Path $releaseDir 'SHA256SUMS.txt'
$releaseRegister = Join-Path $releaseDir 'register_tortoisesvn_sow_tool.bat'
$releaseInstall = Join-Path $releaseDir 'install.bat'
$releaseUninstall = Join-Path $releaseDir 'uninstall.bat'
$releaseContextInstall = Join-Path $releaseDir 'install_context_menu.bat'
$releaseContextUninstall = Join-Path $releaseDir 'uninstall_context_menu.bat'
$releaseStage = Join-Path $releaseDir '_package'
$publishDir = Join-Path (Join-Path 'C:\GM15\design\design' $commonSoftwareDir) 'excel_merge_tool'

if (Test-Path $tmpBuild) { Remove-Item -Recurse -Force $tmpBuild }
if (Test-Path $tmpDist) { Remove-Item -Recurse -Force $tmpDist }

& $python -m PyInstaller --noconsole --onefile --hidden-import branch_submit --name sow_merge_tool --distpath $tmpDist --workpath $tmpBuild --specpath $tmpBuild sow_merge_tool.py

$builtExe = Join-Path $tmpDist 'sow_merge_tool.exe'
if (-not (Test-Path $builtExe)) {
  throw "Build failed: $builtExe not found."
}

if (-not (Test-Path $stableDist)) { New-Item -ItemType Directory -Path $stableDist | Out-Null }
if (-not (Test-Path $archiveDist)) { New-Item -ItemType Directory -Path $archiveDist | Out-Null }
if (-not (Test-Path $releaseDir)) { New-Item -ItemType Directory -Path $releaseDir | Out-Null }

Copy-Item -Force $builtExe $stableExe

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$archiveExe = Join-Path $archiveDist ("sow_merge_tool_{0}_{1}.exe" -f $buildTag, $ts)
Copy-Item -Force $builtExe $archiveExe

Copy-Item -Force $builtExe $releaseExe
Copy-Item -Force (Join-Path $PSScriptRoot 'register_tortoisesvn_sow_tool.bat') $releaseRegister
Copy-Item -Force (Join-Path $PSScriptRoot 'install_context_menu.bat') $releaseContextInstall
Copy-Item -Force (Join-Path $PSScriptRoot 'uninstall_context_menu.bat') $releaseContextUninstall

if (Test-Path $releaseUsageSource) {
  Copy-Item -Force $releaseUsageSource $releaseUsageMd
}

$usageTxt = @"
sow_merge_tool release notes ($appVersion / $buildTag)

Supported formats:
- .xlsx
- .xlsm

Main files:
- sow_merge_tool.exe
- sow_merge_tool_release.zip
- $usageStem.md
- SHA256SUMS.txt

Log files:
1) %TEMP%\sow_merge_tool_debug.log
2) %TEMP%\sow_merge_tool_launch_trace.log
3) Working copy root\sow_update.log
"@
Set-Content -Path $releaseUsageTxt -Value $usageTxt -Encoding UTF8

if (Test-Path $releaseStage) { Remove-Item -Recurse -Force $releaseStage }
New-Item -ItemType Directory -Path $releaseStage | Out-Null

$packageFiles = @(
  $releaseExe,
  $releaseRegister,
  $releaseInstall,
  $releaseUninstall,
  $releaseContextInstall,
  $releaseContextUninstall,
  $releaseUsageMd,
  $releaseUsageTxt
)

foreach ($file in $packageFiles) {
  if (Test-Path $file) {
    Copy-Item -Force $file $releaseStage
  }
}

if (Test-Path $releaseZip) { Remove-Item -Force $releaseZip }
Compress-Archive -Path (Join-Path $releaseStage '*') -DestinationPath $releaseZip -Force
Copy-Item -Force $releaseZip $releaseCompatZip

$versionSlug = ($appVersion -replace '[^0-9A-Za-z._-]', '_')
$tagSlug = ($buildTag -replace '[^0-9A-Za-z._-]', '_')
$historyZip = Join-Path $releaseDir ("sow_merge_tool_release_{0}_{1}.zip" -f $versionSlug, $tagSlug)
Copy-Item -Force $releaseZip $historyZip

$hashFiles = @(
  $releaseExe,
  $releaseZip,
  $releaseCompatZip,
  $releaseUsageMd,
  $releaseUsageTxt,
  $releaseRegister,
  $releaseInstall,
  $releaseUninstall,
  $releaseContextInstall,
  $releaseContextUninstall
) | Where-Object { Test-Path $_ }

$hashLines = foreach ($file in $hashFiles) {
  $hash = (Get-FileHash $file -Algorithm SHA256).Hash.ToLowerInvariant()
  '{0} *{1}' -f $hash, (Split-Path $file -Leaf)
}
Set-Content -Path $releaseSha -Value $hashLines -Encoding UTF8

if (-not $SkipPublish) {
  if (-not (Test-Path $publishDir)) {
    New-Item -ItemType Directory -Path $publishDir | Out-Null
  }

  $publishMap = @(
    @{ Source = $releaseExe; Target = (Join-Path $publishDir 'sow_merge_tool.exe') },
    @{ Source = $releaseZip; Target = (Join-Path $publishDir 'sow_merge_tool_release.zip') },
    @{ Source = $releaseCompatZip; Target = (Join-Path $publishDir 'excel_merge_tool.zip') },
    @{ Source = $releaseSha; Target = (Join-Path $publishDir 'SHA256SUMS.txt') },
    @{ Source = $releaseUsageMd; Target = (Join-Path $publishDir ("{0}.md" -f $usageStem)) },
    @{ Source = $releaseUsageTxt; Target = (Join-Path $publishDir ("{0}.txt" -f $usageStem)) },
    @{ Source = $releaseRegister; Target = (Join-Path $publishDir 'register_tortoisesvn_sow_tool.bat') },
    @{ Source = $releaseInstall; Target = (Join-Path $publishDir 'install.bat') },
    @{ Source = $releaseUninstall; Target = (Join-Path $publishDir 'uninstall.bat') },
    @{ Source = $releaseContextInstall; Target = (Join-Path $publishDir 'install_context_menu.bat') },
    @{ Source = $releaseContextUninstall; Target = (Join-Path $publishDir 'uninstall_context_menu.bat') }
  )

  foreach ($item in $publishMap) {
    if (Test-Path $item.Source) {
      Copy-Item -Force $item.Source $item.Target
    }
  }
}

Write-Host "Build complete (stable): $stableExe" -ForegroundColor Green
Write-Host "Archived copy: $archiveExe" -ForegroundColor Green
Write-Host "Release exe: $releaseExe" -ForegroundColor Green
Write-Host "Release zip: $releaseZip" -ForegroundColor Green
Write-Host "SHA256SUMS: $releaseSha" -ForegroundColor Green
if ($SkipPublish) {
  Write-Host 'Publish sync skipped by -SkipPublish.' -ForegroundColor Yellow
} else {
  Write-Host "Published to: $publishDir" -ForegroundColor Green
}

$doGitFlow = $GitCommit -or $GitPush
if ($doGitFlow) {
  $gitRoot = Get-GitText rev-parse --show-toplevel
  if (-not $gitRoot) {
    throw 'Current directory is not a git repository.'
  }

  $currentBranch = $GitBranch
  if (-not $currentBranch) {
    $currentBranch = Get-GitText rev-parse --abbrev-ref HEAD
  }
  if (-not $currentBranch -or $currentBranch -eq 'HEAD') {
    throw 'Detached HEAD is not supported for release git flow. Please checkout a branch first.'
  }
  if (-not $GitInclude -or $GitInclude.Count -eq 0) {
    throw 'Release git flow requires explicit -GitInclude paths; refusing to stage the whole dirty worktree.'
  }

  $statusBefore = (& git status --short) | Out-String
  Write-Host "Git status before staging:" -ForegroundColor Cyan
  if ([string]::IsNullOrWhiteSpace($statusBefore)) {
    Write-Host '  (clean)' -ForegroundColor DarkGray
  } else {
    Write-Host $statusBefore.TrimEnd()
  }

  $gitArgs = @('add', '--') + @($GitInclude)
  Invoke-Git @gitArgs

  & git diff --cached --quiet
  $hasStagedChanges = ($LASTEXITCODE -ne 0)
  if ($hasStagedChanges) {
    $finalCommitMessage = $GitCommitMessage
    if ([string]::IsNullOrWhiteSpace($finalCommitMessage)) {
      $finalCommitMessage = "release: $appVersion ($buildTag)"
    }
    if ($GitCommit -or $GitPush) {
      Invoke-Git commit -m $finalCommitMessage
      Write-Host "Git commit created on branch: $currentBranch" -ForegroundColor Green
    }
  } else {
    Write-Host 'No staged changes detected; skip git commit.' -ForegroundColor Yellow
  }

  if ($GitPush) {
    Invoke-Git push $GitRemote $currentBranch
    Write-Host "Git push complete: $GitRemote/$currentBranch" -ForegroundColor Green
  } else {
    Write-Host 'Git push skipped. Use -GitPush to push this release.' -ForegroundColor Yellow
  }
}
