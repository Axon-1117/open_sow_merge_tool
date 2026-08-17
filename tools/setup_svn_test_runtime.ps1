param()

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)

$repo = (Get-Location).Path
$version = '1.14.5'
$archiveName = "Slik-Subversion-$version-x64.zip"
$downloadUrl = "https://sliksvn.com/pub/$archiveName"
$expectedSha256 = '77D4FE02999DDA3BDC3A20E86243AE6EDE99AAF072B4C12B0CDEDB54D88E954A'
$toolRoot = Join-Path $repo ".local\tools\SlikSVN-$version-x64"
$archive = Join-Path (Split-Path -Parent $toolRoot) $archiveName
$msi = Join-Path $toolRoot "Slik-Subversion-$version-x64.msi"
$portable = Join-Path $toolRoot 'portable'
$bin = Join-Path $portable 'PFiles\bin'
$svn = Join-Path $bin 'svn.exe'
$svnadmin = Join-Path $bin 'svnadmin.exe'

if (-not ((Test-Path -LiteralPath $svn) -and (Test-Path -LiteralPath $svnadmin))) {
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $toolRoot) | Out-Null
  if (-not (Test-Path -LiteralPath $archive)) {
    Write-Host "下载无界面 SVN 测试运行时：$downloadUrl"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $archive
  }
  $actualSha256 = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash
  if ($actualSha256 -ne $expectedSha256) {
    throw "SlikSVN 下载包 SHA256 不匹配：expected=$expectedSha256 actual=$actualSha256"
  }
  if (-not (Test-Path -LiteralPath $msi)) {
    New-Item -ItemType Directory -Force -Path $toolRoot | Out-Null
    Expand-Archive -LiteralPath $archive -DestinationPath $toolRoot -Force
  }
  New-Item -ItemType Directory -Force -Path $portable | Out-Null
  $arguments = @('/a', ('"' + $msi + '"'), '/qn', ('TARGETDIR="' + $portable + '"'))
  $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
  if ($process.ExitCode -ne 0) {
    throw "SlikSVN 便携解包失败：msiexec exit=$($process.ExitCode)"
  }
}

if (-not ((Test-Path -LiteralPath $svn) -and (Test-Path -LiteralPath $svnadmin))) {
  throw "无界面 SVN 测试运行时不完整：$bin"
}

$reportedVersion = (& $svn --version --quiet).Trim()
if (-not $reportedVersion.StartsWith($version)) {
  throw "SVN 测试运行时版本不符合预期：$reportedVersion"
}

Write-Host "SVN 测试运行时：$bin ($reportedVersion)" -ForegroundColor Green
Write-Output $bin
