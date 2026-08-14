param(
  [ValidateSet('Fast', 'Full', 'Gui', 'Native', 'Adversarial')]
  [string]$Profile = 'Fast',
  [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)

$repo = (Get-Location).Path
$python = Join-Path $repo '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
  throw "Virtual environment not found: $python. Run tools\bootstrap.ps1 first."
}

# Keep synthetic fixtures outside the system temp root. The application uses
# the system temp root as one signal for SVN-generated sidecars; putting test
# fixtures there changes the behavior the tests are meant to observe.
$testRoot = Join-Path $repo 'tmp\test_tmp'
New-Item -ItemType Directory -Force -Path $testRoot | Out-Null
$env:SOW_TEST_TMPDIR = $testRoot
$env:PYTHONUTF8 = '1'
$env:SOW_SKIP_REAL_WC_TESTS = if ($Profile -eq 'Native') { '0' } else { '1' }

function Invoke-PythonFile {
  param([string]$Path)
  $psi = [Diagnostics.ProcessStartInfo]::new()
  $psi.FileName = $python
  $psi.Arguments = '"' + $Path + '"'
  $psi.WorkingDirectory = $repo
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.Environment['SOW_TEST_TMPDIR'] = $testRoot
  $psi.Environment['PYTHONUTF8'] = '1'
  $psi.Environment['SOW_SKIP_REAL_WC_TESTS'] = $env:SOW_SKIP_REAL_WC_TESTS
  $process = [Diagnostics.Process]::new()
  $process.StartInfo = $psi
  [void]$process.Start()
  $watch = [Diagnostics.Stopwatch]::StartNew()
  $finished = $process.WaitForExit($TimeoutSeconds * 1000)
  if (-not $finished) {
    $process.Kill($true)
    $process.WaitForExit()
    throw "Timeout after $TimeoutSeconds seconds: $Path"
  }
  $watch.Stop()
  $stdout = $process.StandardOutput.ReadToEnd()
  $stderr = $process.StandardError.ReadToEnd()
  if ($process.ExitCode -ne 0) {
    throw "Failed ($($process.ExitCode)) $Path`n$stdout`n$stderr"
  }
  Write-Host ("PASS {0} ({1:N2}s)" -f (Split-Path -Leaf $Path), $watch.Elapsed.TotalSeconds) -ForegroundColor Green
}

$testScriptRoot = Join-Path $repo 'tests\legacy'
$smoke = @(Get-ChildItem -LiteralPath $testScriptRoot -File -Filter '_smoke_test*.py' | Sort-Object Name)
if ($Profile -in @('Fast', 'Full', 'Adversarial')) {
  foreach ($test in $smoke) { Invoke-PythonFile $test.FullName }
}

if ($Profile -in @('Full', 'Adversarial')) {
  $pytest = Join-Path $repo '.venv\Scripts\pytest.exe'
  if (Test-Path -LiteralPath $pytest) {
    & $pytest -q
    if ($LASTEXITCODE -ne 0) { throw "pytest failed with exit code $LASTEXITCODE" }
  }
}

if ($Profile -eq 'Gui') {
  & $python (Join-Path $testScriptRoot 'ui_test_scenarios.py')
  if ($LASTEXITCODE -ne 0) { throw "GUI tests failed with exit code $LASTEXITCODE" }
}

if ($Profile -eq 'Native') {
  & $python (Join-Path $testScriptRoot '_gui_self_test_branch_submit_workbench.py')
  if ($LASTEXITCODE -ne 0) { throw "Native GUI test failed with exit code $LASTEXITCODE" }
}

Write-Host "Test profile $Profile passed." -ForegroundColor Green
