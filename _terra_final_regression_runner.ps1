$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$evidenceDir = Join-Path $PSScriptRoot 'openspec\changes\accelerate-large-sheet-comparison-ui\evidence\terra-final-regression'
New-Item -ItemType Directory -Force -Path $evidenceDir | Out-Null
Remove-Item -LiteralPath (Join-Path $evidenceDir 'progress.jsonl'),(Join-Path $evidenceDir 'summary.json') -Force -ErrorAction SilentlyContinue

function Get-DescendantPids([int]$RootPid) {
    $seen = [System.Collections.Generic.HashSet[int]]::new()
    $frontier = [System.Collections.Generic.Queue[int]]::new()
    $frontier.Enqueue($RootPid)
    while ($frontier.Count -gt 0) {
        $parent = $frontier.Dequeue()
        $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$parent" -ErrorAction SilentlyContinue)
        foreach ($child in $children) {
            $childId = [int]$child.ProcessId
            if ($seen.Add($childId)) { $frontier.Enqueue($childId) }
        }
    }
    return @($seen)
}

function Stop-ExactProcessTree([int]$RootPid) {
    $processIds = @(Get-DescendantPids $RootPid) + @($RootPid)
    foreach ($processId in ($processIds | Sort-Object -Descending -Unique)) {
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -ne $process) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }
    }
    return $processIds
}

function Get-Sha256([string]$Path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (($sha.ComputeHash([System.IO.File]::ReadAllBytes($Path)) | ForEach-Object { $_.ToString('x2') }) -join '').ToUpperInvariant()
    } finally {
        $sha.Dispose()
    }
}

$tests = @(
    @{ name = '_smoke_test.py'; timeout = 240 },
    @{ name = '_smoke_test_snapshot_oracle_parity.py'; timeout = 300 },
    @{ name = '_smoke_test_real_snapshot_direct_oracle.py'; timeout = 360 },
    @{ name = '_smoke_test_row_alignment_adversarial.py'; timeout = 240 },
    @{ name = '_smoke_test_column_alignment_matrix.py'; timeout = 240 },
    @{ name = '_smoke_test_2way_row_replay.py'; timeout = 300 },
    @{ name = '_smoke_test_merge_three_way_row_align.py'; timeout = 300 },
    @{ name = '_smoke_test_manual_merge_row_insert.py'; timeout = 360 },
    @{ name = '_smoke_test_row_delete_redo_fidelity.py'; timeout = 360 },
    @{ name = '_smoke_test_cross_workbook_row_style_replay.py'; timeout = 360 },
    @{ name = '_smoke_test_column_native_save_replay.py'; timeout = 360 },
    @{ name = '_smoke_test_operation_overlay.py'; timeout = 300 },
    @{ name = '_fidelity_actual_app_actions.py'; timeout = 420 },
    @{ name = '_fidelity_region_diag.py'; timeout = 300 },
    @{ name = '_smoke_test_large_only_diff_row_insert.py'; timeout = 300 },
    @{ name = '_smoke_test_formula_cache_undo.py'; timeout = 300 },
    @{ name = '_smoke_test_2way_formula_cache_save.py'; timeout = 300 },
    @{ name = '_smoke_test_blank_shared_formula_b_save.py'; timeout = 300 },
    @{ name = '_fidelity_formula_cache_diag.py'; timeout = 300 },
    @{ name = '_smoke_test_save_and_diff_fidelity.py'; timeout = 300 },
    @{ name = '_smoke_test_zip_save_failure_retry.py'; timeout = 300 },
    @{ name = '_smoke_test_native_reopen_failure_retry.py'; timeout = 360 },
    @{ name = '_large_sheet_excel_fidelity_gate.py'; timeout = 600 },
    @{ name = '_smoke_test_xlsm_support.py'; timeout = 300 },
    @{ name = '_smoke_test_svn_merge_role_semantics.py'; timeout = 300 },
    @{ name = '_smoke_test_svn_conflict_detection.py'; timeout = 300 },
    @{ name = '_smoke_test_svn_author_diagnostics.py'; timeout = 300 },
    @{ name = '_gui_self_test_exact_sheet_readiness.py'; timeout = 300 },
    @{ name = '_gui_self_test_sheet_diff_state.py'; timeout = 300 },
    @{ name = '_gui_self_test_sheet_diff_state_3way.py'; timeout = 300 },
    @{ name = '_gui_self_test_large_virtual_viewport.py'; timeout = 300 },
    @{ name = '_gui_self_test_large_overlay_batch.py'; timeout = 300 },
    @{ name = '_gui_self_test_only_diff.py'; timeout = 300 },
    @{ name = '_gui_self_test_only_diff_progress_modal.py'; timeout = 300 },
    @{ name = '_gui_self_test_binary_identical_fastpath.py'; timeout = 300 },
    @{ name = '_gui_self_test_structural_sheet_cache_isolation.py'; timeout = 300 }
)

$results = @()
$failed = $false
foreach ($test in $tests) {
    $name = $test.name
    $safeName = [IO.Path]::GetFileNameWithoutExtension($name)
    $stdout = Join-Path $evidenceDir "$safeName.stdout.log"
    $stderr = Join-Path $evidenceDir "$safeName.stderr.log"
    Remove-Item -LiteralPath $stdout,$stderr -Force -ErrorAction SilentlyContinue
    $started = [DateTime]::UtcNow
    $child = Start-Process -FilePath 'python' -ArgumentList @($name) -WorkingDirectory $PSScriptRoot -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $timedOut = $false
    while (-not $child.HasExited -and (([DateTime]::UtcNow - $started).TotalSeconds -lt [double]$test.timeout)) {
        Start-Sleep -Milliseconds 100
        $child.Refresh()
    }
    if (-not $child.HasExited) {
        $timedOut = $true
        $cleaned = @(Stop-ExactProcessTree $child.Id)
        Start-Sleep -Milliseconds 200
        $child.Refresh()
    } else {
        $cleaned = @()
    }
    $exitCode = $null
    if ($child.HasExited) {
        $child.WaitForExit()
        $exitCode = [int]$child.ExitCode
    }
    $elapsed = [math]::Round(([DateTime]::UtcNow - $started).TotalSeconds, 3)
    $result = [ordered]@{
        name = $name; pid = $child.Id; exit_code = $exitCode; timeout_seconds = $test.timeout
        timed_out = $timedOut; elapsed_seconds = $elapsed; child_cleanup_pids = $cleaned
        stdout = $stdout; stderr = $stderr; status = $(if (-not $timedOut -and $exitCode -eq 0) { 'PASS' } else { 'FAIL' })
    }
    $results += [pscustomobject]$result
    ($result | ConvertTo-Json -Compress) | Add-Content -LiteralPath (Join-Path $evidenceDir 'progress.jsonl')
    Write-Host ("{0} {1} {2:N3}s pid={3}" -f $result.status, $name, $elapsed, $child.Id)
    if ($result.status -ne 'PASS') { $failed = $true; break }
}

$summary = [ordered]@{
    schema = 'terra-final-regression-v1'; created_utc = [DateTime]::UtcNow.ToString('o')
    production_sha256 = Get-Sha256 (Join-Path $PSScriptRoot 'sow_merge_tool.py')
    results = $results; passed = @($results | Where-Object status -eq 'PASS').Count
    failed = @($results | Where-Object status -ne 'PASS').Count; complete = ((-not $failed) -and $results.Count -eq $tests.Count)
}
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $evidenceDir 'summary.json') -Encoding utf8
if ($failed) { exit 1 }
