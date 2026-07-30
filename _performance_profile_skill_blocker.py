"""Fresh-process Skill action and native-save blocker profiler.

The default mode never launches Excel.  ``--native-save`` is intentionally
opt-in and injects Stopwatch markers into the existing PowerShell replay at
runtime; it does not change production save behavior.
"""

from __future__ import annotations

import argparse
import gc
import json
import os
import time
import tracemalloc

import psutil

import sow_merge_tool as smt
import _ux_5_3_final_acceptance as ux


DEFAULT_FIXTURE_ROOT = (
    r"C:\Users\dd\AppData\Local\Temp\sow_ux_5_3_20260723_001"
    r"\cases\Skill"
)
SHEET = "SkillTimeline@design"


def _replace_once(script: str, anchor: str, replacement: str, label: str) -> str:
    count = script.count(anchor)
    if count != 1:
        raise RuntimeError(
            f"native timing anchor {label!r} matched {count} times; refusing to run"
        )
    return script.replace(anchor, replacement, 1)


def _instrument_native_powershell(
    script: str,
    *,
    double_gc: bool = False,
    release_struct_rcw: bool = False,
    bounded_copy: bool = False,
    force_exit: bool = False,
) -> str:
    """Inject cumulative Stopwatch markers without changing Excel operations."""
    init = (
        "$xl=$null;$wb=$null;$wbMine=$null;$wbBase=$null;$wbTheirs=$null;"
        "$wbCheck=$null;"
    )
    script = _replace_once(
        script,
        init + "try{",
        init
        + "$sowExcelBefore=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id});"
        + "$sowExcelPid=$null;"
        + "$sowSuccess=$false;"
        + "$sowSw=[Diagnostics.Stopwatch]::StartNew();"
        + "function SowMark([string]$name){Write-Output "
        + "('SOW_NATIVE_PHASE|'+$name+'|'+[math]::Round($sowSw.Elapsed.TotalMilliseconds,3))};"
        + "SowMark 'payload_ready';try{",
        "init",
    )
    if bounded_copy:
        full_column_ranges = (
            "  $srcRange=$srcWs.Range($srcWs.Columns.Item($srcFirst),"
            "$srcWs.Columns.Item($srcLast)).EntireColumn;"
            "  $dstRange=$ws.Range($ws.Columns.Item($anchor),"
            "$ws.Columns.Item($last)).EntireColumn;"
        )
        bounded_ranges = (
            "  $sowCopyLastRow=[Math]::Max(1,[int]($srcWs.UsedRange.Row+"
            "$srcWs.UsedRange.Rows.Count-1));"
            "  $srcRange=$srcWs.Range($srcWs.Cells.Item(1,$srcFirst),"
            "$srcWs.Cells.Item($sowCopyLastRow,$srcLast));"
            "  $dstRange=$ws.Range($ws.Cells.Item(1,$anchor),"
            "$ws.Cells.Item($sowCopyLastRow,$last));"
        )
        script = _replace_once(
            script, full_column_ranges, bounded_ranges, "bounded-copy"
        )

    range_release = (
        "if($srcRange -ne $null){[void][Runtime.InteropServices.Marshal]::"
        "FinalReleaseComObject($srcRange);$srcRange=$null};"
        "if($dstRange -ne $null){[void][Runtime.InteropServices.Marshal]::"
        "FinalReleaseComObject($dstRange);$dstRange=$null};"
    ) if release_struct_rcw else ""
    column_release = (
        "if($srcColumn -ne $null){[void][Runtime.InteropServices.Marshal]::"
        "FinalReleaseComObject($srcColumn);$srcColumn=$null};"
        "if($dstColumn -ne $null){[void][Runtime.InteropServices.Marshal]::"
        "FinalReleaseComObject($dstColumn);$dstColumn=$null};"
    ) if release_struct_rcw else ""
    sheet_release = (
        "if($srcWs -ne $null){[void][Runtime.InteropServices.Marshal]::"
        "FinalReleaseComObject($srcWs);$srcWs=$null};"
        "if($ws -ne $null){[void][Runtime.InteropServices.Marshal]::"
        "FinalReleaseComObject($ws);$ws=$null};"
    ) if release_struct_rcw else ""
    anchors = (
        (
            "$xl=New-Object -ComObject Excel.Application;",
            "$xl=New-Object -ComObject Excel.Application;"
            "$sowExcelPid=@(Get-Process EXCEL -ErrorAction SilentlyContinue | "
            "Where-Object {$sowExcelBefore -notcontains $_.Id} | "
            "Select-Object -First 1 -ExpandProperty Id);"
            "Write-Output ('SOW_NATIVE_EXCEL_PID|'+[string]$sowExcelPid);"
            "SowMark 'com_activated';",
            "com",
        ),
        (
            "try{$xl.CalculateBeforeSave=$false}catch{};",
            "try{$xl.CalculateBeforeSave=$false}catch{};SowMark 'configured';",
            "configured",
        ),
        (
            "$wb=$xl.Workbooks.Open($src,0,$false);",
            "$wb=$xl.Workbooks.Open($src,0,$false);SowMark 'target_open';",
            "target-open",
        ),
        (
            "if($theirsSrc){$wbTheirs=$xl.Workbooks.Open($theirsSrc,0,$true)};",
            "if($theirsSrc){$wbTheirs=$xl.Workbooks.Open($theirsSrc,0,$true)};"
            "SowMark 'sources_open';",
            "sources-open",
        ),
        (
            "[void]$dstRange.Insert();continue",
            "[void]$dstRange.Insert();"
            "SowMark ('struct_insert_cols_'+[string]$op.order);"
            + range_release
            + sheet_release
            + "continue",
            "insert-cols",
        ),
        (
            "[void]$dstRange.Delete();continue",
            "[void]$dstRange.Delete();"
            "SowMark ('struct_delete_cols_'+[string]$op.order);"
            + range_release
            + sheet_release
            + "continue",
            "delete-cols",
        ),
        (
            "[void]$srcRange.Copy($dstRange);",
            "[void]$srcRange.Copy($dstRange);"
            "SowMark ('struct_copy_data_'+[string]$op.order);"
            + range_release,
            "copy-data",
        ),
        (
            "try{$dstColumn.Hidden=$srcColumn.Hidden}catch{};  };",
            "try{$dstColumn.Hidden=$srcColumn.Hidden}catch{};"
            + column_release
            + "  };"
            "SowMark ('struct_width_hidden_'+[string]$op.order);",
            "width-hidden",
        ),
        (
            "try{$xl.CutCopyMode=0}catch{}};",
            "try{$xl.CutCopyMode=0}catch{};"
            "SowMark ('struct_cut_copy_mode_'+[string]$op.order);"
            + sheet_release
            + "SowMark ('struct_copy_cols_'+[string]$op.order)};"
            "SowMark 'structural_replay';",
            "structural",
        ),
        (
            "if($wbMine -ne $null){$wbMine.Close($false);"
            "[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbMine);"
            "$wbMine=$null};",
            "if($wbMine -ne $null){$wbMine.Close($false);"
            "[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbMine);"
            "$wbMine=$null};SowMark 'sources_closed';",
            "sources-closed",
        ),
        (
            "};$wb.SaveCopyAs($out);",
            "};SowMark 'cell_ops';$wb.SaveCopyAs($out);SowMark 'save_copy_as';",
            "cell-save",
        ),
        (
            "if($wb -ne $null){$wb.Close($false);"
            "[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb);"
            "$wb=$null};$wbCheck=$xl.Workbooks.Open($out,0,$true);",
            "if($wb -ne $null){$wb.Close($false);"
            "[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb);"
            "$wb=$null};SowMark 'target_closed';"
            "$wbCheck=$xl.Workbooks.Open($out,0,$true);SowMark 'output_reopened';",
            "reopen",
        ),
        (
            "[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbCheck);"
            "$wbCheck=$null;",
            "[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbCheck);"
            "$wbCheck=$null;$sowSuccess=$true;SowMark 'output_reopen_closed';",
            "reopen-close",
        ),
        (
            "if($xl -ne $null){try{$xl.Quit()}catch{};"
            "try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}"
            "catch{}};",
            "if($xl -ne $null){try{$xl.Quit()}catch{};SowMark 'excel_quit_called';"
            "try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}"
            "catch{}};SowMark 'excel_app_released';"
            + (
                "if($sowSuccess){[Console]::Out.WriteLine('SOW_NATIVE_FORCE_EXIT|1');"
                "[Console]::Out.Flush();[Environment]::Exit(0)};"
                if force_exit else ""
            ),
            "excel-quit",
        ),
    )
    for anchor, replacement, label in anchors:
        script = _replace_once(script, anchor, replacement, label)
    gc_pass = "[GC]::Collect();[GC]::WaitForPendingFinalizers();"
    script = _replace_once(
        script,
        gc_pass + "};",
        gc_pass
        + (gc_pass if double_gc else "")
        + "SowMark 'quit_gc_complete';"
        + "$sowPresentBeforeExit=0;"
        + "if($sowExcelPid){try{if(Get-Process -Id $sowExcelPid -ErrorAction Stop)"
        + "{$sowPresentBeforeExit=1}}catch{}};"
        + "Write-Output ('SOW_NATIVE_EXCEL_PRESENT_BEFORE_PS_EXIT|'"
        + "+[string]$sowPresentBeforeExit);"
        + "$sowExitWait=[Diagnostics.Stopwatch]::StartNew();"
        + "while($sowExcelPid -and $sowExitWait.Elapsed.TotalSeconds -lt 15)"
        + "{try{[void](Get-Process -Id $sowExcelPid -ErrorAction Stop);"
        + "Start-Sleep -Milliseconds 10}catch{break}};"
        + "Write-Output ('SOW_NATIVE_EXCEL_EXIT_WAIT_MS|'"
        + "+[math]::Round($sowExitWait.Elapsed.TotalMilliseconds,3));"
        + "SowMark 'excel_process_gone';};",
        "gc",
    )
    return script


def _install_native_timing(
    result: dict,
    *,
    double_gc: bool = False,
    release_struct_rcw: bool = False,
    bounded_copy: bool = False,
    force_exit: bool = False,
) -> None:
    original_runner = smt._run_excel_powershell_with_transient_retry
    original_validate = smt._validate_xlsx_package
    original_copy2 = smt.shutil.copy2

    def _runner(script: str, timeout: int = 180):
        instrumented = _instrument_native_powershell(
            script,
            double_gc=double_gc,
            release_struct_rcw=release_struct_rcw,
            bounded_copy=bounded_copy,
            force_exit=force_exit,
        )
        started = time.perf_counter()
        completed = original_runner(instrumented, timeout=timeout)
        result["powershell_wall_ms"] = round(
            (time.perf_counter() - started) * 1000, 3
        )
        phases = []
        for line in str(completed.stdout or "").splitlines():
            if not line.startswith("SOW_NATIVE_PHASE|"):
                continue
            _prefix, name, elapsed = line.split("|", 2)
            phases.append({"name": name, "cumulative_ms": float(elapsed)})
        for line in str(completed.stdout or "").splitlines():
            if line.startswith("SOW_NATIVE_EXCEL_PID|"):
                raw_pid = line.split("|", 1)[1].strip()
                result["excel_pid"] = int(raw_pid) if raw_pid else None
            elif line.startswith("SOW_NATIVE_EXCEL_PRESENT_BEFORE_PS_EXIT|"):
                result["excel_present_before_ps_exit"] = bool(
                    int(line.split("|", 1)[1])
                )
            elif line.startswith("SOW_NATIVE_EXCEL_EXIT_WAIT_MS|"):
                result["excel_exit_wait_ms"] = float(line.split("|", 1)[1])
        result["native_phases"] = phases
        if force_exit:
            excel_pid = result.get("excel_pid")
            result["force_exit_excel_present_at_ps_return"] = bool(
                excel_pid and psutil.pid_exists(int(excel_pid))
            )
            exit_started = time.perf_counter()
            while (
                excel_pid
                and psutil.pid_exists(int(excel_pid))
                and time.perf_counter() - exit_started < 15.0
            ):
                time.sleep(0.01)
            result["force_exit_excel_wait_ms"] = round(
                (time.perf_counter() - exit_started) * 1000, 3
            )
            result["force_exit_excel_remaining"] = bool(
                excel_pid and psutil.pid_exists(int(excel_pid))
            )
        return completed

    def _validate(path: str):
        started = time.perf_counter()
        try:
            return original_validate(path)
        finally:
            result.setdefault("package_validate_ms", []).append(
                round((time.perf_counter() - started) * 1000, 3)
            )

    def _copy2(source, destination, *args, **kwargs):
        started = time.perf_counter()
        copied = original_copy2(source, destination, *args, **kwargs)
        if "native_sources_" in str(destination):
            result.setdefault("immutable_stage", []).append({
                "side_source": os.path.basename(str(source)),
                "bytes": os.path.getsize(destination),
                "ms": round((time.perf_counter() - started) * 1000, 3),
            })
        return copied

    smt._run_excel_powershell_with_transient_retry = _runner
    smt._validate_xlsx_package = _validate
    smt.shutil.copy2 = _copy2


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--fixture-root", default=DEFAULT_FIXTURE_ROOT)
    parser.add_argument("--native-save", action="store_true")
    parser.add_argument("--measure-memory", action="store_true")
    parser.add_argument("--double-gc", action="store_true")
    parser.add_argument("--release-struct-rcw", action="store_true")
    parser.add_argument("--bounded-copy", action="store_true")
    parser.add_argument("--force-exit", action="store_true")
    args = parser.parse_args()

    mine = os.path.join(args.fixture_root, "mine", "Skill.xlsx")
    theirs = os.path.join(args.fixture_root, "theirs", "Skill.xlsx")
    if not os.path.isfile(mine) or not os.path.isfile(theirs):
        raise FileNotFoundError(f"Skill fixture is incomplete: {args.fixture_root}")

    app = None
    result = {
        "pid": os.getpid(),
        "mode": "native-save" if args.native_save else "actions-only",
        "fixture_root": args.fixture_root,
    }
    try:
        app = smt.SowMergeApp(mine, theirs)
        app.root.withdraw()
        view = ux._force_full_view(ux._wait_for_view(app, SHEET, timeout=180.0))
        ux.wait_edit_ready(app, timeout=180.0)
        ux._wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
        ux.only_diff_metrics(view)

        process = psutil.Process()
        gc.collect()
        rss_before = process.memory_info().rss
        if args.measure_memory:
            tracemalloc.start()
            trace_before, _unused = tracemalloc.get_traced_memory()
        else:
            trace_before = 0
        actions = ux.apply_all_structural(
            view, app, source_side="B", validate_undo=True
        )
        gc.collect()
        if args.measure_memory:
            trace_live, trace_peak = tracemalloc.get_traced_memory()
            tracemalloc.stop()
        else:
            trace_live = trace_peak = 0
        rss_live = process.memory_info().rss
        result["actions"] = actions
        result["action_memory"] = {
            "rss_live_delta_bytes": rss_live - rss_before,
            "trace_live_delta_bytes": (
                trace_live - trace_before if args.measure_memory else None
            ),
            "trace_peak_bytes": trace_peak if args.measure_memory else None,
        }

        if args.native_save:
            timing = {}
            _install_native_timing(
                timing,
                double_gc=args.double_gc,
                release_struct_rcw=args.release_struct_rcw,
                bounded_copy=args.bounded_copy,
                force_exit=args.force_exit,
            )
            started = time.perf_counter()
            output = app.build_manual_merge_output_file()
            timing["build_total_ms"] = round(
                (time.perf_counter() - started) * 1000, 3
            )
            timing["output"] = {
                "path": output,
                "bytes": os.path.getsize(output),
            }
            timing["probe"] = ux.workbook_probe(
                output,
                SHEET,
                markers=("__UX_INS1_R1", "__UX_INS2_R1"),
            )
            result["native_save"] = timing
        print(json.dumps(result, ensure_ascii=False, sort_keys=True), flush=True)
    finally:
        ux.close_app(app)


if __name__ == "__main__":
    main()
