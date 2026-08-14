"""One-process real-workbook benchmark/oracle for logical column actions.

Run each case in three fresh processes for nearest-rank P95 evidence, e.g.:
  python _performance_benchmark_column_action_real.py --case Skill --oracle
"""

from __future__ import annotations

import argparse
import gc
import json
import os
import time
import tracemalloc

import psutil

import sow_merge_tool as mod


ROOT = r"C:\tmp\column_alignment_baseline"
CASES = {
    "Guide": ("TGuideStep@design", 12, "B", "A"),
    "Skill": ("SkillTimeline@design", 12, "B", "A"),
    "Dungeon": ("Dungeon@design", 23, "B", "A"),
}


def _pump(root, seconds=0.03):
    deadline = time.time() + seconds
    while time.time() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.005)


def _wait_for_view(app, sheet, timeout=120.0):
    app.nb.select(app._sheet_containers[sheet])
    deadline = time.time() + timeout
    while time.time() < deadline:
        _pump(app.root)
        view = app.sheet_views.get(sheet)
        if view is not None and getattr(view, "_data_ready", False):
            return view
    raise RuntimeError(f"view did not become ready: {sheet}")


def _exact_full_map_same_model(view):
    """Force the ordinary block scanner without rebuilding row/column models."""
    ws_a = view.app.ws_a_val(view.sheet)
    ws_b = view.app.ws_b_val(view.sheet)
    ws_a_edit = view.app.ws_a_edit(view.sheet)
    ws_b_edit = view.app.ws_b_edit(view.sheet)
    view.pair_diff_cols = {}
    view.pair_base_diff_cols = {}
    view.pair_text_a = {}
    view.pair_text_b = {}
    view.pair_text_base = {}
    view._precompute_large_diff_by_blocks(
        ws_a,
        ws_b,
        ws_a_edit,
        ws_b_edit,
        int(getattr(view, "_effective_max_row_a", ws_a.max_row or 1)),
        int(getattr(view, "_effective_max_row_b", ws_b.max_row or 1)),
    )
    return {idx: set(cols) for idx, cols in view.pair_diff_cols.items() if cols}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--case", choices=tuple(CASES), required=True)
    parser.add_argument("--oracle", action="store_true")
    parser.add_argument("--memory-layers", type=int, default=0)
    parser.add_argument("--profile-methods", action="store_true")
    parser.add_argument("--full-view-action", action="store_true")
    args = parser.parse_args()

    method_timings = {}
    if args.profile_methods:
        for method_name in (
            "_column_action_workbook_snapshot",
            "_prepare_column_copy_payload",
            "_apply_column_copy_payload",
            "_rebuild_column_comparison_cache_from_worksheets",
            "_prescan_col_widths",
            "_try_apply_column_action_diff_seed",
            "_restore_column_action_workbook_snapshot",
            "refresh",
        ):
            original = getattr(mod.SheetView, method_name)

            def _profiled(self, *method_args, __name=method_name, __original=original, **method_kwargs):
                method_started = time.perf_counter()
                try:
                    return __original(self, *method_args, **method_kwargs)
                finally:
                    method_timings.setdefault(__name, []).append(
                        time.perf_counter() - method_started
                    )

            setattr(mod.SheetView, method_name, _profiled)

        original_shift = mod.SheetView._shift_worksheet_columns_fast

        def _profiled_shift(*method_args, **method_kwargs):
            method_started = time.perf_counter()
            try:
                return original_shift(*method_args, **method_kwargs)
            finally:
                method_timings.setdefault("_shift_worksheet_columns_fast", []).append(
                    time.perf_counter() - method_started
                )

        mod.SheetView._shift_worksheet_columns_fast = staticmethod(_profiled_shift)

    sheet, logical_col, source, target = CASES[args.case]
    case_root = os.path.join(ROOT, args.case)
    original = os.path.join(case_root, "original.xlsx")
    changed = os.path.join(case_root, "insert2_delete1.xlsx")
    started = time.perf_counter()
    app = mod.SowMergeApp(original, changed)
    result = {"case": args.case, "pid": os.getpid()}
    try:
        view = _wait_for_view(app, sheet)
        result["first_ready_seconds"] = time.perf_counter() - started
        preload_started = time.perf_counter()
        app._ensure_edit_loaded()
        result["edit_ready_seconds"] = time.perf_counter() - preload_started
        view.only_diff_var.set(1)
        exact_started = time.perf_counter()
        view.refresh(row_only=None, rescan=True)
        result["exact_scan_seconds"] = time.perf_counter() - exact_started
        result["initial_exact"] = bool(view._pair_diff_full_exact)
        if args.full_view_action:
            view.only_diff_var.set(0)
            view.refresh(row_only=None, rescan=False)
        view._select_column_block_by_logical_col(logical_col, source)
        plan = view._plan_selected_column_block_action(source, target)
        if args.profile_methods:
            result["profile_before_slots"] = [
                (
                    slot.logical_idx + 1,
                    slot.mine_col,
                    slot.theirs_col,
                    slot.state,
                    slot.confidence.score,
                    slot.confidence.ambiguous,
                    slot.confidence.reason,
                )
                for slot in view._active_column_projection().model.slots
                if abs((slot.logical_idx + 1) - logical_col) <= 3
            ]

        if args.memory_layers > 0:
            process = psutil.Process()
            gc.collect()
            tracemalloc.start()
            trace_before, _peak_before = tracemalloc.get_traced_memory()
            rss_before = process.memory_info().rss
            snapshots = [
                view._column_action_workbook_snapshot(target, plan)
                for _index in range(args.memory_layers)
            ]
            gc.collect()
            trace_live, trace_peak = tracemalloc.get_traced_memory()
            rss_live = process.memory_info().rss
            seed_bytes = [
                len(
                    ((snapshot.get("column_diff_seed") or {}).get("pair_diff_bitmap") or {}).get(
                        "bits", b""
                    )
                )
                for snapshot in snapshots
            ]
            comparison_states = [snapshot.get("column_comparison_state") for snapshot in snapshots]
            snapshots.clear()
            gc.collect()
            trace_after, _peak_after = tracemalloc.get_traced_memory()
            rss_after = process.memory_info().rss
            tracemalloc.stop()
            result["memory"] = {
                "layers": args.memory_layers,
                "rss_live_delta_bytes": rss_live - rss_before,
                "rss_after_clear_delta_bytes": rss_after - rss_before,
                "trace_live_delta_bytes": trace_live - trace_before,
                "trace_after_clear_delta_bytes": trace_after - trace_before,
                "trace_peak_bytes": trace_peak,
                "seed_bitmap_bytes": seed_bytes,
                "comparison_state_all_none": all(state is None for state in comparison_states),
            }

        apply_started = time.perf_counter()
        applied = view._apply_selected_column_block(source, target)
        result["apply_seconds"] = time.perf_counter() - apply_started
        result["action_kind"] = applied.action_kind
        result["seed"] = dict(view._column_diff_seed_last)
        result["display_row_count_after_apply"] = len(view.display_rows)
        if args.profile_methods:
            result["profile_after_slots"] = [
                (
                    slot.logical_idx + 1,
                    slot.mine_col,
                    slot.theirs_col,
                    slot.state,
                    slot.confidence.score,
                    slot.confidence.ambiguous,
                    slot.confidence.reason,
                )
                for slot in view._active_column_projection().model.slots
                if abs((slot.logical_idx + 1) - logical_col) <= 3
            ]
        if args.oracle:
            seeded = {idx: set(cols) for idx, cols in view.pair_diff_cols.items() if cols}
            oracle_started = time.perf_counter()
            oracle = _exact_full_map_same_model(view)
            result["oracle_seconds"] = time.perf_counter() - oracle_started
            result["oracle_equal"] = seeded == oracle
            if seeded != oracle:
                mismatches = [
                    idx for idx in range(len(view.row_pairs))
                    if seeded.get(idx, set()) != oracle.get(idx, set())
                ]
                result["oracle_mismatch_count"] = len(mismatches)
                result["oracle_mismatch_sample"] = mismatches[:20]
                raise AssertionError(result)
        undo_started = time.perf_counter()
        view._undo_last_action()
        result["undo_seconds"] = time.perf_counter() - undo_started
        result["undo_seed"] = dict(view._column_diff_seed_last)
        if args.profile_methods:
            result["profile_methods"] = method_timings
        print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    finally:
        app._shutdown_root()


if __name__ == "__main__":
    main()
