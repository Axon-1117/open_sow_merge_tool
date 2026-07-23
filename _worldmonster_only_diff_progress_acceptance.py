"""Isolated real-file acceptance for only-diff progress, tab lock, and C width."""

from __future__ import annotations

import json
import os
import shutil
import tempfile
import time

import sow_merge_tool as smt
from _ux_5_3_final_acceptance import (
    _pump,
    _wait_for_stable_projection,
    _wait_for_view,
    wait_edit_ready,
    wait_view_ready,
)


FIXTURE_ROOT = (
    r"C:\Users\dd\AppData\Local\Temp"
    r"\sow_ux_5_3_20260723_001\cases\WorldMonster"
)
SHEET = "WorldMonster@design"


def _rect(widget):
    return (
        int(widget.winfo_rootx()),
        int(widget.winfo_rooty()),
        int(widget.winfo_width()),
        int(widget.winfo_height()),
    )


def main():
    mine = os.path.join(FIXTURE_ROOT, "mine", "WorldMonster.xlsx")
    theirs = os.path.join(FIXTURE_ROOT, "theirs", "WorldMonster.xlsx")
    base = os.path.join(FIXTURE_ROOT, "base", "WorldMonster.xlsx")
    for path in (mine, theirs, base):
        if not os.path.isfile(path):
            raise FileNotFoundError(path)

    settings_dir = tempfile.mkdtemp(prefix="sow-wm-only-diff-")
    settings_path = os.path.join(settings_dir, "settings.json")
    with open(settings_path, "w", encoding="utf-8") as stream:
        json.dump({"only_diff": 0}, stream)
    original_settings_path = smt._SETTINGS_PATH
    smt._SETTINGS_PATH = settings_path
    app = None
    try:
        app = smt.SowMergeApp(mine, theirs, base_path=base)
        app._intended_window_state = "normal"
        app.root.state("normal")
        app.root.geometry("1600x900")
        app.nb.select(app._sheet_containers[SHEET])
        view = _wait_for_view(app, SHEET, timeout=240.0)
        wait_edit_ready(app, timeout=240.0)
        wait_view_ready(view, timeout=240.0)
        _wait_for_stable_projection(view, timeout=240.0, stable_for=0.4)

        view.only_diff_var.set(0)
        view._last_only_diff_value = 0
        view._only_diff_preview_full = False
        view._refresh_mode_switch_preserving_selection(rescan=False)
        view._invalidate_only_diff_snapshot_cache()
        view._only_diff_rows_exact = False
        view._pair_diff_full_exact = True
        view._base_diff_full_exact = True
        view._refresh_interaction_gate()
        _pump(app.root, 0.2)
        assert view._lifecycle_state == "READY", view._lifecycle_state

        ready_rect = _rect(view.only_diff_cb)
        c_width = int(view.cursor_cmp.winfo_width())
        c_available = int(view.c_area.winfo_width()) - int(
            view.cursor_cmp_ln.winfo_width()
        )
        main_width = int(view.left.winfo_width())
        other_sheet = next(
            (name for name in app.display_sheets if name != SHEET),
            None,
        )

        # Exercise the real WorldMonster worker and the real Cancel button
        # before the successful retry.
        stable_full_rows = tuple(view.display_rows)
        stable_full_x = float(view.left.xview()[0])
        cancel_started = time.perf_counter()
        view.only_diff_cb.invoke()
        cancel_popup_ms = None
        cancel_progress = None
        cancel_deadline = time.monotonic() + 30.0
        while time.monotonic() < cancel_deadline:
            _pump(app.root, 0.02)
            win = app._only_diff_progress_win
            if win is not None and win.winfo_viewable():
                if cancel_popup_ms is None:
                    cancel_popup_ms = (
                        time.perf_counter() - cancel_started
                    ) * 1000.0
                cancel_progress = float(
                    app._only_diff_progress_bar.cget("value")
                )
                if 1.0 <= cancel_progress <= 70.0:
                    break
        assert cancel_popup_ms is not None
        assert cancel_progress is not None
        canceled_seq = int(view._only_diff_async_build_seq)
        button = app._only_diff_progress_cancel_btn
        cancel_click_started = time.perf_counter()
        button.invoke()
        cancel_click_ms = (
            time.perf_counter() - cancel_click_started
        ) * 1000.0
        _pump(app.root, 0.1)

        broker_stop_started = time.perf_counter()
        broker_deadline = time.monotonic() + 2.0
        while time.monotonic() < broker_deadline:
            _pump(app.root, 0.02)
            with app._exact_broker_lock:
                broker_running = bool(app._exact_broker_running)
            if not broker_running:
                break
        broker_stop_ms = (
            time.perf_counter() - broker_stop_started
        ) * 1000.0
        _pump(app.root, 0.3)
        cancel_view_stable = bool(
            tuple(view.display_rows) == stable_full_rows
            and abs(float(view.left.xview()[0]) - stable_full_x) < 0.01
        )
        stale_publish_rejected = bool(
            int(view._only_diff_async_build_seq) > canceled_seq
            and not view._has_valid_only_diff_snapshot_cache()
        )
        cancel_state_ok = bool(
            app._only_diff_progress_owner is None
            and int(view.only_diff_var.get()) == 0
            and not view._only_diff_async_building
            and view._lifecycle_state == "READY"
        )
        switch_after_cancel = True
        if other_sheet is not None:
            app._select_tab(other_sheet)
            _pump(app.root, 0.05)
            switch_after_cancel = app.selected_sheet == other_sheet
            app._select_tab(SHEET)
            _pump(app.root, 0.08)
            view = app.sheet_views[SHEET]

        app._ui_heartbeat_max_gap = 0.0
        app._ui_heartbeat_last = time.perf_counter()
        started = time.perf_counter()
        view.only_diff_cb.invoke()
        callback_ms = (time.perf_counter() - started) * 1000.0

        popup_seen_ms = None
        progress_values = []
        tab_lock_checked = False
        owner_sheet_while_locked = None
        deadline = time.monotonic() + 240.0
        while time.monotonic() < deadline:
            _pump(app.root, 0.02)
            win = app._only_diff_progress_win
            if win is not None and win.winfo_viewable():
                if popup_seen_ms is None:
                    popup_seen_ms = (time.perf_counter() - started) * 1000.0
                try:
                    progress_values.append(
                        float(app._only_diff_progress_bar.cget("value"))
                    )
                except Exception:
                    pass
                if not tab_lock_checked and other_sheet is not None:
                    app.nb.select(app._sheet_containers[other_sheet])
                    _pump(app.root, 0.03)
                    owner_sheet_while_locked = app.nb.tab(
                        app.nb.select(),
                        "text",
                    )
                    app._select_tab(other_sheet)
                    _pump(app.root, 0.03)
                    tab_lock_checked = True
            if (
                not view._only_diff_async_building
                and not view._mode_switch_pending
                and app._only_diff_progress_owner is None
                and view._has_valid_only_diff_snapshot_cache()
            ):
                break
        total_ms = (time.perf_counter() - started) * 1000.0
        ready_after = _rect(view.only_diff_cb)
        result = {
            "callback_ms": round(callback_ms, 2),
            "popup_seen_ms": (
                None if popup_seen_ms is None else round(popup_seen_ms, 2)
            ),
            "total_ms": round(total_ms, 2),
            "heartbeat_max_gap_ms": round(
                float(app._ui_heartbeat_max_gap) * 1000.0,
                2,
            ),
            "progress_sample_count": len(progress_values),
            "progress_min": min(progress_values) if progress_values else None,
            "progress_max": max(progress_values) if progress_values else None,
            "tab_lock_checked": tab_lock_checked,
            "owner_sheet_while_locked": owner_sheet_while_locked,
            "cache_valid": view._has_valid_only_diff_snapshot_cache(),
            "diff_rows": len(view._only_diff_rows_cache or ()),
            "checkbox_rect_before": ready_rect,
            "checkbox_rect_after": ready_after,
            "checkbox_max_drift_px": max(
                abs(a - b) for a, b in zip(ready_rect, ready_after)
            ),
            "c_width": c_width,
            "c_available": c_available,
            "main_width": main_width,
            "cancel": {
                "popup_seen_ms": round(cancel_popup_ms, 2),
                "progress_at_cancel": round(cancel_progress, 2),
                "click_ms": round(cancel_click_ms, 2),
                "broker_stop_ms": round(broker_stop_ms, 2),
                "state_ok": cancel_state_ok,
                "view_stable": cancel_view_stable,
                "stale_publish_rejected": stale_publish_rejected,
                "sheet_switch_restored": switch_after_cancel,
            },
        }
        result["pass"] = bool(
            result["callback_ms"] < 100.0
            and result["popup_seen_ms"] is not None
            and result["popup_seen_ms"] <= 100.0
            and result["heartbeat_max_gap_ms"] <= 200.0
            and result["progress_sample_count"] > 0
            and result["progress_max"] > result["progress_min"]
            and (
                other_sheet is None
                or result["owner_sheet_while_locked"] == SHEET
            )
            and result["cache_valid"]
            and result["diff_rows"] == 1
            and result["checkbox_max_drift_px"] <= 2
            and result["c_width"] >= result["c_available"] - 12
            and result["c_width"] > result["main_width"] * 2.0
            and result["cancel"]["popup_seen_ms"] <= 100.0
            and 1.0 <= result["cancel"]["progress_at_cancel"] <= 70.0
            and result["cancel"]["click_ms"] < 100.0
            and result["cancel"]["broker_stop_ms"] <= 2000.0
            and result["cancel"]["state_ok"]
            and result["cancel"]["view_stable"]
            and result["cancel"]["stale_publish_rejected"]
            and result["cancel"]["sheet_switch_restored"]
        )
        print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
        if not result["pass"]:
            raise AssertionError(result)
    finally:
        smt._SETTINGS_PATH = original_settings_path
        if app is not None:
            app._shutdown_root()
        shutil.rmtree(settings_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
