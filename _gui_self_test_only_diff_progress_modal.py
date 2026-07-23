"""GUI regression: only-diff progress owns input, supports Cancel, and does not reflow."""

import os
import sys
import time

from openpyxl import Workbook

from _test_temp_utils import make_temp_dir


def _make_book(path: str, changed: bool):
    wb = Workbook()
    for sheet_idx, sheet_name in enumerate(("WorldMonster", "OtherSheet")):
        ws = wb.active if sheet_idx == 0 else wb.create_sheet()
        ws.title = sheet_name
        for row in range(1, 81):
            for col in range(1, 13):
                value = f"{sheet_name}-R{row}-C{col}"
                if changed and sheet_name == "WorldMonster" and row == 60 and col == 7:
                    value = "changed"
                ws.cell(row=row, column=col).value = value
    wb.save(path)


def _pump(root, seconds=0.15):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.01)


def _rect(widget):
    return (
        int(widget.winfo_rootx()),
        int(widget.winfo_rooty()),
        int(widget.winfo_width()),
        int(widget.winfo_height()),
    )


def main():
    root_dir = make_temp_dir(prefix="sow_only_diff_modal_")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")
    _make_book(mine, changed=False)
    _make_book(theirs, changed=True)

    sys.path.insert(0, r"D:\Tools\sow_merge_tool_proj")
    import sow_merge_tool as mod

    old_settings_path = mod._SETTINGS_PATH
    mod._SETTINGS_PATH = os.path.join(root_dir, "settings.json")
    app = None
    try:
        app = mod.SowMergeApp(mine, theirs)
        app._intended_window_state = "normal"
        app.root.state("normal")
        app.root.geometry("1300x800")
        _pump(app.root, 0.4)
        deadline = time.monotonic() + 8.0
        while time.monotonic() < deadline:
            with app._compute_lock:
                background_done = (
                    not app._compute_queue and not app._compute_inflight
                )
            if background_done:
                break
            _pump(app.root, 0.05)
        assert background_done, "background Sheet preparation did not finish"
        sheet = "WorldMonster"
        app.nb.select(app._sheet_containers[sheet])
        _pump(app.root)
        view = app.sheet_views[sheet]
        view.only_diff_var.set(0)
        view._last_only_diff_value = 0
        view.refresh(row_only=None, rescan=True)
        view._data_ready = True
        view._cache_formula_aware = True
        view._pair_diff_full_exact = True
        view._only_diff_rows_exact = False
        view._invalidate_only_diff_snapshot_cache()
        view._refresh_interaction_gate()
        _pump(app.root)
        assert view._lifecycle_state == "READY", view._lifecycle_state

        original_start = view._start_async_large_only_diff_build
        original_cache = view._cache_only_diff_rows_from_exact_pair_maps
        starts = []

        def _fake_start(*, user_initiated=False):
            assert user_initiated is True
            view._only_diff_async_build_seq += 1
            seq = int(view._only_diff_async_build_seq)
            starts.append(seq)
            view._only_diff_async_building = True
            view._only_diff_async_requested_value = 1
            view._only_diff_async_build_key = view._current_only_diff_cache_key()
            view._only_diff_preview_full = True
            app._claim_priority_exact(view, seq)
            view._set_only_diff_pending_info()
            view._refresh_interaction_gate()
            app._begin_only_diff_progress(view, seq)
            return True

        view._start_async_large_only_diff_build = _fake_start
        view._cache_only_diff_rows_from_exact_pair_maps = lambda: False

        ready_rect = _rect(view.only_diff_cb)
        view.only_diff_cb.invoke()
        _pump(app.root)
        assert starts == [starts[0]]
        seq = starts[0]
        assert int(view.only_diff_var.get()) == 1
        assert str(view.only_diff_cb.cget("state")) == "disabled"
        assert str(view.only_diff_cb.cget("text")) == "只看差异内容"
        assert app._only_diff_progress_owner == (view, seq)
        popup = app._only_diff_progress_win
        assert popup is not None and popup.winfo_exists()
        assert popup.grab_current() == popup
        assert "WorldMonster" in popup.title()
        pending_rect = _rect(view.only_diff_cb)
        assert max(abs(a - b) for a, b in zip(ready_rect, pending_rect)) <= 2, (
            ready_rect,
            pending_rect,
        )

        app._update_only_diff_progress(
            view,
            seq,
            "正在生成精确差异",
            40,
            80,
        )
        _pump(app.root, 0.05)
        assert float(app._only_diff_progress_bar.cget("value")) == 50.0
        assert "40 / 80" in app._only_diff_progress_detail_var.get()
        assert "50.0%" in app._only_diff_progress_detail_var.get()

        # Disabled widget input and a defensive direct callback cannot start or
        # cancel a second request.
        view.only_diff_cb.invoke()
        view._toggle_only_diff()
        _pump(app.root, 0.05)
        assert starts == [seq]
        assert int(view.only_diff_var.get()) == 1

        owner_tab = app.nb.select()
        other_tab = app._sheet_containers["OtherSheet"]
        try:
            app.nb.select(other_tab)
        except Exception:
            pass
        _pump(app.root, 0.05)
        assert app.nb.select() == owner_tab
        assert app.selected_sheet == "WorldMonster"
        app._select_tab("OtherSheet")
        _pump(app.root, 0.05)
        assert app.nb.select() == owner_tab

        view._cancel_only_diff_calculation(seq)
        _pump(app.root)
        assert app._only_diff_progress_owner is None
        assert app._only_diff_progress_win is not None
        assert not bool(app._only_diff_progress_win.winfo_viewable())
        assert int(view.only_diff_var.get()) == 0
        assert view._only_diff_async_building is False
        assert str(view.only_diff_cb.cget("state")) == "normal"
        canceled_rect = _rect(view.only_diff_cb)
        assert max(abs(a - b) for a, b in zip(ready_rect, canceled_rect)) <= 2, (
            ready_rect,
            canceled_rect,
        )

        # An old generation cannot reopen/update UI, and tabs are usable again.
        app._update_only_diff_progress(view, seq, "stale", 80, 80)
        assert app._only_diff_progress_owner is None
        app._select_tab("OtherSheet")
        _pump(app.root, 0.1)
        assert app.selected_sheet == "OtherSheet"

        view._start_async_large_only_diff_build = original_start
        view._cache_only_diff_rows_from_exact_pair_maps = original_cache
        app._select_tab("WorldMonster")
        _pump(app.root, 0.1)
        view.only_diff_var.set(1)
        view._only_diff_request_origin_value = 0
        view._invalidate_only_diff_snapshot_cache()
        assert view._start_async_large_only_diff_build(user_initiated=True)
        deadline = time.monotonic() + 15.0
        popup_seen = False
        while time.monotonic() < deadline:
            _pump(app.root, 0.02)
            popup_seen = popup_seen or (
                app._only_diff_progress_win is not None
                and bool(app._only_diff_progress_win.winfo_viewable())
            )
            if (
                not view._only_diff_async_building
                and app._only_diff_progress_owner is None
            ):
                break
        assert popup_seen
        assert view._has_valid_only_diff_snapshot_cache()
        assert int(view.only_diff_var.get()) == 1
        assert view._lifecycle_state == "READY", view._lifecycle_state
        print("GUI_SELF_TEST_ONLY_DIFF_PROGRESS_MODAL_OK")
    finally:
        mod._SETTINGS_PATH = old_settings_path
        if app is not None:
            app._shutdown_root()


if __name__ == "__main__":
    main()
