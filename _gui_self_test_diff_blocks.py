import os
import time

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_book(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def _pump(root, seconds=0.2):
    deadline = time.time() + seconds
    while time.time() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.01)


def _prepare_view(app, *, only_diff=True):
    _pump(app.root)
    view = app.sheet_views["Data"]
    view._suppress_bg_apply = True
    view.only_diff_var.set(1 if only_diff else 0)
    view.refresh(row_only=None, rescan=True)
    _pump(app.root, 0.05)
    return view


def _test_two_way_block_presentation(root_dir):
    mine = os.path.join(root_dir, "two-mine.xlsx")
    theirs = os.path.join(root_dir, "two-theirs.xlsx")
    rows_a = [
        [idx, f"same-{idx}"]
        + [f"payload-{col}-{idx}-" + ("x" * 80) for col in range(3, 13)]
        for idx in range(1, 11)
    ]
    rows_b = [list(row) for row in rows_a]
    for excel_row in (2, 3, 7, 8):
        rows_b[excel_row - 1][1] = f"changed-{excel_row}"
    _make_book(mine, rows_a)
    _make_book(theirs, rows_b)

    app = mod.SowMergeApp(mine, theirs)
    try:
        app.root.state("normal")
        app.root.geometry("800x600")
        view = _prepare_view(app)
        blocks = view._ensure_full_diff_blocks()
        assert [(b.start_pair_idx, b.end_pair_idx) for b in blocks] == [(1, 2), (6, 7)]
        assert "1/2" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert "待处理 2" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert view.diff_block_status.winfo_manager() == "pack"
        assert view.left_ln.get("1.0", "1.end").lstrip().startswith("[1]")
        assert view.left_ln.get("3.0", "3.end").lstrip().startswith("[2]")

        # Hover is a preview only and must not switch the active block counter.
        initial_status = view.diff_block_status_var.get()
        view.hover_pair_idx = 6
        view._update_diff_nav_state()
        assert view.diff_block_status_var.get() == initial_status

        # Row-header hover text and row-only redraw both preserve the marker.
        rn_w = view._rownum_render_width()
        view._set_row_header_text(
            view.left_ln,
            1,
            view._format_main_row_header(1, "A", rn_w, arrow="->"),
        )
        assert view.left_ln.get("1.0", "1.end").lstrip().startswith("[1]")
        assert "blockmarker" in view.left_ln.tag_names("1.0")
        view.refresh(row_only=2, rescan=False)
        assert view.left_ln.get("1.0", "1.end").lstrip().startswith("[1]")

        for widget in (view.left, view.right, view.left_ln, view.right_ln):
            assert int(widget.tag_cget("blockstart", "spacing1")) == 8
            starts = [str(index) for index in widget.tag_ranges("blockstart")]
            assert starts[:2] == ["3.0", "4.0"], (widget, starts)

        # Resolve one member, then adopt the stable block. The resolved member
        # is skipped and the next non-contiguous block must stay untouched.
        assert view._copy_selected_row("B2A", override_pair_idx=1, override_cols={2})
        assert app.ws_a_edit("Data").cell(row=2, column=2).value == "changed-2"
        view._select_line(view.row_to_line[1])
        view.hover_pair_idx = 6
        view._copy_selected_region("B2A")
        assert app.ws_a_edit("Data").cell(row=3, column=2).value == "changed-3"
        assert app.ws_a_edit("Data").cell(row=7, column=2).value == "same-7"
        assert "1/2" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert "待处理 1" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert "已处理" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert view.selected_pair_idx == 1

        view._undo_last_action()
        assert app.ws_a_edit("Data").cell(row=3, column=2).value == "same-3"
        assert [(b.start_pair_idx, b.end_pair_idx) for b in view._ensure_full_diff_blocks()] == [(1, 2), (6, 7)]

        view._sync_main_x_to_frac(0.35)
        _pump(app.root, 0.05)
        saved_x = float(view.left.xview()[0])
        assert saved_x > 0.01, view.left.xview()
        view._goto_next_diff_block()
        assert view.selected_pair_idx == 6, view.selected_pair_idx
        assert "2/2" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert str(view.next_diff_btn.cget("state")) == "disabled"
        assert str(view.prev_diff_btn.cget("state")) == "normal"
        assert abs(float(view.left.xview()[0]) - saved_x) < 0.02, (saved_x, view.left.xview())

        view.only_diff_var.set(0)
        view._refresh_mode_switch_preserving_selection(rescan=False)
        _pump(app.root, 0.05)
        assert not view.diff_block_status.winfo_manager()
        assert not view.left.tag_ranges("blockstart")
        assert not view.left_ln.tag_ranges("blockmarker")
    finally:
        app._shutdown_root()


def _test_three_way_base_only_and_processed_status(root_dir):
    base = os.path.join(root_dir, "three-base.xlsx")
    mine = os.path.join(root_dir, "three-mine.xlsx")
    theirs = os.path.join(root_dir, "three-theirs.xlsx")
    rows_base = [[idx, f"same-{idx}"] for idx in range(1, 9)]
    rows_mine = [list(row) for row in rows_base]
    rows_theirs = [list(row) for row in rows_base]
    rows_base[1][1] = "base-old"
    rows_mine[1][1] = "shared-new"
    rows_theirs[1][1] = "shared-new"
    rows_theirs[6][1] = "theirs-change"
    rows_theirs.append([9, "theirs-only"])
    _make_book(base, rows_base)
    _make_book(mine, rows_mine)
    _make_book(theirs, rows_theirs)

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, base_path=base)
    try:
        view = _prepare_view(app)
        blocks = view._ensure_full_diff_blocks()
        assert [(b.start_pair_idx, b.end_pair_idx) for b in blocks] == [(1, 1), (6, 6), (8, 8)]
        assert view.pair_base_diff_cols.get(1), view.pair_base_diff_cols
        assert view.pair_diff_cols.get(8) == {-1}, view.pair_diff_cols.get(8)
        for widget in (view.left, view.base, view.right, view.left_ln, view.base_ln, view.right_ln):
            assert "blockstart" in widget.tag_names("2.0"), widget
            assert "blockstart" in widget.tag_names("3.0"), widget

        view.selected_pair_idx = 1
        view._last_selected_line = 1
        view.pair_base_diff_cols[1] = set()
        view._invalidate_render_cache()
        view._update_diff_nav_state()
        assert "1/3" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert "待处理 2" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert "已处理" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
    finally:
        app._shutdown_root()


def _test_large_cross_render_limit_navigation(root_dir):
    mine = os.path.join(root_dir, "large-mine.xlsx")
    theirs = os.path.join(root_dir, "large-theirs.xlsx")
    rows_a = []
    rows_b = []
    for idx in range(1, 2101):
        tail = [f"payload-{col}-{idx}-" + ("x" * 24) for col in range(3, 11)]
        rows_a.append([idx, f"same-{idx}"] + tail)
        changed = idx <= 450 or 700 <= idx <= 1100 or 1800 <= idx <= 1801
        rows_b.append([idx, f"changed-{idx}" if changed else f"same-{idx}"] + tail)
    _make_book(mine, rows_a)
    _make_book(theirs, rows_b)

    app = mod.SowMergeApp(mine, theirs)
    try:
        app.root.state("normal")
        app.root.geometry("800x600")
        view = _prepare_view(app)
        blocks = view._ensure_full_diff_blocks()
        assert len(blocks) == 3, [(b.start_pair_idx, b.end_pair_idx) for b in blocks]
        assert blocks[2].start_pair_idx == 1799
        assert len(view._full_display_rows) == 853, len(view._full_display_rows)
        assert len(view.display_rows) < 800, len(view.display_rows)

        rescan_calls = []
        original_refresh = view.refresh

        def _tracked_refresh(row_only, rescan):
            rescan_calls.append(bool(rescan))
            return original_refresh(row_only, rescan)

        view.refresh = _tracked_refresh
        view._sync_main_x_to_frac(0.35)
        _pump(app.root, 0.05)
        saved_x = float(view.left.xview()[0])
        assert saved_x > 0.01, view.left.xview()
        started = time.perf_counter()
        view._goto_full_diff_block(2)
        elapsed = time.perf_counter() - started
        assert view.selected_pair_idx == 1799, view.selected_pair_idx
        assert 1799 in view.row_to_line
        assert len(view.display_rows) == 853, len(view.display_rows)
        assert True not in rescan_calls, rescan_calls
        assert abs(float(view.left.xview()[0]) - saved_x) < 0.02, (saved_x, view.left.xview())
        assert "3/3" in view.diff_block_status_var.get(), view.diff_block_status_var.get()
        assert elapsed < 8.0, elapsed
    finally:
        app._shutdown_root()


def main():
    root_dir = make_temp_dir("sow_diff_blocks_")
    _test_two_way_block_presentation(root_dir)
    _test_three_way_base_only_and_processed_status(root_dir)
    _test_large_cross_render_limit_navigation(root_dir)
    print("GUI_SELF_TEST_DIFF_BLOCKS_OK")


if __name__ == "__main__":
    main()
