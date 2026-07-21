"""GUI self-test: verifies main pane diffcell highlight tags on 2-way and 3-way views."""

import os
import time

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_xlsx(path: str, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value
    wb.save(path)
    wb.close()


def _pump(root, seconds: float):
    end = time.time() + seconds
    while time.time() < end:
        root.update_idletasks()
        root.update()
        time.sleep(0.02)


def _ensure_view(app, sheet: str):
    app.nb.select(app._sheet_containers[sheet])
    view = None
    for _ in range(200):
        _pump(app.root, 0.05)
        view = app.sheet_views.get(sheet)
        if view is not None and getattr(view, "_data_ready", False):
            return view
    raise AssertionError(f"sheet view not ready: {sheet}")


def _assert_diffcell_priority(widget, line_no: int, start_col: int):
    tags = list(widget.tag_names(f"{line_no}.{start_col}"))
    assert "diffcell" in tags, f"missing diffcell at {line_no}.{start_col}: tags={tags}"
    assert "diffrow" not in tags, f"diffrow should be cleared under diffcell at {line_no}.{start_col}: tags={tags}"


def _find_pair_line(view, ra: int | None, rb: int | None):
    for pair_idx, pair in enumerate(view.row_pairs):
        if pair == (ra, rb):
            line_no = view.row_to_line.get(pair_idx)
            if line_no is not None:
                return pair_idx, line_no
    raise AssertionError(f"pair not found: {(ra, rb)} in {view.row_pairs}")


def _turn_on_grid(view):
    try:
        view.grid_overlay_var.set(1)
        view._toggle_grid_overlay()
    except Exception:
        pass


def _assert_main_diffcell_for_col(view, line_no: int, col_idx: int, *, include_base: bool):
    left_line = view.left.get(f"{line_no}.0", f"{line_no}.end")
    right_line = view.right.get(f"{line_no}.0", f"{line_no}.end")
    left_spans = view._spans_for_line(left_line)
    right_spans = view._spans_for_line(right_line)
    assert col_idx in left_spans, f"left span missing col {col_idx}: {left_spans}"
    assert col_idx in right_spans, f"right span missing col {col_idx}: {right_spans}"
    left_start = left_spans[col_idx][0]
    right_start = right_spans[col_idx][0]
    _assert_diffcell_priority(view.left, line_no, left_start)
    _assert_diffcell_priority(view.right, line_no, right_start)
    if include_base:
        base_line = view.base.get(f"{line_no}.0", f"{line_no}.end")
        base_spans = view._spans_for_line(base_line)
        assert col_idx in base_spans, f"base span missing col {col_idx}: {base_spans}"
        base_start = base_spans[col_idx][0]
        _assert_diffcell_priority(view.base, line_no, base_start)


def _case_2way():
    td_a = make_temp_dir("sow_gui_main_diffcell_a_")
    td_b = make_temp_dir("sow_gui_main_diffcell_b_")
    file_a = os.path.join(td_a, "same.xlsx")
    file_b = os.path.join(td_b, "same.xlsx")
    rows_a = [
        ["A", "B", "C", "D", "E"],
        ["world_boss_box_guaranteed_num", "int32", "desc", "50", ""],
    ]
    rows_b = [
        ["A", "B", "C", "D", "E"],
        ["world_boss_box_guaranteed_num", "int32", "desc", "67", ""],
    ]
    _make_xlsx(file_a, rows_a)
    _make_xlsx(file_b, rows_b)

    app = mod.SowMergeApp(file_a, file_b, raw_base=file_a, raw_mine=file_b)
    try:
        view = _ensure_view(app, "S")
        _turn_on_grid(view)
        view.only_diff_var.set(0)
        view.refresh(row_only=None, rescan=True)
        _pump(app.root, 0.2)
        _pair_idx, line_no = _find_pair_line(view, 2, 2)
        _assert_main_diffcell_for_col(view, line_no, 4, include_base=False)
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass


def _case_3way():
    td = make_temp_dir("sow_gui_main_diffcell_3way_")
    base = os.path.join(td, "base.xlsx")
    mine = os.path.join(td, "mine.xlsx")
    theirs = os.path.join(td, "theirs.xlsx")
    merged = os.path.join(td, "merged.xlsx")
    rows_base = [
        ["A", "B", "C", "D", "E"],
        ["world_boss_box_guaranteed_num", "int32", "desc", "50", ""],
    ]
    rows_mine = [
        ["A", "B", "C", "D", "E"],
        ["world_boss_box_guaranteed_num", "int32", "desc", "67", ""],
    ]
    rows_theirs = [
        ["A", "B", "C", "D", "E"],
        ["world_boss_box_guaranteed_num", "int32", "desc", "67", ""],
    ]
    _make_xlsx(base, rows_base)
    _make_xlsx(mine, rows_mine)
    _make_xlsx(theirs, rows_theirs)

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        view = _ensure_view(app, "S")
        _turn_on_grid(view)
        view.refresh(row_only=None, rescan=True)
        _pump(app.root, 0.2)
        _pair_idx, line_no = _find_pair_line(view, 2, 2)
        _assert_main_diffcell_for_col(view, line_no, 4, include_base=True)
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass


def main():
    _case_2way()
    _case_3way()
    print("GUI_SELF_TEST_MAIN_DIFFCELL_HIGHLIGHT_OK")


if __name__ == "__main__":
    main()
