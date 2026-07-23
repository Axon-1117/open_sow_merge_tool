"""Focused GUI regression for region-mode target resolution and interaction."""

from __future__ import annotations

import os
import time

from openpyxl import Workbook

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


def _make_book(path: str, rows) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    for row in rows:
        worksheet.append(list(row))
    workbook.save(path)
    workbook.close()


def _pump(root, seconds=0.08) -> None:
    deadline = time.time() + seconds
    while time.time() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.005)


def _open_view(mine_rows, theirs_rows):
    root_dir = make_temp_dir("sow_region_mode_interaction_")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")
    _make_book(mine, mine_rows)
    _make_book(theirs, theirs_rows)
    app = smt.SowMergeApp(mine, theirs)
    app.root.deiconify()
    app.root.geometry("900x760")
    app.nb.select(app._sheet_containers["Data"])
    deadline = time.time() + 15.0
    view = None
    while time.time() < deadline:
        _pump(app.root, 0.03)
        view = app.sheet_views.get("Data")
        if view is not None and getattr(view, "_data_ready", False):
            break
    assert view is not None and view._data_ready
    view._suppress_bg_apply = True
    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)
    _pump(app.root)
    return app, view


def _ordinary_diff_rows():
    mine = [(idx, f"same-{idx}") for idx in range(1, 11)]
    theirs = [list(row) for row in mine]
    for excel_row in (2, 3, 7, 8):
        theirs[excel_row - 1][1] = f"changed-{excel_row}"
    return mine, theirs


def _block_pairs(block):
    return tuple(int(value) for value in block.pair_indices)


def _clear_region_anchor(view, *, insertion_pair_idx=None) -> None:
    view.clear_explicit_cell_selection()
    view.selected_pair_idx = None
    view.hover_pair_idx = None
    view.hover_col_idx = None
    view.hover_side = None
    view._last_cursor_cmp_pair_idx = None
    view._last_selected_line = None
    if insertion_pair_idx is not None:
        line = int(view.row_to_line[insertion_pair_idx])
        for widget in (view.left, view.base, view.right):
            widget.mark_set("insert", f"{line}.0")


def _assert_region_buttons_enabled(view) -> None:
    view._set_copy_scope_mode("region")
    assert str(view.use_left_btn.cget("state")) == "normal"
    assert str(view.use_right_btn.cget("state")) == "normal"
    assert "区域" in str(view.use_left_btn.cget("text"))
    assert "区域" in str(view.use_right_btn.cget("text"))


def _column_values(worksheet, column=1):
    return [
        worksheet.cell(row=row, column=column).value
        for row in range(1, worksheet.max_row + 1)
    ]


def test_region_target_resolver_uses_nearest_block_and_earlier_tie_break():
    mine, theirs = _ordinary_diff_rows()
    app, view = _open_view(mine, theirs)
    try:
        assert view._logical_diff_pair_block_for_pair(1) == [1, 2]
        assert view._logical_diff_pair_block_for_pair(6) == [6, 7]

        block, relocated = view._resolve_region_action_target("B2A", 4)
        assert _block_pairs(block) == (1, 2), _block_pairs(block)
        assert relocated is True

        block, relocated = view._resolve_region_action_target("B2A", 6)
        assert _block_pairs(block) == (6, 7), _block_pairs(block)
        assert relocated is False

        block, relocated = view._resolve_region_action_target("B2A", None)
        assert _block_pairs(block) == (1, 2), _block_pairs(block)
        assert relocated is True
    finally:
        app._shutdown_root()


def test_region_target_resolver_filters_blocks_by_copy_direction():
    mine = [("id-1",), ("id-2",), ("mine-only",), ("id-3",), ("id-4",), ("id-5",), ("id-6",)]
    theirs = [("id-1",), ("id-2",), ("id-3",), ("id-4",), ("id-5",), ("theirs-only",), ("id-6",)]
    app, view = _open_view(mine, theirs)
    try:
        mine_only_pair = next(
            idx for idx, (row_a, row_b) in enumerate(view.row_pairs)
            if row_a is not None and row_b is None
        )
        theirs_only_pair = next(
            idx for idx, (row_a, row_b) in enumerate(view.row_pairs)
            if row_a is None and row_b is not None
        )
        block, relocated = view._resolve_region_action_target("B2A", mine_only_pair)
        assert _block_pairs(block) == (theirs_only_pair,)
        assert relocated is True
        block, relocated = view._resolve_region_action_target("A2B", theirs_only_pair)
        assert _block_pairs(block) == (mine_only_pair,)
        assert relocated is True
        _assert_region_buttons_enabled(view)
    finally:
        app._shutdown_root()


def test_explicit_applicable_region_writes_immediately():
    mine, theirs = _ordinary_diff_rows()
    app, view = _open_view(mine, theirs)
    try:
        _assert_region_buttons_enabled(view)
        target_pair = 1
        view._select_line(view.row_to_line[target_pair])
        undo_before = len(app.undo_stack)

        view._run_copy_action_by_mode("B2A")
        _pump(app.root)

        assert view.selected_pair_idx == target_pair
        assert app.ws_a_edit("Data").cell(2, 2).value == "changed-2"
        assert app.ws_a_edit("Data").cell(3, 2).value == "changed-3"
        assert len(app.undo_stack) == undo_before + 1
        assert not view.pair_diff_cols.get(1) and not view.pair_diff_cols.get(2)
        assert view.pair_diff_cols.get(6) and view.pair_diff_cols.get(7)
    finally:
        app._shutdown_root()


def test_explicit_direction_inapplicable_region_never_crosses_block():
    mine = [("id-1",), ("id-2",), ("mine-only",), ("id-3",), ("id-4",), ("id-5",), ("id-6",)]
    theirs = [("id-1",), ("id-2",), ("id-3",), ("id-4",), ("id-5",), ("theirs-only",), ("id-6",)]
    app, view = _open_view(mine, theirs)
    try:
        _assert_region_buttons_enabled(view)
        mine_only_pair = next(
            idx for idx, (row_a, row_b) in enumerate(view.row_pairs)
            if row_a is not None and row_b is None
        )
        theirs_only_pair = next(
            idx for idx, (row_a, row_b) in enumerate(view.row_pairs)
            if row_a is None and row_b is not None
        )
        mine_before = _column_values(app.ws_a_edit("Data"))
        undo_before = len(app.undo_stack)

        # A deliberate selection in a direction-inapplicable visual block is
        # authoritative: B2A must not silently jump to the later theirs-only block.
        view._select_line(view.row_to_line[mine_only_pair])
        view.info.configure(text="")
        view._run_copy_action_by_mode("B2A")
        _pump(app.root)
        assert view.selected_pair_idx == mine_only_pair
        assert view.selected_pair_idx != theirs_only_pair
        assert _column_values(app.ws_a_edit("Data")) == mine_before
        assert len(app.undo_stack) == undo_before
        message = str(view.info.cget("text"))
        assert "区域" in message and "不能" in message, message

        # With no explicit selection, fallback is allowed, but the first click
        # only locates and prompts. The second click performs the insertion.
        _clear_region_anchor(view, insertion_pair_idx=mine_only_pair)
        view.info.configure(text="")
        view._run_copy_action_by_mode("B2A")
        _pump(app.root)
        assert view.selected_pair_idx == theirs_only_pair
        assert _column_values(app.ws_a_edit("Data")) == mine_before
        assert len(app.undo_stack) == undo_before
        message = str(view.info.cget("text"))
        assert "区域" in message and "再次" in message, message

        view._run_copy_action_by_mode("B2A")
        _pump(app.root)
        assert "theirs-only" in _column_values(app.ws_a_edit("Data"))
        assert len(app.undo_stack) == undo_before + 1
    finally:
        app._shutdown_root()


def test_region_fallback_first_click_only_locates_second_click_applies_and_undo_reapplies():
    mine, theirs = _ordinary_diff_rows()
    app, view = _open_view(mine, theirs)
    try:
        _assert_region_buttons_enabled(view)
        # Pair 5 is an equal row nearest to the later block (6, 7).
        _clear_region_anchor(view, insertion_pair_idx=5)
        undo_before = len(app.undo_stack)
        view.info.configure(text="")
        view._run_copy_action_by_mode("B2A")
        _pump(app.root)

        target_pair = 6
        target_line = int(view.row_to_line[target_pair])
        assert view.selected_pair_idx == target_pair
        assert view.hover_pair_idx == target_pair
        assert view._last_cursor_cmp_pair_idx == target_pair
        for widget in (view.left, view.right):
            assert int(str(widget.index("insert")).split(".")[0]) == target_line
            assert widget.bbox(f"{target_line}.0") is not None
        assert "changed-7" in view.cursor_cmp.get("1.0", "end")
        assert app.ws_a_edit("Data").cell(7, 2).value == "same-7"
        assert app.ws_a_edit("Data").cell(8, 2).value == "same-8"
        assert view.pair_diff_cols.get(1) and view.pair_diff_cols.get(2)
        assert view.pair_diff_cols.get(6) and view.pair_diff_cols.get(7)
        assert len(app.undo_stack) == undo_before
        message = str(view.info.cget("text"))
        assert "区域" in message and "再次" in message, message

        # The located block is now an explicit selection, so the second click
        # applies it without another relocation or intermediate undo entry.
        view._run_copy_action_by_mode("B2A")
        _pump(app.root)
        assert view.selected_pair_idx == target_pair
        assert app.ws_a_edit("Data").cell(7, 2).value == "changed-7"
        assert app.ws_a_edit("Data").cell(8, 2).value == "changed-8"
        assert not view.pair_diff_cols.get(6) and not view.pair_diff_cols.get(7)
        assert len(app.undo_stack) == undo_before + 1

        view._undo_last_action()
        _pump(app.root)
        assert app.ws_a_edit("Data").cell(7, 2).value == "same-7"
        assert app.ws_a_edit("Data").cell(8, 2).value == "same-8"
        assert view.pair_diff_cols.get(6) and view.pair_diff_cols.get(7)

        # Reapply is the existing redo-equivalent interaction; it must keep the
        # selected target and must not jump to the earlier still-different block.
        view._run_copy_action_by_mode("B2A")
        _pump(app.root)
        assert view.selected_pair_idx == target_pair
        assert app.ws_a_edit("Data").cell(7, 2).value == "changed-7"
        assert app.ws_a_edit("Data").cell(8, 2).value == "changed-8"
        assert len(app.undo_stack) == undo_before + 1
        _assert_region_buttons_enabled(view)
    finally:
        app._shutdown_root()


def test_region_action_with_no_applicable_diff_is_nonmodal_silent_noop():
    rows = [(idx, f"same-{idx}") for idx in range(1, 6)]
    app, view = _open_view(rows, rows)
    original_showerror = smt.messagebox.showerror
    bells = []
    errors = []
    try:
        _assert_region_buttons_enabled(view)
        view.root.bell = lambda: bells.append(True)
        smt.messagebox.showerror = lambda *args, **kwargs: errors.append((args, kwargs))
        _clear_region_anchor(view, insertion_pair_idx=0)
        undo_before = list(app.undo_stack)
        mine_before = [
            app.ws_a_edit("Data").cell(row=row, column=2).value
            for row in range(1, 6)
        ]

        assert view._resolve_region_action_target("B2A", None) is None
        view.info.configure(text="")
        view._run_copy_action_by_mode("B2A")
        _pump(app.root)
        message = str(view.info.cget("text"))
        assert "区域" in message and ("无" in message or "没有" in message), message
        assert bells == [], "no-applicable-region feedback must not ring the bell"
        assert errors == [], errors
        assert app.undo_stack == undo_before
        assert [
            app.ws_a_edit("Data").cell(row=row, column=2).value
            for row in range(1, 6)
        ] == mine_before
        _assert_region_buttons_enabled(view)
    finally:
        smt.messagebox.showerror = original_showerror
        app._shutdown_root()


def main():
    tests = (
        test_region_target_resolver_uses_nearest_block_and_earlier_tie_break,
        test_region_target_resolver_filters_blocks_by_copy_direction,
        test_explicit_applicable_region_writes_immediately,
        test_explicit_direction_inapplicable_region_never_crosses_block,
        test_region_fallback_first_click_only_locates_second_click_applies_and_undo_reapplies,
        test_region_action_with_no_applicable_diff_is_nonmodal_silent_noop,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"PASS: region mode interaction regression ({len(tests)} tests)")


if __name__ == "__main__":
    main()
