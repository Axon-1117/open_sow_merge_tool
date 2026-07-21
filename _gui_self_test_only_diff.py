"""GUI self-test: verifies 'Only show diffs' actually filters rows.

Run:
  .venv\\Scripts\\python.exe _gui_self_test_only_diff.py

No desktop automation required.
"""

import os
import time
from openpyxl import Workbook
from _test_temp_utils import make_temp_dir


def _make_xlsx(path: str, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, v in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = v
    wb.save(path)


def _pump(root, loops: int = 20, delay: float = 0.01):
    for _ in range(loops):
        root.update_idletasks()
        root.update()
        time.sleep(delay)


def main():
    # 5 rows, only row 3 differs
    a_rows = [
        ["h1", "h2"],
        [1, 1],
        [2, 2],
        [3, 3],
        [4, 4],
    ]
    b_rows = [
        ["h1", "h2"],
        [1, 1],
        [2, 999],  # diff at row 3
        [3, 3],
        [4, 4],
    ]

    td1 = make_temp_dir(prefix="sow_merge_gui_test_onlydiff_a_")
    td2 = make_temp_dir(prefix="sow_merge_gui_test_onlydiff_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, a_rows)
    _make_xlsx(fb, b_rows)

    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)
    sheet = app.common_sheets[0]
    # Ensure view is created (lazy)
    view = app.sheet_views.get(sheet)
    if view is None:
        # simulate tab selection to trigger lazy creation
        app.nb.select(app._sheet_containers[sheet])
        try:
            app.root.update_idletasks(); app.root.update()
        except Exception:
            pass
        view = app.sheet_views[sheet]

    # Full mode
    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)
    full_count = len(view.display_rows)
    assert full_count == 5, f"Expected 5 rows shown, got {full_count}"

    # Hover a non-diff row so toggle must not keep stale hover-driven C区 state.
    view.hover_pair_idx = view.row_a_to_pair_idx.get(2)
    view.hover_col_idx = 1
    view.hover_side = "A"
    view._last_cursor_cmp_pair_idx = view.hover_pair_idx
    view._update_cursor_lines()
    _pump(app.root, 5)
    before_hover = view.cursor_cmp.get("1.0", "1.end")
    assert "1" in before_hover, f"Expected C区 to show hovered non-diff row before toggle, got {before_hover!r}"

    # Select the diff row and remember a cell selection; only-diff toggle should
    # preserve this logical selection instead of clearing it.
    pair_idx_row3 = view.row_a_to_pair_idx.get(3)
    assert pair_idx_row3 is not None, "Expected row 3 to map to a pair index"
    diff_line_full = view.row_to_line.get(pair_idx_row3)
    assert diff_line_full is not None, "Expected diff row to be visible in full mode"
    view._highlight_selected_line(diff_line_full)
    view.selected_pair_idx = pair_idx_row3
    view.selected_excel_row_a = 3
    view.selected_excel_row_b = 3
    view.selected_excel_row = 3
    view._set_main_selected_cell(diff_line_full, 2)
    view._cursor_cmp_sel_col = 2
    view._cursor_cmp_sel_line = 1
    view._update_cursor_lines()
    _pump(app.root, 5)

    # Only diff
    view.only_diff_var.set(1)
    view._toggle_only_diff()
    _pump(app.root, 10)
    diff_count = len(view.display_rows)
    assert diff_count == 1, f"Expected 1 diff row shown, got {diff_count}; display_rows={view.display_rows}"
    assert view.display_rows[0] == pair_idx_row3, \
        f"Expected diff row pair index {pair_idx_row3}, got {view.display_rows}"
    assert view.has_explicit_cell_selection(), "Expected explicit selection to survive only-diff toggle"
    assert view.selected_pair_idx == pair_idx_row3, \
        f"Expected selected pair to stay on diff row, got {view.selected_pair_idx}"
    diff_line_only = view.row_to_line.get(pair_idx_row3)
    assert diff_line_only == 1, f"Expected diff row to remap to line 1 in only-diff mode, got {diff_line_only}"
    assert view._main_sel_line == diff_line_only and view._main_sel_col == 2, \
        f"Expected main selected cell to remap to visible diff line, got {(view._main_sel_line, view._main_sel_col)}"
    assert view._cursor_cmp_sel_col == 2, f"Expected C区 selected col to stay on 2, got {view._cursor_cmp_sel_col}"
    after_lines = [view.cursor_cmp.get("1.0", "1.end"), view.cursor_cmp.get("2.0", "2.end")]
    assert "2" in after_lines[0] and "999" in after_lines[1], \
        f"Expected C区 to show diff row after toggle, got {after_lines}"

    # Toggle back to full mode: rows should restore immediately without requiring scroll.
    view.only_diff_var.set(0)
    view._toggle_only_diff()
    _pump(app.root, 5)
    assert len(view.display_rows) == 5, \
        f"Expected full rows to restore immediately after disabling only-diff, got {view.display_rows}"
    full_lines = [view.left.get(f"{i}.0", f"{i}.end") for i in range(1, 6)]
    assert "4" in full_lines[4], f"Expected row 5 to render immediately after disabling only-diff, got {full_lines}"

    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_ONLY_DIFF_OK")


if __name__ == "__main__":
    main()
