"""GUI self-test: C-area horizontal scrollbars must reflect and drive main xview.

Run:
  .venv\\Scripts\\python.exe _gui_self_test_c_hscroll_bar.py
"""

import os
import sys
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


def main():
    cols = 20
    header = [f"H{i}" for i in range(1, cols + 1)]
    row_a = [f"A{i:02d}-" + ("x" * 40) for i in range(1, cols + 1)]
    row_b = list(row_a)
    row_b[10] = "DIFF-" + ("y" * 80)

    td1 = make_temp_dir(prefix="sow_merge_gui_test_c_hscroll_a_")
    td2 = make_temp_dir(prefix="sow_merge_gui_test_c_hscroll_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, [header, row_a])
    _make_xlsx(fb, [header, row_b])

    sys.path.insert(0, r"D:\Tools\sow_merge_tool_proj")
    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)
    app.root.geometry("700x450")
    sheet = app.common_sheets[0]
    view = app.sheet_views.get(sheet)
    if view is None:
        app.nb.select(app._sheet_containers[sheet])
        app.root.update_idletasks()
        app.root.update()
        view = app.sheet_views[sheet]

    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)
    for _ in range(8):
        app.root.update_idletasks()
        app.root.update()

    view._sync_main_x_to_frac(0.6)
    view._sync_c_x_to_frac(0.6)
    for _ in range(4):
        app.root.update_idletasks()
        app.root.update()

    left_first = float((view.left.xview() or (0.0, 1.0))[0])
    cursor_first = float((view.cursor_hsb.get() or (0.0, 1.0))[0])
    cell_first = float((view.cell_cmp_hsb.get() or (0.0, 1.0))[0])
    assert abs(left_first - 0.6) < 0.03, f"Expected main pane near 0.6, got {left_first:.6f}"
    assert abs(cursor_first - left_first) < 0.03, (
        f"Expected cursor_hsb to mirror main xview; left={left_first:.6f} cursor={cursor_first:.6f}"
    )
    assert abs(cell_first - left_first) < 0.03, (
        f"Expected cell_cmp_hsb to mirror main xview; left={left_first:.6f} cell={cell_first:.6f}"
    )

    view._xview_cursor_cmp("moveto", "0.2")
    for _ in range(4):
        app.root.update_idletasks()
        app.root.update()
    after_cursor_drive = float((view.left.xview() or (0.0, 1.0))[0])
    assert abs(after_cursor_drive - 0.2) < 0.03, (
        f"Expected C top scrollbar to drive main xview; got {after_cursor_drive:.6f}"
    )

    view._xview_cell_cmp("moveto", "0.4")
    for _ in range(4):
        app.root.update_idletasks()
        app.root.update()
    after_cell_drive = float((view.left.xview() or (0.0, 1.0))[0])
    assert abs(after_cell_drive - 0.4) < 0.03, (
        f"Expected C bottom scrollbar to drive main xview; got {after_cell_drive:.6f}"
    )

    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_C_HSCROLL_BAR_OK")


if __name__ == "__main__":
    main()
