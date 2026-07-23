"""GUI self-test: C-area uses full width and keeps its own horizontal scroll.

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
    app._intended_window_state = "normal"
    app.root.state("normal")
    app.root.geometry("1200x700")
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

    c_width = int(view.cursor_cmp.winfo_width())
    header_width = int(view.cursor_cmp_colhdr.winfo_width())
    main_width = int(view.left.winfo_width())
    available_width = int(view.c_area.winfo_width()) - int(view.cursor_cmp_ln.winfo_width())
    assert c_width >= available_width - 12, (
        f"C body did not consume lower-pane width: c={c_width} available={available_width}"
    )
    assert c_width > main_width * 1.5, (
        f"C body should be substantially wider than one main pane: c={c_width} main={main_width}"
    )
    assert abs(header_width - c_width) <= 4, (
        f"C header/body widths diverged: header={header_width} body={c_width}"
    )

    view._sync_main_x_to_frac(0.6)
    view._sync_c_x_to_frac(0.6)
    for _ in range(4):
        app.root.update_idletasks()
        app.root.update()

    left_first = float((view.left.xview() or (0.0, 1.0))[0])
    cursor_first = float((view.cursor_hsb.get() or (0.0, 1.0))[0])
    cursor_view_first = float((view.cursor_cmp.xview() or (0.0, 1.0))[0])
    cell_first = float((view.cell_cmp_hsb.get() or (0.0, 1.0))[0])
    assert abs(left_first - 0.6) < 0.03, f"Expected main pane near 0.6, got {left_first:.6f}"
    assert abs(cursor_first - cursor_view_first) < 0.03, (
        "Expected cursor_hsb to report C's wider viewport; "
        f"c_view={cursor_view_first:.6f} cursor={cursor_first:.6f}"
    )
    assert abs(cell_first - left_first) < 0.03, (
        f"Expected cell_cmp_hsb to mirror main xview; left={left_first:.6f} cell={cell_first:.6f}"
    )

    main_before_c_drive = float((view.left.xview() or (0.0, 1.0))[0])
    view._xview_cursor_cmp("moveto", "0.2")
    for _ in range(4):
        app.root.update_idletasks()
        app.root.update()
    after_cursor_drive = float((view.left.xview() or (0.0, 1.0))[0])
    c_after_drive = float((view.cursor_cmp.xview() or (0.0, 1.0))[0])
    assert abs(after_cursor_drive - main_before_c_drive) < 0.01, (
        "C scrollbar unexpectedly moved the main pane; "
        f"before={main_before_c_drive:.6f} after={after_cursor_drive:.6f}"
    )
    assert abs(c_after_drive - 0.2) < 0.03, (
        f"Expected C scrollbar to move C itself; got {c_after_drive:.6f}"
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
