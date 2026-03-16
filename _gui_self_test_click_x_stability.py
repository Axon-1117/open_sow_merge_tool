"""GUI self-test: clicking a far-right visible cell must not reset horizontal xview.

Run:
  .venv\\Scripts\\python.exe _gui_self_test_click_x_stability.py
"""

import os
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
    # Build wide rows with a diff at a high column (W=23) to mimic the bug scenario.
    cols = 30
    header = [f"H{i}" for i in range(1, cols + 1)]
    row_a = [f"A{i:02d}-" + ("x" * 8) for i in range(1, cols + 1)]
    row_b = list(row_a)
    row_b[22] = "DIFF-W-COL"  # col 23

    td1 = make_temp_dir(prefix="sow_merge_gui_test_click_x_a_")
    td2 = make_temp_dir(prefix="sow_merge_gui_test_click_x_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, [header, row_a])
    _make_xlsx(fb, [header, row_b])

    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)
    sheet = app.common_sheets[0]
    view = app.sheet_views.get(sheet)
    if view is None:
        app.nb.select(app._sheet_containers[sheet])
        try:
            app.root.update_idletasks(); app.root.update()
        except Exception:
            pass
        view = app.sheet_views[sheet]

    view.only_diff_var.set(0)
    view.refresh(row_only=None, rescan=True)

    try:
        app.root.update_idletasks(); app.root.update()
    except Exception:
        pass

    # Ensure target line is materialized in viewport before reading bbox.
    try:
        view.left.see("2.0")
        view.right.see("2.0")
        app.root.update_idletasks(); app.root.update()
    except Exception:
        pass

    # Move to a non-zero horizontal position first.
    view._sync_main_x_to_frac(0.55)
    view._sync_c_x_to_frac(0.55)
    try:
        app.root.update_idletasks(); app.root.update()
    except Exception:
        pass

    before = float((view.left.xview() or (0.0, 1.0))[0])

    # Simulate a click near the right side; avoid bbox dependency in headless Tk runs.
    click_x = max(10, int(view.left.winfo_width()) - 30)
    click_y = 5

    class E:
        pass

    e = E()
    e.x = int(click_x)
    e.y = int(click_y)

    clicked_idx = view.left.index(f"@{e.x},{e.y}")
    clicked_col = int(str(clicked_idx).split(".")[1])
    if clicked_col <= 2:
        raise AssertionError(f"Expected far-right click column > 2, got {clicked_idx} (x={e.x}, y={e.y})")

    # Go through the actual click handler path used by Button-1 binding.
    view._on_click_with_arrow(view.left, e, "A2B")
    # Real UI also runs ButtonRelease callback; mimic it.
    view._update_cursor_lines()

    try:
        app.root.update_idletasks(); app.root.update()
    except Exception:
        pass

    after_left = float((view.left.xview() or (0.0, 1.0))[0])
    after_right = float((view.right.xview() or (0.0, 1.0))[0])
    after_c = float((view.cursor_cmp.xview() or (0.0, 1.0))[0])

    assert abs(after_left - before) < 0.03, (
        f"xview regressed after click: before={before:.6f} after_left={after_left:.6f}"
    )
    assert abs(after_right - after_left) < 0.02, (
        f"main panes out of sync: left={after_left:.6f} right={after_right:.6f}"
    )
    assert abs(after_c - after_left) < 0.03, (
        f"C pane out of sync: left={after_left:.6f} c={after_c:.6f}"
    )

    try:
        app.root.destroy()
    except Exception:
        pass

    print("GUI_SELF_TEST_CLICK_X_STABILITY_OK")


if __name__ == "__main__":
    main()
