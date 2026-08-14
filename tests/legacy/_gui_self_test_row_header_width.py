import os
import time

from openpyxl import Workbook

from _test_temp_utils import make_temp_dir


def _make_xlsx(path: str, rows: int):
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for r in range(1, rows + 1):
        ws.cell(row=r, column=1).value = f"R{r}"
    wb.save(path)
    wb.close()


def _pump(root, loops: int = 20, delay: float = 0.01):
    for _ in range(loops):
        root.update_idletasks()
        root.update()
        time.sleep(delay)


def main():
    td = make_temp_dir(prefix="sow_gui_rowhdr_width_")
    fa = os.path.join(td, "a.xlsx")
    fb = os.path.join(td, "b.xlsx")
    _make_xlsx(fa, 12050)
    _make_xlsx(fb, 12050)

    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)
    try:
        app.nb.select(app._sheet_containers["S1"])
        _pump(app.root, 60)
        view = app.sheet_views.get("S1")
        assert view is not None, "Expected sheet view to be created"
        if not getattr(view, "_data_ready", False):
            view.refresh(row_only=None, rescan=True)
            _pump(app.root, 10)

        expected_digits = len(str(12050))
        assert int(view.left_ln.cget("width")) >= expected_digits + 1, view.left_ln.cget("width")
        assert int(view.base_ln.cget("width")) >= expected_digits + 1, view.base_ln.cget("width")
        assert int(view.right_ln.cget("width")) >= expected_digits + 1, view.right_ln.cget("width")
        assert int(view.cursor_cmp_ln.cget("width")) >= expected_digits + 1, view.cursor_cmp_ln.cget("width")

        pair = view.row_pairs[12049]
        view._render_cursor_row_headers(pair, is_three=False)
        _pump(app.root, 5)
        first_line = view.cursor_cmp_ln.get("1.0", "1.end").strip()
        second_line = view.cursor_cmp_ln.get("2.0", "2.end").strip()
        assert first_line == "12050", first_line
        assert second_line == "12050", second_line
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("GUI_SELF_TEST_ROW_HEADER_WIDTH_OK")


if __name__ == "__main__":
    main()
