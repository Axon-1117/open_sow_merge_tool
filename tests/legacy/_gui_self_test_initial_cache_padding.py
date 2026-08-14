"""GUI self-test: initial background cache apply must not paint false grey padding columns."""

import os
import time

from openpyxl import Workbook

from _test_temp_utils import make_temp_dir


def _make_xlsx(path: str, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(r_idx, c_idx).value = value
    wb.save(path)
    wb.close()


def _pump(root, loops: int = 20, delay: float = 0.02):
    for _ in range(loops):
        root.update_idletasks()
        root.update()
        time.sleep(delay)


def main():
    a_rows = [
        ["A", "B", "C", "D", "E"],
        ["row1", "same", "same", "same", "same"],
        ["row2", "same", "same", "same", "same"],
    ]
    b_rows = [
        ["A", "B", "C", "D", "E"],
        ["row1", "same", "same", "same", "same"],
        ["row2", "same", "same", "DIFF", "same"],
    ]

    td1 = make_temp_dir(prefix="sow_cache_pad_a_")
    td2 = make_temp_dir(prefix="sow_cache_pad_b_")
    fa = os.path.join(td1, "same.xlsx")
    fb = os.path.join(td2, "same.xlsx")
    _make_xlsx(fa, a_rows)
    _make_xlsx(fb, b_rows)

    import sow_merge_tool as mod

    app = mod.SowMergeApp(fa, fb)
    try:
        sheet = app.common_sheets[0]
        app.nb.select(app._sheet_containers[sheet])

        view = None
        for _ in range(150):
            _pump(app.root, 2)
            view = app.sheet_views.get(sheet)
            if view is not None and getattr(view, "_data_ready", False):
                break
        assert view is not None, "Expected initial sheet view to be created"
        assert getattr(view, "_data_ready", False), "Expected background cache apply to finish on initial open"
        assert view.col_max_a == 5, f"Expected col_max_a=5 after cache apply, got {view.col_max_a}"
        assert view.col_max_b == 5, f"Expected col_max_b=5 after cache apply, got {view.col_max_b}"

        # Initial open should not mispaint paddingcol grey background on fully populated sheets.
        line_no = 2
        line_left = view.left.get(f"{line_no}.0", f"{line_no}.end")
        line_right = view.right.get(f"{line_no}.0", f"{line_no}.end")
        spans_left = view._spans_for_line(line_left)
        spans_right = view._spans_for_line(line_right)
        for col_idx in (2, 3, 4, 5):
            s_left, e_left = spans_left[col_idx]
            pos_left = s_left + 1 if (e_left - s_left) > 1 else s_left
            tags_left = set(view.left.tag_names(f"{line_no}.{pos_left}"))
            assert "paddingcol" not in tags_left, \
                f"Unexpected left paddingcol on initial open at col {col_idx}: {tags_left}"

            s_right, e_right = spans_right[col_idx]
            pos_right = s_right + 1 if (e_right - s_right) > 1 else s_right
            tags_right = set(view.right.tag_names(f"{line_no}.{pos_right}"))
            assert "paddingcol" not in tags_right, \
                f"Unexpected right paddingcol on initial open at col {col_idx}: {tags_right}"
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("GUI_SELF_TEST_INITIAL_CACHE_PADDING_OK")


if __name__ == "__main__":
    main()
