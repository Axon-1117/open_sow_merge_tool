import os
import time

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_book(path: str, rows: list[object]):
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for idx, value in enumerate(rows, start=1):
        ws.cell(row=idx, column=1).value = value
    wb.save(path)
    wb.close()


def _pump(root, seconds: float):
    end = time.time() + seconds
    while time.time() < end:
        root.update_idletasks()
        root.update()
        time.sleep(0.02)


def main():
    root_dir = make_temp_dir("sow_only_diff_minimap_")
    file_a = os.path.join(root_dir, "a.xlsx")
    file_b = os.path.join(root_dir, "b.xlsx")

    rows_a = [f"R{i:03d}" for i in range(1, 101)]
    rows_b = list(rows_a)
    rows_b.insert(80, "")

    _make_book(file_a, rows_a)
    _make_book(file_b, rows_b)

    app = mod.SowMergeApp(file_a, file_b)
    try:
        try:
            app.root.geometry("1400x900")
        except Exception:
            pass
        _pump(app.root, 1.0)

        app.nb.select(app._sheet_containers["S1"])
        view = None
        for _ in range(200):
            _pump(app.root, 0.05)
            view = app.sheet_views.get("S1")
            if view is not None and getattr(view, "_data_ready", False):
                break

        assert view is not None, "sheet view not created"
        assert view._data_ready, "background diff cache did not become ready"

        insert_pair = None
        for pair_idx, (ra, rb) in enumerate(view.row_pairs):
            if ra is None and rb is not None and rb == 81:
                insert_pair = pair_idx
                break
        assert insert_pair is not None, f"inserted blank row not aligned: {view.row_pairs}"
        assert view.pair_diff_cols.get(insert_pair) == {-1}, view.pair_diff_cols.get(insert_pair)

        view.only_diff_var.set(1)
        view._toggle_only_diff()
        view.refresh(row_only=None, rescan=False)
        _pump(app.root, 0.3)

        assert view.display_rows == [insert_pair], view.display_rows

        view._update_diff_maps()
        _pump(app.root, 0.2)
        canvas_h = max(1, view.vdiff_map.winfo_height())
        red_items = [
            item
            for item in view.vdiff_map.find_all()
            if str(view.vdiff_map.itemcget(item, "fill")).lower() == "#ff2d2d"
        ]
        assert red_items, "diff minimap has no red marker"
        y1 = min(view.vdiff_map.coords(item)[1] for item in red_items)
        assert y1 > canvas_h * 0.6, f"diff marker too high: y1={y1}, h={canvas_h}"
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_ONLY_DIFF_MINIMAP_OK")


if __name__ == "__main__":
    main()
