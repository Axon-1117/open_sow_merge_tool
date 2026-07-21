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
    root_dir = make_temp_dir("sow_large_onlydiff_insert_")
    file_a = os.path.join(root_dir, "a.xlsx")
    file_b = os.path.join(root_dir, "b.xlsx")

    rows_a = [f"R{i:04d}" for i in range(1, 2401)]
    rows_b = list(rows_a)
    rows_b[751] = "R0752_MOD"
    rows_b.insert(752, "R0753_NEW")

    _make_book(file_a, rows_a)
    _make_book(file_b, rows_b)

    app = mod.SowMergeApp(file_a, file_b)
    try:
        _pump(app.root, 1.0)
        app.nb.select(app._sheet_containers["S1"])
        view = None
        for _ in range(200):
            _pump(app.root, 0.05)
            view = app.sheet_views.get("S1")
            if view is not None and getattr(view, "_data_ready", False):
                break

        assert view is not None, "sheet view not created"
        view.force_align_var.set(1)
        view._toggle_force_align()
        before_toggle_rows = list(view.display_rows)
        view.only_diff_var.set(1)
        t0 = time.time()
        view._toggle_only_diff()
        toggle_cost = time.time() - t0
        assert toggle_cost < 1.5, f"Expected only-diff toggle to stay responsive, got {toggle_cost:.3f}s"
        assert list(view.display_rows) == before_toggle_rows, "Expected full view to stay visible while async only-diff builds"
        for _ in range(300):
            _pump(app.root, 0.05)
            if (not getattr(view, "_only_diff_async_building", False)) and len(view.display_rows) <= 10:
                break
        assert not getattr(view, "_only_diff_async_building", False), "Expected async only-diff build to finish"

        mod_pair = None
        insert_pair = None
        for pair_idx, (ra, rb) in enumerate(view.row_pairs):
            if ra == 752 and rb == 752:
                mod_pair = pair_idx
            if ra is None and rb == 753:
                insert_pair = pair_idx
        assert mod_pair is not None, view.row_pairs
        assert insert_pair is not None, view.row_pairs
        assert view.pair_diff_cols.get(mod_pair), view.pair_diff_cols.get(mod_pair)
        assert view.pair_diff_cols.get(insert_pair) == {-1}, view.pair_diff_cols.get(insert_pair)
        assert mod_pair in view.display_rows, view.display_rows[:20]
        assert insert_pair in view.display_rows, view.display_rows[:20]

        view.only_diff_var.set(0)
        view._toggle_only_diff()
        _pump(app.root, 0.1)
        view.only_diff_var.set(1)
        t1 = time.time()
        view._toggle_only_diff()
        cached_toggle_cost = time.time() - t1
        assert cached_toggle_cost < 0.5, f"Expected cached only-diff toggle to be near-instant, got {cached_toggle_cost:.3f}s"
        _pump(app.root, 0.1)
        assert mod_pair in view.display_rows and insert_pair in view.display_rows, view.display_rows[:20]
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_LARGE_ONLY_DIFF_ROW_INSERT_OK")


if __name__ == "__main__":
    main()
