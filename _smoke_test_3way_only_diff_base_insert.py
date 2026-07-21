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
    root_dir = make_temp_dir("sow_3way_onlydiff_base_insert_")
    base = os.path.join(root_dir, "base.xlsx")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")
    merged = os.path.join(root_dir, "merged.xlsx")

    _make_book(base, ["id", "A"])
    _make_book(mine, ["id", "A", "B"])
    _make_book(theirs, ["id", "A", "B"])

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        _pump(app.root, 1.2)
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
        view.refresh(row_only=None, rescan=True)

        insert_pair = None
        for pair_idx, (ra, rb) in enumerate(view.row_pairs):
            if ra == 3 and rb == 3:
                insert_pair = pair_idx
                break
        assert insert_pair is not None, view.row_pairs
        assert view._base_row_for_pair(insert_pair, view.row_pairs[insert_pair]) is None
        assert not view.pair_diff_cols.get(insert_pair), view.pair_diff_cols.get(insert_pair)
        assert view.pair_base_diff_cols.get(insert_pair) == {-1}, view.pair_base_diff_cols.get(insert_pair)

        view.only_diff_var.set(1)
        view._toggle_only_diff()
        view.refresh(row_only=None, rescan=False)
        _pump(app.root, 0.2)

        assert insert_pair in view.display_rows, view.display_rows
        assert view._pair_has_visual_diff(insert_pair), view._visual_diff_cols_for_pair(insert_pair)

        view.only_diff_var.set(0)
        view._toggle_only_diff()
        _pump(app.root, 0.2)
        assert len(view.display_rows) == len(view.row_pairs), \
            f"Expected full row list to restore immediately after disabling only-diff, got {view.display_rows}"
        assert view.left.get("3.0", "3.end"), "Expected row 3 text to render immediately after disabling only-diff"
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_3WAY_ONLY_DIFF_BASE_INSERT_OK")


if __name__ == "__main__":
    main()
