import os
import time

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_book(path: str, rows: list[list[object]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lang"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value
    wb.save(path)
    wb.close()


def _pump(root, seconds: float):
    end = time.time() + seconds
    while time.time() < end:
        root.update_idletasks()
        root.update()
        time.sleep(0.02)


def main():
    root_dir = make_temp_dir("sow_large_3way_merge_open_")
    base = os.path.join(root_dir, "base.xlsx")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")
    merged = os.path.join(root_dir, "merged.xlsx")

    row_count = 6200
    header = ["id", "text", "type", "group", "note"]
    rows_base = [header]
    rows_mine = [header.copy()]
    rows_theirs = [header.copy()]

    for idx in range(1, row_count + 1):
        row = [f"k_{idx}", f"text_{idx}", "lang", f"g_{idx % 5}", f"n_{idx}"]
        rows_base.append(list(row))
        rows_mine.append(list(row))
        rows_theirs.append(list(row))

    # base-only diff: mine/theirs stay equal, but differ from base
    base_only_excel_row = 5901
    rows_base[base_only_excel_row - 1][1] = "base_old_only"
    rows_mine[base_only_excel_row - 1][1] = "shared_new_only"
    rows_theirs[base_only_excel_row - 1][1] = "shared_new_only"

    # A/B direct diff
    ab_diff_excel_row = 6001
    rows_mine[ab_diff_excel_row - 1][4] = "mine_only_value"
    rows_theirs[ab_diff_excel_row - 1][4] = "theirs_only_value"

    _make_book(base, rows_base)
    _make_book(mine, rows_mine)
    _make_book(theirs, rows_theirs)

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        _pump(app.root, 0.5)
        app.nb.select(app._sheet_containers["Lang"])
        _pump(app.root, 0.1)
        view = app.sheet_views["Lang"]

        view.only_diff_var.set(1)
        view._data_ready = False
        view.pair_text_a = {}
        view.pair_text_b = {}
        view.pair_diff_cols = {}
        view.pair_base_diff_cols = {}

        t0 = time.time()
        view.refresh(row_only=None, rescan=True)
        refresh_sec = time.time() - t0

        base_only_pair = view.row_a_to_pair_idx.get(base_only_excel_row)
        ab_pair = view.row_a_to_pair_idx.get(ab_diff_excel_row)
        assert base_only_pair is not None, "Expected base-only row pair to exist"
        assert ab_pair is not None, "Expected A/B diff row pair to exist"
        assert base_only_pair in view._full_display_rows, "Expected base-only diff row in only-diff rows"
        assert ab_pair in view._full_display_rows, "Expected A/B diff row in only-diff rows"
        assert view.pair_base_diff_cols.get(base_only_pair) == {2}, view.pair_base_diff_cols.get(base_only_pair)
        assert view.pair_diff_cols.get(base_only_pair) in (set(), None), view.pair_diff_cols.get(base_only_pair)
        assert view.pair_diff_cols.get(ab_pair) == {5}, view.pair_diff_cols.get(ab_pair)
        assert refresh_sec < 20.0, f"Expected large-sheet 3-way only-diff refresh to stay below 20s, got {refresh_sec:.2f}s"
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_LARGE_3WAY_MERGE_OPEN_OK")


if __name__ == "__main__":
    main()
