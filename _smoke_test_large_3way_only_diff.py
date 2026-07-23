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


def _open_view(base: str, mine: str, theirs: str, merged: str):
    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    _pump(app.root, 1.0)
    app.nb.select(app._sheet_containers["S1"])
    view = None
    for _ in range(220):
        _pump(app.root, 0.05)
        view = app.sheet_views.get("S1")
        if view is not None and getattr(view, "_data_ready", False):
            break
    assert view is not None, "sheet view not created"
    view.force_align_var.set(1)
    view._toggle_force_align()
    return app, view


def _case_theirs_insert_large():
    root_dir = make_temp_dir("sow_large_3way_insert_")
    base = os.path.join(root_dir, "base.xlsx")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")
    merged = os.path.join(root_dir, "merged.xlsx")

    base_rows = [f"R{i:04d}" for i in range(1, 2401)]
    mine_rows = list(base_rows)
    theirs_rows = list(base_rows)
    theirs_rows[751] = "R0752_MOD"
    theirs_rows.insert(752, "R0753_NEW")

    _make_book(base, base_rows)
    _make_book(mine, mine_rows)
    _make_book(theirs, theirs_rows)

    app, view = _open_view(base, mine, theirs, merged)
    try:
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

        targets = []
        for idx in view.display_rows:
            ra, rb = view.row_pairs[idx]
            if (ra is not None and 748 <= ra <= 756) or (rb is not None and 748 <= rb <= 756):
                targets.append((idx, ra, rb, view._base_row_for_pair(idx, (ra, rb)), view.pair_diff_cols.get(idx), view.pair_base_diff_cols.get(idx)))
        assert (751, 752, 752, 752, {1}, set()) in targets, targets
        assert (752, None, 753, None, {-1}, set()) in targets, targets
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass


def _case_base_only_insert_large():
    root_dir = make_temp_dir("sow_large_3way_base_insert_")
    base = os.path.join(root_dir, "base.xlsx")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")
    merged = os.path.join(root_dir, "merged.xlsx")

    base_rows = [f"R{i:04d}" for i in range(1, 2401)]
    mine_rows = list(base_rows)
    theirs_rows = list(base_rows)
    mine_rows.insert(752, "R0753_NEW")
    theirs_rows.insert(752, "R0753_NEW")

    _make_book(base, base_rows)
    _make_book(mine, mine_rows)
    _make_book(theirs, theirs_rows)

    app, view = _open_view(base, mine, theirs, merged)
    try:
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

        targets = []
        for idx in view.display_rows:
            ra, rb = view.row_pairs[idx]
            if (ra is not None and 748 <= ra <= 756) or (rb is not None and 748 <= rb <= 756):
                # Large-sheet async results intentionally keep the exact A/B
                # map sparse: a missing key and an explicit empty set both
                # mean that Mine and Theirs have no row-content difference.
                targets.append((idx, ra, rb, view._base_row_for_pair(idx, (ra, rb)), view.pair_diff_cols.get(idx, set()), view.pair_base_diff_cols.get(idx), view._visual_diff_cols_for_pair(idx)))
        assert (752, 753, 753, None, set(), {-1}, {-1}) in targets, targets
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass


def main():
    _case_theirs_insert_large()
    _case_base_only_insert_large()
    print("SMOKE_LARGE_3WAY_ONLY_DIFF_OK")


if __name__ == "__main__":
    main()
