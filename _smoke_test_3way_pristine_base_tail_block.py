import os
import time

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_book(path: str, rows: list[list[object]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value
    wb.save(path)
    wb.close()


def _pump(root, loops: int = 20, delay: float = 0.02):
    for _ in range(loops):
        root.update_idletasks()
        root.update()
        time.sleep(delay)


def main():
    root_dir = make_temp_dir("sow_pristine_tail_block_")
    base = os.path.join(root_dir, "base.xlsx")
    mine = os.path.join(root_dir, "mine.xlsx")
    theirs = os.path.join(root_dir, "theirs.xlsx")

    base_rows = [["id", "name"], [1, "a"], [2, "b"], [3, "c"], [4, "d"], [5, "e"]]
    mine_rows = [["id", "name"], [1, "a"], [2, "b"], [3, "c"], ["local", "mine-only"]]
    theirs_rows = [["id", "name"], [1, "a"], [2, "b"], [3, "c"], [4, "d"], [5, "e"]]

    _make_book(base, base_rows)
    _make_book(mine, mine_rows)
    _make_book(theirs, theirs_rows)

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, base_path=base)
    try:
        view = app.sheet_views.get("S1")
        if view is None:
            app.nb.select(app._sheet_containers["S1"])
            _pump(app.root, 40)
            view = app.sheet_views.get("S1")
        assert view is not None, "Expected S1 view"
        if not getattr(view, "_data_ready", False):
            view.refresh(row_only=None, rescan=True)
            _pump(app.root, 10)

        expected_pairs = [(1, 1), (2, 2), (3, 3), (4, 4), (None, 5), (5, None), (None, 6)]
        assert view.row_pairs == expected_pairs, view.row_pairs

        payload = view._cmp_tooltip_payload_by_pair_col(4, 1, force_panel=True)
        assert payload is not None, "Expected tooltip payload for theirs-only row"
        text = payload[0]
        assert "base[5]: 4" in text, text
        assert "mine[-]: <missing>" in text, text
        assert "theirs[5]: 4" in text, text
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_3WAY_PRISTINE_BASE_TAIL_BLOCK_OK")


if __name__ == "__main__":
    main()
