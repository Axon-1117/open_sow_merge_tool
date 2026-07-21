import os
import time

from openpyxl import Workbook, load_workbook

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


def main():
    root = make_temp_dir("sow_manual_row_insert_")
    base = os.path.join(root, "base.xlsx")
    mine = os.path.join(root, "mine.xlsx")
    theirs = os.path.join(root, "theirs.xlsx")
    merged = os.path.join(root, "merged.xlsx")

    _make_book(base, [["id"], ["A"], ["C"]])
    _make_book(mine, [["id"], ["A"], ["C"]])
    _make_book(theirs, [["id"], ["A"], ["B"], ["C"]])

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        for _ in range(200):
            if app._edit_loaded_event.is_set():
                break
            app.root.update_idletasks()
            app.root.update()
            time.sleep(0.02)

        view = app.sheet_views.get("S1")
        if view is None:
            app.nb.select(app._sheet_containers["S1"])
            for _ in range(50):
                app.root.update_idletasks()
                app.root.update()
                time.sleep(0.02)
            view = app.sheet_views["S1"]

        view.force_align_var.set(1)
        view._toggle_force_align()
        view.refresh(row_only=None, rescan=True)

        insert_pair = None
        for pair_idx, (ra, rb) in enumerate(view.row_pairs):
            if ra is None and rb is not None and app.ws_b_val("S1").cell(rb, 1).value == "B":
                insert_pair = pair_idx
                break

        assert insert_pair is not None, f"did not find inserted row: {view.row_pairs}"
        assert view._copy_selected_row("B2A", override_pair_idx=insert_pair), "row insert copy failed"

        out = app.build_manual_merge_output_file()
        wb = load_workbook(out, data_only=False)
        try:
            ws = wb["S1"]
            values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
            assert values == ["id", "A", "B", "C"], f"unexpected merged rows: {values}"
        finally:
            wb.close()
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_MANUAL_ROW_INSERT_OK")


if __name__ == "__main__":
    main()
