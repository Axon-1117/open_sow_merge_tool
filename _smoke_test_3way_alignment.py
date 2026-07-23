import os
import time

from openpyxl import Workbook, load_workbook

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


def main():
    root = make_temp_dir("sow_3way_align_")

    base = os.path.join(root, "base.xlsx")
    mine = os.path.join(root, "mine.xlsx")
    theirs = os.path.join(root, "theirs.xlsx")
    merged = os.path.join(root, "merged.xlsx")
    _make_book(base, ["id", "A", "C"])
    _make_book(mine, ["id", "A", "B", "C"])
    _make_book(theirs, ["id", "A", "X"])

    conflicts, cmap = mod._scan_three_way_conflicts(base, mine, theirs)
    assert not conflicts, (conflicts, cmap)

    _make_book(theirs, ["id", "A", "B", "C"])
    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        for _ in range(60):
            app.root.update_idletasks()
            app.root.update()
            time.sleep(0.02)

        app.nb.select(app._sheet_containers["S1"])
        for _ in range(60):
            app.root.update_idletasks()
            app.root.update()
            time.sleep(0.02)

        view = app.sheet_views["S1"]
        view.force_align_var.set(1)
        view._toggle_force_align()
        view.refresh(row_only=None, rescan=True)

        insert_pair = None
        for pair_idx, (ra, rb) in enumerate(view.row_pairs):
            if ra is not None and app.ws_a_val("S1").cell(ra, 1).value == "B":
                insert_pair = pair_idx
                break
        assert insert_pair is not None, view.row_pairs
        base_line = view._build_base_line(insert_pair).strip()
        assert "此侧" in base_line, repr(base_line)

        # The C-area cell action must use mine->base row mapping. Mine row 4
        # corresponds to base row 3 after the inserted B row.
        mapped_pair = view.row_a_to_pair_idx[4]
        app.ws_a_val("S1").cell(row=4, column=1).value = "wrong"
        app.ws_a_edit("S1").cell(row=4, column=1).value = "wrong"
        view._copy_single_cell_by_pair(mapped_pair, "BASE2A", 1)
        assert app.ws_a_val("S1").cell(row=4, column=1).value == "C"

        assert view._copy_selected_row("BASE2A", override_pair_idx=insert_pair), "BASE2A delete failed"
        out = app.build_manual_merge_output_file()
        wb = load_workbook(out, data_only=False)
        try:
            ws = wb["S1"]
            values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
            assert values == ["id", "A", "C"], values
        finally:
            wb.close()
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("SMOKE_3WAY_ALIGNMENT_OK")


if __name__ == "__main__":
    main()
