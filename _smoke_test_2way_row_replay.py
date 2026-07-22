import os
import time

from openpyxl import Workbook, load_workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_book(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def main():
    root = make_temp_dir("sow_2way_row_replay_")
    mine = os.path.join(root, "mine.xlsx")
    theirs = os.path.join(root, "theirs.xlsx")
    _make_book(mine, [["id", "formula"], ["A", None], ["NEW", "=A3"], ["C", None]])
    _make_book(theirs, [["id", "formula"], ["A", None], ["C", None]])

    app = mod.SowMergeApp(mine, theirs)
    original_excel_builder = mod._build_manual_merge_output_with_excel
    captured = {}
    try:
        view = None
        for _ in range(120):
            app.root.update_idletasks()
            app.root.update()
            view = app.sheet_views.get("S1")
            if view is not None and view._data_ready:
                break
            time.sleep(0.02)
        assert view is not None and view._data_ready
        view.force_align_var.set(1)
        view._toggle_force_align()
        insert_pair = next(
            pair_idx for pair_idx, (row_a, row_b) in enumerate(view.row_pairs)
            if row_a == 3 and row_b is None
        )
        assert view._copy_selected_row("A2B", override_pair_idx=insert_pair)
        assert app.manual_b_row_ops[0]["source_side"] == "A"
        assert app.manual_b_cell_ops[("S1", 3, 2)] == "=A3"

        def _fake_excel(src, out, manual_ops, row_ops, sheet_ops=None, source_paths=None):
            captured["src"] = src
            captured["row_ops"] = list(row_ops or [])
            captured["source_paths"] = dict(source_paths or {})
            return mod._build_manual_merge_output_with_openpyxl(
                src,
                out,
                manual_ops,
                row_ops,
                sheet_ops=sheet_ops,
                source_paths=source_paths,
            )

        mod._build_manual_merge_output_with_excel = _fake_excel
        out = app.build_manual_b_output_file()
        assert captured["src"] == theirs
        assert captured["source_paths"]["A"] == mine
        assert captured["row_ops"][0]["source_side"] == "A"

        wb = load_workbook(out, data_only=False)
        try:
            ws = wb["S1"]
            assert [ws.cell(row=row, column=1).value for row in range(1, 5)] == ["id", "A", "NEW", "C"]
            assert ws["B3"].value == "=A3"
        finally:
            wb.close()
    finally:
        mod._build_manual_merge_output_with_excel = original_excel_builder
        app._shutdown_root()

    print("SMOKE_2WAY_ROW_REPLAY_OK")


if __name__ == "__main__":
    main()
