import os
import time

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_xlsx(path: str, sheets: dict[str, list[list[object]]]):
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(title=name)
        for r_idx, row in enumerate(rows, start=1):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx).value = value
    wb.save(path)
    wb.close()


def _pump(root, seconds: float = 1.0):
    end = time.time() + seconds
    while time.time() < end:
        root.update_idletasks()
        root.update()
        time.sleep(0.02)


def main():
    td = make_temp_dir("sow_state_3way_")
    base = os.path.join(td, "base.xlsx")
    mine = os.path.join(td, "mine.xlsx")
    theirs = os.path.join(td, "theirs.xlsx")
    merged = os.path.join(td, "merged.xlsx")

    base_sheets = {
        "S_same": [["id", "val"], [1, "same"]],
        "S_base_diff": [["id", "val"], [1, "old"]],
        "S_ab_diff": [["id", "val"], [1, "base"]],
    }
    mine_sheets = {
        "S_same": [["id", "val"], [1, "same"]],
        "S_base_diff": [["id", "val"], [1, "new"]],
        "S_ab_diff": [["id", "val"], [1, "mine"]],
        "S_new_common": [["id", "val"], [1, "new-sheet"]],
    }
    theirs_sheets = {
        "S_same": [["id", "val"], [1, "same"]],
        "S_base_diff": [["id", "val"], [1, "new"]],
        "S_ab_diff": [["id", "val"], [1, "theirs"]],
        "S_new_common": [["id", "val"], [1, "new-sheet"]],
    }

    _make_xlsx(base, base_sheets)
    _make_xlsx(mine, mine_sheets)
    _make_xlsx(theirs, theirs_sheets)

    app = mod.SowMergeApp(mine, theirs, merge_mode=True, merged_path=merged, base_path=base)
    try:
        _pump(app.root, 4.0)
        st = app.sheet_diff_state
        assert st.get("S_same") == 0, f"Expected S_same=0, got {st.get('S_same')}"
        assert st.get("S_base_diff") == 2, f"Expected S_base_diff=2, got {st.get('S_base_diff')}"
        assert st.get("S_ab_diff") == 2, f"Expected S_ab_diff=2, got {st.get('S_ab_diff')}"
        assert st.get("S_new_common") == 2, f"Expected S_new_common=2, got {st.get('S_new_common')}"
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("GUI_SELF_TEST_SHEET_DIFF_STATE_3WAY_OK")


if __name__ == "__main__":
    main()
