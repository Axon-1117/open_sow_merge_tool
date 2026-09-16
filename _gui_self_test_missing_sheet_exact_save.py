"""Missing-Sheet background comparison must publish a usable exact surface."""
import os
import tempfile
import time
from openpyxl import Workbook
import sow_merge_tool as sm
from _gui_self_test_exact_sheet_readiness import _shutdown_app, _wait_selected_full_detail


def make_book(path, include):
    wb = Workbook()
    wb.active.title = "Common"
    wb.active.append(["id", "value"])
    wb.active.append([1, "same"])
    if include:
        ws = wb.create_sheet("ResidentLink@design")
        ws.append(["id", "value"])
        ws.append([7, "present-side"])
    wb.save(path)
    wb.close()


def main():
    with tempfile.TemporaryDirectory(prefix="sow-missing-exact-") as root:
        for case, sides in (("mine-only", (True, False, False)),
                            ("theirs-only", (False, True, False)),
                            ("base-only", (False, False, True))):
            paths = [os.path.join(root, case + side + ".xlsx") for side in ("A", "B", "BASE")]
            for path, include in zip(paths, sides):
                make_book(path, include)
            app = None
            try:
                app = sm.SowMergeApp(paths[0], paths[1], base_path=paths[2], merge_mode=True,
                                    initial_sheet="ResidentLink@design")
                view = _wait_selected_full_detail(app, "ResidentLink@design", deadline=time.monotonic() + 45)
                assert app._sheet_exact_entry(view.sheet)["state"] == sm._SHEET_EXACT_CHANGED
                assert any(view.pair_diff_cols.values())
                if sides[0]:
                    assert "present-side" in view.left.get("1.0", "end")
                if sides[1]:
                    assert "present-side" in view.right.get("1.0", "end")
                if sides[2]:
                    assert view._base_row_for_pair(1) == 2
                print("PASS", case, flush=True)
            finally:
                _shutdown_app(app)


if __name__ == "__main__":
    main()
