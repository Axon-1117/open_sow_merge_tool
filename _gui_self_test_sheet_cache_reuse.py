import os
import sys
import tempfile
import time

from openpyxl import Workbook

import sow_merge_tool as sm


def _make_book(path: str, suffix: str):
    wb = Workbook()
    first = wb.active
    first.title = "S1"
    second = wb.create_sheet("S2")
    for ws in (first, second):
        ws.append(["id", "value"])
        for row in range(1, 401):
            ws.append([row, f"{ws.title}-{row}-{suffix}"])
    wb.save(path)
    wb.close()


def main():
    with tempfile.TemporaryDirectory(prefix="sow_sheet_cache_") as tmp:
        left = os.path.join(tmp, "Book.xlsx")
        right = os.path.join(tmp, "Book-right.xlsx")
        _make_book(left, "A")
        _make_book(right, "B")

        app = sm.SowMergeApp(left, right)
        deadline = time.monotonic() + 20
        while time.monotonic() < deadline:
            app.root.update()
            if "S2" in app._sheet_cache_store:
                break
            time.sleep(0.02)
        assert "S2" in app._sheet_cache_store, "background result for unopened Sheet was discarded"
        assert app.sheet_views.get("S2") is None

        started = time.monotonic()
        app._select_tab("S2")
        app.root.update()
        elapsed = time.monotonic() - started
        view = app.sheet_views.get("S2")
        assert view is not None and view._data_ready
        assert "S2" not in app._sheet_cache_store
        assert elapsed < 0.5, f"retained cache was not applied immediately: {elapsed:.3f}s"
        with app._compute_lock:
            assert "S2" not in app._compute_queue
            assert "S2" not in app._compute_inflight

        app._shutdown_root()
    print("GUI_SELF_TEST_SHEET_CACHE_REUSE_OK")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"GUI_SELF_TEST_SHEET_CACHE_REUSE_FAIL: {exc}", file=sys.stderr)
        raise
