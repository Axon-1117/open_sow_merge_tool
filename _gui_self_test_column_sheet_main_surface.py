"""GUI regression: retained @column cache must remain visible after tab creation."""

import os
import tempfile
import time

from openpyxl import Workbook

import sow_merge_tool as sm


def _write(path, side):
    wb = Workbook()
    first = wb.active
    first.title = "Data@design"
    first.append(["id@id", "value"])
    first.append([1, "same"])
    config = wb.create_sheet("GeomancyConfig@column")
    config.append(["field@pm", "type", "comment", "default", "value"])
    config.append(["daily_max", "int32", "same", 5, 5])
    config.append([
        "radar_rate", "int32", f"{side} comment",
        2000 if side == "base" else None,
        2000 if side == "base" else 0,
    ])
    config.append([
        "special_radar_num", "repeated int32", f"{side} special",
        "5|10" if side == "base" else None,
        "5|10" if side == "base" else None,
    ])
    wb.save(path)
    wb.close()


def _pump(app, seconds=0.04):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        app.root.update_idletasks()
        app.root.update()
        time.sleep(0.005)


def _wait(app, predicate, timeout, label):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        _pump(app)
        if predicate():
            return
    raise AssertionError(f"timeout: {label}")


def main():
    sheet = "GeomancyConfig@column"
    with tempfile.TemporaryDirectory(prefix="sow_column_surface_") as tmp:
        base = os.path.join(tmp, "base.xlsx")
        mine = os.path.join(tmp, "mine.xlsx")
        _write(base, "base")
        _write(mine, "mine")
        app = sm.SowMergeApp(base, mine, initial_sheet="Data@design")
        try:
            _wait(
                app,
                lambda: (
                    sheet in app._sheet_cache_store
                    and app.sheet_views.get(sheet) is None
                ),
                20.0,
                "hidden @column exact cache",
            )
            app.nb.select(app._sheet_containers[sheet])
            _wait(
                app,
                lambda: (
                    app.selected_sheet == sheet
                    and sheet in app.sheet_views
                    and app._is_sheet_exact_current(sheet)
                    and not app.sheet_views[sheet]._pending_exact_render
                ),
                15.0,
                "selected @column exact surface",
            )
            # Run idle callbacks after the exact surface was installed. The
            # historical warmup callback cleared the six main Text widgets at
            # precisely this point while leaving C-area diff state intact.
            _pump(app, 0.25)
            view = app.sheet_views[sheet]
            assert view.display_rows == [2, 3], view.display_rows
            assert "radar_rate" in view.left.get("1.0", "end")
            assert "radar_rate" in view.right.get("1.0", "end")
            assert "special_radar_num" in view.left.get("1.0", "end")
            assert "special_radar_num" in view.right.get("1.0", "end")
        finally:
            app._shutdown_root()
    print("GUI_COLUMN_SHEET_MAIN_SURFACE_OK")


if __name__ == "__main__":
    main()
