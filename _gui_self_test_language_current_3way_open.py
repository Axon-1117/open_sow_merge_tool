"""GUI acceptance for the current Language 3-way conflict inputs."""

from __future__ import annotations

import os
import shutil
import time

import sow_merge_tool as sm
from _test_temp_utils import make_temp_dir


SHEET = "default@design@na_TLanguageCn"
BASE = (
    "D:/Tools/sow_merge_tool_proj/tmp/"
    "language_3way_unresolved_repro_20260828_183244/base.xlsx"
)
MINE = "C:/GM15/design/sheets/common/Language.xlsx"
THEIRS = (
    "D:/Tools/sow_merge_tool_proj/tmp/"
    "language_3way_unresolved_repro_20260828_183244/theirs.xlsx"
)


def _wait(app, predicate, label, timeout=90.0):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        app.root.update_idletasks()
        app.root.update()
        if predicate():
            return
        time.sleep(0.02)
    raise AssertionError(label)


def main():
    temp_dir = make_temp_dir("sow_language_current_3way_open_")
    merged = os.path.join(temp_dir, "Language-merged.xlsx")
    shutil.copy2(MINE, merged)
    app = sm.SowMergeApp(
        MINE,
        THEIRS,
        merge_mode=True,
        merged_path=merged,
        base_path=BASE,
        initial_sheet=SHEET,
    )
    try:
        _wait(
            app,
            lambda: (
                SHEET in app.sheet_views
                and app.selected_sheet == SHEET
                and app._is_sheet_exact_current(SHEET)
                and bool(app.sheet_views[SHEET]._prepared_complete)
                and bool(app.sheet_views[SHEET]._data_ready)
                and not bool(app.sheet_views[SHEET]._pending_exact_render)
            ),
            "Language exact GUI surface did not become ready",
        )
        view = app.sheet_views[SHEET]
        cache = view._active_column_comparison_cache()
        assert not cache.unresolved_cols, cache.unresolved_cols
        assert cache.structural_diff_cols == frozenset(range(7, 27))
        assert len(view.row_pairs) == 20897, len(view.row_pairs)
        assert view._derive_lifecycle_state() != "UNRESOLVED"
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass
    print("GUI_LANGUAGE_CURRENT_3WAY_OPEN_OK")


if __name__ == "__main__":
    main()
