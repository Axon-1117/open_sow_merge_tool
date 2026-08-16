"""Headed smoke test for visible workbook paths and aligned compare cards."""

from __future__ import annotations

import os
import tempfile
import time

from openpyxl import Workbook

import sow_merge_tool as smt


def _write_book(path: str, value: str) -> None:
    workbook = Workbook()
    workbook.active["A1"] = value
    workbook.save(path)
    workbook.close()


def _wait_for_view(app, timeout: float = 12.0):
    app.nb.select(app._sheet_containers["Sheet"])
    deadline = time.time() + timeout
    while time.time() < deadline:
        app.root.update()
        view = app.sheet_views.get("Sheet")
        if view is not None:
            return view
        time.sleep(0.01)
    raise AssertionError("merge path GUI view was not created")


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="sow-visible-paths-") as temp_root:
        base_dir = os.path.join(temp_root, "源分支 修改前")
        mine_dir = os.path.join(temp_root, "目标分支 修改后预览")
        os.makedirs(base_dir)
        os.makedirs(mine_dir)
        base_path = os.path.join(base_dir, "Language.xlsx")
        mine_path = os.path.join(mine_dir, "Language.xlsx")
        _write_book(base_path, "修改前")
        _write_book(mine_path, "修改后")

        app = smt.SowMergeApp(
            base_path,
            mine_path,
            raw_base=base_path,
            raw_mine=mine_path,
        )
        try:
            view = _wait_for_view(app)
            app.root.update_idletasks()
            assert os.path.normpath(base_path) in view.path_file_label_a.cget("text")
            assert os.path.normpath(mine_path) in view.path_file_label_b.cget("text")
            assert not view.path_card_base.winfo_manager(), "two-way mode must not reserve a blank Base card"
            assert int(view.path_card_a.grid_info()["column"]) == 0
            assert int(view.path_card_b.grid_info()["column"]) == 1
            assert abs(view.path_card_a.winfo_width() - view.path_card_b.winfo_width()) <= 4
            view._copy_visible_pane_path("b")
            assert app.root.clipboard_get() == mine_path
            assert view.save_b_btn.cget("text") == "保存修改后文件"
        finally:
            app._shutdown_root()
    print("PASS: visible workbook paths and aligned compare cards")


if __name__ == "__main__":
    main()
