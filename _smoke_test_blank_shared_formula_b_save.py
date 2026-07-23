"""Blank Base-side cells must safely clear shared formulas on B-side save."""

from __future__ import annotations

import os
import time
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import Workbook, load_workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _make_book(path: str, *, formulas: bool):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet.append(["ID", "名称", "输入", "公式"])
    worksheet.append(["id-1", "one", 10, "=C2" if formulas else None])
    worksheet.append(["id-2", "two", 20, "=C3" if formulas else None])
    workbook.save(path)
    workbook.close()


def _inject_shared_formula(path: str):
    part = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path, "r") as source:
        infos = source.infolist()
        payloads = {info.filename: source.read(info.filename) for info in infos}
    root = ET.fromstring(payloads[part])
    q = lambda name: f"{{{SHEET_NS}}}{name}"
    cells = {
        node.attrib.get("r"): node
        for node in root.iter(q("c"))
        if node.attrib.get("r")
    }
    master = cells["D2"].find(q("f"))
    follower = cells["D3"].find(q("f"))
    assert master is not None and follower is not None
    master.attrib.clear()
    master.attrib.update({"t": "shared", "ref": "D2:D3", "si": "0"})
    master.text = "C2"
    follower.attrib.clear()
    follower.attrib.update({"t": "shared", "si": "0"})
    follower.text = None
    payloads[part] = ET.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )
    rewritten = path + ".rewrite"
    with zipfile.ZipFile(rewritten, "w") as target:
        for info in infos:
            target.writestr(info, payloads[info.filename])
    os.replace(rewritten, path)


def _pump_until_ready(app, sheet: str):
    app.nb.select(app._sheet_containers[sheet])
    deadline = time.time() + 30.0
    view = None
    while time.time() < deadline:
        app.root.update_idletasks()
        app.root.update()
        view = app.sheet_views.get(sheet)
        if view is not None and view._derive_lifecycle_state() == "READY":
            return view
        time.sleep(0.01)
    raise AssertionError(
        f"view did not become READY: {getattr(view, '_lifecycle_state', None)}"
    )


def main():
    root = make_temp_dir("sow_blank_shared_formula_")
    base = os.path.join(root, "base.xlsx")
    mine = os.path.join(root, "mine.xlsx")
    _make_book(base, formulas=False)
    _make_book(mine, formulas=True)
    _inject_shared_formula(mine)

    app = mod.SowMergeApp(base, mine)
    output = None
    try:
        view = _pump_until_ready(app, "Data")
        pair_idx = view.row_a_to_pair_idx[2]
        assert view.pair_diff_cols[pair_idx] == {4}
        view._copy_single_cell_by_pair(pair_idx, "A2B", 4)
        assert app.ws_b_edit("Data").cell(2, 4).value is None
        assert ("Data", 2, 4) in app.manual_b_cell_ops

        native_save_enabled = mod._EXCEL_NATIVE_SAVE_ON_MERGE
        mod._EXCEL_NATIVE_SAVE_ON_MERGE = False
        try:
            output = app.build_manual_b_output_file()
        finally:
            mod._EXCEL_NATIVE_SAVE_ON_MERGE = native_save_enabled
        saved = load_workbook(output, read_only=True, data_only=False)
        try:
            assert saved["Data"]["D2"].value is None
            assert saved["Data"]["D3"].value == "=C3"
        finally:
            saved.close()
    finally:
        app._shutdown_root()
        if output and os.path.exists(output):
            os.remove(output)
    print("SMOKE_BLANK_SHARED_FORMULA_B_SAVE_OK")


if __name__ == "__main__":
    main()
