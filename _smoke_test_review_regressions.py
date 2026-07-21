import os
import threading
import time
import xml.etree.ElementTree as ET

from openpyxl import Workbook, load_workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _test_only_diff_region_boundaries():
    class DummyView:
        display_rows = [1, 4, 5, 9]
        _diff_blocks_cache = None

        @staticmethod
        def _pair_has_visual_diff(_pair_idx):
            return True

    view = DummyView()
    blocks = mod.SheetView._compute_diff_blocks(view)
    assert blocks == [(1, 1), (2, 3), (4, 4)], blocks


def _test_logical_region_extends_beyond_render_limit():
    class DummyView:
        # The UI rendered only a small window inside a much larger logical block.
        display_rows = [3, 4, 5]
        row_pairs = [(idx + 1, idx + 1) for idx in range(12)]

        @staticmethod
        def _pair_has_visual_diff(pair_idx):
            return 2 <= pair_idx <= 9

    view = DummyView()
    block = mod.SheetView._logical_diff_pair_block_for_line(view, 2)
    assert block == list(range(2, 10)), block


def _test_tail_identical_append_stays_paired():
    wb_mine = Workbook()
    wb_theirs = Workbook()
    try:
        ws_mine = wb_mine.active
        ws_theirs = wb_theirs.active
        ws_mine.append(["base"])
        ws_mine.append(["same append"])
        ws_theirs.append(["base"])
        ws_theirs.append(["same append"])
        pairs = mod._split_tail_independent_append_pairs(
            [(1, 1), (2, 2)],
            {1: 1},
            {1: 1},
            ws_mine,
            ws_theirs,
            1,
        )
        assert pairs == [(1, 1), (2, 2)], pairs

        ws_theirs["A2"] = "different append"
        ws_mine["A3"] = "mine append 2"
        ws_theirs["A3"] = "theirs append 2"
        pairs = mod._split_tail_independent_append_pairs(
            [(1, 1), (2, 2), (3, 3)],
            {1: 1},
            {1: 1},
            ws_mine,
            ws_theirs,
            1,
        )
        assert pairs == [(1, 1), (None, 2), (None, 3), (2, None), (3, None)], pairs
    finally:
        wb_mine.close()
        wb_theirs.close()


def _test_shared_formula_is_not_destroyed():
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xml = f"""
    <worksheet xmlns="{ns}">
      <sheetData>
        <row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="7">SUM(B1:C1)</f><v>3</v></c></row>
        <row r="2"><c r="A2"><f t="shared" si="7"/><v>7</v></c></row>
      </sheetData>
    </worksheet>
    """
    root = ET.fromstring(xml)
    mod._sheet_xml_set_cell(root, 1, 1, "=SUM(B1:C1)", 9)
    f = root.find(f".//{{{ns}}}c[@r='A1']/{{{ns}}}f")
    v = root.find(f".//{{{ns}}}c[@r='A1']/{{{ns}}}v")
    assert f is not None and f.attrib == {"t": "shared", "ref": "A1:A2", "si": "7"}, f.attrib
    assert v is not None and v.text == "9"

    # A cache-only adoption must preserve even a shared member formula whose
    # formula text is stored only on the group master.
    mod._sheet_xml_set_cell(
        root,
        2,
        1,
        "=IGNORED_FOR_CACHE_ONLY()",
        11,
        preserve_existing_formula=True,
    )
    member_f = root.find(f".//{{{ns}}}c[@r='A2']/{{{ns}}}f")
    member_v = root.find(f".//{{{ns}}}c[@r='A2']/{{{ns}}}v")
    assert member_f is not None and member_f.attrib == {"t": "shared", "si": "7"}
    assert member_f.text in (None, "")
    assert member_v is not None and member_v.text == "11"

    try:
        mod._sheet_xml_set_cell(root, 1, 1, "=SUM(B1:D1)", 4)
    except RuntimeError as exc:
        assert "shared formula" in str(exc)
    else:
        raise AssertionError("unsafe shared-formula replacement was not rejected")


def _test_formula_noop_filter():
    root = make_temp_dir("sow_formula_noop_")
    path = os.path.join(root, "source.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=SUM(B1:C1)"
    ws["B1"] = 1
    ws["C1"] = 2
    wb.save(path)
    wb.close()

    ops = {
        ("S1", 1, 1): "=SUM(B1:C1)",
        ("S1", 1, 2): 10,
    }
    filtered = mod._filter_noop_manual_ops(path, ops)
    assert ("S1", 1, 1) not in filtered, filtered
    assert filtered[("S1", 1, 2)] == 10

    text_out = os.path.join(root, "text-output.xlsx")
    mod._build_manual_merge_xlsx_via_zip(
        path,
        text_out,
        {("S1", 2, 1): "1.0"},
    )
    wb = load_workbook(text_out, data_only=False)
    try:
        assert wb["S1"]["A2"].value == "1.0"
        assert isinstance(wb["S1"]["A2"].value, str)
    finally:
        wb.close()


def _test_large_alignment_fast_and_exact():
    sig_base = [f"id-{idx}" for idx in range(20000)]
    sig_theirs = list(sig_base)
    sig_theirs[16153:16158] = [f"changed-{idx}" for idx in range(5)]
    started = time.perf_counter()
    pairs = mod._compute_row_pairs_from_signatures(sig_base, sig_theirs)
    elapsed = time.perf_counter() - started
    assert len(pairs) == 20000, len(pairs)
    assert pairs[0] == (1, 1) and pairs[-1] == (20000, 20000)
    assert elapsed < 2.0, elapsed


def _test_stable_copy_waits_for_complete_zip():
    root = make_temp_dir("sow_partial_svn_export_")
    complete = os.path.join(root, "complete.xlsx")
    partial = os.path.join(root, "WorldMonster.xlsx-revBASE.tmp.xlsx")
    wb = Workbook()
    wb.active["A1"] = "complete"
    wb.save(complete)
    wb.close()

    with open(complete, "rb") as src:
        payload = src.read()
    split_at = len(payload) // 2
    with open(partial, "wb") as dst:
        dst.write(payload[:split_at])
        dst.flush()

    def _finish_export():
        time.sleep(0.35)
        with open(partial, "ab") as dst:
            dst.write(payload[split_at:])
            dst.flush()

    writer = threading.Thread(target=_finish_export, daemon=True)
    writer.start()
    stable = mod._ensure_stable_copy(partial)
    writer.join(timeout=2.0)
    assert stable != partial
    assert os.path.getsize(stable) == len(payload)
    assert mod._workbook_package_ready(stable)


def _test_background_recalc_policy_is_explicit():
    old_always = mod._AUTO_RECALC_FORMULAS_ALWAYS
    old_missing = mod._AUTO_RECALC_MISSING_CACHE
    old_recalc = mod._recalc_and_prepare_val_path
    calls = []
    try:
        mod._AUTO_RECALC_FORMULAS_ALWAYS = False
        mod._AUTO_RECALC_MISSING_CACHE = False
        mod._recalc_and_prepare_val_path = lambda path: calls.append(path) or "recalc.xlsx"
        assert mod._maybe_recalc_and_prepare_val_path("formula.xlsx", force=False) is None
        assert not calls
        assert mod._maybe_recalc_and_prepare_val_path("formula.xlsx", force=True) == "recalc.xlsx"
        assert calls == ["formula.xlsx"]
    finally:
        mod._AUTO_RECALC_FORMULAS_ALWAYS = old_always
        mod._AUTO_RECALC_MISSING_CACHE = old_missing
        mod._recalc_and_prepare_val_path = old_recalc


def main():
    _test_only_diff_region_boundaries()
    _test_logical_region_extends_beyond_render_limit()
    _test_tail_identical_append_stays_paired()
    _test_shared_formula_is_not_destroyed()
    _test_formula_noop_filter()
    _test_large_alignment_fast_and_exact()
    _test_stable_copy_waits_for_complete_zip()
    _test_background_recalc_policy_is_explicit()
    print("SMOKE_REVIEW_REGRESSIONS_OK")


if __name__ == "__main__":
    main()
