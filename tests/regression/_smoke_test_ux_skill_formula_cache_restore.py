from __future__ import annotations

import os
import tempfile

import sow_merge_tool as smt
from _ux_5_3_final_acceptance import (
    _restore_skill_formula_caches,
    _same_formula_cache_plan,
)


SHEET = "SkillTimeline@design"


def _seed_formula_caches(path, formulas, caches):
    raw_path = path + ".raw.xlsx"
    wb = smt.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws["A1"] = 1
    formula_ops = {}
    cached_values = {}
    for coordinate, formula in formulas.items():
        ws[coordinate] = formula
        cell = ws[coordinate]
        key = (SHEET, int(cell.row), int(cell.column))
        formula_ops[key] = formula
        cached_values[key] = caches[coordinate]
    wb.save(raw_path)
    wb.close()

    smt._build_manual_merge_xlsx_via_zip(
        raw_path,
        path,
        formula_ops,
        cached_values=cached_values,
        cache_only_keys=set(formula_ops),
    )
    os.remove(raw_path)


def main():
    with tempfile.TemporaryDirectory(prefix="sow-skill-cache-restore-") as temp_dir:
        mine = os.path.join(temp_dir, "mine.xlsx")
        theirs = os.path.join(temp_dir, "theirs.xlsx")
        _seed_formula_caches(
            mine,
            {"B1": "=A1+1", "C1": "=A1+2"},
            {"B1": 2, "C1": 3},
        )
        _seed_formula_caches(
            theirs,
            {"B1": "=A1+1", "C1": "=A1+3"},
            {"B1": 999, "C1": 777},
        )

        formula_ops, _cached_values, before = _same_formula_cache_plan(
            mine, theirs, SHEET
        )
        assert set(formula_ops) == {(SHEET, 1, 2)}, formula_ops
        assert before["cache_diff_count"] == 1, before
        assert before["cache_diff_columns"] == {"B": 1}, before

        restored = _restore_skill_formula_caches(
            mine,
            theirs,
            SHEET,
            validate_excel_reopen=False,
        )
        assert restored["applied"] is True, restored
        assert restored["before"]["cache_diff_count"] == 1, restored
        assert restored["after"]["cache_diff_count"] == 0, restored
        assert restored["package_validation"]["valid"] is True, restored

        wb_formula = smt.load_workbook(theirs, data_only=False, read_only=False)
        wb_cached = smt.load_workbook(theirs, data_only=True, read_only=False)
        try:
            assert wb_formula[SHEET]["B1"].value == "=A1+1"
            assert wb_formula[SHEET]["C1"].value == "=A1+3"
            assert wb_cached[SHEET]["B1"].value == 2
            assert wb_cached[SHEET]["C1"].value == 777
        finally:
            wb_formula.close()
            wb_cached.close()

        second = _restore_skill_formula_caches(
            mine,
            theirs,
            SHEET,
            validate_excel_reopen=False,
        )
        assert second["applied"] is False, second
        assert second["before"]["cache_diff_count"] == 0, second
        assert second["after"]["cache_diff_count"] == 0, second
        print("PASS: Skill fixture cache restore is exact-formula-only, package-safe, and idempotent")


if __name__ == "__main__":
    main()
