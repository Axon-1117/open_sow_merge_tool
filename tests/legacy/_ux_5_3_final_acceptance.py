from __future__ import annotations

import argparse
import gc
import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
import traceback

import sow_merge_tool as smt
from _gui_self_test_logical_column_actions import (
    _force_full_view,
    _model_snapshot,
    _pump,
    _selection_snapshot,
    _wait_for_stable_projection,
    _wait_for_view,
)


ROOT = r"C:\Users\dd\AppData\Local\Temp\sow_ux_5_3_20260723_001"
SOURCE_ROOT = r"C:\GM15\design\sheets\develop"
REPORT = os.path.join(ROOT, "acceptance_report.json")

CASES = {
    "Guide": {"sheet": "TGuideStep@design"},
    "Skill": {"sheet": "SkillTimeline@design"},
    "Dungeon": {"sheet": "Dungeon@design"},
    "WorldMonster": {"sheet": "WorldMonster@design"},
}


def now_ms():
    return time.perf_counter() * 1000.0


def sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def close_wb(wb):
    try:
        wb.close()
    except Exception:
        pass


def close_app(app):
    if app is None:
        return
    try:
        app._is_closing = True
    except Exception:
        pass
    for name in ("_interactive_action_event", "_priority_diff_event", "_edit_loaded_event", "_initial_sheet_ready_event"):
        try:
            getattr(app, name).set()
        except Exception:
            pass
    try:
        smt._wbs_close(
            getattr(app, "_wb_a_val", None), getattr(app, "_wb_b_val", None), getattr(app, "_wb_base_val", None),
            getattr(app, "_wb_a_edit", None), getattr(app, "_wb_b_edit", None), getattr(app, "_wb_base_edit", None),
        )
    except Exception:
        pass
    try:
        app.root.destroy()
    except Exception:
        pass
    gc.collect()


def copy_same_name(src, side_dir, basename):
    os.makedirs(side_dir, exist_ok=True)
    dst = os.path.join(side_dir, basename)
    shutil.copy2(src, dst)
    return dst


def mutate_copy(src, dst, mutator):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    wb = smt.load_workbook(dst, data_only=False)
    try:
        details = mutator(wb)
        smt._atomic_save_wb(wb, dst)
    finally:
        close_wb(wb)
    return details


def _exact_formula_identity(value):
    """Return an exact formula identity suitable for fixture cache restoration."""
    special = smt._special_formula_signature(value)
    if special is not None:
        return ("SPECIAL", type(value).__name__, repr(special), smt._formula_text(value))
    if isinstance(value, str) and value.startswith("="):
        return ("FORMULA", value)
    return None


def _same_formula_cache_plan(mine_path, theirs_path, sheet, sample_limit=12):
    """Collect same-coordinate/exact-formula cache values and their differences."""
    workbooks = []
    try:
        mine_edit = smt.load_workbook(mine_path, data_only=False, read_only=False)
        workbooks.append(mine_edit)
        theirs_edit = smt.load_workbook(theirs_path, data_only=False, read_only=False)
        workbooks.append(theirs_edit)
        mine_values = smt.load_workbook(mine_path, data_only=True, read_only=False)
        workbooks.append(mine_values)
        theirs_values = smt.load_workbook(theirs_path, data_only=True, read_only=False)
        workbooks.append(theirs_values)

        mine_formula_values = smt._formula_edit_value_map(mine_edit[sheet])
        theirs_formula_values = smt._formula_edit_value_map(theirs_edit[sheet])
        mine_value_ws = mine_values[sheet]
        theirs_value_ws = theirs_values[sheet]
        formula_ops = {}
        cached_values = {}
        differences = []
        differing_columns = {}
        for row_idx, col_idx in sorted(set(mine_formula_values) & set(theirs_formula_values)):
            mine_formula = mine_formula_values[(row_idx, col_idx)]
            theirs_formula = theirs_formula_values[(row_idx, col_idx)]
            mine_identity = _exact_formula_identity(mine_formula)
            if mine_identity is None or mine_identity != _exact_formula_identity(theirs_formula):
                continue
            key = (sheet, int(row_idx), int(col_idx))
            mine_cached = mine_value_ws.cell(row=row_idx, column=col_idx).value
            theirs_cached = theirs_value_ws.cell(row=row_idx, column=col_idx).value
            formula_ops[key] = theirs_formula
            cached_values[key] = mine_cached
            mine_key = smt._merge_cmp_value(mine_cached)
            theirs_key = smt._merge_cmp_value(theirs_cached)
            if mine_key == theirs_key:
                continue
            column_letter = smt.get_column_letter(col_idx)
            differing_columns[column_letter] = differing_columns.get(column_letter, 0) + 1
            if len(differences) < sample_limit:
                differences.append({
                    "coordinate": f"{column_letter}{row_idx}",
                    "mine": mine_key,
                    "theirs": theirs_key,
                })

        diff_count = sum(differing_columns.values())
        summary = {
            "same_coordinate_exact_formula_count": len(formula_ops),
            "cache_diff_count": diff_count,
            "cache_diff_columns": dict(sorted(differing_columns.items())),
            "samples": differences,
        }
        return formula_ops, cached_values, summary
    finally:
        for wb in workbooks:
            close_wb(wb)


def _restore_skill_formula_caches(
    mine_path,
    theirs_path,
    sheet,
    *,
    validate_excel_reopen=True,
):
    """Undo Excel recalculation noise without changing formula or structure XML."""
    formula_ops, cached_values, before = _same_formula_cache_plan(
        mine_path, theirs_path, sheet
    )
    result = {
        "strategy": "same-coordinate exact-formula ZIP cache-only restore",
        "applied": bool(before["cache_diff_count"]),
        "before": before,
        "cache_only_key_count": len(formula_ops),
    }
    if not before["cache_diff_count"]:
        result["after"] = before
        result["package_validation"] = {"valid": True, "reason": "unchanged"}
        if validate_excel_reopen:
            reopen_started = now_ms()
            reopened = smt._excel_reopen_validate(theirs_path)
            result["excel_read_only_reopen"] = {
                "valid": bool(reopened),
                "elapsed_ms": round(now_ms() - reopen_started, 2),
            }
            if not reopened:
                raise RuntimeError("Skill unchanged fixture failed Excel read-only reopen")
        return result

    fd, patched_path = tempfile.mkstemp(
        prefix="skill-cache-restore-",
        suffix=".xlsx",
        dir=os.path.dirname(theirs_path),
    )
    os.close(fd)
    try:
        smt._build_manual_merge_xlsx_via_zip(
            theirs_path,
            patched_path,
            formula_ops,
            cached_values=cached_values,
            cache_only_keys=set(formula_ops),
        )
        valid, reason = smt._validate_xlsx_package(patched_path)
        result["package_validation"] = {"valid": bool(valid), "reason": reason}
        if not valid:
            raise RuntimeError(f"Skill cache-only fixture package rejected: {reason}")
        os.replace(patched_path, theirs_path)
    finally:
        try:
            if os.path.exists(patched_path):
                os.remove(patched_path)
        except OSError:
            pass

    after_workbook = smt.load_workbook(
        theirs_path, data_only=True, read_only=False
    )
    try:
        after_ws = after_workbook[sheet]
        after_columns = {}
        after_samples = []
        for (_sheet, row_idx, col_idx), mine_cached in cached_values.items():
            theirs_cached = after_ws.cell(row=row_idx, column=col_idx).value
            mine_key = smt._merge_cmp_value(mine_cached)
            theirs_key = smt._merge_cmp_value(theirs_cached)
            if mine_key == theirs_key:
                continue
            column_letter = smt.get_column_letter(col_idx)
            after_columns[column_letter] = after_columns.get(column_letter, 0) + 1
            if len(after_samples) < 12:
                after_samples.append({
                    "coordinate": f"{column_letter}{row_idx}",
                    "mine": mine_key,
                    "theirs": theirs_key,
                })
        after = {
            "same_coordinate_exact_formula_count": len(formula_ops),
            "cache_diff_count": sum(after_columns.values()),
            "cache_diff_columns": dict(sorted(after_columns.items())),
            "samples": after_samples,
        }
    finally:
        close_wb(after_workbook)
    result["after"] = after
    if after["cache_diff_count"]:
        raise RuntimeError(
            "Skill same-formula cached-value restoration did not converge: "
            f"{after['cache_diff_count']} differences remain"
        )
    if validate_excel_reopen:
        reopen_started = now_ms()
        reopened = smt._excel_reopen_validate(theirs_path)
        result["excel_read_only_reopen"] = {
            "valid": bool(reopened),
            "elapsed_ms": round(now_ms() - reopen_started, 2),
        }
        if not reopened:
            raise RuntimeError("Skill restored fixture failed Excel read-only reopen")
    return result


def excel_column_fixture(src, dst, sheet, *, restore_skill_formula_caches=False):
    """Create the structural branch in real Excel so formula references follow Excel rules."""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    wb = smt.load_workbook(src, data_only=False, read_only=True)
    try:
        ws = wb[sheet]
        before_shape = (ws.max_row, ws.max_column)
        deleted_headers = [ws.cell(1, col).value for col in (22, 23)]
    finally:
        close_wb(wb)
    p = str(dst).replace("'", "''")
    s = str(sheet).replace("'", "''")
    ps = (
        "$ErrorActionPreference='Stop';"
        f"$p='{p}';$sheetName='{s}';"
        "$xl=$null;$wb=$null;$ws=$null;"
        "try{"
        "$xl=New-Object -ComObject Excel.Application;"
        "$xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.AskToUpdateLinks=$false;$xl.EnableEvents=$false;"
        "try{$xl.Calculation=-4135}catch{};try{$xl.CalculateBeforeSave=$false}catch{};"
        "$wb=$xl.Workbooks.Open($p,0,$false);"
        "try{$xl.Calculation=-4135}catch{};try{$xl.CalculateBeforeSave=$false}catch{};"
        "$ws=$wb.Worksheets.Item($sheetName);"
        "$ws.Columns.Item(12).Insert();$ws.Columns.Item(12).Insert();"
        "$rows=[int]$ws.UsedRange.Rows.Count;"
        "$data=[System.Array]::CreateInstance([object],@($rows,2));"
        "for($r=0;$r -lt $rows;$r++){"
        "$data.SetValue(('__UX_INS1_R'+($r+1)),$r,0);"
        "$data.SetValue(('__UX_INS2_R'+($r+1)),$r,1);"
        "};"
        "$range=$ws.Range($ws.Cells.Item(1,12),$ws.Cells.Item($rows,13));"
        "$range.Value2=$data;"
        "$ws.Columns.Item(24).Delete();$ws.Columns.Item(24).Delete();"
        "$wb.Save();"
        "}finally{"
        "if($range-ne $null){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)}catch{}};"
        "if($ws-ne $null){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws)}catch{}};"
        "if($wb-ne $null){try{$wb.Close($false)}catch{};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}catch{}};"
        "if($xl-ne $null){try{$xl.Quit()}catch{};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}catch{}};"
        "[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers();"
        "};"
    )
    started = now_ms()
    completed = smt._run_excel_powershell_with_transient_retry(ps, timeout=300)
    if completed.returncode != 0:
        raise RuntimeError(f"Excel COM fixture failed: {completed.stderr.strip()}")
    cache_restore = None
    if restore_skill_formula_caches:
        cache_restore = _restore_skill_formula_caches(src, dst, sheet)
    probe = workbook_probe(dst, sheet, markers=("__UX_INS1_R1", "__UX_INS2_R1"))
    details = {
        "engine": "Excel COM",
        "before": before_shape,
        "after": (probe["rows"], probe["cols"]),
        "insert_cols": 2,
        "delete_cols": 2,
        "deleted_original_headers": deleted_headers,
        "marker_headers": ["__UX_INS1_R1", "__UX_INS2_R1"],
        "formula_count_after": probe["formula_count"],
        "excel_fixture_ms": round(now_ms() - started, 2),
    }
    if cache_restore is not None:
        details["formula_cache_restore"] = cache_restore
    return details


def formula_semantic_summary(mine_path, theirs_path, sheet, sample_limit=12):
    def inventory(path):
        wb = smt.load_workbook(path, data_only=False, read_only=True)
        try:
            ws = wb[sheet]
            formulas = {}
            digest = hashlib.sha256()
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    special = smt._special_formula_signature(value)
                    if special is not None:
                        normalized = "SPECIAL:" + repr(special)
                    elif isinstance(value, str) and value.startswith("="):
                        normalized = value
                    else:
                        continue
                    formulas[cell.coordinate] = normalized
                    digest.update(cell.coordinate.encode("utf-8"))
                    digest.update(b"\0")
                    digest.update(normalized.encode("utf-8", errors="replace"))
                    digest.update(b"\n")
            return formulas, digest.hexdigest().upper()
        finally:
            close_wb(wb)

    mine, mine_digest = inventory(mine_path)
    theirs, theirs_digest = inventory(theirs_path)
    changed = []
    for coordinate in sorted(set(mine) | set(theirs)):
        if mine.get(coordinate) == theirs.get(coordinate):
            continue
        if len(changed) < sample_limit:
            changed.append({"coordinate": coordinate, "mine": mine.get(coordinate), "theirs": theirs.get(coordinate)})
    return {
        "mine_formula_count": len(mine), "theirs_formula_count": len(theirs),
        "mine_formula_digest": mine_digest, "theirs_formula_digest": theirs_digest,
        "coordinate_formula_diff_count": sum(1 for coordinate in set(mine) | set(theirs) if mine.get(coordinate) != theirs.get(coordinate)),
        "samples": changed,
    }


def prepare():
    os.makedirs(ROOT, exist_ok=True)
    result = {"root": ROOT, "sources": {}, "cases": {}}
    for stem in ("Guide", "Skill", "Dungeon", "WorldMonster"):
        src = os.path.join(SOURCE_ROOT, stem + ".xlsx")
        result["sources"][stem] = {"path": src, "sha256": sha256(src), "size": os.path.getsize(src)}

    # Guide: base, locally modified/uncommitted mine, remotely modified/committed theirs.
    stem = "Guide"
    basename = stem + ".xlsx"
    src = os.path.join(SOURCE_ROOT, basename)
    case_root = os.path.join(ROOT, "cases", stem)
    base = copy_same_name(src, os.path.join(case_root, "base"), basename)

    def mutate_mine(wb):
        ws = wb[CASES[stem]["sheet"]]
        before = (ws.max_row, ws.max_column)
        edits = {
            (10, 1): "UX_MINE_CONFLICT",
            (11, 2): "UX_MINE_EDIT_R11",
            (12, 3): "UX_MINE_EDIT_R12",
            (13, 4): "UX_MINE_EDIT_R13",
        }
        for (row, col), value in edits.items():
            ws.cell(row=row, column=col).value = value
        ws.insert_rows(20, 3)
        for offset in range(3):
            for col in range(1, 5):
                ws.cell(row=20 + offset, column=col).value = f"UX_MINE_INSERT_R{offset + 1}_C{col}"
        ws.delete_rows(40, 2)
        return {"before": before, "after": (ws.max_row, ws.max_column), "edits": len(edits), "insert_rows": 3, "delete_rows": 2}

    mine = os.path.join(case_root, "mine", basename)
    mine_details = mutate_copy(src, mine, mutate_mine)

    def mutate_theirs(wb):
        ws = wb[CASES[stem]["sheet"]]
        before = (ws.max_row, ws.max_column)
        original_headers = [ws.cell(1, col).value for col in (22, 23)]
        edits = {
            (10, 1): "UX_THEIRS_CONFLICT",
            (11, 2): "UX_THEIRS_EDIT_R11",
            (12, 3): "UX_THEIRS_EDIT_R12",
            (13, 4): "UX_THEIRS_EDIT_R13",
        }
        for (row, col), value in edits.items():
            ws.cell(row=row, column=col).value = value
        ws.insert_rows(25, 4)
        for offset in range(4):
            for col in range(1, 5):
                ws.cell(row=25 + offset, column=col).value = f"UX_THEIRS_INSERT_R{offset + 1}_C{col}"
        ws.delete_rows(50, 3)
        ws.insert_cols(12, 2)
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=12).value = f"__UX_INS1_R{row}"
            ws.cell(row=row, column=13).value = f"__UX_INS2_R{row}"
        ws.delete_cols(24, 2)  # original columns 22-23 after the two-column insertion
        return {
            "before": before, "after": (ws.max_row, ws.max_column), "edits": len(edits),
            "insert_rows": 4, "delete_rows": 3, "insert_cols": 2, "delete_cols": 2,
            "deleted_original_headers": original_headers, "marker_headers": ["__UX_INS1_R1", "__UX_INS2_R1"],
        }

    theirs = os.path.join(case_root, "theirs", basename)
    theirs_details = mutate_copy(src, theirs, mutate_theirs)
    result["cases"][stem] = {
        "base": base, "mine": mine, "theirs": theirs,
        "hashes": {"base": sha256(base), "mine": sha256(mine), "theirs": sha256(theirs)},
        "mine_fixture": mine_details, "theirs_fixture": theirs_details,
    }

    # Pure structural branch variants for formula-heavy Skill and wide Dungeon.
    for stem in ("Skill", "Dungeon"):
        basename = stem + ".xlsx"
        src = os.path.join(SOURCE_ROOT, basename)
        case_root = os.path.join(ROOT, "cases", stem)
        mine = copy_same_name(src, os.path.join(case_root, "mine"), basename)

        theirs = os.path.join(case_root, "theirs", basename)
        fixture_details = excel_column_fixture(
            src,
            theirs,
            CASES[stem]["sheet"],
            restore_skill_formula_caches=(stem == "Skill"),
        )
        mine_reopen_started = now_ms()
        mine_excel_reopen = smt._excel_reopen_validate(mine)
        mine_reopen_ms = now_ms() - mine_reopen_started
        theirs_reopen_started = now_ms()
        theirs_excel_reopen = smt._excel_reopen_validate(theirs)
        theirs_reopen_ms = now_ms() - theirs_reopen_started
        if not mine_excel_reopen or not theirs_excel_reopen:
            raise RuntimeError(
                f"Excel fixture reopen failed for {stem}: mine={mine_excel_reopen} theirs={theirs_excel_reopen}"
            )
        fixture_details["independent_excel_reopen"] = {
            "mine": bool(mine_excel_reopen), "mine_ms": round(mine_reopen_ms, 2),
            "theirs": bool(theirs_excel_reopen), "theirs_ms": round(theirs_reopen_ms, 2),
        }
        fixture_details["formula_semantics"] = formula_semantic_summary(
            mine, theirs, CASES[stem]["sheet"]
        )
        result["cases"][stem] = {
            "mine": mine, "theirs": theirs,
            "hashes": {"mine": sha256(mine), "theirs": sha256(theirs)},
            "fixture": fixture_details,
        }

    # Large real workbook with independent local/remote cell edits.
    stem = "WorldMonster"
    basename = stem + ".xlsx"
    src = os.path.join(SOURCE_ROOT, basename)
    case_root = os.path.join(ROOT, "cases", stem)

    def mutate_world_mine(wb):
        ws = wb[CASES[stem]["sheet"]]
        before = ws.cell(3000, 2).value
        ws.cell(3000, 2).value = "UX_WM_MINE_LOCAL_UNCOMMITTED"
        return {"cell": "B3000", "before": before, "after": ws.cell(3000, 2).value, "shape": (ws.max_row, ws.max_column)}

    def mutate_world_theirs(wb):
        ws = wb[CASES[stem]["sheet"]]
        before = ws.cell(3000, 3).value
        ws.cell(3000, 3).value = "UX_WM_THEIRS_COMMITTED"
        return {"cell": "C3000", "before": before, "after": ws.cell(3000, 3).value, "shape": (ws.max_row, ws.max_column)}

    mine = os.path.join(case_root, "mine", basename)
    theirs = os.path.join(case_root, "theirs", basename)
    mine_details = mutate_copy(src, mine, mutate_world_mine)
    theirs_details = mutate_copy(src, theirs, mutate_world_theirs)
    result["cases"][stem] = {
        "mine": mine, "theirs": theirs,
        "hashes": {"mine": sha256(mine), "theirs": sha256(theirs)},
        "mine_fixture": mine_details, "theirs_fixture": theirs_details,
    }

    with open(REPORT, "w", encoding="utf-8") as stream:
        json.dump({"prepare": result}, stream, ensure_ascii=False, indent=2, default=str)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str), flush=True)


def load_report():
    if not os.path.exists(REPORT):
        return {}
    with open(REPORT, "r", encoding="utf-8") as stream:
        return json.load(stream)


def save_section(name, value):
    report = load_report()
    report[name] = value
    with open(REPORT, "w", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2, default=str)


def wait_edit_ready(app, timeout=180.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        _pump(app.root, 0.05)
        if app._edit_loaded_event.is_set() and app._edit_workbooks_ready():
            return True
    raise TimeoutError("editable workbook preload did not finish")


def wait_view_ready(view, timeout=240.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        _pump(view.root, 0.05)
        if view._derive_lifecycle_state() == "READY":
            return True
    raise TimeoutError(
        f"Sheet lifecycle did not reach READY: {view._derive_lifecycle_state()}"
    )


def remaining_conflicts(app):
    return sum(len(cols) for rows in app.merge_conflict_cells_by_sheet.values() for cols in rows.values())


def view_metrics(view):
    projection = view._ensure_column_projection_current("5.3验收")
    structural = sorted(int(v) for v in view.column_comparison_cache.structural_diff_cols)
    unresolved = sorted(int(v) for v in view.column_comparison_cache.unresolved_cols)
    blocks = []
    for block in projection.model.blocks:
        if not view._column_block_is_structural(block):
            continue
        blocks.append({
            "ordinal": block.ordinal,
            "logical": [block.start_slot_idx + 1, block.end_slot_idx + 1],
            "slots": list(block.slot_indices),
        })
    line = view.left_colhdr.get("1.0", "1.end")
    spans = view._spans_for_line(line)
    misses = []
    for logical_col in range(1, projection.slot_count + 1):
        span = spans.get(logical_col)
        if not span:
            misses.append({"logical": logical_col, "reason": "missing-span"})
            continue
        midpoint = (int(span[0]) + int(span[1]) - 1) // 2
        mapped = view._col_from_char(midpoint)
        if mapped != logical_col:
            misses.append({"logical": logical_col, "mapped": mapped})
    visual_diff_pairs = sum(1 for idx in range(len(view.row_pairs)) if view._pair_has_visual_diff(idx))
    missing_a = sum(1 for a, _b in view.row_pairs if a is None)
    missing_b = sum(1 for _a, b in view.row_pairs if b is None)
    return {
        "physical": {"mine_cols": view.col_max_a, "theirs_cols": view.col_max_b, "max_rows": view.max_row},
        "logical_slots": projection.slot_count,
        "structural_cols": structural,
        "structural_slot_details": [
            {
                "logical": logical_col,
                "mine_col": projection.slot(logical_col).mine_col,
                "base_col": projection.slot(logical_col).base_col,
                "theirs_col": projection.slot(logical_col).theirs_col,
                "state": projection.slot(logical_col).state,
                "ambiguous": bool(projection.slot(logical_col).confidence.ambiguous),
            }
            for logical_col in structural
        ],
        "unresolved_cols": unresolved,
        "structural_blocks": blocks,
        "row_pairs": len(view.row_pairs), "mine_missing_pairs": missing_a, "theirs_missing_pairs": missing_b,
        "visual_diff_pairs": visual_diff_pairs,
        "header_hit_test_misses": misses,
    }


def only_diff_metrics(view):
    if bool(view.only_diff_var.get()):
        view.only_diff_var.set(0)
        view._toggle_only_diff()
        mode_deadline = time.time() + 10.0
        while bool(getattr(view, "_mode_switch_pending", False)) and time.time() < mode_deadline:
            _pump(view.root, 0.02)
    try:
        state_before = str(view.root.state())
        geometry_before = str(view.root.geometry())
    except Exception:
        state_before = ""
        geometry_before = ""
    app = view.app
    app._ui_heartbeat_max_gap = 0.0
    app._ui_heartbeat_last = time.perf_counter()
    view.only_diff_var.set(1)
    started = now_ms()
    view._toggle_only_diff()
    callback_elapsed = now_ms() - started
    deadline = time.time() + 240.0
    while time.time() < deadline:
        _pump(view.root, 0.05)
        if (
            not bool(getattr(view, "_only_diff_async_building", False))
            and not bool(getattr(view, "_mode_switch_pending", False))
            and view._has_valid_only_diff_snapshot_cache()
        ):
            break
    elapsed = now_ms() - started
    try:
        state_after = str(view.root.state())
        geometry_after = str(view.root.geometry())
    except Exception:
        state_after = ""
        geometry_after = ""
    result = {
        "toggle_ms": round(elapsed, 2),
        "callback_ms": round(callback_elapsed, 2),
        "heartbeat_max_gap_ms": round(
            float(getattr(app, "_ui_heartbeat_max_gap", 0.0)) * 1000.0,
            2,
        ),
        "window_state_before": state_before,
        "window_state_after": state_after,
        "window_geometry_before": geometry_before,
        "window_geometry_after": geometry_after,
        "display_rows": len(view.display_rows),
        "snapshot_only_diff": bool(view.snapshot_only_diff),
        "async_building": bool(getattr(view, "_only_diff_async_building", False)),
        "cache_valid": bool(view._has_valid_only_diff_snapshot_cache()),
        "cached_diff_rows": len(getattr(view, "_only_diff_rows_cache", None) or []),
    }
    view.only_diff_var.set(0)
    view._toggle_only_diff()
    mode_deadline = time.time() + 10.0
    while bool(getattr(view, "_mode_switch_pending", False)) and time.time() < mode_deadline:
        _pump(view.root, 0.02)
    return result


def action_button_state(view):
    result = {"status": view.column_action_status_var.get()}
    for attr in ("use_mine_col_btn", "use_base_col_btn", "use_theirs_col_btn"):
        widget = getattr(view, attr, None)
        if widget is not None:
            result[attr] = {"text": widget.cget("text"), "state": str(widget.cget("state"))}
    return result


def apply_all_structural(view, app, *, source_side="B", validate_undo=True):
    actions = []
    undo_checked = False
    for _iteration in range(12):
        projection = view._ensure_column_projection_current("5.3列块动作")
        structural_cols = set(int(v) for v in view.column_comparison_cache.structural_diff_cols)
        blocks = [
            block for block in projection.model.blocks
            if view._column_block_is_structural(block)
            and any((int(slot_idx) + 1) in structural_cols for slot_idx in block.slot_indices)
        ]
        if not blocks:
            break
        block = blocks[0]
        logical_col = int(block.start_slot_idx) + 1
        selected = view._select_column_block_by_logical_col(logical_col, source_side)
        if selected is None:
            raise AssertionError(f"cannot select structural column L{logical_col}")
        selected_live = view._selected_column_block()
        debug_values = {
            "requested_logical": logical_col,
            "selected_ordinal": getattr(selected_live, "ordinal", None),
            "selected_slots": list(getattr(selected_live, "slot_indices", ()) or ()),
            "source_values": [projection.physical_col(source_side, int(idx) + 1) for idx in getattr(selected_live, "slot_indices", ())],
            "target_values": [projection.physical_col("A", int(idx) + 1) for idx in getattr(selected_live, "slot_indices", ())],
        }
        print("ACTION_DEBUG " + json.dumps(debug_values, ensure_ascii=False), flush=True)
        before_model = _model_snapshot(view) if validate_undo and not undo_checked else None
        before_selection = _selection_snapshot(view) if validate_undo and not undo_checked else None
        buttons = action_button_state(view)
        started = now_ms()
        plan = view._apply_selected_column_block(source_side, confirm_unresolved=True)
        apply_ms = now_ms() - started
        _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
        item = {
            "plan": {
                "kind": plan.action_kind, "source": plan.source_side, "target": plan.target_side,
                "logical": [plan.logical_start, plan.logical_end], "count": plan.count,
                "source_physical_cols": list(plan.source_physical_cols),
                "target_physical_cols": list(plan.target_physical_cols),
                "target_anchor": plan.target_physical_anchor,
            },
            "apply_ms": round(apply_ms, 2), "buttons_before": buttons,
        }
        if validate_undo and not undo_checked:
            action = app.undo_stack.pop()
            started = now_ms()
            ok = view._undo_column_action(action)
            undo_ms = now_ms() - started
            _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
            item["undo"] = {
                "ok": bool(ok), "undo_ms": round(undo_ms, 2),
                "model_exact": _model_snapshot(view) == before_model,
                "selection_exact": _selection_snapshot(view) == before_selection,
                "status": view.column_action_status_var.get(),
            }
            # Reapply the exact selection restored by undo.
            started = now_ms()
            reapplied = view._apply_selected_column_block(source_side, confirm_unresolved=True)
            item["reapply_ms"] = round(now_ms() - started, 2)
            item["reapply_kind"] = reapplied.action_kind
            _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
            undo_checked = True
        actions.append(item)
    remaining = sorted(int(v) for v in view.column_comparison_cache.structural_diff_cols)
    if remaining:
        raise AssertionError(f"structural diffs remain after actions: {remaining}")
    return actions


def apply_first_structural_with_undo(view, app, *, source_side="B"):
    projection = view._ensure_column_projection_current("5.3首个列块动作")
    structural_cols = set(int(v) for v in view.column_comparison_cache.structural_diff_cols)
    blocks = [
        block for block in projection.model.blocks
        if view._column_block_is_structural(block)
        and any((int(slot_idx) + 1) in structural_cols for slot_idx in block.slot_indices)
    ]
    if not blocks:
        raise AssertionError("no structural column block available")
    logical_col = int(blocks[0].start_slot_idx) + 1
    if view._select_column_block_by_logical_col(logical_col, source_side) is None:
        raise AssertionError(f"cannot select structural column L{logical_col}")
    before_model = _model_snapshot(view)
    before_selection = _selection_snapshot(view)
    buttons = action_button_state(view)
    started = now_ms()
    plan = view._apply_selected_column_block(source_side, confirm_unresolved=True)
    apply_ms = now_ms() - started
    _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
    after_first = view_metrics(view)
    action = app.undo_stack.pop()
    started = now_ms()
    undo_ok = view._undo_column_action(action)
    undo_ms = now_ms() - started
    _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
    undo_model_exact = _model_snapshot(view) == before_model
    undo_selection_exact = _selection_snapshot(view) == before_selection
    started = now_ms()
    reapplied = view._apply_selected_column_block(source_side, confirm_unresolved=True)
    reapply_ms = now_ms() - started
    _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
    after_reapply = view_metrics(view)
    return {
        "plan": {
            "kind": plan.action_kind, "source": plan.source_side, "target": plan.target_side,
            "logical": [plan.logical_start, plan.logical_end], "count": plan.count,
            "source_physical_cols": list(plan.source_physical_cols),
            "target_physical_cols": list(plan.target_physical_cols),
            "target_anchor": plan.target_physical_anchor,
        },
        "buttons_before": buttons, "apply_ms": round(apply_ms, 2),
        "after_first_apply": after_first,
        "undo": {
            "ok": bool(undo_ok), "undo_ms": round(undo_ms, 2),
            "model_exact": undo_model_exact, "selection_exact": undo_selection_exact,
        },
        "reapply_ms": round(reapply_ms, 2), "reapply_kind": reapplied.action_kind,
        "after_reapply": after_reapply,
    }


def workbook_probe(path, sheet, markers=()):
    started = now_ms()
    wb = smt.load_workbook(path, data_only=False, read_only=True)
    open_ms = now_ms() - started
    try:
        ws = wb[sheet]
        formula_count = 0
        marker_hits = {marker: [] for marker in markers}
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formula_count += 1
                if value in marker_hits and len(marker_hits[value]) < 5:
                    marker_hits[value].append(cell.coordinate)
        return {
            "reopen_ms": round(open_ms, 2), "rows": ws.max_row, "cols": ws.max_column,
            "formula_count": formula_count, "marker_hits": marker_hits,
        }
    finally:
        close_wb(wb)


def guide_output_contract(base_path, mine_path, theirs_path, output_path, sheet, sample_limit=12):
    """Verify every Guide mutation that the scripted user explicitly authorized."""
    workbooks = []
    try:
        base_wb = smt.load_workbook(base_path, data_only=False, read_only=False)
        workbooks.append(base_wb)
        mine_wb = smt.load_workbook(mine_path, data_only=False, read_only=False)
        workbooks.append(mine_wb)
        theirs_wb = smt.load_workbook(theirs_path, data_only=False, read_only=False)
        workbooks.append(theirs_wb)
        output_wb = smt.load_workbook(output_path, data_only=False, read_only=False)
        workbooks.append(output_wb)
        base_ws = base_wb[sheet]
        mine_ws = mine_wb[sheet]
        theirs_ws = theirs_wb[sheet]
        output_ws = output_wb[sheet]

        expected_mine_markers = {
            "A10": "UX_MINE_CONFLICT",
            "B11": "UX_MINE_EDIT_R11",
            "C12": "UX_MINE_EDIT_R12",
            "D13": "UX_MINE_EDIT_R13",
        }
        for row_offset in range(3):
            for col_idx in range(1, 5):
                coordinate = f"{smt.get_column_letter(col_idx)}{20 + row_offset}"
                expected_mine_markers[coordinate] = (
                    f"UX_MINE_INSERT_R{row_offset + 1}_C{col_idx}"
                )
        mine_marker_mismatches = [
            {"coordinate": coordinate, "expected": expected, "actual": mine_ws[coordinate].value}
            for coordinate, expected in expected_mine_markers.items()
            if mine_ws[coordinate].value != expected
        ]

        payload_mismatches = []
        payload_diff_count = 0
        for row_idx in range(1, mine_ws.max_row + 1):
            for mine_col in range(1, mine_ws.max_column + 1):
                output_col = mine_col if mine_col < 12 else mine_col + 2
                expected = mine_ws.cell(row=row_idx, column=mine_col).value
                if row_idx == 10 and mine_col == 1:
                    expected = theirs_ws.cell(row=10, column=1).value
                actual = output_ws.cell(row=row_idx, column=output_col).value
                if smt._merge_cmp_value(expected) == smt._merge_cmp_value(actual):
                    continue
                payload_diff_count += 1
                if len(payload_mismatches) < sample_limit:
                    payload_mismatches.append({
                        "mine_coordinate": f"{smt.get_column_letter(mine_col)}{row_idx}",
                        "output_coordinate": f"{smt.get_column_letter(output_col)}{row_idx}",
                        "expected": expected,
                        "actual": actual,
                    })

        marker_mismatches = []
        marker_mismatch_count = 0
        for row_idx in range(1, output_ws.max_row + 1):
            expected_l = f"__UX_INS1_R{row_idx}"
            expected_m = f"__UX_INS2_R{row_idx}"
            actual_l = output_ws.cell(row=row_idx, column=12).value
            actual_m = output_ws.cell(row=row_idx, column=13).value
            if actual_l != expected_l or actual_m != expected_m:
                marker_mismatch_count += 1
                if len(marker_mismatches) < sample_limit:
                    marker_mismatches.append({
                        "row": row_idx,
                        "expected": [expected_l, expected_m],
                        "actual": [actual_l, actual_m],
                    })

        unresolved_mine_choices = {
            coordinate: output_ws[coordinate].value
            for coordinate in ("B11", "C12", "D13")
        }
        authorized_conflict_choice = output_ws["A10"].value
        theirs_row_markers_found = 0
        for row in output_ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("UX_THEIRS_INSERT_"):
                    theirs_row_markers_found += 1

        shape_contract = {
            "base": [base_ws.max_row, base_ws.max_column],
            "mine": [mine_ws.max_row, mine_ws.max_column],
            "theirs": [theirs_ws.max_row, theirs_ws.max_column],
            "output": [output_ws.max_row, output_ws.max_column],
            "mine_row_insert_delete_exact": (
                mine_ws.max_row == base_ws.max_row + 3 - 2
            ),
            "theirs_row_insert_delete_exact": (
                theirs_ws.max_row == base_ws.max_row + 4 - 3
            ),
            "theirs_column_insert_delete_exact": (
                theirs_ws.max_column == base_ws.max_column + 2 - 2
            ),
            "output_authorized_column_insert_exact": (
                output_ws.max_column == mine_ws.max_column + 2
            ),
        }
        contract_pass = bool(
            not mine_marker_mismatches and
            payload_diff_count == 0 and
            not marker_mismatches and
            authorized_conflict_choice == "UX_THEIRS_CONFLICT" and
            unresolved_mine_choices == {
                "B11": "UX_MINE_EDIT_R11",
                "C12": "UX_MINE_EDIT_R12",
                "D13": "UX_MINE_EDIT_R13",
            } and
            theirs_row_markers_found == 0 and
            all(value for key, value in shape_contract.items() if key.endswith("_exact"))
        )
        return {
            "pass": contract_pass,
            "shape_contract": shape_contract,
            "mine_marker_mismatches": mine_marker_mismatches,
            "authorized_conflict_choice": authorized_conflict_choice,
            "unresolved_mine_choices": unresolved_mine_choices,
            "theirs_unselected_row_marker_count": theirs_row_markers_found,
            "authorized_inserted_column_rows": output_ws.max_row - marker_mismatch_count,
            "inserted_column_marker_mismatch_count": marker_mismatch_count,
            "inserted_column_marker_mismatches": marker_mismatches,
            "unexpected_payload_diff_count": payload_diff_count,
            "unexpected_payload_diff_samples": payload_mismatches,
        }
    finally:
        for wb in workbooks:
            close_wb(wb)


def copy_output(temp_output, final_output):
    os.makedirs(os.path.dirname(final_output), exist_ok=True)
    shutil.copy2(temp_output, final_output)
    return final_output


def guide():
    case_root = os.path.join(ROOT, "cases", "Guide")
    base = os.path.join(case_root, "base", "Guide.xlsx")
    mine = os.path.join(case_root, "mine", "Guide.xlsx")
    theirs = os.path.join(case_root, "theirs", "Guide.xlsx")
    result = {"paths": {"base": base, "mine": mine, "theirs": theirs}}
    control_started = now_ms()
    control_conflicts, control_map = smt._scan_three_way_conflicts(base, base, base)
    result["identical_control_scan_ms"] = round(now_ms() - control_started, 2)
    result["identical_control_conflict_count"] = len(control_conflicts)
    result["identical_control_conflict_sheets"] = sorted(control_map)
    scan_started = now_ms()
    conflicts, conflict_map = smt._scan_three_way_conflicts(base, mine, theirs)
    result["three_way_scan_ms"] = round(now_ms() - scan_started, 2)
    result["conflict_count"] = len(conflicts)
    result["conflict_sheets"] = {sheet: {str(row): sorted(cols) for row, cols in rows.items()} for sheet, rows in conflict_map.items()}
    result["same_cell_conflict_detected"] = any(sheet == CASES["Guide"]["sheet"] and row == 10 and col == 1 for sheet, row, col, _a, _b in conflicts)

    app = None
    try:
        started = now_ms()
        app = smt.SowMergeApp(
            mine, theirs, merge_mode=True, base_path=base,
            raw_mine=mine, raw_base=base, raw_theirs=theirs,
            merge_conflict_cells_by_sheet=conflict_map, merge_conflict_mode=True,
        )
        result["constructor_ms"] = round(now_ms() - started, 2)
        app.root.withdraw()
        view = _force_full_view(_wait_for_view(app, CASES["Guide"]["sheet"], timeout=180.0))
        wait_edit_ready(app)
        _wait_for_stable_projection(view, timeout=180.0, stable_for=0.5)
        result["ready_ms"] = round(now_ms() - started, 2)
        result["view_before"] = view_metrics(view)
        result["only_diff"] = only_diff_metrics(view)
        result["actions"] = []
        try:
            result["actions"].append(apply_first_structural_with_undo(view, app, source_side="B"))
        except Exception as action_exc:
            result["column_action_error"] = repr(action_exc)
            result["column_action_traceback"] = traceback.format_exc()
        result["view_after_columns"] = view_metrics(view)
        result["only_diff_after_column_action"] = only_diff_metrics(view)

        conflict_before = remaining_conflicts(app)
        pair_idx = next((idx for idx, pair in enumerate(view.row_pairs) if pair == (10, 10)), None)
        if pair_idx is None:
            raise AssertionError("Guide row 10 mine/theirs pair not found")
        started_cell = now_ms()
        view._copy_single_cell_by_pair(pair_idx, "B2A", 1)
        _pump(app.root, 0.2)
        result["conflict_action"] = {
            "pair_idx": pair_idx, "direction": "B2A", "logical_col": 1,
            "ms": round(now_ms() - started_cell, 2),
            "before": conflict_before, "after": remaining_conflicts(app),
            "value_after": app.ws_a_edit(CASES["Guide"]["sheet"]).cell(10, 1).value,
        }

        save_started = now_ms()
        temp_output = app.build_manual_merge_output_file()
        result["save_ms"] = round(now_ms() - save_started, 2)
        final_output = copy_output(temp_output, os.path.join(case_root, "output", "Guide.xlsx"))
        package_ok, package_error = smt._validate_xlsx_package(final_output)
        result["output"] = {
            "path": final_output, "sha256": sha256(final_output), "package_ok": package_ok,
            "package_error": package_error, "size": os.path.getsize(final_output),
        }
        probe = workbook_probe(
            final_output, CASES["Guide"]["sheet"],
            markers=("UX_THEIRS_CONFLICT", "UX_MINE_EDIT_R11", "UX_MINE_INSERT_R1_C1", "__UX_INS1_R1", "__UX_INS2_R1"),
        )
        result["reopen"] = probe
        result["reopen"]["row10_col1"] = probe["marker_hits"].get("UX_THEIRS_CONFLICT")
        result["output_contract"] = guide_output_contract(
            base,
            mine,
            theirs,
            final_output,
            CASES["Guide"]["sheet"],
        )
        column_action_ok = bool(
            result["actions"] and not result.get("column_action_error") and
            result["view_after_columns"]["structural_cols"] == [] and
            all(
                item["undo"]["ok"] and item["undo"]["model_exact"] and item["undo"]["selection_exact"]
                for item in result["actions"]
            )
        )
        result["column_action_pass"] = column_action_ok
        only_diff_after_ok = bool(
            result["only_diff_after_column_action"]["cache_valid"] and
            result["only_diff_after_column_action"]["cached_diff_rows"] < 100
        )
        result["only_diff_after_column_action_pass"] = only_diff_after_ok
        # Native cross-workbook full-column replay has an Excel lifecycle
        # floor that is separate from cell-only/ZIP saves.  The small native
        # tier is calibrated at 7.5 s; formula-dense native cases below use
        # the 12 s single-run hard ceiling recorded in the change baseline.
        result["save_performance_limit_ms"] = 7500.0
        result["save_performance_pass"] = (
            result["save_ms"] <= result["save_performance_limit_ms"]
        )
        result["pass"] = bool(
            result["identical_control_conflict_count"] == 0 and
            result["same_cell_conflict_detected"] and
            result["conflict_action"]["after"] == result["conflict_action"]["before"] - 1 and
            result["conflict_action"]["value_after"] == "UX_THEIRS_CONFLICT" and
            column_action_ok and only_diff_after_ok and result["save_performance_pass"] and
            package_ok and result["output_contract"]["pass"] and
            all(probe["marker_hits"].get(marker) for marker in ("UX_THEIRS_CONFLICT", "UX_MINE_EDIT_R11", "UX_MINE_INSERT_R1_C1", "__UX_INS1_R1", "__UX_INS2_R1"))
        )
    except Exception as exc:
        result["pass"] = False
        result["error"] = repr(exc)
        result["traceback"] = traceback.format_exc()
    finally:
        close_app(app)
    save_section("Guide", result)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str), flush=True)
    if not result.get("pass"):
        sys.exit(2)


def two_way(stem):
    case_root = os.path.join(ROOT, "cases", stem)
    mine = os.path.join(case_root, "mine", stem + ".xlsx")
    theirs = os.path.join(case_root, "theirs", stem + ".xlsx")
    sheet = CASES[stem]["sheet"]
    result = {"paths": {"mine": mine, "theirs": theirs}, "sheet": sheet}
    result["theirs_probe"] = workbook_probe(theirs, sheet, markers=("__UX_INS1_R1", "__UX_INS2_R1"))
    app = None
    try:
        started = now_ms()
        app = smt.SowMergeApp(mine, theirs)
        result["constructor_ms"] = round(now_ms() - started, 2)
        app.root.withdraw()
        view = _force_full_view(_wait_for_view(app, sheet, timeout=240.0))
        wait_edit_ready(app, timeout=240.0)
        wait_view_ready(view, timeout=240.0)
        _wait_for_stable_projection(view, timeout=240.0, stable_for=0.5)
        result["ready_ms"] = round(now_ms() - started, 2)
        result["view_before"] = view_metrics(view)
        result["only_diff"] = only_diff_metrics(view)
        result["actions"] = apply_all_structural(view, app, source_side="B", validate_undo=True)
        result["view_after"] = view_metrics(view)
        result["only_diff_after_actions"] = only_diff_metrics(view)
        save_started = now_ms()
        temp_output = app.build_manual_merge_output_file()
        result["save_ms"] = round(now_ms() - save_started, 2)
        final_output = copy_output(temp_output, os.path.join(case_root, "output", stem + ".xlsx"))
        package_ok, package_error = smt._validate_xlsx_package(final_output)
        result["output"] = {
            "path": final_output, "sha256": sha256(final_output), "package_ok": package_ok,
            "package_error": package_error, "size": os.path.getsize(final_output),
        }
        result["reopen"] = workbook_probe(final_output, sheet, markers=("__UX_INS1_R1", "__UX_INS2_R1"))
        actions_have_insert2 = any(item["plan"]["kind"] == "insert_copy" and item["plan"]["count"] == 2 for item in result["actions"])
        actions_have_delete2 = any(item["plan"]["kind"] == "delete" and item["plan"]["count"] == 2 for item in result["actions"])
        row_alignment_ok = (
            result["view_before"]["mine_missing_pairs"] == 0 and
            result["view_before"]["theirs_missing_pairs"] == 0
        )
        only_diff_ok = bool(
            result["only_diff_after_actions"]["cache_valid"] and
            (
                result["only_diff_after_actions"]["cached_diff_rows"] == 0
                if stem == "Skill"
                else result["only_diff_after_actions"]["cached_diff_rows"] > 0
            )
        )
        result["row_alignment_pass"] = row_alignment_ok
        result["only_diff_expectation"] = "0 rows (formula-stable pure structure)" if stem == "Skill" else ">0 rows (Excel-rewritten formula semantics must remain visible)"
        result["only_diff_pass"] = only_diff_ok
        result["save_performance_limit_ms"] = 12000.0
        result["save_performance_pass"] = (
            result["save_ms"] <= result["save_performance_limit_ms"]
        )
        result["pass"] = bool(
            result["view_before"]["header_hit_test_misses"] == [] and
            row_alignment_ok and only_diff_ok and
            actions_have_insert2 and actions_have_delete2 and result["view_after"]["structural_cols"] == [] and
            all(not item.get("undo") or (item["undo"]["ok"] and item["undo"]["model_exact"] and item["undo"]["selection_exact"]) for item in result["actions"]) and
            result["save_performance_pass"] and package_ok and result["reopen"]["rows"] == result["theirs_probe"]["rows"] and
            result["reopen"]["cols"] == result["theirs_probe"]["cols"] and
            result["reopen"]["formula_count"] == result["theirs_probe"]["formula_count"] and
            all(result["reopen"]["marker_hits"].get(marker) for marker in ("__UX_INS1_R1", "__UX_INS2_R1"))
        )
    except Exception as exc:
        result["pass"] = False
        result["error"] = repr(exc)
        result["traceback"] = traceback.format_exc()
    finally:
        close_app(app)
    save_section(stem, result)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str), flush=True)
    if not result.get("pass"):
        sys.exit(2)


def worldmonster():
    stem = "WorldMonster"
    case_root = os.path.join(ROOT, "cases", stem)
    mine = os.path.join(case_root, "mine", stem + ".xlsx")
    theirs = os.path.join(case_root, "theirs", stem + ".xlsx")
    source = os.path.join(SOURCE_ROOT, stem + ".xlsx")
    base = os.path.join(case_root, "base", stem + ".xlsx")
    os.makedirs(os.path.dirname(base), exist_ok=True)
    if not os.path.exists(base) or sha256(base) != sha256(source):
        shutil.copy2(source, base)
    sheet = CASES[stem]["sheet"]
    protected_paths = (source, base, mine, theirs)
    protected_hashes_before = {path: sha256(path) for path in protected_paths}
    result = {
        "paths": {
            "source": source,
            "base": base,
            "mine": mine,
            "theirs": theirs,
        },
        "sheet": sheet,
    }
    app = None
    original_settings_path = smt._SETTINGS_PATH
    isolated_settings_dir = tempfile.mkdtemp(prefix="sow-worldmonster-settings-")
    isolated_settings_path = os.path.join(isolated_settings_dir, "settings.json")
    with open(isolated_settings_path, "w", encoding="utf-8") as stream:
        json.dump({"only_diff": 1}, stream)
    smt._SETTINGS_PATH = isolated_settings_path
    try:
        started = now_ms()
        app = smt.SowMergeApp(mine, theirs, base_path=base)
        result["constructor_ms"] = round(now_ms() - started, 2)
        app.root.withdraw()
        app.nb.select(app._sheet_containers[sheet])
        view_deadline = time.time() + 10.0
        view = app.sheet_views.get(sheet)
        while view is None and time.time() < view_deadline:
            _pump(app.root, 0.01)
            view = app.sheet_views.get(sheet)
        if view is None:
            raise AssertionError(f"WorldMonster view was not created: {sheet}")
        preload_state = {
            "editable_ready_at_request": bool(
                app._edit_loaded_event.is_set() and app._edit_workbooks_ready()
            ),
            "requested_checked": int(view.only_diff_var.get()) == 1,
            "lifecycle": view._derive_lifecycle_state(),
            "checkbox_state": str(view.only_diff_cb.cget("state")),
            "checkbox_text": str(view.only_diff_cb.cget("text")),
        }
        other_sheet = next(
            (candidate for candidate in app.display_sheets if candidate != sheet),
            None,
        )
        if other_sheet is not None:
            app.nb.select(app._sheet_containers[other_sheet])
            _pump(app.root, 0.05)
            preload_state["switched_to"] = app.nb.tab(app.nb.select(), "text")
            app.nb.select(app._sheet_containers[sheet])
            _pump(app.root, 0.05)
            preload_state["switched_back_to"] = app.nb.tab(app.nb.select(), "text")
        result["preload_only_diff"] = preload_state
        view = _wait_for_view(app, sheet, timeout=240.0)
        wait_edit_ready(app, timeout=240.0)
        wait_view_ready(view, timeout=240.0)
        _wait_for_stable_projection(view, timeout=240.0, stable_for=0.5)
        result["ready_ms"] = round(now_ms() - started, 2)
        result["view_before"] = view_metrics(view)
        result["only_diff"] = only_diff_metrics(view)
        result["only_diff_pass"] = bool(
            result["only_diff"]["cache_valid"] and
            result["only_diff"]["cached_diff_rows"] == 1 and
            result["only_diff"]["callback_ms"] < 100.0 and
            result["only_diff"]["heartbeat_max_gap_ms"] <= 200.0 and
            result["only_diff"]["window_state_before"] ==
            result["only_diff"]["window_state_after"] and
            result["only_diff"]["window_geometry_before"] ==
            result["only_diff"]["window_geometry_after"]
        )
        pair_idx = next((idx for idx, pair in enumerate(view.row_pairs) if pair == (3000, 3000)), None)
        if pair_idx is None:
            raise AssertionError("WorldMonster row 3000 pair not found")
        mine_local_before = app.ws_a_edit(sheet).cell(3000, 2).value
        started_cell = now_ms()
        view._copy_single_cell_by_pair(pair_idx, "B2A", 3)
        _pump(app.root, 0.2)
        adopted = app.ws_a_edit(sheet).cell(3000, 3).value
        cell_apply_ms = now_ms() - started_cell
        undo_started = now_ms()
        view._undo_last_action()
        _pump(app.root, 0.2)
        undo_ms = now_ms() - undo_started
        value_after_undo = app.ws_a_edit(sheet).cell(3000, 3).value
        started_cell = now_ms()
        view._copy_single_cell_by_pair(pair_idx, "B2A", 3)
        _pump(app.root, 0.2)
        reapply_ms = now_ms() - started_cell
        result["cell_action"] = {
            "pair_idx": pair_idx, "direction": "B2A", "logical_col": 3,
            "apply_ms": round(cell_apply_ms, 2), "adopted": adopted,
            "undo_ms": round(undo_ms, 2), "value_after_undo": value_after_undo,
            "reapply_ms": round(reapply_ms, 2),
            "mine_local_preserved_in_memory": app.ws_a_edit(sheet).cell(3000, 2).value == mine_local_before,
        }
        save_started = now_ms()
        temp_output = app.build_manual_merge_output_file()
        result["save_ms"] = round(now_ms() - save_started, 2)
        final_output = copy_output(temp_output, os.path.join(case_root, "output", stem + ".xlsx"))
        package_ok, package_error = smt._validate_xlsx_package(final_output)
        result["output"] = {
            "path": final_output, "sha256": sha256(final_output), "package_ok": package_ok,
            "package_error": package_error, "size": os.path.getsize(final_output),
        }
        result["reopen"] = workbook_probe(
            final_output, sheet, markers=("UX_WM_MINE_LOCAL_UNCOMMITTED", "UX_WM_THEIRS_COMMITTED")
        )
        protected_hashes_after = {path: sha256(path) for path in protected_paths}
        result["protected_hashes"] = {
            "before": protected_hashes_before,
            "after": protected_hashes_after,
            "unchanged": protected_hashes_after == protected_hashes_before,
        }
        preload_gate_pass = bool(
            not result["preload_only_diff"]["editable_ready_at_request"] and
            result["preload_only_diff"]["requested_checked"] and
            result["preload_only_diff"]["checkbox_state"] == "disabled" and
            result["preload_only_diff"]["checkbox_text"] == "只看差异内容" and
            (
                other_sheet is None or (
                    result["preload_only_diff"].get("switched_to") == other_sheet and
                    result["preload_only_diff"].get("switched_back_to") == sheet
                )
            )
        )
        result["preload_gate_pass"] = preload_gate_pass
        result["pass"] = bool(
            result["view_before"]["header_hit_test_misses"] == [] and
            result["view_before"]["structural_cols"] == [] and
            preload_gate_pass and
            result["only_diff_pass"] and
            result["only_diff"]["display_rows"] == 1 and
            adopted == "UX_WM_THEIRS_COMMITTED" and
            value_after_undo != "UX_WM_THEIRS_COMMITTED" and
            result["cell_action"]["mine_local_preserved_in_memory"] and
            result["protected_hashes"]["unchanged"] and
            package_ok and
            all(result["reopen"]["marker_hits"].get(marker) for marker in ("UX_WM_MINE_LOCAL_UNCOMMITTED", "UX_WM_THEIRS_COMMITTED"))
        )
    except Exception as exc:
        result["pass"] = False
        result["error"] = repr(exc)
        result["traceback"] = traceback.format_exc()
    finally:
        close_app(app)
        smt._SETTINGS_PATH = original_settings_path
        shutil.rmtree(isolated_settings_dir, ignore_errors=True)
    save_section(stem, result)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str), flush=True)
    if not result.get("pass"):
        sys.exit(2)


def svn_conflict():
    case_root = os.path.join(ROOT, "cases", "Guide")
    wc = os.path.join(ROOT, "svn_wc")
    os.makedirs(wc, exist_ok=True)
    target = os.path.join(wc, "Guide.xlsx")
    left = target + ".merge-left.r100"
    right = target + ".merge-right.r101"
    mine_artifact = target + ".mine"
    shutil.copy2(os.path.join(case_root, "mine", "Guide.xlsx"), target)
    shutil.copy2(os.path.join(case_root, "base", "Guide.xlsx"), left)
    shutil.copy2(os.path.join(case_root, "theirs", "Guide.xlsx"), right)
    shutil.copy2(os.path.join(case_root, "mine", "Guide.xlsx"), mine_artifact)
    started = now_ms()
    detected = smt._detect_svn_conflict_files(target)
    detect_ms = now_ms() - started
    has = smt._has_svn_conflict_artifacts(target)
    found = smt._find_conflict_in_dir(wc)
    result = {
        "transport": "canonical isolated SVN conflict artifacts; no real repository mutation",
        "cli_available": bool(shutil.which("svn")),
        "paths": {"target_local_uncommitted": target, "base_r100": left, "theirs_committed_r101": right, "mine_artifact": mine_artifact},
        "hash_roles": {"target_mine": sha256(target), "base": sha256(left), "theirs": sha256(right)},
        "detect_ms": round(detect_ms, 2), "has_artifacts": has,
        "detected": list(detected) if detected else None, "find_conflict_in_dir": found,
    }
    expected = [os.path.abspath(left), os.path.abspath(target), os.path.abspath(right), os.path.abspath(target)]
    result["pass"] = bool(has and detected and [os.path.abspath(x) for x in detected] == expected and os.path.abspath(found) == os.path.abspath(target))
    save_section("SVN", result)
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    if not result["pass"]:
        sys.exit(2)


def sheet_diff_summary(path_a, path_b, sheet, limit=20):
    wb_a = smt.load_workbook(path_a, data_only=False, read_only=True)
    wb_b = smt.load_workbook(path_b, data_only=False, read_only=True)
    try:
        ws_a = wb_a[sheet]
        ws_b = wb_b[sheet]
        max_row = max(ws_a.max_row, ws_b.max_row)
        max_col = max(ws_a.max_column, ws_b.max_column)
        count = 0
        samples = []
        rows_a = ws_a.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True)
        rows_b = ws_b.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True)
        def comparable(value):
            signature = smt._special_formula_signature(value)
            return ("SPECIAL_FORMULA", signature) if signature is not None else value
        for row_idx, (row_a, row_b) in enumerate(zip(rows_a, rows_b), start=1):
            for col_idx, (value_a, value_b) in enumerate(zip(row_a, row_b), start=1):
                if comparable(value_a) == comparable(value_b):
                    continue
                count += 1
                if len(samples) < limit:
                    samples.append({"row": row_idx, "col": col_idx, "a": value_a, "b": value_b})
        return {"diff_count": count, "samples": samples, "shape_a": [ws_a.max_row, ws_a.max_column], "shape_b": [ws_b.max_row, ws_b.max_column]}
    finally:
        close_wb(wb_a)
        close_wb(wb_b)


def post_validate():
    result = {}
    output_paths = {
        "Guide": os.path.join(ROOT, "cases", "Guide", "output", "Guide.xlsx"),
        "Skill": os.path.join(ROOT, "cases", "Skill", "output", "Skill.xlsx"),
        "Dungeon": os.path.join(ROOT, "cases", "Dungeon", "output", "Dungeon.xlsx"),
        "WorldMonster": os.path.join(ROOT, "cases", "WorldMonster", "output", "WorldMonster.xlsx"),
    }
    for stem, path in output_paths.items():
        started = now_ms()
        ok = smt._excel_reopen_validate(path) if os.path.exists(path) else False
        result.setdefault(stem, {})["excel_reopen_ok"] = bool(ok)
        result[stem]["excel_reopen_ms"] = round(now_ms() - started, 2)
    result["Guide"]["output_contract"] = guide_output_contract(
        os.path.join(ROOT, "cases", "Guide", "base", "Guide.xlsx"),
        os.path.join(ROOT, "cases", "Guide", "mine", "Guide.xlsx"),
        os.path.join(ROOT, "cases", "Guide", "theirs", "Guide.xlsx"),
        output_paths["Guide"],
        CASES["Guide"]["sheet"],
    )
    result["Skill"]["output_vs_theirs"] = sheet_diff_summary(
        output_paths["Skill"], os.path.join(ROOT, "cases", "Skill", "theirs", "Skill.xlsx"), CASES["Skill"]["sheet"]
    )
    result["Dungeon"]["output_vs_theirs"] = sheet_diff_summary(
        output_paths["Dungeon"], os.path.join(ROOT, "cases", "Dungeon", "theirs", "Dungeon.xlsx"), CASES["Dungeon"]["sheet"]
    )
    result["WorldMonster"]["output_vs_mine"] = sheet_diff_summary(
        output_paths["WorldMonster"], os.path.join(ROOT, "cases", "WorldMonster", "mine", "WorldMonster.xlsx"), CASES["WorldMonster"]["sheet"]
    )
    wm_diff = result["WorldMonster"]["output_vs_mine"]
    result["pass"] = bool(
        all(item.get("excel_reopen_ok") for item in result.values() if isinstance(item, dict) and "excel_reopen_ok" in item) and
        result["Guide"]["output_contract"]["pass"] and
        result["Skill"]["output_vs_theirs"]["diff_count"] == 0 and
        result["Dungeon"]["output_vs_theirs"]["diff_count"] == 0 and
        wm_diff["diff_count"] == 1 and wm_diff["samples"] and
        wm_diff["samples"][0]["row"] == 3000 and wm_diff["samples"][0]["col"] == 3 and
        wm_diff["samples"][0]["a"] == "UX_WM_THEIRS_COMMITTED"
    )
    save_section("post_validate", result)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str), flush=True)
    if not result["pass"]:
        sys.exit(2)


def final_hashes():
    report = load_report()
    before = report.get("prepare", {}).get("sources", {})
    result = {}
    all_same = True
    for stem in ("Guide", "Skill", "Dungeon", "WorldMonster"):
        path = os.path.join(SOURCE_ROOT, stem + ".xlsx")
        current = sha256(path)
        expected = before.get(stem, {}).get("sha256")
        same = current == expected
        all_same = all_same and same
        result[stem] = {"path": path, "before": expected, "after": current, "same": same}
    result["pass"] = all_same
    save_section("original_hash_recheck", result)
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    if not all_same:
        sys.exit(2)


def main():
    smt.messagebox.showerror = lambda *args, **kwargs: print("SHOWERROR", args, flush=True)
    smt.messagebox.showwarning = lambda *args, **kwargs: print("SHOWWARNING", args, flush=True)
    smt.messagebox.showinfo = lambda *args, **kwargs: print("SHOWINFO", args, flush=True)
    smt.messagebox.askyesno = lambda *args, **kwargs: True
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("prepare", "guide", "skill", "dungeon", "worldmonster", "svn", "post", "hashes"))
    args = parser.parse_args()
    if args.mode == "prepare":
        prepare()
    elif args.mode == "guide":
        guide()
    elif args.mode == "skill":
        two_way("Skill")
    elif args.mode == "dungeon":
        two_way("Dungeon")
    elif args.mode == "worldmonster":
        worldmonster()
    elif args.mode == "svn":
        svn_conflict()
    elif args.mode == "post":
        post_validate()
    else:
        final_hashes()


if __name__ == "__main__":
    main()
