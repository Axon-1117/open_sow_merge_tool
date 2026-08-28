"""Regression for cached formula keys in ConditionData-style Sheets."""

from __future__ import annotations

import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


_SHEET = "ConditionData@design"
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _write_book(
    path: Path,
    *,
    first_key: int,
    first_label: str,
    cache_formulas: bool,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = _SHEET
    sheet.append(("payload", "key@constid", "component", "label"))
    sheet.append(("string", "string", "int", "string"))
    sheet.append(("descriptor", "descriptor-key", "component", "label"))
    sheet.append((None, None, None, None))
    sheet.append((None, "=C5", first_key, first_label))
    sheet.append((None, "=C6", 20000, "common"))
    workbook.save(path)
    workbook.close()
    if cache_formulas:
        _set_numeric_formula_caches(path, {"B5": first_key, "B6": 20000})


def _set_numeric_formula_caches(path: Path, values: dict[str, int]) -> None:
    with zipfile.ZipFile(path, "r") as source:
        members = {name: source.read(name) for name in source.namelist()}
    root = ET.fromstring(members["xl/worksheets/sheet1.xml"])
    cells = {
        cell.attrib.get("r"): cell
        for cell in root.iter(_NS + "c")
    }
    for coordinate, value in values.items():
        cell = cells[coordinate]
        assert cell.find(_NS + "f") is not None
        cached = cell.find(_NS + "v")
        if cached is None:
            cached = ET.SubElement(cell, _NS + "v")
        cached.text = str(value)
    members["xl/worksheets/sheet1.xml"] = ET.tostring(
        root, encoding="utf-8", xml_declaration=True,
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in members.items():
            target.writestr(name, payload)


def _snapshot(path: Path, side: str):
    return sm._stream_selected_sheet_snapshot(
        str(path), str(path), _SHEET, side,
    )


def _key(snapshot, physical_row: int | None):
    if physical_row is None:
        return None
    return snapshot.rows[int(physical_row) - 1].cells[1].cached_value


def _assert_cached_formula_keys_are_authoritative(root: Path) -> None:
    mine_path = root / "mine.xlsx"
    base_path = root / "base.xlsx"
    theirs_path = root / "theirs.xlsx"
    _write_book(
        mine_path, first_key=13002, first_label="mine-only",
        cache_formulas=True,
    )
    _write_book(
        base_path, first_key=13003, first_label="base-and-theirs",
        cache_formulas=True,
    )
    _write_book(
        theirs_path, first_key=13003, first_label="base-and-theirs",
        cache_formulas=True,
    )

    mine, base, theirs = (
        _snapshot(mine_path, "A"),
        _snapshot(base_path, "BASE"),
        _snapshot(theirs_path, "B"),
    )
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert not result.unresolved

    mismatched = [
        (mine_row, theirs_row)
        for mine_row, theirs_row in result.row_pairs
        if mine_row is not None
        and theirs_row is not None
        and _key(mine, mine_row) not in (None, "")
        and _key(theirs, theirs_row) not in (None, "")
        and _key(mine, mine_row) != _key(theirs, theirs_row)
    ]
    assert not mismatched, mismatched

    mine_only = next(
        index for index, (mine_row, theirs_row) in enumerate(result.row_pairs)
        if _key(mine, mine_row) == 13002 and theirs_row is None
    )
    theirs_owner = next(
        index for index, (mine_row, theirs_row) in enumerate(result.row_pairs)
        if mine_row is None and _key(theirs, theirs_row) == 13003
    )
    common = next(
        index for index, (mine_row, theirs_row) in enumerate(result.row_pairs)
        if _key(mine, mine_row) == 20000 and _key(theirs, theirs_row) == 20000
    )
    assert result.base_rows_by_pair[mine_only] is None
    assert _key(base, result.base_rows_by_pair[theirs_owner]) == 13003
    assert _key(base, result.base_rows_by_pair[common]) == 20000
    assert result.pair_diff_cols[mine_only] == frozenset((-1,))
    assert result.pair_diff_cols[theirs_owner] == frozenset((-1,))


def _assert_missing_formula_cache_fails_closed(root: Path) -> None:
    mine_path = root / "missing-cache-mine.xlsx"
    theirs_path = root / "missing-cache-theirs.xlsx"
    _write_book(
        mine_path, first_key=30001, first_label="mine",
        cache_formulas=False,
    )
    _write_book(
        theirs_path, first_key=30001, first_label="theirs",
        cache_formulas=False,
    )
    mine = _snapshot(mine_path, "A")
    theirs = _snapshot(theirs_path, "B")
    assert sm._snapshot_declared_records(mine) is None
    assert sm._snapshot_declared_records(theirs) is None
    assert sm._compare_selected_sheet_snapshots(mine, theirs).unresolved


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="sow_condition_formula_key_") as raw:
        root = Path(raw)
        _assert_cached_formula_keys_are_authoritative(root)
        _assert_missing_formula_cache_fails_closed(root)
    print("SMOKE_CONDITION_FORMULA_KEY_ALIGNMENT_OK")


if __name__ == "__main__":
    main()
