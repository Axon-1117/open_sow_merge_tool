"""Regression for Language-style undeclared payload columns and clear states."""

from __future__ import annotations

import hashlib

import sow_merge_tool as sm


SHEET = "default@design@na_TLanguageCn"
WIDTH = 12


def _cell(value=None, formula=None):
    formula_value = value if formula is None else formula
    formula_type = "f" if isinstance(formula_value, str) and formula_value.startswith("=") else (
        "b" if isinstance(formula_value, bool) else
        "n" if isinstance(formula_value, (int, float)) or formula_value is None else
        "s"
    )
    cached_type = (
        "b" if isinstance(value, bool) else
        "n" if isinstance(value, (int, float)) or value is None else
        "s"
    )
    return sm.SnapshotCell(
        value,
        cached_type,
        formula_value,
        formula_type,
        "formula" if formula_type == "f" else "literal",
        False,
    )


def _row(number, values, formulas=None):
    formulas = formulas or {}
    cells = tuple(
        _cell(value, formulas.get(index, value))
        for index, value in enumerate(values, start=1)
    )
    return sm.SnapshotRow(number, cells, sm._snapshot_row_hash(cells))


def _snapshot(side, records, *, tail_override=None):
    declarations = (
        "ID@const", "txt_ll", "version", "tx@pm", "display@pm", "note@pm",
        "", "", "", "", "", "",
    )
    types = ("string",) * 6 + ("",) * 6
    rows = [
        _row(1, declarations),
        _row(2, types),
    ]
    tail_override = tail_override or {}
    for offset, (key, text) in enumerate(records, start=3):
        tail = [None, None, None, None, None, None]
        if key == "K2":
            tail = [None, 5, True, True, None, True]
        for col, value in tail_override.get(key, {}).items():
            tail[int(col) - 7] = value
        values = [key, text, "v1", "pm", "display", "note", *tail]
        formulas = {9: "=IF(A%s=\"K2\",TRUE,FALSE)" % offset} if key == "K2" else {}
        rows.append(_row(offset, values, formulas))
    fields = tuple(
        sm.SnapshotField(
            col,
            declarations[col - 1],
            types[col - 1],
            frozenset(("const",)) if col == 1 else frozenset(),
        )
        for col in range(1, WIDTH + 1)
    )
    digest = hashlib.sha256(repr((side, records)).encode()).digest()
    version = sm.SheetSnapshotVersion(
        parser=1,
        topology_generation=0,
        mutation_generation=0,
        file_size=len(digest),
        file_mtime_ns=1,
    )
    return sm.SheetSnapshot(
        side,
        SHEET,
        version,
        len(rows),
        WIDTH,
        fields,
        tuple(rows),
    )


def _fixtures(*, mine_tail=None, theirs_tail=None):
    base = _snapshot("base", (("K1", "base"), ("K2", "same"), ("K3", "same")))
    mine = _snapshot(
        "mine",
        (("K1", "mine"), ("K2", "same"), ("K3", "same"), ("K4", "mine-add")),
        tail_override=mine_tail,
    )
    theirs = _snapshot(
        "theirs",
        (("K1", "theirs"), ("K2", "same"), ("K3", "same"),
         ("K5", "theirs-add-1"), ("K6", "theirs-add-2")),
        tail_override=theirs_tail,
    )
    return mine, base, theirs


def _without_duplicate_tail(snapshot, side):
    rows = []
    for row in snapshot.rows:
        cells = tuple(row.cells[:6])
        rows.append(
            sm.SnapshotRow(
                row.physical_row,
                cells,
                sm._snapshot_row_hash(cells),
            )
        )
    return sm.SheetSnapshot(
        side,
        snapshot.sheet,
        snapshot.version,
        snapshot.max_row,
        6,
        tuple(snapshot.fields[:6]),
        tuple(rows),
    )


def _with_cleared_columns(snapshot, side, columns):
    cleared = {int(column) for column in columns}
    rows = []
    for row in snapshot.rows:
        cells = tuple(
            _cell() if index in cleared else cell
            for index, cell in enumerate(row.cells, start=1)
        )
        rows.append(
            sm.SnapshotRow(
                row.physical_row,
                cells,
                sm._snapshot_row_hash(cells),
            )
        )
    fields = tuple(
        sm.SnapshotField(field.physical_col, "", "", frozenset())
        if field.physical_col in cleared else field
        for field in snapshot.fields
    )
    return sm.SheetSnapshot(
        side,
        snapshot.sheet,
        snapshot.version,
        snapshot.max_row,
        snapshot.max_col,
        fields,
        tuple(rows),
    )


def _empty_snapshot(snapshot, side, width=1):
    cells = tuple(_cell() for _ in range(width))
    return sm.SheetSnapshot(
        side,
        snapshot.sheet,
        snapshot.version,
        1,
        width,
        tuple(
            sm.SnapshotField(column, "", "", frozenset())
            for column in range(1, width + 1)
        ),
        (sm.SnapshotRow(1, cells, sm._snapshot_row_hash(cells)),),
    )


def test_language_keyed_duplicate_tail_resolves_with_blank_independent_additions():
    mine, base, theirs = _fixtures()
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert not result.unresolved
    assert len(result.row_pairs) == 8
    assert result.column_cache.model.confidence.reason == (
        "snapshot-duplicate-base-anchored-payload-proof"
    )
    assert not result.column_cache.unresolved_cols
    assert not result.column_cache.structural_diff_cols
    assert sum(bool(cols) for cols in result.conflict_cols) == 1
    assert sum(len(cols) for cols in result.conflict_cols) == 1

    prepared = sm._snapshot_result_to_sheet_cache_immutable(
        SHEET, result, mine, theirs, base, has_base=True
    )
    assert prepared["prepared_complete"]
    assert len(prepared["row_pairs"]) == 8
    assert len(prepared["mine_to_base_row"]) == 5
    assert len(prepared["theirs_to_base_row"]) == 5


def test_language_keyed_duplicate_tail_two_way_uses_same_proof():
    mine, base, _theirs = _fixtures()
    result = sm._compare_selected_sheet_snapshots(mine, base)
    assert not result.unresolved
    assert result.column_cache.model.confidence.reason == (
        "snapshot-duplicate-keyed-payload-proof"
    )


def test_language_keyed_duplicate_tail_rejects_nonblank_one_sided_payload():
    mine, base, theirs = _fixtures(mine_tail={"K4": {8: "unsafe-new-tail"}})
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert result.unresolved


def test_language_keyed_duplicate_tail_accepts_base_anchored_branch_edit():
    mine, base, theirs = _fixtures(theirs_tail={"K2": {10: False}})
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert not result.unresolved
    assert any(10 in cols for cols in result.pair_diff_cols)
    assert not any(10 in cols for cols in result.conflict_cols)


def test_language_complete_duplicate_tail_deleted_in_mine_is_resolved():
    full_mine, base, theirs = _fixtures()
    mine = _without_duplicate_tail(full_mine, "mine")
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert not result.unresolved
    assert result.column_cache.unresolved_cols == frozenset()
    assert result.column_cache.structural_diff_cols == frozenset(range(7, 13))
    tail_slots = tuple(result.column_cache.model.slots[6:])
    assert all(
        slot.mine_col is None
        and slot.base_col == slot.logical_idx + 1
        and slot.theirs_col == slot.logical_idx + 1
        and slot.state == "mine-deleted"
        and not slot.confidence.ambiguous
        for slot in tail_slots
    )


def test_language_complete_duplicate_tail_delete_versus_modify_stays_unresolved():
    full_mine, base, theirs = _fixtures(
        theirs_tail={"K2": {10: False}}
    )
    mine = _without_duplicate_tail(full_mine, "mine")
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert result.unresolved


def test_language_two_way_whole_columns_cleared_is_resolved():
    _mine, base, _theirs = _fixtures()
    cleared = _with_cleared_columns(base, "theirs", (4, 5, 6, 9))
    result = sm._compare_selected_sheet_snapshots(base, cleared)
    assert not result.unresolved
    assert result.column_cache.unresolved_cols == frozenset()
    assert result.column_cache.structural_diff_cols == frozenset()
    assert result.column_cache.model.confidence.reason == (
        "snapshot-two-way-same-ordinal-clear-proof"
    )
    changed = set().union(*result.pair_diff_cols)
    assert {4, 5, 6, 9}.issubset(changed)


def test_language_two_way_partial_nonblank_replacement_stays_unresolved():
    _mine, base, _theirs = _fixtures()
    changed = _with_cleared_columns(base, "theirs", (4, 5, 6, 9))
    rows = list(changed.rows)
    cells = list(rows[2].cells)
    cells[3] = _cell("replacement")
    rows[2] = sm.SnapshotRow(3, tuple(cells), sm._snapshot_row_hash(tuple(cells)))
    changed = sm.SheetSnapshot(
        changed.side,
        changed.sheet,
        changed.version,
        changed.max_row,
        changed.max_col,
        changed.fields,
        tuple(rows),
    )
    result = sm._compare_selected_sheet_snapshots(base, changed)
    assert result.unresolved


def test_two_way_whole_sheet_content_cleared_is_resolved_on_either_side():
    _mine, base, _theirs = _fixtures()
    empty = _empty_snapshot(base, "empty")
    for left, right in ((base, empty), (empty, base)):
        result = sm._compare_selected_sheet_snapshots(left, right)
        assert not result.unresolved
        assert result.column_cache.unresolved_cols == frozenset()
        assert result.column_cache.model.confidence.reason == (
            "snapshot-two-way-empty-side-proof"
        )
        assert len(result.row_pairs) == base.max_row
        assert any(result.pair_diff_cols)
        prepared = sm._snapshot_result_to_sheet_cache_immutable(
            SHEET, result, left, right, None
        )
        assert prepared["prepared_complete"]
        assert prepared["has_diff"]


if __name__ == "__main__":
    tests = (
        test_language_keyed_duplicate_tail_resolves_with_blank_independent_additions,
        test_language_keyed_duplicate_tail_two_way_uses_same_proof,
        test_language_keyed_duplicate_tail_rejects_nonblank_one_sided_payload,
        test_language_keyed_duplicate_tail_accepts_base_anchored_branch_edit,
        test_language_complete_duplicate_tail_deleted_in_mine_is_resolved,
        test_language_complete_duplicate_tail_delete_versus_modify_stays_unresolved,
        test_language_two_way_whole_columns_cleared_is_resolved,
        test_language_two_way_partial_nonblank_replacement_stays_unresolved,
        test_two_way_whole_sheet_content_cleared_is_resolved_on_either_side,
    )
    for test in tests:
        test()
        print("PASS", test.__name__)
