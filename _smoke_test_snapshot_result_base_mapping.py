"""Pure contract for result-owned immutable Base row mappings.

The adapter consumes an already exact ``SnapshotComparisonResult``.  These
fixtures deliberately contain no Workbook, Worksheet, GUI, or child process;
they prove that an exact top-level Mine/Base/Theirs result remains authoritative
when the independently aligned child gaps would be conservative.
"""

from __future__ import annotations

from dataclasses import replace
import hashlib

import sow_merge_tool as sm


def _cell(value, *, cached_type="n", formula_value=None):
    if formula_value is not None:
        return sm.SnapshotCell(
            value, cached_type, formula_value, "f", "formula", False,
        )
    return sm.SnapshotCell(
        value, cached_type, None, cached_type, "literal", False,
    )


_BLANK = _cell(None)


def _row(physical_row: int, cells) -> sm.SnapshotRow:
    payload = tuple(
        (
            cell.cached_value,
            cell.cached_type,
            cell.formula_value,
            cell.formula_type,
            cell.formula_kind,
            cell.external_link,
        )
        for cell in cells
    )
    digest = hashlib.sha256(repr(payload).encode("utf-8")).hexdigest()
    return sm.SnapshotRow(int(physical_row), tuple(cells), digest)


def _real_like_duplicate_snapshot(side: str) -> sm.SheetSnapshot:
    """Return the 3,169-row Monster-shaped duplicate-field snapshot.

    It preserves the real release topology important to this adapter contract:
    2 header rows, 2,964 typed owners, 203 blank-key continuations, and blank
    duplicate schema fields at physical 27/31.  The comparison remains purely
    immutable and deliberately has no worksheet backing object.
    """
    headers = []
    types = []
    fields = []
    for physical_col in range(1, 32):
        if physical_col == 1:
            declaration, type_declaration = "id@id", "int32"
        elif physical_col in (27, 31):
            declaration, type_declaration = "", ""
        else:
            declaration, type_declaration = f"field_{physical_col}", "string"
        headers.append(declaration if declaration else None)
        types.append(type_declaration if type_declaration else None)
        fields.append(sm.SnapshotField(
            physical_col=physical_col,
            declaration=declaration,
            type_declaration=type_declaration,
            markers=frozenset(("id",)) if physical_col == 1 else frozenset(),
        ))

    rows = [
        _row(1, tuple(_cell(value, cached_type="s") if value is not None else _BLANK for value in headers)),
        _row(2, tuple(_cell(value, cached_type="s") if value is not None else _BLANK for value in types)),
    ]
    physical_row = 3
    for owner in range(1, 2965):
        owner_cells = (_cell(owner, cached_type="n"),) + (_BLANK,) * 30
        rows.append(_row(physical_row, owner_cells))
        physical_row += 1
        if owner <= 203:
            rows.append(_row(physical_row, (_BLANK,) * 31))
            physical_row += 1
    assert physical_row == 3170 and len(rows) == 3169
    return sm.SheetSnapshot(
        side=str(side),
        sheet="MonsterGroup@design",
        version=sm.SheetSnapshotVersion(1, 0, 0, 1, 1),
        max_row=len(rows),
        max_col=31,
        fields=tuple(fields),
        rows=tuple(rows),
    )


def _mini_snapshot(side: str, row_count: int = 6) -> sm.SheetSnapshot:
    fields = (
        sm.SnapshotField(1, "id@id", "int32", frozenset(("id",))),
    )
    rows = []
    for row in range(1, row_count + 1):
        if row == 1:
            cell = _cell("id@id", cached_type="s")
        elif row == 2:
            cell = _cell("int32", cached_type="s")
        else:
            cell = _cell(row, cached_type="n")
        rows.append(_row(row, (cell,)))
    return sm.SheetSnapshot(
        side=str(side),
        sheet="BaseMapping",
        version=sm.SheetSnapshotVersion(1, 0, 0, 1, 1),
        max_row=row_count,
        max_col=1,
        fields=fields,
        rows=tuple(rows),
    )


def _manual_result(mine, pairs, base_rows, pair_diffs, base_diffs, conflicts=None):
    cache = sm._physical_identity_snapshot_comparison(mine, three_way=True).column_cache
    pair_count = len(pairs)
    if conflicts is None:
        conflicts = tuple(frozenset() for _ in range(pair_count))
    return sm.SnapshotComparisonResult(
        tuple(pairs),
        tuple(base_rows),
        cache,
        tuple(frozenset(values) for values in pair_diffs),
        tuple(frozenset(values) for values in base_diffs),
        tuple(frozenset(values) for values in conflicts),
        False,
    )


def _assert_unresolved_without_targets(result, mine, theirs, base):
    cache = sm._snapshot_result_to_sheet_cache_immutable(
        "BaseMapping", result, mine, theirs, base, has_base=True,
    )
    assert not cache["prepared_complete"]
    assert cache["unresolved_reason"]
    assert "mine_to_base_row" not in cache
    assert "pair_base_row_override" not in cache


def _test_real_like_exact_result_does_not_realign_child_gaps():
    mine = _real_like_duplicate_snapshot("mine")
    base = _real_like_duplicate_snapshot("base")
    theirs = _real_like_duplicate_snapshot("theirs")
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert not result.unresolved
    assert len(result.row_pairs) == len(result.base_rows_by_pair) == 3169
    assert not result.column_cache.unresolved_cols

    original_align = sm._align_selected_sheet_snapshots
    unexpected_calls = []

    def _forbid_child_realign(*args, **kwargs):
        unexpected_calls.append((args, kwargs))
        raise AssertionError("adapter must use result-owned Base row mappings")

    sm._align_selected_sheet_snapshots = _forbid_child_realign
    try:
        cache = sm._snapshot_result_to_sheet_cache_immutable(
            "MonsterGroup@design", result, mine, theirs, base, has_base=True,
        )
    finally:
        sm._align_selected_sheet_snapshots = original_align
    assert not unexpected_calls
    assert cache["prepared_complete"]
    assert len(cache["row_pairs"]) == 3169
    assert cache["mine_to_base_row"][3] == 3
    assert cache["theirs_to_base_row"][3169] == 3169
    assert cache["pair_base_row_override"][0] == 1
    assert cache["pair_base_row_override"][3168] == 3169


def _test_structural_none_semantics_and_reorder_remain_actionable():
    mine = _mini_snapshot("mine")
    theirs = _mini_snapshot("theirs")
    base = _mini_snapshot("base")
    pairs = ((1, 1), (2, 2), (3, 3), (4, None), (None, 4), (5, 5), (6, 6))
    base_rows = (1, 2, None, 3, 4, 5, 6)
    result = _manual_result(
        mine,
        pairs,
        base_rows,
        ((), (), (), (-1,), (-1,), (), ()),
        ((), (), (-1,), (), (-1,), (), ()),
    )
    mappings = sm._snapshot_result_base_row_mappings(result, mine, theirs, base)
    assert mappings is not None
    mine_to_base, theirs_to_base, overrides = mappings
    assert mine_to_base == {1: 1, 2: 2, 4: 3, 5: 5, 6: 6}
    assert theirs_to_base == {1: 1, 2: 2, 4: 4, 5: 5, 6: 6}
    assert overrides == {0: 1, 1: 2, 3: 3, 4: 4, 5: 5, 6: 6}
    cache = sm._snapshot_result_to_sheet_cache_immutable(
        "BaseMapping", result, mine, theirs, base, has_base=True,
    )
    assert cache["prepared_complete"]
    assert cache["mine_to_base_row"] == mine_to_base
    assert cache["theirs_to_base_row"] == theirs_to_base
    assert cache["pair_base_row_override"] == overrides

    reorder = _manual_result(
        mine,
        ((1, 1), (2, 2), (3, 4), (4, 3), (5, 5), (6, 6)),
        (1, 2, 3, 4, 5, 6),
        ((), (), (), (), (), ()),
        ((), (), (), (), (), ()),
    )
    reordered = sm._snapshot_result_base_row_mappings(reorder, mine, theirs, base)
    assert reordered is not None
    assert reordered[0][3] == 3
    assert reordered[1][4] == 3
    assert reordered[2] == {index: index + 1 for index in range(6)}


def _test_malformed_or_conflicting_result_mappings_fail_closed():
    mine = _mini_snapshot("mine")
    theirs = _mini_snapshot("theirs")
    base = _mini_snapshot("base")
    valid = _manual_result(
        mine,
        ((1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6)),
        (1, 2, 3, 4, 5, 6),
        ((), (), (), (), (), ()),
        ((), (), (), (), (), ()),
    )
    bad_results = (
        replace(valid, base_rows_by_pair=valid.base_rows_by_pair[:-1]),
        replace(valid, base_rows_by_pair=(1, 2, 3, 4, 5, 99)),
        replace(valid, base_rows_by_pair=(1, 1, 3, 4, 5, 6)),
        replace(valid, row_pairs=((1, 1), (1, 2), (3, 3), (4, 4), (5, 5), (6, 6))),
        replace(valid, row_pairs=((1, 1), (2, 1), (3, 3), (4, 4), (5, 5), (6, 6))),
        replace(valid, conflict_cols=valid.conflict_cols[:-1]),
        replace(valid, row_pairs=((1, 1), (None, None), (3, 3), (4, 4), (5, 5), (6, 6))),
        replace(valid, pair_diff_cols=(frozenset(),) * 6, row_pairs=((1, None), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6))),
        replace(valid, pair_base_diff_cols=(frozenset(),) * 6, base_rows_by_pair=(None, 2, 3, 4, 5, 6)),
        replace(
            valid,
            row_pairs=((1, 1), (2, None), (3, 3), (4, 4), (5, 5), (6, 6)),
            pair_diff_cols=(frozenset(), frozenset((-1,)), frozenset(), frozenset(), frozenset(), frozenset()),
        ),
        replace(
            valid,
            base_rows_by_pair=(1, None, 3, 4, 5, 6),
            pair_base_diff_cols=(frozenset(), frozenset((-1,)), frozenset(), frozenset(), frozenset(), frozenset()),
        ),
    )
    for malformed in bad_results:
        assert sm._snapshot_result_base_row_mappings(malformed, mine, theirs, base) is None
        _assert_unresolved_without_targets(malformed, mine, theirs, base)

    one_sided_extra = _manual_result(
        mine,
        ((1, 1), (2, 2), (3, None), (4, 4), (5, 5), (6, 6)),
        (1, 2, 3, 4, 5, 6),
        ((), (), (-1, 1), (), (), ()),
        ((), (), (), (), (), ()),
    )
    base_one_sided_extra = _manual_result(
        mine,
        ((1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6)),
        (1, 2, None, 4, 5, 6),
        ((), (), (), (), (), ()),
        ((), (), (-1, 1), (), (), ()),
    )
    conflict_with_missing_side = _manual_result(
        mine,
        ((1, 1), (2, 2), (3, None), (4, 4), (5, 5), (6, 6)),
        (1, 2, 3, 4, 5, 6),
        ((), (), (-1,), (), (), ()),
        ((), (), (), (), (), ()),
        ((), (), (1,), (), (), ()),
    )
    for malformed in (one_sided_extra, base_one_sided_extra, conflict_with_missing_side):
        assert sm._snapshot_result_base_row_mappings(malformed, mine, theirs, base) is None
        _assert_unresolved_without_targets(malformed, mine, theirs, base)

    two_way = _manual_result(
        mine,
        ((1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6)),
        (None, None, None, None, None, None),
        ((), (), (), (), (), ()),
        ((), (), (), (), (), ()),
    )
    assert sm._snapshot_result_base_row_mappings(two_way, mine, theirs, None) == ({}, {}, {})
    assert sm._snapshot_result_base_row_mappings(
        replace(two_way, pair_base_diff_cols=(frozenset(), frozenset((1,)), frozenset(), frozenset(), frozenset(), frozenset())),
        mine, theirs, None,
    ) is None
    assert sm._snapshot_result_base_row_mappings(
        replace(two_way, conflict_cols=(frozenset(), frozenset((1,)), frozenset(), frozenset(), frozenset(), frozenset())),
        mine, theirs, None,
    ) is None


def main():
    _test_real_like_exact_result_does_not_realign_child_gaps()
    _test_structural_none_semantics_and_reorder_remain_actionable()
    _test_malformed_or_conflicting_result_mappings_fail_closed()
    print("snapshot result Base mapping contract: PASS")


if __name__ == "__main__":
    main()
