from __future__ import annotations

import os
import tempfile
from datetime import date

from openpyxl import Workbook

import sow_merge_tool as smt


def _key(name: str) -> smt.ColumnModelCacheKey:
    return smt.ColumnModelCacheKey(name, 1, 1)


def test_two_way_insert_does_not_cascade() -> None:
    mine = [
        ("A", "B", "C", "D"),
        ("a1", "b1", "c1", "d1"),
        ("a2", "b2", "c2", "d2"),
    ]
    theirs = [
        ("A", "X", "Y", "B", "C", "D"),
        ("a1", "x1", "y1", "b1", "c1", "d1"),
        ("a2", "x2", "y2", "b2", "c2", "d2"),
    ]
    cache = smt.build_logical_column_comparison_cache_2way(
        _key("insert"), mine, theirs, mine, theirs, mine_max_col=4, theirs_max_col=6
    )
    result = smt.compare_logical_row_2way(
        cache, mine[1], theirs[1], mine[1], theirs[1], mine_row=2, theirs_row=2
    )
    inserted = {
        slot.logical_idx + 1
        for slot in cache.model.slots
        if slot.mine_col is None and slot.theirs_col is not None
    }
    assert inserted == result.diff_cols, (cache.model.slots, result)
    assert len(result.diff_cols) == 2


def test_same_order_value_edit_is_actionable_but_ambiguity_stays_blocked() -> None:
    mine = [("x", "y"), (1, 2)]
    theirs = [("x", "Y"), (1, 2)]
    cache = smt.build_logical_column_comparison_cache_2way(
        _key("same-order-edit"),
        mine,
        theirs,
        mine,
        theirs,
        mine_max_col=2,
        theirs_max_col=2,
    )
    assert cache.unresolved_cols == frozenset(), cache.model.slots
    assert [slot.state for slot in cache.model.slots] == ["retained", "retained"]
    assert cache.model.slots[1].confidence.reason == "same-ordinal-content-evidence"
    row = smt.compare_logical_row_2way(
        cache,
        mine[0],
        theirs[0],
        mine[0],
        theirs[0],
        mine_row=1,
        theirs_row=1,
    )
    assert row.diff_cols == frozenset((2,)), row

    # Entirely replaced same-ordinal content has no stable evidence and remains
    # unresolved; a physical-order fallback must not become an editable mapping.
    replaced_mine = [("A", "Business-B", "Z"), (1, "old-1", 9), (2, "old-2", 10)]
    replaced_theirs = [("A", "Different-Q", "Z"), (1, "new-x", 9), (2, "new-y", 10)]
    replaced = smt.build_logical_column_comparison_cache_2way(
        _key("same-order-replaced"),
        replaced_mine,
        replaced_theirs,
        replaced_mine,
        replaced_theirs,
        mine_max_col=3,
        theirs_max_col=3,
    )
    assert 2 in replaced.unresolved_cols, replaced.model.slots

    # Indistinguishable blank/duplicate columns retain their conservative gate
    # even though equal cells exist at the same physical ordinals.
    ambiguous = [("A", None, None, "Z"), (1, None, None, 9), (2, "", "", 10)]
    duplicate_cache = smt.build_logical_column_comparison_cache_2way(
        _key("same-order-duplicates"),
        ambiguous,
        ambiguous,
        ambiguous,
        ambiguous,
        mine_max_col=4,
        theirs_max_col=4,
    )
    assert {2, 3}.issubset(duplicate_cache.unresolved_cols), duplicate_cache.model.slots


def test_formula_translation_uses_logical_slot() -> None:
    model = smt.ColumnModel.from_slots(
        _key("formula"),
        (
            smt.ColumnSlot(0, mine_col=1, theirs_col=1),
            smt.ColumnSlot(1, mine_col=None, theirs_col=2, state="inserted"),
            smt.ColumnSlot(2, mine_col=2, theirs_col=3),
            smt.ColumnSlot(3, mine_col=3, theirs_col=4),
        ),
    )
    cache = smt.LogicalColumnComparisonCache(model=model, structural_diff_cols=frozenset((2,)))
    result = smt.compare_logical_row_2way(
        cache,
        (1, 2, None),
        (1, "new", 2, None),
        (1, 2, "=B1*2"),
        (1, "new", 2, "=C1*2"),
        mine_row=1,
        theirs_row=1,
    )
    assert result.diff_cols == frozenset((2,)), result

    # A reference before the insertion boundary does not move when Excel
    # inserts the neighbouring column.  Normalization must therefore map
    # referenced columns through the logical model instead of translating the
    # whole formula solely from the formula cell's physical position.
    before_boundary = smt.compare_logical_row_2way(
        cache,
        (1, 2, None),
        (1, "new", 2, None),
        (1, 2, "=A1*2"),
        (1, "new", 2, "=A1*2"),
        mine_row=1,
        theirs_row=1,
    )
    assert before_boundary.diff_cols == frozenset((2,)), before_boundary


def test_formula_reference_canonicalization_boundaries() -> None:
    model = smt.ColumnModel.from_slots(
        _key("formula-reference-boundaries"),
        (
            smt.ColumnSlot(0, mine_col=1, theirs_col=1),
            smt.ColumnSlot(1, mine_col=None, theirs_col=2, state="inserted"),
            smt.ColumnSlot(2, mine_col=2, theirs_col=3),
            smt.ColumnSlot(3, mine_col=3, theirs_col=4),
        ),
    )
    cache = smt.LogicalColumnComparisonCache(
        model=model,
        structural_diff_cols=frozenset((2,)),
    )
    mine_map = {1: 1, 2: 3, 3: 4}
    theirs_map = {1: 1, 2: 2, 3: 3, 4: 4}

    helper_cases = (
        ("=$B$1", "=$C$1"),
        ("=SUM(B1:C2)", "=SUM(C1:D2)"),
        ("=SUM(B:B)", "=SUM(C:C)"),
    )
    for mine_formula, expected in helper_cases:
        assert smt._canonicalize_formula_column_references(
            mine_formula, mine_map
        ) == expected
        assert smt._canonicalize_formula_column_references(
            expected, theirs_map
        ) == expected

        compared = smt.compare_logical_row_2way(
            cache,
            (1, 2, None),
            (1, "new", 2, None),
            (1, 2, mine_formula),
            (1, "new", 2, expected),
            mine_row=1,
            theirs_row=1,
        )
        assert compared.diff_cols == frozenset((2,)), (
            mine_formula,
            expected,
            compared,
        )

    conservative = (
        "=Sheet2!B1",
        "=SUM(Table1[Amount])",
        "=NamedRange+1",
    )
    for formula in conservative:
        assert smt._canonicalize_formula_column_references(
            formula, mine_map
        ) == formula
        compared = smt.compare_logical_row_2way(
            cache,
            (1, 2, None),
            (1, "new", 2, None),
            (1, 2, formula),
            (1, "new", 2, formula),
            mine_row=1,
            theirs_row=1,
        )
        assert compared.diff_cols == frozenset((2,)), (formula, compared)

    # Unprovable qualified references that genuinely differ stay visible.
    qualified_difference = smt.compare_logical_row_2way(
        cache,
        (1, 2, None),
        (1, "new", 2, None),
        (1, 2, "=Sheet2!B1"),
        (1, "new", 2, "=Sheet2!C1"),
        mine_row=1,
        theirs_row=1,
    )
    assert qualified_difference.diff_cols == frozenset((2, 4)), qualified_difference


def test_formula_normalization_and_mapping_caches_preserve_semantics() -> None:
    normal_cache = smt._normalize_normal_formula_text_cached
    canonical_cache = smt._canonicalize_normal_formula_cached
    normal_cache.cache_clear()
    canonical_cache.cache_clear()

    # Normalization stays case-insensitive outside string literals while
    # retaining strings, escaped quotes, and intersection whitespace exactly.
    formula = '=sum(A1:B2,"A""b") C:C'
    assert smt._norm_formula_text(formula) == 'SUM(A1:B2,"A""b") C:C'
    assert smt._norm_formula_text(formula) == 'SUM(A1:B2,"A""b") C:C'
    assert normal_cache.cache_info().hits >= 1
    assert normal_cache.cache_info().maxsize == 8192
    assert canonical_cache.cache_info().maxsize == 16384

    identity = {1: 1, 2: 2, 3: 3}
    inserted = {1: 1, 2: 3, 3: 4}
    original_tokenizer = smt.Tokenizer
    try:
        smt.Tokenizer = lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("identity mapping must not tokenize")
        )
        assert smt._canonicalize_formula_column_references("=B1", identity) == "=B1"
    finally:
        smt.Tokenizer = original_tokenizer

    # The immutable mapping is part of the cache key. The same raw formula has
    # distinct logical meaning after an inserted physical column.
    assert smt._canonicalize_formula_column_references("=B1", inserted) == "=C1"
    assert smt._canonicalize_formula_column_references("=B1", identity) == "=B1"
    for conservative in (
        "=Sheet2!B1",
        "=[Book.xlsx]Sheet2!B1",
        "=SUM(Table1[B])",
        "=NamedRange+1",
    ):
        assert smt._canonicalize_formula_column_references(
            conservative, inserted
        ) == conservative

    array_formula = smt.ArrayFormula("A1:A2", "=ROW(A1:A2)")
    data_table = smt.DataTableFormula("B1:C3", r1="A1")
    assert smt._canonicalize_formula_column_references(array_formula, inserted) is array_formula
    assert smt._canonicalize_formula_column_references(data_table, inserted) is data_table


def test_same_formula_cache_and_inserted_column_semantics_stay_visible() -> None:
    identity_model = smt.ColumnModel.from_slots(
        _key("same-formula-cache"),
        (smt.ColumnSlot(0, mine_col=1, theirs_col=1),),
    )
    identity_cache = smt.LogicalColumnComparisonCache(model=identity_model)
    cache_difference = smt.compare_logical_row_2way(
        identity_cache,
        (10,),
        (20,),
        ("=A1",),
        ("=A1",),
        mine_row=1,
        theirs_row=1,
    )
    assert cache_difference.diff_cols == frozenset((1,)), cache_difference

    inserted_model = smt.ColumnModel.from_slots(
        _key("same-raw-formula-different-logical-reference"),
        (
            smt.ColumnSlot(0, mine_col=1, theirs_col=1),
            smt.ColumnSlot(1, mine_col=None, theirs_col=2, state="inserted"),
            smt.ColumnSlot(2, mine_col=2, theirs_col=3),
            smt.ColumnSlot(3, mine_col=3, theirs_col=4),
        ),
    )
    inserted_cache = smt.LogicalColumnComparisonCache(
        model=inserted_model,
        structural_diff_cols=frozenset((2,)),
    )
    logical_difference = smt.compare_logical_row_2way(
        inserted_cache,
        (1, 2, None),
        (1, "new", 2, None),
        (1, 2, "=B1"),
        (1, "new", 2, "=B1"),
        mine_row=1,
        theirs_row=1,
    )
    assert logical_difference.diff_cols == frozenset((2, 4)), logical_difference


def test_display_only_helper_matches_legacy_self_comparison() -> None:
    cases = (
        (None, None),
        (0, 0),
        ("#DIV/0!", "#DIV/0!"),
        (date(2026, 7, 23), date(2026, 7, 23)),
        (42, "=A1"),
        (None, "=A1"),
        ("=A1", "=A1"),
        (7, smt.ArrayFormula("A1:A2", "=ROW(A1:A2)")),
        (None, smt.ArrayFormula("A1:A2", "=ROW(A1:A2)")),
        (8, smt.DataTableFormula("B1:C3", r1="A1")),
        (None, smt.DataTableFormula("B1:C3", r1="A1")),
    )
    for value, edit in cases:
        expected, _unused, equal = smt._cell_display_and_equal_from_values(
            value, value, edit, edit
        )
        assert equal
        assert smt._cell_display_from_values(value, edit) == expected, (value, edit)


def test_three_way_delete_modify_and_competing_insert() -> None:
    base = [
        ("A", "B", "C"),
        ("a1", "b1", "c1"),
        ("a2", "b2", "c2"),
    ]
    mine = [
        ("A", "C"),
        ("a1", "c1"),
        ("a2", "c2"),
    ]
    theirs = [
        ("A", "B", "C"),
        ("a1", "b-changed", "c1"),
        ("a2", "b2", "c2"),
    ]
    cache = smt.build_logical_column_comparison_cache_3way(
        _key("delete-modify"),
        mine,
        base,
        theirs,
        mine,
        base,
        theirs,
        mine_max_col=2,
        base_max_col=3,
        theirs_max_col=3,
    )
    mine_changed = set()
    theirs_changed = set()
    for row_idx in range(len(base)):
        row_result = smt.compare_logical_row_3way(
            cache,
            mine[row_idx],
            base[row_idx],
            theirs[row_idx],
            mine[row_idx],
            base[row_idx],
            theirs[row_idx],
            mine_row=row_idx + 1,
            base_row=row_idx + 1,
            theirs_row=row_idx + 1,
        )
        mine_changed.update(row_result.mine_changed_cols)
        theirs_changed.update(row_result.theirs_changed_cols)
    analysis = smt.classify_logical_columns_3way(
        cache,
        mine_changed_cols=mine_changed,
        theirs_changed_cols=theirs_changed,
    )
    assert any(
        conflict.kind == "delete-versus-modify"
        for conflict in analysis.structural_conflicts
    ), analysis

    mine_insert = [("A", "Mine New", "B"), ("a", "m", "b")]
    theirs_insert = [("A", "Theirs New", "B"), ("a", "t", "b")]
    base_insert = [("A", "B"), ("a", "b")]
    competing_cache = smt.build_logical_column_comparison_cache_3way(
        _key("competing"),
        mine_insert,
        base_insert,
        theirs_insert,
        mine_insert,
        base_insert,
        theirs_insert,
        mine_max_col=3,
        base_max_col=2,
        theirs_max_col=3,
    )
    competing = smt.classify_logical_columns_3way(competing_cache)
    assert any(
        conflict.kind == "competing-insertion"
        for conflict in competing.structural_conflicts
    ), competing


def _save_book(path: str, rows) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in rows:
        ws.append(list(row))
    wb.save(path)
    wb.close()


def test_end_to_end_three_way_scanner() -> None:
    with tempfile.TemporaryDirectory(prefix="smt-logical-column-") as temp_dir:
        base_path = os.path.join(temp_dir, "base.xlsx")
        mine_path = os.path.join(temp_dir, "mine.xlsx")
        theirs_path = os.path.join(temp_dir, "theirs.xlsx")
        base = [
            ("A", "B", "C"),
            ("a1", "b1", "c1"),
            ("a2", "b2", "c2"),
            ("a3", "b3", "c3"),
        ]
        mine = [(row[0], row[2]) for row in base]
        theirs = list(base)
        theirs[2] = ("a2", "b2-modified", "c2")
        _save_book(base_path, base)
        _save_book(mine_path, mine)
        _save_book(theirs_path, theirs)
        conflicts, conflict_map = smt._scan_three_way_conflicts(
            base_path, mine_path, theirs_path
        )
        analysis = smt._LAST_THREE_WAY_COLUMN_ANALYSIS["Data"]
        assert conflicts and conflict_map.get("Data"), (conflicts, conflict_map)
        assert any(
            conflict.kind == "delete-versus-modify"
            for conflict in analysis.structural_conflicts
        ), analysis


def test_row_and_column_structure_do_not_cross_align() -> None:
    with tempfile.TemporaryDirectory(prefix="smt-row-column-") as temp_dir:
        base_path = os.path.join(temp_dir, "base.xlsx")
        mine_path = os.path.join(temp_dir, "mine.xlsx")
        theirs_path = os.path.join(temp_dir, "theirs.xlsx")
        base = [("A", "B", "C", "D", "E")] + [
            (f"id-{i}", f"b-{i}", f"c-{i}", f"d-{i}", f"e-{i}")
            for i in range(1, 13)
        ]
        mine = [("A", "X", "Y", "B", "D", "E")] + [
            (f"id-{i}", f"x-{i}", f"y-{i}", f"b-{i}", f"d-{i}", f"e-{i}")
            for i in range(1, 13)
        ]
        mine.insert(5, ("id-new", "x-new", "y-new", "b-new", "d-new", "e-new"))
        theirs = [tuple(row) for row in base]
        theirs[8] = (*theirs[8][:-1], "e-8-edited")
        _save_book(base_path, base)
        _save_book(mine_path, mine)
        _save_book(theirs_path, theirs)
        conflicts, conflict_map = smt._scan_three_way_conflicts(
            base_path, mine_path, theirs_path
        )
        assert conflicts == [] and conflict_map == {}, (conflicts, conflict_map)
        analysis = smt._LAST_THREE_WAY_COLUMN_ANALYSIS["Data"]
        assert [state.state for state in analysis.states] == [
            "retained",
            "inserted",
            "inserted",
            "retained",
            "mine-deleted",
            "retained",
            "modified",
        ], analysis


if __name__ == "__main__":
    test_two_way_insert_does_not_cascade()
    test_same_order_value_edit_is_actionable_but_ambiguity_stays_blocked()
    test_formula_translation_uses_logical_slot()
    test_formula_reference_canonicalization_boundaries()
    test_formula_normalization_and_mapping_caches_preserve_semantics()
    test_same_formula_cache_and_inserted_column_semantics_stay_visible()
    test_display_only_helper_matches_legacy_self_comparison()
    test_three_way_delete_modify_and_competing_insert()
    test_end_to_end_three_way_scanner()
    test_row_and_column_structure_do_not_cross_align()
    print("SMOKE_TEST_LOGICAL_COLUMN_COMPARISON_OK")
