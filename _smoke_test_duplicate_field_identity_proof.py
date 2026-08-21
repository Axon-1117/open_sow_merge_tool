"""Pure immutable-snapshot contract for Decision 14 duplicate fields.

The fixtures never create a Worksheet.  They model only the immutable snapshot
payload the selected-sheet comparator receives, so an accepted proof cannot
accidentally depend on editable workbooks or a legacy worksheet fallback.
"""

from __future__ import annotations

import hashlib

import sow_merge_tool as sm


_LAYOUTS = {
    # One repeated blank identity is an interior run; the other is a tail run
    # bounded by END.  They intentionally share the same blank declaration.
    "interior_end": (
        ("id@id", "left", "", "", "right", "", ""),
        ("string", "string", "", "", "string", "", ""),
    ),
    # START is equally valid, but only as a virtual outer boundary.
    "start": (
        ("", "", "id@id", "left", "right"),
        ("", "", "string", "string", "string"),
    ),
    "missing_member": (
        ("id@id", "left", "", "", "right", ""),
        ("string", "string", "", "", "string", ""),
    ),
    # Same duplicate count, but the interior/tail run partition differs.
    "run_mismatch": (
        ("id@id", "left", "", "right", "", "", ""),
        ("string", "string", "", "string", "", "", ""),
    ),
    "anchor_reordered": (
        ("id@id", "right", "", "", "left", "", ""),
        ("string", "string", "", "", "string", "", ""),
    ),
    # Base moves the complete interior duplicate segment without changing its
    # local left/right anchors or width.  The proof must retain the distinct
    # Base physical coordinates instead of guessing Mine's coordinates.
    "base_offset_mine": (
        ("id@id", "prefix", "left", "", "", "right", "suffix", "", ""),
        ("string", "string", "string", "", "", "string", "string", "", ""),
    ),
    "base_offset_base": (
        ("id@id", "left", "", "", "right", "prefix", "suffix", "", ""),
        ("string", "string", "", "", "string", "string", "string", "", ""),
    ),
    # A formula-bearing unique field is separated from the duplicate runs by
    # the stable anchor at physical col3.
    "formula_outside": (
        ("id@id", "formula-anchor", "stable", "left", "", "", "right", "", ""),
        ("string", "string", "string", "string", "", "", "string", "", ""),
    ),
}


def _cell(value, *, cached_type=None, formula_value=None):
    if formula_value is not None:
        return sm.SnapshotCell(
            value,
            str(cached_type or "n"),
            formula_value,
            "f",
            "formula",
            False,
        )
    if value is None:
        return sm.SnapshotCell(
            None, str(cached_type or "n"), None, "n", "literal", False
        )
    cell_type = str(cached_type or "s")
    return sm.SnapshotCell(value, cell_type, value, cell_type, "literal", False)


def _row(physical_row, values, *, type_overrides, formula_overrides):
    cells = tuple(
        _cell(
            value,
            cached_type=type_overrides.get((physical_row, column)),
            formula_value=formula_overrides.get((physical_row, column)),
        )
        for column, value in enumerate(values, start=1)
    )
    digest = hashlib.sha256(
        repr(tuple(
            (
                cell.cached_value,
                cell.cached_type,
                cell.formula_value,
                cell.formula_type,
                cell.formula_kind,
                cell.external_link,
            )
            for cell in cells
        )).encode("utf-8")
    ).hexdigest()
    return sm.SnapshotRow(physical_row, cells, digest)


def _snapshot(
    side,
    *,
    layout="interior_end",
    ids=("id-01", "id-02", "id-03"),
    value_overrides=None,
    type_overrides=None,
    formula_overrides=None,
):
    headers, types = _LAYOUTS[layout]
    value_overrides = dict(value_overrides or {})
    type_overrides = dict(type_overrides or {})
    formula_overrides = dict(formula_overrides or {})
    # The immutable stream represents a physically empty header/type cell as
    # ``None`` while retaining ``""`` in SnapshotField schema text.  Model
    # that distinction so this fixture proves a real blank column rather than
    # a literal empty-string payload.
    header_cells = tuple(value if value != "" else None for value in headers)
    type_cells = tuple(value if value != "" else None for value in types)
    rows = [_row(
        1, header_cells, type_overrides=type_overrides,
        formula_overrides=formula_overrides,
    )]
    rows.append(_row(
        2, type_cells, type_overrides=type_overrides,
        formula_overrides=formula_overrides,
    ))
    for physical_row, record_id in enumerate(ids, start=3):
        values = []
        for header in headers:
            if header == "id@id":
                values.append(record_id)
            elif header == "left":
                values.append("left-value")
            elif header == "right":
                values.append("right-value")
            else:
                values.append(None)
        for column, value in enumerate(values, start=1):
            values[column - 1] = value_overrides.get((physical_row, column), value)
        rows.append(_row(
            physical_row, values, type_overrides=type_overrides,
            formula_overrides=formula_overrides,
        ))
    fields = tuple(
        sm.SnapshotField(
            physical_col=column,
            declaration=header,
            type_declaration=types[column - 1],
            markers=frozenset(("id",)) if "@id" in header else frozenset(),
        )
        for column, header in enumerate(headers, start=1)
    )
    return sm.SheetSnapshot(
        side=str(side),
        sheet="DuplicateFields",
        version=sm.SheetSnapshotVersion(1, 0, 0, 1, 1),
        max_row=len(rows),
        max_col=len(headers),
        fields=fields,
        rows=tuple(rows),
    )


def _assert_exact(result, mine, theirs, base=None):
    assert not result.unresolved
    assert not result.column_cache.unresolved_cols
    slots = tuple(result.column_cache.model.slots)
    expected = set(range(1, mine.max_col + 1))
    assert {slot.mine_col for slot in slots} == expected
    assert {slot.theirs_col for slot in slots} == expected
    if base is not None:
        assert {slot.base_col for slot in slots} == expected
        assert all(row is not None for row in result.base_rows_by_pair)


def _assert_proof_pairs_share_one_logical_slot(result, proof):
    for mine_col, theirs_col, base_col in proof.pairs:
        matches = [
            slot for slot in result.column_cache.model.slots
            if slot.mine_col == mine_col
            and slot.theirs_col == theirs_col
            and slot.base_col == base_col
        ]
        assert len(matches) == 1, (mine_col, theirs_col, base_col, matches)


def _assert_unresolved(mine, theirs, base=None):
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert result.unresolved
    immutable_cache = sm._snapshot_result_to_sheet_cache_immutable(
        "DuplicateFields", result, mine, theirs, base, has_base=base is not None
    )
    assert immutable_cache["unresolved_reason"]
    assert not immutable_cache["prepared_complete"]
    # The unresolved adapter intentionally does not carry an operation mapping.
    assert "row_a_to_pair_idx" not in immutable_cache


def _test_two_way_blank_interior_end_and_start_proofs():
    mine = _snapshot("mine")
    theirs = _snapshot("theirs")
    pending = sm._align_selected_sheet_snapshots(mine, theirs)
    assert pending.unresolved and pending.duplicate_field_proof is not None
    result = sm._compare_selected_sheet_snapshots(mine, theirs)
    _assert_exact(result, mine, theirs)
    _assert_proof_pairs_share_one_logical_slot(
        result, pending.duplicate_field_proof
    )

    start_mine = _snapshot("mine", layout="start")
    start_theirs = _snapshot("theirs", layout="start")
    start_pending = sm._align_selected_sheet_snapshots(start_mine, start_theirs)
    start_result = sm._compare_selected_sheet_snapshots(start_mine, start_theirs)
    _assert_exact(start_result, start_mine, start_theirs)
    _assert_proof_pairs_share_one_logical_slot(
        start_result, start_pending.duplicate_field_proof
    )


def _test_three_way_blank_proof_preserves_base_coordinates():
    mine = _snapshot("mine")
    base = _snapshot("base")
    theirs = _snapshot("theirs")
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    _assert_exact(result, mine, theirs, base)
    assert all(not cols for cols in result.conflict_cols)
    proof = sm._build_snapshot_duplicate_field_identity_proof(
        mine,
        theirs,
        sm._align_selected_sheet_snapshots(mine, theirs),
        base,
        sm._align_selected_sheet_snapshots(mine, base),
        sm._align_selected_sheet_snapshots(theirs, base),
    )
    assert proof is not None
    _assert_proof_pairs_share_one_logical_slot(result, proof)


def _test_three_way_base_physical_offset_fails_closed_without_target():
    mine = _snapshot("mine", layout="base_offset_mine")
    theirs = _snapshot("theirs", layout="base_offset_mine")
    base = _snapshot("base", layout="base_offset_base")
    alignment = sm._align_selected_sheet_snapshots(mine, theirs)
    mine_base = sm._align_selected_sheet_snapshots(mine, base)
    theirs_base = sm._align_selected_sheet_snapshots(theirs, base)
    proof = sm._build_snapshot_duplicate_field_identity_proof(
        mine, theirs, alignment, base, mine_base, theirs_base
    )
    assert proof is not None
    assert any(
        occurrence.mine_col != occurrence.base_col
        for occurrence in proof.occurrences
    )
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    assert result.unresolved
    # Moving ``prefix`` creates an A/B-only insertion and a Base-only delete.
    # The strict final cache gate must keep this structural evidence terminal;
    # no duplicate proof may invent an actionable all-side target.
    slots = tuple(result.column_cache.model.slots)
    assert any(
        (slot.mine_col, slot.base_col, slot.theirs_col, slot.state)
        == (2, None, 2, "inserted")
        for slot in slots
    )
    assert any(
        (slot.mine_col, slot.base_col, slot.theirs_col, slot.state)
        == (None, 6, None, "both-deleted")
        for slot in slots
    )
    immutable_cache = sm._snapshot_result_to_sheet_cache_immutable(
        "DuplicateFields", result, mine, theirs, base, has_base=True
    )
    assert immutable_cache["unresolved_reason"]
    assert "row_a_to_pair_idx" not in immutable_cache


def _test_three_way_same_position_same_gap_formula_diff_keeps_targets():
    mine = _snapshot(
        "mine",
        layout="interior_end",
        formula_overrides={(3, 5): "=1+1"},
    )
    theirs = _snapshot(
        "theirs",
        layout="interior_end",
        formula_overrides={(3, 5): "=1+1"},
    )
    base = _snapshot(
        "base",
        layout="interior_end",
        formula_overrides={(3, 5): "=1+2"},
    )
    original_builder = sm.build_logical_column_comparison_cache_3way

    def _same_gap_builder(*args, **kwargs):
        cache = original_builder(*args, **kwargs)
        three_way = cache.three_way_alignment
        assert three_way is not None
        mine_to_base = sm.ColumnAlignmentResult(
            three_way.mine_to_base.model,
            ((1, 1), (2, 2)),
            three_way.mine_to_base.fallback_slot_indices,
            three_way.mine_to_base.used_physical_fallback,
            three_way.mine_to_base.fallback_reason,
        )
        theirs_to_base = sm.ColumnAlignmentResult(
            three_way.theirs_to_base.model,
            ((1, 1), (2, 2)),
            three_way.theirs_to_base.fallback_slot_indices,
            three_way.theirs_to_base.used_physical_fallback,
            three_way.theirs_to_base.fallback_reason,
        )
        alignment = sm.ColumnAlignment3WayResult(
            cache.model,
            mine_to_base,
            theirs_to_base,
            three_way.fallback_slot_indices,
            three_way.used_physical_fallback,
            three_way.fallback_reason,
        )
        return sm.LogicalColumnComparisonCache(
            model=cache.model,
            three_way_alignment=alignment,
            structural_diff_cols=cache.structural_diff_cols,
            unresolved_cols=cache.unresolved_cols,
        )

    try:
        sm.build_logical_column_comparison_cache_3way = _same_gap_builder
        result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    finally:
        sm.build_logical_column_comparison_cache_3way = original_builder
    _assert_exact(result, mine, theirs, base)
    target = next(
        slot for slot in result.column_cache.model.slots
        if (slot.mine_col, slot.base_col, slot.theirs_col) == (5, 5, 5)
    )
    logical_col = target.logical_idx + 1
    assert any(logical_col in cols for cols in result.pair_base_diff_cols)
    proof = sm._build_snapshot_duplicate_field_identity_proof(
        mine,
        theirs,
        sm._align_selected_sheet_snapshots(mine, theirs),
        base,
        sm._align_selected_sheet_snapshots(mine, base),
        sm._align_selected_sheet_snapshots(theirs, base),
    )
    assert proof is not None
    groups = sm._snapshot_duplicate_proof_pre_gap_groups(
        proof,
        result.column_cache.three_way_alignment,
        mine_max_col=mine.max_col,
        theirs_max_col=theirs.max_col,
        base_max_col=base.max_col,
    )
    assert groups is not None
    assert any(
        {item.occurrence.mine_col for item in group} == {3, 4, 6, 7}
        and group[0].mine_cache_bounds == (2, 8)
        and group[0].base_cache_bounds == (2, 8)
        for group in groups
    )


def _three_way_exact_top_fast_path_fixture():
    mine = _snapshot("mine", layout="interior_end")
    theirs = _snapshot("theirs", layout="interior_end")
    base = _snapshot("base", layout="interior_end")
    proof = sm._build_snapshot_duplicate_field_identity_proof(
        mine,
        theirs,
        sm._align_selected_sheet_snapshots(mine, theirs),
        base,
        sm._align_selected_sheet_snapshots(mine, base),
        sm._align_selected_sheet_snapshots(theirs, base),
    )
    assert proof is not None
    return mine, theirs, base, proof


def _force_exact_top_slots(cache, proof):
    proof_pairs = set(proof.pairs)
    confidence = sm.ColumnMappingConfidence(
        1.0, False, "injected-exact-top-cache", ("contract",)
    )
    return [
        _slot_like(slot, state="retained", confidence=confidence)
        if any(sm._snapshot_duplicate_proof_slot_matches(
            slot, pair, three_way=True
        ) for pair in proof_pairs) else slot
        for slot in cache.model.slots
    ], confidence


def _test_three_way_exact_top_cache_accepts_asymmetric_child_gaps():
    mine, theirs, base, proof = _three_way_exact_top_fast_path_fixture()
    original_builder = sm.build_logical_column_comparison_cache_3way

    def _asymmetric_builder(*args, **kwargs):
        cache = original_builder(*args, **kwargs)
        slots, confidence = _force_exact_top_slots(cache, proof)
        return _three_way_cache_with_slots(
            cache,
            slots,
            # Mine/Base has one tail gap while Theirs/Base has an intervening
            # physical anchor.  The fast path must prove the already exact top
            # triples rather than deriving a shared child pre-gap.
            mine_to_base_anchor_pairs=((1, 1), (2, 2)),
            theirs_to_base_anchor_pairs=((1, 1), (2, 2), (5, 5)),
            structural_diff_cols=frozenset(),
            unresolved_cols=frozenset(),
            model_confidence=confidence,
            fallback_slot_indices=(),
            used_physical_fallback=False,
            fallback_reason="",
        )

    try:
        sm.build_logical_column_comparison_cache_3way = _asymmetric_builder
        result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    finally:
        sm.build_logical_column_comparison_cache_3way = original_builder
    _assert_exact(result, mine, theirs, base)
    _assert_proof_pairs_share_one_logical_slot(result, proof)
    # The helper correctly cannot provide a common child pre-gap; accepting
    # this result therefore proves the fast path did not invoke a rebuild.
    assert sm._snapshot_duplicate_proof_pre_gap_groups(
        proof,
        result.column_cache.three_way_alignment,
        mine_max_col=mine.max_col,
        theirs_max_col=theirs.max_col,
        base_max_col=base.max_col,
    ) is None


def _test_three_way_nonpending_top_cache_failures_remain_unresolved():
    mine, theirs, base, proof = _three_way_exact_top_fast_path_fixture()
    original_builder = sm.build_logical_column_comparison_cache_3way

    def _broken(mode):
        def _builder(*args, **kwargs):
            cache = original_builder(*args, **kwargs)
            slots, confidence = _force_exact_top_slots(cache, proof)
            if mode == "crosswire":
                first, second = proof.pairs[:2]
                first_index = next(
                    index for index, slot in enumerate(slots)
                    if sm._snapshot_duplicate_proof_slot_matches(
                        slot, first, three_way=True)
                )
                second_index = next(
                    index for index, slot in enumerate(slots)
                    if sm._snapshot_duplicate_proof_slot_matches(
                        slot, second, three_way=True)
                )
                first_slot, second_slot = slots[first_index], slots[second_index]
                slots[first_index] = _slot_like(
                    first_slot, base_col=second_slot.base_col
                )
                slots[second_index] = _slot_like(
                    second_slot, base_col=first_slot.base_col
                )
                return _three_way_cache_with_slots(
                    cache, slots, structural_diff_cols=frozenset(),
                    unresolved_cols=frozenset(), model_confidence=confidence,
                    fallback_slot_indices=(),
                    used_physical_fallback=False,
                    fallback_reason="",
                )
            if mode == "nonretained":
                first = proof.pairs[0]
                index = next(
                    index for index, slot in enumerate(slots)
                    if sm._snapshot_duplicate_proof_slot_matches(
                        slot, first, three_way=True)
                )
                slots[index] = _slot_like(
                    slots[index], state="inserted", confidence=confidence
                )
                return _three_way_cache_with_slots(
                    cache, slots, structural_diff_cols=frozenset(),
                    unresolved_cols=frozenset(), model_confidence=confidence,
                    fallback_slot_indices=(),
                    used_physical_fallback=False,
                    fallback_reason="",
                )
            if mode == "fallback_only":
                # The model and every slot are otherwise exact and bijective.
                # A top-level physical fallback alone must reject the direct
                # fast path rather than being treated as a harmless diagnostic.
                return _three_way_cache_with_slots(
                    cache, slots, structural_diff_cols=frozenset(),
                    unresolved_cols=frozenset(), model_confidence=confidence,
                    fallback_slot_indices=(slots[0].logical_idx,),
                    used_physical_fallback=True,
                    fallback_reason="unresolved-three-way-column-range",
                )
            return _three_way_cache_with_slots(
                cache, slots, structural_diff_cols=frozenset((1,)),
                unresolved_cols=frozenset(), model_confidence=confidence,
                fallback_slot_indices=(),
                used_physical_fallback=False,
                fallback_reason="",
            )
        return _builder

    try:
        for mode in ("crosswire", "nonretained", "fallback_only", "structural"):
            sm.build_logical_column_comparison_cache_3way = _broken(mode)
            _assert_unresolved(mine, theirs, base)
    finally:
        sm.build_logical_column_comparison_cache_3way = original_builder


def _test_pre_gap_groups_do_not_cross_original_anchor_boundaries():
    mine = _snapshot("mine")
    theirs = _snapshot("theirs")
    pending = sm._align_selected_sheet_snapshots(mine, theirs)
    proof = pending.duplicate_field_proof
    assert proof is not None
    row_pairs = pending.row_pairs
    mine_payloads = tuple(sm._snapshot_row_payload(mine, row) for row, _ in row_pairs)
    theirs_payloads = tuple(sm._snapshot_row_payload(theirs, row) for _, row in row_pairs)
    cache = sm.build_logical_column_comparison_cache_2way(
        sm.ColumnModelCacheKey(mine.sheet, 0, 0),
        tuple(payload[0] for payload in mine_payloads),
        tuple(payload[0] for payload in theirs_payloads),
        tuple(payload[1] for payload in mine_payloads),
        tuple(payload[1] for payload in theirs_payloads),
        mine_max_col=mine.max_col,
        theirs_max_col=theirs.max_col,
    )
    groups = sm._snapshot_duplicate_proof_pre_gap_groups(
        proof,
        cache.two_way_alignment,
        mine_max_col=mine.max_col,
        theirs_max_col=theirs.max_col,
    )
    assert groups is not None
    physical_groups = {
        tuple(item.occurrence.mine_col for item in group)
        for group in groups
    }
    assert physical_groups == {(3, 4), (6, 7)}, physical_groups
    assert any(
        group[0].theirs_cache_bounds[1] == theirs.max_col + 1
        for group in groups if tuple(item.occurrence.mine_col for item in group) == (6, 7)
    )


def _test_two_stage_formula_gap_preserves_nonproof_cell_diffs():
    """A gap-level formula cause may not hide a resolved unique member."""
    mine = _snapshot("mine", formula_overrides={(3, 5): "=1+1"})
    theirs = _snapshot("theirs", formula_overrides={(3, 5): "=1+2"})
    pending = sm._align_selected_sheet_snapshots(mine, theirs)
    proof = pending.duplicate_field_proof
    assert proof is not None
    original_builder = sm.build_logical_column_comparison_cache_2way

    def _formula_gap_builder(*args, **kwargs):
        cache = original_builder(*args, **kwargs)
        proof_mine_cols = {mine_col for mine_col, _theirs_col, _base_col in proof.pairs}
        slots = [
            _slot_like(
                slot,
                state="unresolved",
                confidence=sm.ColumnMappingConfidence(
                    0.0,
                    True,
                    "inherited-formula-gap",
                    cause_codes=(
                        sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,
                        sm.COLUMN_MAPPING_CAUSE_FORMULA_MISMATCH,
                    ),
                ),
            ) if slot.mine_col in proof_mine_cols else slot
            for slot in cache.model.slots
        ]
        # This makes the full physical range 3..7 one original cache gap;
        # col5 is then the unique nonproof member whose formula difference is
        # valid provenance for the duplicate slots' inherited cause.
        return _two_way_cache_with_slots(
            cache,
            slots,
            global_ambiguous=True,
            model_reason="unresolved-column-range",
            model_cause_codes=(),
            anchor_pairs=((1, 1), (2, 2)),
        )

    try:
        sm.build_logical_column_comparison_cache_2way = _formula_gap_builder
        result = sm._compare_selected_sheet_snapshots(mine, theirs)
        _assert_exact(result, mine, theirs)
        assert all(
            slot.state == "retained" and not slot.confidence.ambiguous
            for slot in result.column_cache.model.slots
            if slot.logical_idx + 1 == 5
        )
        assert any(5 in cols for cols in result.pair_diff_cols), (
            "the unique formula-different interval member must remain a cell diff"
        )
        target = next(
            slot for slot in result.column_cache.model.slots
            if slot.mine_col == 5
        )
        assert (target.mine_col, target.theirs_col, target.state) == (5, 5, "retained")

        # The same inherited formula cause remains fail-closed when the sole
        # non-proof interval member has not independently retained a mapping.
        def _broken_interval_builder(*args, **kwargs):
            broken = _formula_gap_builder(*args, **kwargs)
            slots = list(broken.model.slots)
            index = next(idx for idx, slot in enumerate(slots) if slot.mine_col == 5)
            slots[index] = _slot_like(
                slots[index],
                state="unresolved",
                confidence=sm.ColumnMappingConfidence(
                    0.0,
                    True,
                    "injected-nonproof-interval-failure",
                    cause_codes=(sm.COLUMN_MAPPING_CAUSE_LOW_CONFIDENCE,),
                ),
            )
            return _two_way_cache_with_slots(
                broken,
                slots,
                global_ambiguous=True,
                model_reason="unresolved-column-range",
                model_cause_codes=(),
            )

        sm.build_logical_column_comparison_cache_2way = _broken_interval_builder
        _assert_unresolved(mine, theirs)
    finally:
        sm.build_logical_column_comparison_cache_2way = original_builder


def _test_formula_mismatch_must_belong_to_the_same_original_pre_gap():
    mine = _snapshot(
        "mine",
        layout="formula_outside",
        formula_overrides={(3, 2): "=1+1"},
    )
    theirs = _snapshot(
        "theirs",
        layout="formula_outside",
        formula_overrides={(3, 2): "=1+2"},
    )
    pending = sm._align_selected_sheet_snapshots(mine, theirs)
    proof = pending.duplicate_field_proof
    assert proof is not None
    original_builder = sm.build_logical_column_comparison_cache_2way

    def _unrelated_formula_builder(*args, **kwargs):
        cache = original_builder(*args, **kwargs)
        proof_mine_cols = {pair[0] for pair in proof.pairs}
        slots = [
            _slot_like(
                slot,
                state="unresolved",
                confidence=sm.ColumnMappingConfidence(
                    0.0,
                    True,
                    "injected-unrelated-formula-gap",
                    cause_codes=(
                        sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,
                        sm.COLUMN_MAPPING_CAUSE_FORMULA_MISMATCH,
                    ),
                ),
            ) if slot.mine_col in proof_mine_cols else slot
            for slot in cache.model.slots
        ]
        # Preserve the real anchors: each duplicate run has its own gap and
        # neither contains a formula signal.  The same cause code must not
        # borrow provenance from a different gap.
        return _two_way_cache_with_slots(
            cache,
            slots,
            global_ambiguous=True,
            model_reason="unresolved-column-range",
            model_cause_codes=(sm.COLUMN_MAPPING_CAUSE_FORMULA_MISMATCH,),
        )

    try:
        sm.build_logical_column_comparison_cache_2way = _unrelated_formula_builder
        _assert_unresolved(mine, theirs)
    finally:
        sm.build_logical_column_comparison_cache_2way = original_builder


def _test_three_way_base_blank_mismatch_fails_closed():
    mine = _snapshot("mine", layout="base_offset_mine")
    theirs = _snapshot("theirs", layout="base_offset_mine")
    # Base physical col3 is one of the offset interior duplicate members.
    _assert_unresolved(
        mine,
        theirs,
        _snapshot(
            "base",
            layout="base_offset_base",
            value_overrides={(3, 3): "base-only-nonblank"},
        ),
    )


def _test_negative_proofs_remain_unresolved_and_non_actionable():
    mine = _snapshot("mine")
    _assert_unresolved(
        _snapshot("mine", value_overrides={(3, 3): "unexpected"}),
        _snapshot("theirs"),
    )
    _assert_unresolved(
        _snapshot("mine", formula_overrides={(3, 3): "=1+1"}),
        _snapshot("theirs"),
    )
    # Values/formulas are blank on both sides here, but their typed digest is
    # not.  This is distinct from the nonblank rejection above.
    _assert_unresolved(
        mine,
        _snapshot("theirs", type_overrides={(3, 3): "s"}),
    )
    _assert_unresolved(mine, _snapshot("theirs", layout="missing_member"))
    _assert_unresolved(mine, _snapshot("theirs", layout="run_mismatch"))
    _assert_unresolved(mine, _snapshot("theirs", layout="anchor_reordered"))
    _assert_unresolved(
        mine,
        _snapshot("theirs", ids=("id-01", "id-01", "id-03")),
    )
    _assert_unresolved(
        mine,
        _snapshot("theirs"),
        _snapshot("base", type_overrides={(3, 3): "s"}),
    )


def _test_builder_or_final_cache_exception_fails_closed():
    mine = _snapshot("mine")
    theirs = _snapshot("theirs")
    original_builder = sm._build_snapshot_duplicate_field_identity_proof
    original_apply = sm._apply_snapshot_duplicate_field_proof_to_column_cache
    try:
        def _raise_builder(*_args, **_kwargs):
            raise RuntimeError("injected duplicate proof builder failure")

        sm._build_snapshot_duplicate_field_identity_proof = _raise_builder
        _assert_unresolved(mine, theirs)
        sm._build_snapshot_duplicate_field_identity_proof = original_builder

        def _raise_finalizer(*_args, **_kwargs):
            raise RuntimeError("injected final cache failure")

        sm._apply_snapshot_duplicate_field_proof_to_column_cache = _raise_finalizer
        _assert_unresolved(mine, theirs)
    finally:
        sm._build_snapshot_duplicate_field_identity_proof = original_builder
        sm._apply_snapshot_duplicate_field_proof_to_column_cache = original_apply


_UNSET = object()


def _slot_like(
    slot, *, base_col=_UNSET, theirs_col=_UNSET, state=None, confidence=None
):
    return sm.ColumnSlot(
        logical_idx=slot.logical_idx,
        mine_col=slot.mine_col,
        base_col=slot.base_col if base_col is _UNSET else base_col,
        theirs_col=slot.theirs_col if theirs_col is _UNSET else theirs_col,
        state=slot.state if state is None else state,
        confidence=slot.confidence if confidence is None else confidence,
        base_boundary=slot.base_boundary,
        origin_side=slot.origin_side,
    )


def _two_way_cache_with_slots(
    cache,
    slots,
    *,
    global_ambiguous=False,
    fallback_slot_indices=None,
    fallback_reason=None,
    anchor_pairs=None,
    model_reason="injected-final-cache-state",
    model_cause_codes=(),
):
    slots = tuple(slots)
    unresolved = tuple(
        slot.logical_idx for slot in slots
        if slot.state == "unresolved" or slot.confidence.ambiguous
    )
    model = sm.ColumnModel.from_slots(
        cache.model.cache_key,
        slots,
        blocks=sm._build_column_blocks(slots),
        confidence=sm.ColumnMappingConfidence(
            1.0,
            bool(global_ambiguous),
            str(model_reason),
            ("contract-injection",),
            tuple(model_cause_codes),
        ),
    )
    if fallback_slot_indices is None:
        fallback_slot_indices = unresolved
    fallback_slot_indices = tuple(fallback_slot_indices)
    if anchor_pairs is None:
        anchor_pairs = cache.two_way_alignment.anchor_pairs
    anchor_pairs = tuple(anchor_pairs)
    used_physical_fallback = bool(fallback_slot_indices)
    alignment = sm.ColumnAlignmentResult(
        model,
        anchor_pairs,
        fallback_slot_indices,
        used_physical_fallback,
        (
            fallback_reason
            if fallback_reason is not None
            else "unresolved-column-range" if used_physical_fallback else ""
        ),
    )
    return sm.LogicalColumnComparisonCache(
        model=model,
        two_way_alignment=alignment,
        structural_diff_cols=frozenset(),
        unresolved_cols=frozenset(slot.logical_idx + 1 for slot in slots if slot.logical_idx in unresolved),
    )


def _three_way_cache_with_slots(
    cache,
    slots,
    *,
    mine_to_base_anchor_pairs=None,
    theirs_to_base_anchor_pairs=None,
    structural_diff_cols=None,
    unresolved_cols=None,
    model_confidence=None,
    fallback_slot_indices=None,
    used_physical_fallback=None,
    fallback_reason=None,
):
    """Test-only top-cache injection; child alignments remain diagnostics."""
    slots = tuple(slots)
    model = sm.ColumnModel.from_slots(
        cache.model.cache_key,
        slots,
        blocks=sm._build_column_blocks(slots),
        confidence=cache.model.confidence if model_confidence is None else model_confidence,
    )
    original = cache.three_way_alignment
    assert original is not None
    mine_to_base = sm.ColumnAlignmentResult(
        original.mine_to_base.model,
        original.mine_to_base.anchor_pairs if mine_to_base_anchor_pairs is None else tuple(mine_to_base_anchor_pairs),
        original.mine_to_base.fallback_slot_indices,
        original.mine_to_base.used_physical_fallback,
        original.mine_to_base.fallback_reason,
    )
    theirs_to_base = sm.ColumnAlignmentResult(
        original.theirs_to_base.model,
        original.theirs_to_base.anchor_pairs if theirs_to_base_anchor_pairs is None else tuple(theirs_to_base_anchor_pairs),
        original.theirs_to_base.fallback_slot_indices,
        original.theirs_to_base.used_physical_fallback,
        original.theirs_to_base.fallback_reason,
    )
    alignment = sm.ColumnAlignment3WayResult(
        model,
        mine_to_base,
        theirs_to_base,
        (
            original.fallback_slot_indices
            if fallback_slot_indices is None
            else tuple(fallback_slot_indices)
        ),
        (
            original.used_physical_fallback
            if used_physical_fallback is None
            else bool(used_physical_fallback)
        ),
        original.fallback_reason if fallback_reason is None else str(fallback_reason),
    )
    return sm.LogicalColumnComparisonCache(
        model=model,
        three_way_alignment=alignment,
        structural_diff_cols=(
            cache.structural_diff_cols if structural_diff_cols is None
            else frozenset(structural_diff_cols)
        ),
        unresolved_cols=(
            cache.unresolved_cols if unresolved_cols is None
            else frozenset(unresolved_cols)
        ),
    )


def _test_final_cache_unresolved_and_crosswire_fail_closed():
    mine = _snapshot("mine")
    theirs = _snapshot("theirs")
    pending = sm._align_selected_sheet_snapshots(mine, theirs)
    proof = pending.duplicate_field_proof
    assert proof is not None
    row_pairs = pending.row_pairs
    mine_payloads = tuple(sm._snapshot_row_payload(mine, row) for row, _ in row_pairs)
    theirs_payloads = tuple(sm._snapshot_row_payload(theirs, row) for _, row in row_pairs)
    original_builder = sm.build_logical_column_comparison_cache_2way
    seed_cache = original_builder(
        sm.ColumnModelCacheKey(mine.sheet, 0, 0),
        tuple(payload[0] for payload in mine_payloads),
        tuple(payload[0] for payload in theirs_payloads),
        tuple(payload[1] for payload in mine_payloads),
        tuple(payload[1] for payload in theirs_payloads),
        mine_max_col=mine.max_col,
        theirs_max_col=theirs.max_col,
    )
    assert not sm._snapshot_duplicate_proof_cache_is_bijective(
        _two_way_cache_with_slots(
            seed_cache, seed_cache.model.slots, global_ambiguous=True
        ),
        mine,
        theirs,
        None,
    )

    def _broken(mode):
        def _builder(*args, **kwargs):
            cache = original_builder(*args, **kwargs)
            slots = list(cache.model.slots)
            proof_mine_cols = {pair[0] for pair in proof.pairs}
            if mode == "extra_fallback":
                index = next(
                    idx for idx, slot in enumerate(slots)
                    if slot.mine_col == 1
                )
                slots[index] = _slot_like(
                    slots[index],
                    state="unresolved",
                    confidence=sm.ColumnMappingConfidence(
                        0.0,
                        True,
                        "injected-extra-blank-fallback",
                        cause_codes=(sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,),
                    ),
                )
                return _two_way_cache_with_slots(
                    cache, slots, global_ambiguous=True
                )
            if mode == "nonblank_cause":
                index = next(
                    idx for idx, slot in enumerate(slots)
                    if slot.mine_col == 3
                )
                slots[index] = _slot_like(
                    slots[index],
                    state="unresolved",
                    confidence=sm.ColumnMappingConfidence(
                        0.0,
                        True,
                        "injected-nonblank-cause",
                        cause_codes=(sm.COLUMN_MAPPING_CAUSE_LOW_CONFIDENCE,),
                    ),
                )
                return _two_way_cache_with_slots(
                    cache, slots, global_ambiguous=True
                )
            elif mode == "global_ambiguous":
                return _two_way_cache_with_slots(
                    cache,
                    slots,
                    global_ambiguous=True,
                    fallback_slot_indices=(),
                    fallback_reason="",
                    model_reason="unresolved-column-range",
                    model_cause_codes=(sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,),
                )
            elif mode == "wrong_model_reason":
                return _two_way_cache_with_slots(
                    cache,
                    slots,
                    global_ambiguous=True,
                    model_cause_codes=(sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,),
                )
            elif mode == "wrong_model_cause":
                return _two_way_cache_with_slots(
                    cache,
                    slots,
                    global_ambiguous=True,
                    model_reason="unresolved-column-range",
                    model_cause_codes=(sm.COLUMN_MAPPING_CAUSE_LOW_CONFIDENCE,),
                )
            elif mode == "slot_only_inconsistent":
                # The unresolved duplicate slots remain, but the malformed
                # global model claims they are resolved.  The comparator must
                # still reject this pending cache evidence fail-closed.
                return _two_way_cache_with_slots(
                    cache,
                    slots,
                    global_ambiguous=False,
                    model_reason="unresolved-column-range",
                    model_cause_codes=(sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,),
                )
            elif mode in ("model_only_empty", "model_only_formula"):
                # A global empty/formula-only diagnostic is not proof.  With
                # no matching unresolved slot evidence it must not promote a
                # duplicate range, even when fallback indices name it.
                slots = [
                    _slot_like(
                        slot,
                        state="retained",
                        confidence=sm.ColumnMappingConfidence(
                            1.0, False, "injected-model-only"
                        ),
                    ) if slot.mine_col in proof_mine_cols else slot
                    for slot in slots
                ]
                return _two_way_cache_with_slots(
                    cache,
                    slots,
                    global_ambiguous=True,
                    fallback_slot_indices=tuple(
                        slot.logical_idx for slot in cache.model.slots
                        if slot.mine_col in proof_mine_cols
                    ),
                    fallback_reason="unresolved-column-range",
                    model_reason="unresolved-column-range",
                    model_cause_codes=(
                        () if mode == "model_only_empty" else
                        (sm.COLUMN_MAPPING_CAUSE_FORMULA_MISMATCH,)
                    ),
                )
            elif mode == "structural_proof_slot":
                index = next(
                    idx for idx, slot in enumerate(slots)
                    if slot.mine_col == min(proof_mine_cols)
                )
                slots[index] = _slot_like(
                    slots[index],
                    theirs_col=None,
                    state="unresolved",
                    confidence=sm.ColumnMappingConfidence(
                        0.0,
                        True,
                        "injected-structural-proof-gap",
                        cause_codes=(sm.COLUMN_MAPPING_CAUSE_BLANK_COLUMN,),
                    ),
                )
                return _two_way_cache_with_slots(
                    cache, slots, global_ambiguous=True
                )
            else:
                by_mine = {slot.mine_col: slot for slot in slots}
                for index, slot in enumerate(slots):
                    if slot.mine_col == 3:
                        slots[index] = _slot_like(slot, theirs_col=by_mine[4].theirs_col)
                    elif slot.mine_col == 4:
                        slots[index] = _slot_like(slot, theirs_col=by_mine[3].theirs_col)
            return _two_way_cache_with_slots(cache, slots)

        return _builder

    try:
        sm.build_logical_column_comparison_cache_2way = _broken("global_ambiguous")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("wrong_model_reason")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("wrong_model_cause")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("slot_only_inconsistent")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("model_only_empty")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("model_only_formula")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("extra_fallback")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("nonblank_cause")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("crosswire")
        _assert_unresolved(mine, theirs)
        sm.build_logical_column_comparison_cache_2way = _broken("structural_proof_slot")
        _assert_unresolved(mine, theirs)
    finally:
        sm.build_logical_column_comparison_cache_2way = original_builder


def main():
    _test_two_way_blank_interior_end_and_start_proofs()
    _test_three_way_blank_proof_preserves_base_coordinates()
    _test_three_way_base_physical_offset_fails_closed_without_target()
    _test_three_way_same_position_same_gap_formula_diff_keeps_targets()
    _test_three_way_exact_top_cache_accepts_asymmetric_child_gaps()
    _test_three_way_nonpending_top_cache_failures_remain_unresolved()
    _test_pre_gap_groups_do_not_cross_original_anchor_boundaries()
    _test_two_stage_formula_gap_preserves_nonproof_cell_diffs()
    _test_formula_mismatch_must_belong_to_the_same_original_pre_gap()
    _test_three_way_base_blank_mismatch_fails_closed()
    _test_negative_proofs_remain_unresolved_and_non_actionable()
    _test_builder_or_final_cache_exception_fails_closed()
    _test_final_cache_unresolved_and_crosswire_fail_closed()
    print("SMOKE_DUPLICATE_FIELD_IDENTITY_PROOF_OK")


if __name__ == "__main__":
    main()
