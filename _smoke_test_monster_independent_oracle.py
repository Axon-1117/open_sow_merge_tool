"""Pure contract coverage for the independent Monster typed-key grouper.

This deliberately creates immutable snapshots in memory.  It neither opens a
workbook nor calls the production snapshot comparator/alignment/cache path.
"""

from __future__ import annotations

from dataclasses import replace
import hashlib
import json

import _gui_real_changed_revision_benchmark as benchmark
import sow_merge_tool as sm


def _cell(value=None, *, cached_type=None, formula=None, formula_type=None):
    if cached_type is None:
        cached_type = "n" if value is None or isinstance(value, (int, float)) else "s"
    if formula_type is None:
        formula_type = "f" if formula is not None else "n"
    return sm.SnapshotCell(
        value, cached_type, formula, formula_type,
        "formula" if formula is not None else "literal", False,
    )


def _row(physical_row, cells):
    return sm.SnapshotRow(int(physical_row), tuple(cells), f"row-{physical_row}")


def _fields(*markers):
    fields = []
    for index, marker in enumerate(markers, start=1):
        declaration = "id@id" if marker == "id" else "const@const"
        fields.append(sm.SnapshotField(
            index, declaration, "int32", frozenset({marker})
        ))
    fields.append(sm.SnapshotField(len(fields) + 1, "payload", "string", frozenset()))
    return tuple(fields)


def _snapshot(side, groups, *, markers=("id",), physical_start=3):
    """Build rows as ``(owner-key-cells, continuation-payloads, payload)``."""
    fields = _fields(*markers)
    key_count = len(markers)
    rows = [
        _row(1, [_cell("header")] * (key_count + 1)),
        _row(2, [_cell("type")] * (key_count + 1)),
    ]
    physical = int(physical_start)
    for owner_cells, continuations, payload in groups:
        assert len(owner_cells) == key_count
        rows.append(_row(physical, [*owner_cells, payload]))
        physical += 1
        for continuation_payload in continuations:
            rows.append(_row(
                physical,
                [*([_cell()] * key_count), continuation_payload],
            ))
            physical += 1
    return sm.SheetSnapshot(
        str(side), "MonsterGroup@design",
        sm.SheetSnapshotVersion(1, 0, 0, 1, 1),
        max(row.physical_row for row in rows), len(fields), fields, tuple(rows),
    )


_RELEASE_FOCUS_FIELDS = {
    1: ("id@id", "int32", frozenset({"id"})),
    13: ("monster_mark@pm", "string", frozenset()),
    14: ("monster_number", "int32", frozenset()),
    27: ("", "", frozenset()),
    28: ("attack_base@pm", "float32", frozenset()),
    29: ("defense_base@pm", "float32", frozenset()),
    30: ("hp_base@pm", "float32", frozenset()),
    31: ("", "", frozenset()),
}
_RELEASE_DIFF_PROFILE = (
    (13, "monster_mark@pm", "string", 3075),
    (28, "attack_base@pm", "float32", 3075),
    (29, "defense_base@pm", "float32", 3074),
    (30, "hp_base@pm", "float32", 3074),
)


def _full_fields(*, schema_variant=None):
    """Use the target's coordinate-sensitive headers, not generic aliases."""
    focus = dict(_RELEASE_FOCUS_FIELDS)
    if schema_variant == "swap-p13-p14":
        focus[13], focus[14] = focus[14], focus[13]
    elif schema_variant == "wrong-type-p13":
        declaration, _type, markers = focus[13]
        focus[13] = (declaration, "int32", markers)
    elif schema_variant is not None:
        raise AssertionError(f"unknown full fixture schema variant: {schema_variant}")
    return tuple(
        sm.SnapshotField(
            physical_col,
            *(focus.get(
                physical_col,
                (f"field-{physical_col}", "int32", frozenset()),
            )),
        )
        for physical_col in range(1, 32)
    )


def _local_token(cell):
    """Test-local token: do not delegate expected values to the harness."""
    return (
        str(cell.cached_type or ""), repr(cell.cached_value),
        str(cell.formula_type or ""), repr(cell.formula_value),
        str(cell.formula_kind or ""), bool(cell.external_link),
    )


def _manifest_digest(rows, columns):
    canonical = json.dumps(
        {"rows": list(rows), "columns": list(columns)},
        ensure_ascii=False, separators=(",", ":"), sort_keys=True,
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest().upper()


def _full_snapshot(
    side, *, group_lengths=None, variant="mine", key_override=None,
    schema_variant=None,
):
    """Create the fixed 3169x31 Monster contract plus an independent manifest."""
    if group_lengths is None:
        group_lengths = [1] * 203 + [0] * 2761
    assert len(group_lengths) + sum(group_lengths) == 3167
    fields = _full_fields(schema_variant=schema_variant)
    blank = _cell()
    rows = [
        _row(1, [blank] * 31),
        _row(2, [blank] * 31),
    ]
    physical_row = 3
    identity_index = 0
    key_override = dict(key_override or {})
    identity_rows = {}
    identity_order = []
    injected_diffs = []

    def _data_cells(owner_value):
        return [_cell(owner_value)] + [blank] * 30

    def _apply_diff(cells, index, identity, row_number):
        if variant not in {"theirs", "base-different"} or index >= 3075:
            return
        # The target profile has 3,075 diffs at p13/p28 and 3,074 at p29/p30:
        # the first two pairs each omit a different base-stat field.  Total
        # typed-token differences stay 3 + 3 + 3073 * 4 == 12,298.
        columns = [13, 28, 29, 30]
        if index == 0:
            columns.remove(29)
        elif index == 1:
            columns.remove(30)
        for physical_col in columns:
            if physical_col == 13:
                cells[12] = _cell(f"value-{index}")
            elif physical_col == 28:
                cells[27] = _cell(None, formula=f"={index + 1}")
            elif physical_col == 29:
                cells[28] = _cell(f"formula-cache-{index}")
            else:
                cells[29] = _cell(float(index) + 0.25)
        if variant == "base-different" and index == 0:
            cells[1] = _cell("Base must not silently differ from Theirs")
            columns = (*columns, 2)
        for physical_col in columns:
            token = _local_token(cells[physical_col - 1])
            injected_diffs.append({
                "identity": identity,
                "physical_row": int(row_number),
                "physical_col": physical_col,
                "token": token,
            })

    for owner_index, continuation_count in enumerate(group_lengths, start=1):
        owner_value = key_override.get(owner_index, owner_index)
        owner_key = (_local_token(_cell(owner_value)),)
        identity = (owner_key, 0)
        owner = _data_cells(owner_value)
        _apply_diff(owner, identity_index, identity, physical_row)
        rows.append(_row(physical_row, owner))
        identity_rows[identity] = int(physical_row)
        identity_order.append(identity)
        physical_row += 1
        identity_index += 1
        for ordinal in range(1, int(continuation_count) + 1):
            identity = (owner_key, ordinal)
            continuation = [blank] * 31
            _apply_diff(continuation, identity_index, identity, physical_row)
            rows.append(_row(physical_row, continuation))
            identity_rows[identity] = int(physical_row)
            identity_order.append(identity)
            physical_row += 1
            identity_index += 1
    assert physical_row == 3170 and identity_index == 3167
    snapshot = sm.SheetSnapshot(
        str(side), "MonsterGroup@design",
        sm.SheetSnapshotVersion(1, 0, 0, 1, 1),
        3169, 31, fields, tuple(rows),
    )
    distribution = {}
    for count in group_lengths:
        distribution[int(count)] = int(distribution.get(int(count), 0)) + 1
    return snapshot, {
        "identity_rows": identity_rows,
        "identity_order": tuple(identity_order),
        "injected_diffs": tuple(injected_diffs),
        "topology": {
            "owner_count": len(group_lengths),
            "continuation_count": sum(group_lengths),
            "continuation_distribution": tuple(sorted(distribution.items())),
            "max_continuation_ordinal": max(group_lengths or [0]),
            "expanded_identity_count": len(identity_order),
        },
    }


def _full_expected_manifest(mine, theirs, *, base=None):
    """Derive all Oracle expectations from fixture manifests, not Oracle code."""
    mine_order = tuple(mine["identity_order"])
    theirs_order = tuple(theirs["identity_order"])
    assert mine_order == theirs_order
    identities = mine_order
    row_pairs = ((1, 1), (2, 2)) + tuple(
        (mine["identity_rows"][identity], theirs["identity_rows"][identity])
        for identity in identities
    )
    base_rows = None
    if base is not None:
        assert tuple(base["identity_order"]) == identities
        base_rows = (1, 2) + tuple(
            base["identity_rows"][identity] for identity in identities
        )
    columns = tuple(
        (logical, logical, logical if base is not None else None, logical)
        for logical in range(1, 32)
    )
    injected_by_identity = {}
    for item in theirs["injected_diffs"]:
        injected_by_identity.setdefault(item["identity"], set()).add(
            int(item["physical_col"]) - 1
        )
    pair_diff_physical_idx0 = (frozenset(), frozenset()) + tuple(
        frozenset(injected_by_identity.get(identity, set()))
        for identity in identities
    )
    pair_diff_cols = tuple(
        frozenset(int(index0) + 1 for index0 in cols)
        for cols in pair_diff_physical_idx0
    )
    if base is None:
        pair_base_diff_cols = tuple(frozenset() for _ in row_pairs)
        pair_base_diff_physical_idx0 = tuple(frozenset() for _ in row_pairs)
    else:
        base_by_identity = {}
        for item in base["injected_diffs"]:
            base_by_identity.setdefault(item["identity"], set()).add(
                int(item["physical_col"]) - 1
            )
        assert base_by_identity == injected_by_identity
        pair_base_diff_physical_idx0 = (frozenset(), frozenset()) + tuple(
            frozenset(base_by_identity.get(identity, set()))
            for identity in identities
        )
        pair_base_diff_cols = tuple(
            frozenset(int(index0) + 1 for index0 in cols)
            for cols in pair_base_diff_physical_idx0
        )
    mapping_rows = tuple(
        (
            pair_index, mine_row, theirs_row,
            None if base_rows is None else base_rows[pair_index],
        )
        for pair_index, (mine_row, theirs_row) in enumerate(row_pairs)
    )
    pair_coordinate_rows = tuple(
        (
            pair_index,
            tuple(sorted(int(index0) for index0 in pair_diff_physical_idx0[pair_index])),
            tuple(sorted(int(logical_col) for logical_col in pair_diff_cols[pair_index])),
        )
        for pair_index in range(len(row_pairs))
    )
    diff_physical_columns = sorted({
        int(index0) + 1
        for cols in pair_diff_physical_idx0
        for index0 in cols
    })
    expected_profile = [
        {
            "physical_col1": physical_col,
            "snapshot_idx0": physical_col - 1,
            "pair_diff_idx0": physical_col - 1,
            "declaration": declaration,
            "type_declaration": type_declaration,
            "expected_diff_pair_count": changed_pairs,
        }
        for physical_col, declaration, type_declaration, changed_pairs
        in _RELEASE_DIFF_PROFILE
    ]
    physical_pair_counts = {
        physical_col: sum(
            physical_col - 1 in cols for cols in pair_diff_physical_idx0
        )
        for physical_col in diff_physical_columns
    }
    assert physical_pair_counts == {
        item["physical_col1"]: item["expected_diff_pair_count"]
        for item in expected_profile
    }
    return {
        "row_pairs": row_pairs,
        "base_rows_by_pair": tuple(None for _ in row_pairs) if base_rows is None else base_rows,
        "pair_diff_cols": pair_diff_cols,
        "pair_diff_physical_idx0": pair_diff_physical_idx0,
        "pair_base_diff_cols": pair_base_diff_cols,
        "pair_base_diff_physical_idx0": pair_base_diff_physical_idx0,
        "column_slots": columns,
        "mapping_digest": _manifest_digest(mapping_rows, columns),
        "pair_diff_coordinate_digest": _manifest_digest(pair_coordinate_rows, ()),
        "diff_profile": expected_profile,
        "diff_physical_columns": diff_physical_columns,
        "diff_pair_idx0": [physical_col - 1 for physical_col in diff_physical_columns],
        "diff_pair_count": sum(bool(cols) for cols in pair_diff_cols),
        "diff_cell_count": sum(len(cols) for cols in pair_diff_cols),
        "diff_logical_columns": sorted(
            {int(column) for cols in pair_diff_cols for column in cols}
        ),
        "proof_logical_triples": [
            (27, 27, 27 if base is not None else None, 27),
            (31, 31, 31 if base is not None else None, 31),
        ],
        "injected_diffs": tuple(theirs["injected_diffs"]),
    }


def _expect_assertion(label, callback, *, contains=None):
    try:
        callback()
    except AssertionError as exc:
        if contains is not None:
            assert str(contains) in repr(exc), (label, exc)
        return
    raise AssertionError(f"{label}: expected AssertionError")


def _ids(snapshot):
    return benchmark._monster_declared_identity_rows(snapshot)


def _pair(mine, theirs, *, base=None):
    mine_rows, mine_order, _ = _ids(mine)
    theirs_rows, theirs_order, _ = _ids(theirs)
    if base is None:
        benchmark._monster_assert_identity_bijection(
            mine_rows, mine_order, theirs_rows, theirs_order
        )
        return mine_rows, mine_order, theirs_rows, theirs_order
    base_rows, base_order, _ = _ids(base)
    benchmark._monster_assert_identity_bijection(
        mine_rows, mine_order, theirs_rows, theirs_order,
        base_rows=base_rows, base_order=base_order,
    )
    return mine_rows, mine_order, theirs_rows, theirs_order, base_rows, base_order


def _assert_full_oracle_contract():
    """Exercise the actual full independent Oracle, not only its grouper."""
    mine, mine_manifest = _full_snapshot("mine", variant="mine")
    theirs, theirs_manifest = _full_snapshot("theirs", variant="theirs")
    base, base_manifest = _full_snapshot("base", variant="theirs")
    expected_two = _full_expected_manifest(mine_manifest, theirs_manifest)
    expected_three = _full_expected_manifest(
        mine_manifest, theirs_manifest, base=base_manifest
    )
    two_way = benchmark._monster_physical_oracle(mine, theirs)
    three_way = benchmark._monster_physical_oracle(mine, theirs, base)

    assert tuple(two_way.row_pairs) == expected_two["row_pairs"]
    assert tuple(two_way.base_rows_by_pair) == expected_two["base_rows_by_pair"]
    assert tuple(two_way.pair_diff_cols) == expected_two["pair_diff_cols"]
    assert tuple(two_way.pair_diff_physical_idx0) == expected_two["pair_diff_physical_idx0"]
    assert tuple(two_way.pair_base_diff_cols) == expected_two["pair_base_diff_cols"]
    assert tuple(two_way.pair_base_diff_physical_idx0) == expected_two[
        "pair_base_diff_physical_idx0"
    ]
    assert tuple(three_way.row_pairs) == expected_three["row_pairs"]
    assert tuple(three_way.base_rows_by_pair) == expected_three["base_rows_by_pair"]
    assert tuple(three_way.pair_diff_cols) == expected_three["pair_diff_cols"]
    assert tuple(three_way.pair_diff_physical_idx0) == expected_three[
        "pair_diff_physical_idx0"
    ]
    assert tuple(three_way.pair_base_diff_cols) == expected_three["pair_base_diff_cols"]
    assert tuple(three_way.pair_base_diff_physical_idx0) == expected_three[
        "pair_base_diff_physical_idx0"
    ]
    assert all(not cols for cols in three_way.conflict_cols)
    assert two_way.physical_facts["row_key_bijection"] == {
        "pair_count": len(expected_two["row_pairs"]),
        "declared_owner_count": mine_manifest["topology"]["owner_count"],
        "expanded_identity_count": mine_manifest["topology"]["expanded_identity_count"],
        "all_two_sided": True,
    }
    assert two_way.physical_facts["declared_key_topology"] == {
        "declared_key_fields": ((1, "id@id", "int32", ("id",)),),
        **mine_manifest["topology"],
        "leading_blank_count": 0,
        "duplicate_owner_count": 0,
    }
    assert two_way.physical_facts["physical_schema_columns"] == [
        slot[0] for slot in expected_two["column_slots"]
    ]
    assert two_way.physical_facts["blank_duplicate_proof_columns"] == [27, 31]
    assert two_way.physical_facts["diff_profile"] == expected_two["diff_profile"]
    assert two_way.physical_facts["diff_physical_columns"] == expected_two[
        "diff_physical_columns"
    ]
    assert two_way.physical_facts["diff_pair_idx0"] == expected_two["diff_pair_idx0"]
    assert two_way.physical_facts["diff_pair_count"] == expected_two["diff_pair_count"]
    assert two_way.physical_facts["diff_cell_count"] == expected_two["diff_cell_count"]
    assert two_way.physical_facts["diff_logical_columns"] == expected_two["diff_logical_columns"]
    assert two_way.physical_facts["mapping_digest"] == expected_two["mapping_digest"]
    assert two_way.physical_facts["pair_diff_coordinate_digest"] == expected_two[
        "pair_diff_coordinate_digest"
    ]
    assert two_way.physical_facts["proof_logical_triples"] == expected_two["proof_logical_triples"]
    assert three_way.physical_facts["base_equals_theirs"] is True
    assert three_way.physical_facts["conflict_count"] == 0
    assert three_way.physical_facts["mapping_digest"] == expected_three["mapping_digest"]
    assert three_way.physical_facts["pair_diff_coordinate_digest"] == expected_three[
        "pair_diff_coordinate_digest"
    ]
    assert three_way.physical_facts["proof_logical_triples"] == expected_three["proof_logical_triples"]
    assert three_way.physical_facts["diff_pair_count"] == expected_three["diff_pair_count"]
    assert three_way.physical_facts["diff_cell_count"] == expected_three["diff_cell_count"]
    assert three_way.physical_facts["diff_logical_columns"] == expected_three["diff_logical_columns"]

    long_group, _long_manifest = _full_snapshot(
        "long", group_lengths=[2] + [1] * 201 + [0] * 2762, variant="mine"
    )
    _expect_assertion(
        "group length greater than one",
        lambda: benchmark._monster_physical_oracle(long_group, theirs),
        contains="continuation_distribution",
    )
    base_topology, _base_topology_manifest = _full_snapshot(
        "base-topology", variant="theirs", key_override={1: 999_999}
    )
    _expect_assertion(
        "Base typed-key topology mismatch",
        lambda: benchmark._monster_physical_oracle(mine, theirs, base_topology),
        contains="Base declared-key coverage differs",
    )
    base_content, _base_content_manifest = _full_snapshot("base-content", variant="base-different")
    _expect_assertion(
        "Base content mismatch",
        lambda: benchmark._monster_physical_oracle(mine, theirs, base_content),
        contains="base-content-mismatch",
    )
    swapped_mine, _ = _full_snapshot("swapped-mine", schema_variant="swap-p13-p14")
    swapped_theirs, _ = _full_snapshot("swapped-theirs", schema_variant="swap-p13-p14")
    _expect_assertion(
        "release profile rejects same-side p13/p14 swap",
        lambda: benchmark._monster_physical_oracle(swapped_mine, swapped_theirs),
        contains="release-diff-schema-position",
    )
    wrong_type_mine, _ = _full_snapshot("wrong-type-mine", schema_variant="wrong-type-p13")
    wrong_type_theirs, _ = _full_snapshot("wrong-type-theirs", schema_variant="wrong-type-p13")
    _expect_assertion(
        "release profile rejects same-side p13 wrong type",
        lambda: benchmark._monster_physical_oracle(wrong_type_mine, wrong_type_theirs),
        contains="release-diff-schema-unique",
    )


def _exercise_contracts():
    # Positive: physical rows differ across sides but typed owner-key plus
    # continuation ordinal still provides the exact pairing.  A payload token
    # difference remains a physical cell difference rather than a key change.
    mine = _snapshot(
        "mine", [([_cell(0)], [_cell(None, formula="=1")], _cell("mine"))]
    )
    theirs = _snapshot(
        "theirs", [([_cell(0)], [_cell(None, formula="=2")], _cell("theirs"))],
        physical_start=30,
    )
    mine_rows, order, theirs_rows, _ = _pair(mine, theirs)
    assert [mine_rows[item] for item in order] == [3, 4]
    assert [theirs_rows[item] for item in order] == [30, 31]
    mine_by_row = {row.physical_row: row for row in mine.rows}
    theirs_by_row = {row.physical_row: row for row in theirs.rows}
    assert benchmark._physical_cell_token(mine_by_row[3].cells[1]) != (
        benchmark._physical_cell_token(theirs_by_row[30].cells[1])
    )
    assert benchmark._physical_cell_token(mine_by_row[4].cells[1]) != (
        benchmark._physical_cell_token(theirs_by_row[31].cells[1])
    )

    # Zero, empty string, and a formula-only key are owners; only a cell with
    # both cached/formula values absent is a continuation key part.
    scalar = _snapshot("scalar", [
        ([_cell(0)], [], _cell("zero")),
        ([_cell("")], [], _cell("empty")),
        ([_cell(None, formula="=ROW()")], [], _cell("formula")),
    ])
    _scalar_rows, _scalar_order, scalar_topology = _ids(scalar)
    assert scalar_topology["owner_count"] == 3
    assert scalar_topology["continuation_count"] == 0

    # Every malformed key/group topology fails closed.
    leading = _snapshot("leading", [([_cell()], [], _cell("bad")), ([ _cell(1)], [], _cell("ok"))])
    _expect_assertion("leading blank", lambda: _ids(leading))
    duplicate = _snapshot("duplicate", [([_cell(1)], [], _cell("a")), ([_cell(1)], [], _cell("b"))])
    _expect_assertion("duplicate owner", lambda: _ids(duplicate))
    partial = _snapshot(
        "partial", [([_cell(1), _cell()], [], _cell("bad"))], markers=("id", "const")
    )
    _expect_assertion("partial compound key", lambda: _ids(partial))
    missing = replace(mine, rows=(mine.rows[0], mine.rows[1], _row(3, [])))
    _expect_assertion("missing key cell", lambda: _ids(missing))

    typed_mismatch = _snapshot(
        "typed", [([_cell("0")], [_cell("typed-cont")], _cell("same-repr"))]
    )
    _expect_assertion("typed key mismatch", lambda: _pair(mine, typed_mismatch))
    short_group = _snapshot("short", [([_cell(0)], [], _cell("mine"))])
    _expect_assertion("group ordinal mismatch", lambda: _pair(mine, short_group))
    reorder = _snapshot("reorder", [([_cell(2)], [], _cell("b")), ([_cell(1)], [], _cell("a"))])
    normal_order = _snapshot("normal", [([_cell(1)], [], _cell("a")), ([_cell(2)], [], _cell("b"))])
    _expect_assertion("owner reorder", lambda: _pair(normal_order, reorder))
    coverage = _snapshot("coverage", [([_cell(3)], [], _cell("c"))])
    _expect_assertion("owner coverage", lambda: _pair(normal_order, coverage))
    cross_group = _snapshot("cross", [([_cell(0)], [], _cell("a")), ([_cell(1)], [_cell("b-cont")], _cell("b"))])
    _expect_assertion("cross-group continuation", lambda: _pair(mine, cross_group))
    _expect_assertion("Base topology mismatch", lambda: _pair(mine, mine, base=short_group))
    length_two = _snapshot(
        "length-two", [([_cell(0)], [_cell("one"), _cell("two")], _cell("owner"))]
    )
    _expect_assertion("group length greater than one", lambda: _pair(mine, length_two))

    # The separate physical contract catches header/schema/row-envelope drift
    # without invoking a workbook or production alignment path.
    contract = _snapshot("contract", [([_cell(1)], [], _cell("x"))])
    schema = tuple(
        (field.physical_col, field.declaration, field.type_declaration)
        for field in contract.fields
    )
    benchmark._monster_validate_snapshot_contract(
        contract, max_row=3, max_col=2, schema=schema
    )
    _expect_assertion(
        "rowcount drift",
        lambda: benchmark._monster_validate_snapshot_contract(
            replace(contract, max_row=4), max_row=3, max_col=2, schema=schema
        ),
    )
    _expect_assertion(
        "header/schema drift",
        lambda: benchmark._monster_validate_snapshot_contract(
            replace(contract, fields=(
                sm.SnapshotField(1, "changed@id", "int32", frozenset({"id"})),
                contract.fields[1],
            )),
            max_row=3, max_col=2, schema=schema,
        ),
    )
    _expect_assertion(
        "header physical drift",
        lambda: benchmark._monster_validate_snapshot_contract(
            replace(contract, rows=(_row(0, contract.rows[0].cells), *contract.rows[1:])),
            max_row=3, max_col=2, schema=schema,
        ),
    )
    _assert_full_oracle_contract()


def main():
    # The grouper must remain independent of every production compare/align/
    # cache callable.  The synthetic exercise below will explode on any call.
    originals = {}
    def _blocked(name):
        def _raise(*_args, **_kwargs):
            raise AssertionError(f"independent Monster grouper touched {name}")
        return _raise
    for name in dir(sm):
        value = getattr(sm, name)
        if callable(value) and (
            name.startswith("_compare")
            or name.startswith("_align")
            or "cache" in name.lower()
        ):
            originals[name] = value
            setattr(sm, name, _blocked(name))
    try:
        _exercise_contracts()
    finally:
        for name, value in originals.items():
            setattr(sm, name, value)
    print("MONSTER_INDEPENDENT_ORACLE PASS cases=22")


if __name__ == "__main__":
    main()
