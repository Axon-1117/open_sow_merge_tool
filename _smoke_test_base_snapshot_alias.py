"""Deterministic immutable Base-snapshot alias contract.

This is intentionally Tk/workbook/child-free.  It supplies fake paired stream
snapshots backed by disposable byte files, so planner freshness proofs, the
distinct Base wrapper, pickle validation, comparison/adapter parity, and every
fail-closed fallback can be exercised without opening Excel files.
"""

from __future__ import annotations

import hashlib
import os
import pickle
import tempfile

import sow_merge_tool as sm


SHEET = "Alias@design"
GENERATION = 17


def _cell(value, *, cached_type="s", formula_value=None):
    formula_type = "f" if formula_value is not None else cached_type
    return sm.SnapshotCell(
        value, cached_type, formula_value, formula_type,
        "formula" if formula_value is not None else "literal", False,
    )


_BLANK = _cell(None, cached_type="n")


def _row(physical_row: int, cells) -> sm.SnapshotRow:
    payload = tuple(
        (
            cell.cached_value, cell.cached_type, cell.formula_value,
            cell.formula_type, cell.formula_kind, cell.external_link,
        )
        for cell in cells
    )
    return sm.SnapshotRow(
        int(physical_row), tuple(cells),
        hashlib.sha256(repr(payload).encode("utf-8")).hexdigest(),
    )


def _snapshot(side: str, spec: dict, *, mine: bool = False) -> sm.SheetSnapshot:
    fields = (
        sm.SnapshotField(1, "id@id", "int32", frozenset(("id",))),
        sm.SnapshotField(2, "", "", frozenset()),
        sm.SnapshotField(3, "stat", "string", frozenset()),
        sm.SnapshotField(4, "", "", frozenset()),
    )
    headers = (_cell("id@id"), _BLANK, _cell("stat"), _BLANK)
    types = (_cell("int32"), _BLANK, _cell("string"), _BLANK)
    rows = [_row(1, headers), _row(2, types)]
    for physical_row, owner in enumerate((101, 102, 103), start=3):
        value = f"mine-{owner}" if mine and owner == 102 else f"base-{owner}"
        formula = "=10+2" if mine and owner == 102 else "=6+6"
        rows.append(_row(
            physical_row,
            (_cell(owner, cached_type="n"), _BLANK, _cell(value, formula_value=formula), _BLANK),
        ))
    return sm.SheetSnapshot(
        str(side), SHEET,
        sm._selected_sheet_snapshot_version(
            str(spec["value_path"]),
            topology_generation=int(spec.get("topology_generation", GENERATION)),
            mutation_generation=int(spec.get("mutation_generation", 0)),
        ),
        len(rows), len(fields), fields, tuple(rows),
    )


def _write(path: str, payload: bytes) -> str:
    with open(path, "wb") as stream:
        stream.write(payload)
    return path


def _spec(value_path: str, formula_path: str, **overrides) -> dict:
    result = {
        "value_path": value_path,
        "formula_path": formula_path,
        "sheet": SHEET,
        "generation": GENERATION,
        "topology_generation": GENERATION,
        "mutation_generation": 0,
        "parser": 1,
        "schema": "snapshot-field-declarations-v1",
        "loader": sm._SNAPSHOT_ALIAS_LOADER,
        "loader_options": sm._SNAPSHOT_ALIAS_OPTIONS,
    }
    result.update(overrides)
    return result


def _fixture(*, base_source: str = "B"):
    temp = tempfile.TemporaryDirectory(prefix="sow_base_alias_")
    root = temp.name
    a_value = _write(os.path.join(root, "a.value"), b"A-values-v1")
    a_formula = _write(os.path.join(root, "a.formula"), b"A-formulas-v1")
    b_value = _write(os.path.join(root, "b.value"), b"B-values-v1")
    b_formula = _write(os.path.join(root, "b.formula"), b"B-formulas-v1")
    source_value, source_formula = (b_value, b_formula) if base_source == "B" else (a_value, a_formula)
    base_value = _write(os.path.join(root, "base.value"), open(source_value, "rb").read())
    base_formula = _write(os.path.join(root, "base.formula"), open(source_formula, "rb").read())
    inputs = {
        "A": _spec(a_value, a_formula),
        "B": _spec(b_value, b_formula),
        "BASE": _spec(base_value, base_formula),
    }
    snapshots = {
        "A": _snapshot("A", inputs["A"], mine=True),
        "B": _snapshot("B", inputs["B"]),
        "BASE": _snapshot("BASE", inputs["BASE"]),
    }
    return temp, inputs, snapshots


def _reader(snapshots, calls, *, mutate_on_b=None):
    def _read(_value_path, _formula_path, _sheet, side, **_kwargs):
        calls.append(str(side))
        if side == "B" and mutate_on_b is not None:
            mutate_on_b()
        return snapshots[str(side)]
    return _read


def _normalized_result(result):
    return {
        "row_pairs": tuple(result.row_pairs),
        "base_rows": tuple(result.base_rows_by_pair),
        "pair_diffs": tuple(result.pair_diff_cols),
        "base_diffs": tuple(result.pair_base_diff_cols),
        "conflicts": tuple(result.conflict_cols),
        "unresolved": bool(result.unresolved),
        "slots": tuple(
            (slot.logical_idx, slot.mine_col, slot.base_col, slot.theirs_col, slot.state)
            for slot in result.column_cache.model.slots
        ),
    }


def _normalized_cache(cache):
    return {
        key: cache[key]
        for key in (
            "row_pairs", "pair_diff_cols", "pair_base_diff_cols",
            "mine_to_base_row", "theirs_to_base_row", "pair_base_row_override",
            "only_diff_rows", "has_diff", "prepared_complete",
        )
    }


def _test_alias_preference_wrapper_pickle_and_threeway_parity():
    temp, inputs, snapshots = _fixture(base_source="B")
    try:
        calls = []
        read, _metrics, alias, _before = sm._read_snapshot_inputs_with_base_alias(
            inputs, sheet=SHEET, generation=GENERATION,
            stream_snapshot=_reader(snapshots, calls),
        )
        assert calls == ["A", "B"], calls
        assert alias["used"] and alias["source"] == "B" and alias["saved_side"] == "BASE"
        assert read["BASE"] is not read["B"]
        assert read["BASE"].side == "BASE"
        assert read["BASE"].version == sm._selected_sheet_snapshot_version(
            inputs["BASE"]["value_path"], topology_generation=GENERATION,
        )
        assert read["BASE"].rows is read["B"].rows and read["BASE"].fields is read["B"].fields
        restored = pickle.loads(pickle.dumps(read["BASE"], protocol=pickle.HIGHEST_PROTOCOL))
        assert restored is not read["BASE"] and restored.side == "BASE"

        baseline = sm._compare_selected_sheet_snapshots(
            snapshots["A"], snapshots["B"], snapshots["BASE"],
        )
        aliased = sm._compare_selected_sheet_snapshots(read["A"], read["B"], read["BASE"])
        assert not baseline.unresolved and not aliased.unresolved
        assert _normalized_result(aliased) == _normalized_result(baseline)
        # The duplicate blank fields at physical 2/4 must remain proven rather
        # than disappearing behind aliasing, while the ordinary stat diff stays.
        assert any(len(group) > 1 for group in sm._snapshot_field_groups(read["BASE"]).values())
        assert not aliased.column_cache.unresolved_cols
        assert any(3 in cols for cols in aliased.pair_diff_cols)
        baseline_cache = sm._snapshot_result_to_sheet_cache_immutable(
            SHEET, baseline, snapshots["A"], snapshots["B"], snapshots["BASE"], has_base=True,
        )
        alias_cache = sm._snapshot_result_to_sheet_cache_immutable(
            SHEET, aliased, read["A"], read["B"], read["BASE"], has_base=True,
        )
        assert baseline_cache["prepared_complete"] and alias_cache["prepared_complete"]
        assert _normalized_cache(alias_cache) == _normalized_cache(baseline_cache)
        assert alias_cache["pair_base_row_override"] == {index: index + 1 for index in range(5)}
    finally:
        temp.cleanup()


def _test_base_source_preference_and_all_same():
    for base_source, expected in (("A", "A"), ("B", "B")):
        temp, inputs, snapshots = _fixture(base_source=base_source)
        try:
            calls = []
            _read, _metrics, alias, _before = sm._read_snapshot_inputs_with_base_alias(
                inputs, sheet=SHEET, generation=GENERATION,
                stream_snapshot=_reader(snapshots, calls),
            )
            assert alias["used"] and alias["source"] == expected
            assert calls == ["A", "B"]
        finally:
            temp.cleanup()

    temp, inputs, snapshots = _fixture(base_source="B")
    try:
        # Make all three raw pairs identical; B must remain the deterministic
        # candidate even though A also matches.
        for key in ("value_path", "formula_path"):
            with open(inputs["A"][key], "wb") as stream:
                stream.write(open(inputs["B"][key], "rb").read())
        calls = []
        _read, _metrics, alias, _before = sm._read_snapshot_inputs_with_base_alias(
            inputs, sheet=SHEET, generation=GENERATION,
            stream_snapshot=_reader(snapshots, calls),
        )
        assert alias["used"] and alias["source"] == "B"
        assert calls == ["A", "B"]
    finally:
        temp.cleanup()


def _assert_full_base_stream(inputs, snapshots):
    calls = []
    _read, _metrics, alias, _before = sm._read_snapshot_inputs_with_base_alias(
        inputs, sheet=SHEET, generation=GENERATION,
        stream_snapshot=_reader(snapshots, calls),
    )
    assert calls == ["A", "B", "BASE"], calls
    assert not alias["used"]


def _test_byte_and_semantic_mismatches_stream_all_three_sides():
    for raw_key, payload in (("value_path", b"different-base-value"), ("formula_path", b"different-base-formula")):
        temp, inputs, snapshots = _fixture(base_source="B")
        try:
            with open(inputs["BASE"][raw_key], "wb") as stream:
                stream.write(payload)
            _assert_full_base_stream(inputs, snapshots)
        finally:
            temp.cleanup()

    for key, value in (
        ("sheet", "Other@design"),
        ("parser", 2),
        ("schema", "different-schema"),
        ("generation", GENERATION + 1),
        ("topology_generation", GENERATION + 1),
        ("mutation_generation", 1),
        ("loader", "different-loader"),
        ("loader_options", ("different-option",)),
    ):
        temp, inputs, snapshots = _fixture(base_source="B")
        try:
            inputs["BASE"][key] = value
            _assert_full_base_stream(inputs, snapshots)
        finally:
            temp.cleanup()


def _test_pickle_and_builder_exceptions_fall_back_to_full_base_stream():
    temp, inputs, snapshots = _fixture(base_source="B")
    try:
        plan = sm._plan_snapshot_base_alias(inputs, sheet=SHEET, generation=GENERATION)
        wrapper, telemetry = sm._build_snapshot_base_alias(
            plan, snapshots["B"], inputs,
            pickle_dumps=lambda *_args, **_kwargs: (_ for _ in ()).throw(pickle.PicklingError("injected")),
        )
        assert wrapper is None and not telemetry["used"]
        assert "PicklingError" in telemetry["reason"]

        original = sm._build_snapshot_base_alias
        def _raise(*_args, **_kwargs):
            raise RuntimeError("injected builder failure")
        sm._build_snapshot_base_alias = _raise
        try:
            _assert_full_base_stream(inputs, snapshots)
        finally:
            sm._build_snapshot_base_alias = original
    finally:
        temp.cleanup()


def _test_stale_base_path_is_blocked_after_alias_plan():
    temp, inputs, snapshots = _fixture(base_source="B")
    try:
        calls = []
        def _mutate_base():
            with open(inputs["BASE"]["value_path"], "ab") as stream:
                stream.write(b"-changed-after-plan")
        try:
            sm._read_snapshot_inputs_with_base_alias(
                inputs, sheet=SHEET, generation=GENERATION,
                stream_snapshot=_reader(snapshots, calls, mutate_on_b=_mutate_base),
            )
        except RuntimeError as exc:
            assert "input changed during read" in str(exc)
        else:
            raise AssertionError("stale Base path must not publish an alias result")
        assert calls == ["A", "B", "BASE"], calls
    finally:
        temp.cleanup()


def _test_mid_read_a_replacement_blocks_alias_and_full_base_streams():
    for alias_eligible in (True, False):
        temp, inputs, snapshots = _fixture(base_source="B")
        try:
            if not alias_eligible:
                with open(inputs["BASE"]["formula_path"], "wb") as stream:
                    stream.write(b"non-alias-base-formula")
            calls = []
            def _replace_a_after_a_read():
                # The callback runs while B is being requested, so A has
                # already yielded a candidate snapshot.  The shared reader's
                # all-side after proof must reject both alias and 3-stream
                # paths instead of letting that stale A publish.
                with open(inputs["A"]["value_path"], "ab") as stream:
                    stream.write(b"-replaced-between-a-and-b")
            try:
                sm._read_snapshot_inputs_with_base_alias(
                    inputs, sheet=SHEET, generation=GENERATION,
                    stream_snapshot=_reader(snapshots, calls, mutate_on_b=_replace_a_after_a_read),
                )
            except RuntimeError as exc:
                assert "input changed during read" in str(exc)
            else:
                raise AssertionError("mid-read A replacement must never publish a cache")
            expected_calls = ["A", "B"] if alias_eligible else ["A", "B", "BASE"]
            assert calls == expected_calls, (alias_eligible, calls)
        finally:
            temp.cleanup()


def main():
    _test_alias_preference_wrapper_pickle_and_threeway_parity()
    _test_base_source_preference_and_all_same()
    _test_byte_and_semantic_mismatches_stream_all_three_sides()
    _test_pickle_and_builder_exceptions_fall_back_to_full_base_stream()
    _test_stale_base_path_is_blocked_after_alias_plan()
    _test_mid_read_a_replacement_blocks_alias_and_full_base_streams()
    print("base snapshot alias contract: PASS")


if __name__ == "__main__":
    main()
