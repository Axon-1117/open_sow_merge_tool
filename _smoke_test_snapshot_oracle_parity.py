"""Disposable exact parity smoke test for immutable selected-Sheet snapshots."""

from __future__ import annotations

import tempfile
from pathlib import Path

import sow_merge_tool as sm
from _large_sheet_oracle_fixtures import build_adversarial_fixture_set
from _large_sheet_snapshot_oracle import capture_legacy, compare_manifests


_legacy_call_count = 0


class _Args:
    pass


def _legacy(item):
    global _legacy_call_count
    _legacy_call_count += 1
    args = _Args()
    args.mine = item["mine"]
    args.theirs = item["theirs"]
    args.base = item.get("base")
    args.sheet = item["sheet"]
    args.timeout = 45
    return capture_legacy(args)


def _candidate(item):
    sheet = item["sheet"]
    mine = sm._stream_selected_sheet_snapshot(item["mine"], item["mine"], sheet, "A")
    theirs = sm._stream_selected_sheet_snapshot(item["theirs"], item["theirs"], sheet, "B")
    base = (
        sm._stream_selected_sheet_snapshot(item["base"], item["base"], sheet, "BASE")
        if item.get("base") else None
    )
    result = sm._compare_selected_sheet_snapshots(mine, theirs, base)
    return (
        mine,
        theirs,
        base,
        result,
        sm.snapshot_comparison_oracle_manifest(mine, theirs, result, base),
    )


def main():
    # The fixture root is disposable even if a parity assertion raises.
    with tempfile.TemporaryDirectory(prefix="sow_snapshot_oracle_") as temporary:
        cases = build_adversarial_fixture_set(Path(temporary))
        exact_cases = (
            "composite_key", "blank_continuation", "equal_count_insert_delete",
            "reorder", "formula_cache", "three_way_conflict", "column_structure",
        )
        for name in exact_cases:
            _mine, _theirs, _base, result, candidate = _candidate(cases[name])
            assert not result.unresolved, name
            parity = compare_manifests(_legacy(cases[name]), candidate)
            assert parity["exact"], (name, parity["mismatches"])

        # Ambiguous identity is a terminal safety gate, not an exact-result
        # parity case.  The legacy capture intentionally requires an exact
        # READY view and must never be invoked for this known UNRESOLVED case.
        mine, theirs, base, result, candidate = _candidate(cases["duplicate_missing_keys"])
        calls_before_negative = _legacy_call_count
        assert result.unresolved
        assert candidate["unresolved"] is True
        terminal_cache = sm._snapshot_result_to_sheet_cache_immutable(
            str(cases["duplicate_missing_keys"]["sheet"]),
            result,
            mine,
            theirs,
            base,
            has_base=False,
        )
        assert terminal_cache["prepared_complete"] is False
        assert terminal_cache["has_diff"] is True
        assert terminal_cache["completeness"]["mode"] == "snapshot-unresolved"
        assert terminal_cache["completeness"]["ab_diff_exact"] is False
        assert terminal_cache["unresolved_reason"]
        # The unresolved adapter deliberately publishes no row projection,
        # diff map, or operation target that a caller could act on.
        assert set(terminal_cache) == {
            "sheet", "snapshot_engine", "unresolved_reason",
            "prepared_complete", "has_diff", "completeness",
        }
        assert _legacy_call_count == calls_before_negative
    print("SNAPSHOT_ORACLE_PARITY_OK")


if __name__ == "__main__":
    main()
