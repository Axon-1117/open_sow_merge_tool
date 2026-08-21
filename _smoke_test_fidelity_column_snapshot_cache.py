"""Focused case-local immutable snapshot-cache gate regression."""

from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path
from types import SimpleNamespace

import _large_sheet_excel_fidelity_gate as gate
import sow_merge_tool as sm


_CASE = "fidelity-column-snapshot-cache"
_SHEET = "S1"


def _manifest(*, three_way: bool, blocked: bool) -> dict:
    return {
        "sheet": _SHEET,
        "three_way": three_way,
        "columns": ([
            {
                "logical": 1,
                "mine": None,
                "base": None,
                "theirs": 1,
                "state": "unresolved",
                "ambiguous": True,
            }
        ] if blocked else []),
        "records": [],
        "only_diff_rows": [],
    }


def _run_case() -> None:
    temporary = tempfile.TemporaryDirectory(prefix="sow_fidelity_column_snapshot_cache_")
    root = Path(temporary.name)
    primary: BaseException | None = None
    cleanup_errors: list[str] = []
    original_stream = sm._stream_selected_sheet_snapshot
    original_compare = sm._compare_selected_sheet_snapshots
    original_manifest = sm.snapshot_comparison_oracle_manifest
    original_direct = gate.capture_direct_legacy
    original_frozen = gate.capture_legacy
    try:
        base = root / "base.xlsx"
        mine = root / "mine.xlsx"
        theirs = root / "theirs.xlsx"
        mine_conflict = root / "mine-conflict.xlsx"
        for path in (base, mine, theirs, mine_conflict):
            path.write_bytes(path.name.encode("utf-8"))

        stream_calls: list[tuple[str, str, str]] = []
        snapshot_paths: dict[int, str] = {}
        direct_calls: list[tuple[str, str, str]] = []
        frozen_calls: list[tuple[str, str, str, str]] = []
        manifest_modes: list[tuple[bool, bool]] = []

        def canonical(path: Path | str) -> str:
            return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))

        def stream(value_path, formula_path, sheet, side, **_kwargs):
            assert str(value_path) == str(formula_path)
            path = canonical(value_path)
            stream_calls.append((path, str(sheet), str(side)))
            snapshot = sm.SheetSnapshot(
                side=str(side),
                sheet=str(sheet),
                version=sm._selected_sheet_snapshot_version(str(value_path)),
                max_row=1,
                max_col=1,
                fields=(),
                rows=(),
            )
            snapshot_paths[id(snapshot)] = path
            return snapshot

        def compare(mine_snapshot, _theirs_snapshot, _base_snapshot=None):
            return SimpleNamespace(
                unresolved=snapshot_paths[id(mine_snapshot)] == canonical(mine_conflict)
            )

        def snapshot_manifest(mine_snapshot, _theirs_snapshot, result, base_snapshot=None):
            blocked = bool(result.unresolved)
            three_way = base_snapshot is not None
            manifest_modes.append((three_way, blocked))
            return _manifest(three_way=three_way, blocked=blocked)

        def direct_capture(mine_path, theirs_path, sheet):
            direct_calls.append((canonical(mine_path), canonical(theirs_path), str(sheet)))
            return _manifest(three_way=False, blocked=False)

        def frozen_capture(args):
            conflict = canonical(args.mine) == canonical(mine_conflict)
            frozen_calls.append((
                canonical(args.mine),
                canonical(args.theirs),
                canonical(args.base),
                str(args.sheet),
            ))
            return _manifest(three_way=True, blocked=conflict)

        sm._stream_selected_sheet_snapshot = stream
        sm._compare_selected_sheet_snapshots = compare
        sm.snapshot_comparison_oracle_manifest = snapshot_manifest
        gate.capture_direct_legacy = direct_capture
        gate.capture_legacy = frozen_capture

        cache: dict[tuple[str, str, str], sm.SheetSnapshot] = {}
        two_way = gate._assert_direct_pair_parity(
            base,
            theirs,
            _SHEET,
            "cache-2way",
            snapshot_cache=cache,
        )
        normal_legacy, normal_candidate = gate._assert_frozen_three_way_parity(
            mine,
            theirs,
            base,
            _SHEET,
            "cache-normal-3way",
            timeout=90.0,
            snapshot_cache=cache,
        )
        gate._assert_dual_column_conflict_blocked(
            mine_conflict,
            theirs,
            base,
            _SHEET,
            timeout=90.0,
            snapshot_cache=cache,
        )

        assert two_way == _manifest(three_way=False, blocked=False)
        assert normal_legacy == normal_candidate == _manifest(three_way=True, blocked=False)
        assert direct_calls == [(canonical(base), canonical(theirs), _SHEET)]
        assert frozen_calls == [
            (canonical(mine), canonical(theirs), canonical(base), _SHEET),
            (canonical(mine_conflict), canonical(theirs), canonical(base), _SHEET),
        ]
        expected_streams = [
            (canonical(base), _SHEET, "A"),
            (canonical(theirs), _SHEET, "B"),
            (canonical(mine), _SHEET, "A"),
            (canonical(base), _SHEET, "BASE"),
            (canonical(mine_conflict), _SHEET, "A"),
        ]
        assert stream_calls == expected_streams
        assert tuple(cache) == tuple(gate._snapshot_cache_key(*item) for item in expected_streams)
        assert manifest_modes == [(False, False), (True, False), (True, True)]

        base.write_bytes(b"post-cache rewrite")
        try:
            gate._snapshot_manifest(base, theirs, None, _SHEET, snapshot_cache=cache)
        except AssertionError as exc:
            assert "identity/version mismatch" in str(exc)
        else:
            raise AssertionError("post-cache rewrite must fail closed before restream")
        assert stream_calls == expected_streams
    except BaseException as exc:
        primary = exc
        raise
    finally:
        sm._stream_selected_sheet_snapshot = original_stream
        sm._compare_selected_sheet_snapshots = original_compare
        sm.snapshot_comparison_oracle_manifest = original_manifest
        gate.capture_direct_legacy = original_direct
        gate.capture_legacy = original_frozen
        try:
            temporary.cleanup()
        except BaseException as exc:
            cleanup_errors.append(f"temporary cleanup: {type(exc).__name__}: {exc}")
        if os.path.lexists(root):
            cleanup_errors.append(f"owned root remains: {root}")
        if cleanup_errors:
            detail = "; ".join(cleanup_errors)
            if primary is not None:
                primary.add_note(detail)
            else:
                raise AssertionError(detail)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case")
    parser.add_argument("--list-cases", action="store_true")
    args = parser.parse_args()
    if args.list_cases:
        print(_CASE)
        return
    selected = args.case or _CASE
    if selected != _CASE:
        raise SystemExit(f"unknown case: {selected}")
    _run_case()
    print(f"SMOKE_FIDELITY_COLUMN_SNAPSHOT_CACHE_OK {selected}")


if __name__ == "__main__":
    main()
