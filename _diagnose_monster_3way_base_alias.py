"""Read-only real MonsterGroup 3-way Base-alias diagnostic.

This diagnostic constructs no Tk application, editable workbook, legacy
fallback, or workbook mutation.  It drives the production immutable
``_read_snapshot_inputs_with_base_alias`` entry twice against the real revision
exports: once with Base aliasing enabled and once with the planner disabled to
obtain an independent three-stream baseline.  Its only write is one atomically
replaced JSON report outside every protected input.
"""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import json
import os
import tempfile
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import sow_merge_tool as sm


DEFAULT_REVISION_DIR = Path(r"C:\Users\dd\AppData\Local\Temp\sow_revision_export_39264_39265")
DEFAULT_SHEET = "MonsterGroup@design"
EXPECTED_ROWS = 3169
EXPECTED_COLS = 31
PROOF_PHYSICAL_COLS = (27, 31)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _provenance(path: Path) -> dict[str, Any]:
    return {
        "path": str(path.resolve()),
        "exists": path.exists(),
        "size_bytes": path.stat().st_size if path.exists() else None,
        "sha256": _sha256(path) if path.exists() else None,
    }


def _json(value: Any) -> Any:
    """Convert every report value to deliberate, JSON-safe evidence.

    Dataclasses are expanded by fields rather than repr() so a Base wrapper
    version remains comparable evidence if a later error path serializes a
    partially assembled payload.  The final fallback is reserved for unknown
    diagnostic-only objects; it never handles versions, paths, bytes, or
    containers used by this report.
    """
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Path):
        return str(value.resolve())
    if isinstance(value, bytes):
        return {"encoding": "hex", "data": value.hex()}
    if dataclasses.is_dataclass(value) and not isinstance(value, type):
        return {
            field.name: _json(getattr(value, field.name))
            for field in dataclasses.fields(value)
        }
    if isinstance(value, dict):
        return {str(key): _json(item) for key, item in value.items()}
    if isinstance(value, (tuple, list)):
        return [_json(item) for item in value]
    if isinstance(value, (set, frozenset)):
        normalized = [_json(item) for item in value]
        return sorted(
            normalized,
            key=lambda item: json.dumps(item, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
        )
    return {"type": type(value).__name__, "repr": repr(value)}


def _canonical_digest(value: Any) -> str:
    encoded = json.dumps(_json(value), ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest().upper()


def _version_facts(version: sm.SheetSnapshotVersion) -> dict[str, int]:
    """Persist a snapshot version as stable fields, never a repr()."""
    return {
        "parser": int(version.parser),
        "topology_generation": int(version.topology_generation),
        "mutation_generation": int(version.mutation_generation),
        "file_size": int(version.file_size),
        "file_mtime_ns": int(version.file_mtime_ns),
    }


def _signature_facts(signature: Any) -> list[dict[str, Any]] | None:
    """Make the paired value/formula raw signature directly comparable."""
    if signature is None:
        return None
    pairs = tuple(signature)
    if len(pairs) != 2:
        raise ValueError("Base alias input signature must contain value/formula pairs")
    result = []
    for size, digest in pairs:
        result.append({"size_bytes": int(size), "sha256": str(digest)})
    return result


def _atomic_json(path: Path, payload: dict[str, Any]) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_name(path.name + ".partial")
    # Normalize at the one persistence boundary as well as at individual
    # evidence fields.  This keeps a caught comparison/adapter exception
    # reportable even if it carries an unexpected dataclass or path object.
    encoded = json.dumps(_json(payload), ensure_ascii=False, indent=2, sort_keys=True)
    partial.write_text(encoded, encoding="utf-8")
    os.replace(partial, path)


def _is_within(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
    except ValueError:
        return False
    return True


def _validate_output_path(
    raw_output: Path,
    *,
    revision_dir: Path,
    protected_paths: tuple[Path, ...],
) -> Path:
    raw = Path(raw_output)
    if not raw.is_absolute() or raw.suffix.casefold() != ".json":
        raise ValueError("--output must be an absolute .json path")
    output = raw.resolve()
    partial = output.with_name(output.name + ".partial")
    protected = tuple(path.resolve() for path in protected_paths)
    if output in protected or partial in protected:
        raise ValueError("--output may not overwrite a source, diagnostic, or revision input")
    forbidden_roots = {
        revision_dir.resolve(),
        *(path.parent.resolve() for path in protected_paths if path.suffix.casefold() in {".xlsx", ".xlsm"}),
    }
    if any(_is_within(candidate, root) for candidate in (output, partial) for root in forbidden_roots):
        raise ValueError("--output may not be inside a revision export or input directory")
    workspace_reports = Path(__file__).resolve().parent / "benchmark_results"
    temp_root = Path(tempfile.gettempdir()).resolve()
    if not any(_is_within(output, root.resolve()) for root in (workspace_reports, temp_root)):
        raise ValueError("--output must be beneath benchmark_results or the system temporary directory")
    return output


def _spec(path: Path, *, sheet: str, generation: int) -> dict[str, Any]:
    return {
        "value_path": str(path),
        "formula_path": str(path),
        "sheet": str(sheet),
        "generation": int(generation),
        "topology_generation": int(generation),
        "mutation_generation": 0,
        "parser": 1,
        "schema": "snapshot-field-declarations-v1",
        "loader": sm._SNAPSHOT_ALIAS_LOADER,
        "loader_options": sm._SNAPSHOT_ALIAS_OPTIONS,
    }


def _snapshot_inputs(mine: Path, base: Path, theirs: Path, *, sheet: str, generation: int) -> dict[str, dict[str, Any]]:
    return {
        "A": _spec(mine, sheet=sheet, generation=generation),
        "B": _spec(theirs, sheet=sheet, generation=generation),
        "BASE": _spec(base, sheet=sheet, generation=generation),
    }


def _deadline_checker(deadline: float):
    def _check():
        if time.monotonic() > deadline:
            raise TimeoutError("base-alias diagnostic budget exceeded")
    return _check


def _counting_stream(counter: dict[str, int], deadline_check):
    def _stream(value_path, formula_path, sheet, side, **kwargs):
        counter[str(side)] = int(counter.get(str(side), 0)) + 1
        upstream_check = kwargs.get("cancel_check")

        def _check():
            deadline_check()
            if callable(upstream_check):
                upstream_check()

        return sm._stream_selected_sheet_snapshot(
            value_path, formula_path, sheet, side,
            topology_generation=int(kwargs.get("topology_generation", 0)),
            mutation_generation=int(kwargs.get("mutation_generation", 0)),
            cancel_check=_check,
        )
    return _stream


def _run_reader(
    inputs: dict[str, dict[str, Any]],
    *,
    sheet: str,
    generation: int,
    deadline_check,
    disable_alias: bool,
) -> tuple[dict[str, sm.SheetSnapshot], dict[str, float], dict[str, Any], tuple, dict[str, int]]:
    """Invoke the production reader and optionally force its full Base stream."""
    counter: dict[str, int] = {}
    original_plan = sm._plan_snapshot_base_alias
    if disable_alias:
        def _disabled_plan(*args, **kwargs):
            plan = original_plan(*args, **kwargs)
            return dataclasses.replace(
                plan, source_side=None, reason="diagnostic-forced-full-base-stream",
            )
        sm._plan_snapshot_base_alias = _disabled_plan
    try:
        result = sm._read_snapshot_inputs_with_base_alias(
            inputs,
            sheet=str(sheet),
            generation=int(generation),
            cancel_check=deadline_check,
            stream_snapshot=_counting_stream(counter, deadline_check),
        )
    finally:
        sm._plan_snapshot_base_alias = original_plan
    snapshots, timings, alias, before = result
    return snapshots, timings, alias, before, counter


def _slot_triple_summary(cache: sm.LogicalColumnComparisonCache, physical_col: int) -> dict[str, Any]:
    matches = [
        slot for slot in cache.model.slots
        if (
            int(slot.mine_col or -1) == int(physical_col)
            and int(slot.base_col or -1) == int(physical_col)
            and int(slot.theirs_col or -1) == int(physical_col)
        )
    ]
    if len(matches) != 1:
        raise AssertionError(f"proof physical column {physical_col} has {len(matches)} top-cache triples")
    slot = matches[0]
    if slot.state != "retained" or slot.confidence.ambiguous:
        raise AssertionError(f"proof physical column {physical_col} is not retained/exact")
    return {
        "physical_col1": int(physical_col),
        "logical_col1": int(slot.logical_idx) + 1,
        "mine_col1": int(slot.mine_col),
        "base_col1": int(slot.base_col),
        "theirs_col1": int(slot.theirs_col),
        "state": str(slot.state),
        "score": float(slot.confidence.score),
        "cause_codes": list(slot.confidence.cause_codes),
    }


def _cache_top_summary(cache: sm.LogicalColumnComparisonCache) -> dict[str, Any]:
    slots = tuple(cache.model.slots)
    alignment = cache.three_way_alignment
    required = tuple(range(1, EXPECTED_COLS + 1))
    side_bijection = {
        side: tuple(sorted(int(getattr(slot, side)) for slot in slots if getattr(slot, side) is not None)) == required
        for side in ("mine_col", "base_col", "theirs_col")
    }
    return {
        "slot_count": len(slots),
        "model_ambiguous": bool(cache.model.confidence.ambiguous),
        "model_reason": str(cache.model.confidence.reason),
        "model_score": float(cache.model.confidence.score),
        "model_cause_codes": list(cache.model.confidence.cause_codes),
        "unresolved_cols": sorted(int(value) for value in cache.unresolved_cols),
        "structural_diff_cols": sorted(int(value) for value in cache.structural_diff_cols),
        "top_used_physical_fallback": bool(getattr(alignment, "used_physical_fallback", False)),
        "top_fallback_slot_indices": [int(value) for value in getattr(alignment, "fallback_slot_indices", ())],
        "all_retained": all(slot.state == "retained" for slot in slots),
        "all_nonambiguous": all(not slot.confidence.ambiguous for slot in slots),
        "side_bijection": side_bijection,
        "proof_triples": [_slot_triple_summary(cache, col) for col in PROOF_PHYSICAL_COLS],
    }


def _normalized_comparison(result: sm.SnapshotComparisonResult, adapter: dict[str, Any]) -> dict[str, Any]:
    cache = result.column_cache
    slot_map = tuple(
        (slot.logical_idx, slot.mine_col, slot.base_col, slot.theirs_col, slot.state)
        for slot in cache.model.slots
    )
    return {
        "row_pairs": tuple(result.row_pairs),
        "base_rows_by_pair": tuple(result.base_rows_by_pair),
        "pair_diff_cols": tuple(tuple(sorted(values)) for values in result.pair_diff_cols),
        "pair_base_diff_cols": tuple(tuple(sorted(values)) for values in result.pair_base_diff_cols),
        "conflict_cols": tuple(tuple(sorted(values)) for values in result.conflict_cols),
        "slot_map": slot_map,
        "adapter_targets": {
            "mine_to_base_row": tuple(sorted((int(key), int(value)) for key, value in (adapter.get("mine_to_base_row") or {}).items())),
            "theirs_to_base_row": tuple(sorted((int(key), int(value)) for key, value in (adapter.get("theirs_to_base_row") or {}).items())),
            "pair_base_row_override": tuple(sorted((int(key), int(value)) for key, value in (adapter.get("pair_base_row_override") or {}).items())),
        },
    }


def _compare_and_adapt(
    sheet: str,
    snapshots: dict[str, sm.SheetSnapshot],
    deadline_check,
) -> tuple[sm.SnapshotComparisonResult, dict[str, Any], dict[str, Any]]:
    deadline_check()
    result = sm._compare_selected_sheet_snapshots(
        snapshots["A"], snapshots["B"], snapshots["BASE"],
    )
    deadline_check()
    adapter = sm._snapshot_result_to_sheet_cache_immutable(
        str(sheet), result, snapshots["A"], snapshots["B"], snapshots["BASE"], has_base=True,
        cancel_check=deadline_check,
    )
    return result, adapter, _normalized_comparison(result, adapter)


def _hard_assert_exact(
    result: sm.SnapshotComparisonResult,
    adapter: dict[str, Any],
    *,
    label: str,
) -> dict[str, Any]:
    if result.unresolved:
        raise AssertionError(f"{label}: comparison remained unresolved")
    if not bool(adapter.get("prepared_complete")):
        raise AssertionError(f"{label}: immutable adapter was not prepared")
    if len(result.row_pairs) != EXPECTED_ROWS or len(result.base_rows_by_pair) != EXPECTED_ROWS:
        raise AssertionError(f"{label}: expected {EXPECTED_ROWS} row/Base pairs")
    if len(result.column_cache.model.slots) != EXPECTED_COLS:
        raise AssertionError(f"{label}: expected {EXPECTED_COLS} top cache slots")
    if any(result.conflict_cols):
        raise AssertionError(f"{label}: Base=Theirs must produce zero conflicts")
    targets = adapter.get("pair_base_row_override") or {}
    if len(targets) != EXPECTED_ROWS:
        raise AssertionError(f"{label}: missing Base physical targets")
    top = _cache_top_summary(result.column_cache)
    if (
        top["model_ambiguous"]
        or top["unresolved_cols"]
        or top["structural_diff_cols"]
        or top["top_used_physical_fallback"]
        or top["top_fallback_slot_indices"]
        or not top["all_retained"]
        or not top["all_nonambiguous"]
        or not all(top["side_bijection"].values())
    ):
        raise AssertionError(f"{label}: top cache did not remain exact/bijective")
    return {
        "result_exact": True,
        "row_pair_count": len(result.row_pairs),
        "base_row_count": len(result.base_rows_by_pair),
        "conflict_pair_count": sum(bool(values) for values in result.conflict_cols),
        "adapter_prepared_complete": bool(adapter.get("prepared_complete")),
        "base_target_count": len(targets),
        "first_base_target": targets.get(0),
        "last_base_target": targets.get(len(result.row_pairs) - 1),
        "top_cache": top,
    }


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--revision-dir", type=Path, default=DEFAULT_REVISION_DIR)
    parser.add_argument("--mine", type=Path)
    parser.add_argument("--base", type=Path)
    parser.add_argument("--theirs", type=Path)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--budget-seconds", type=float, default=25.0)
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    if not (0.1 < float(args.budget_seconds) <= 29.0):
        raise ValueError("--budget-seconds must be between 0.1 and 29 seconds")
    mine_path = args.mine or args.revision_dir / "Dungeon-r39265.xlsx"
    base_path = args.base or args.revision_dir / "Dungeon-r39264.xlsx"
    theirs_path = args.theirs or args.revision_dir / "Dungeon-r39264.xlsx"
    source_path = Path(__file__).with_name("sow_merge_tool.py")
    diagnostic_path = Path(__file__).resolve()
    protected = {
        "source": source_path,
        "diagnostic_script": diagnostic_path,
        "mine": mine_path,
        "base": base_path,
        "theirs": theirs_path,
    }
    default_output = diagnostic_path.parent / "benchmark_results" / (
        "monster_3way_base_alias_diagnostic_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".json"
    )
    output = _validate_output_path(
        args.output or default_output,
        revision_dir=args.revision_dir,
        protected_paths=tuple(protected.values()),
    )
    started = time.monotonic()
    deadline_check = _deadline_checker(started + float(args.budget_seconds))
    payload: dict[str, Any] = {
        "schema": "monster-3way-base-alias-diagnostic-v1",
        "mode": "3way",
        "roles": {
            "mine": {"revision": 39265, "side": "A"},
            "base": {"revision": 39264, "side": "BASE"},
            "theirs": {"revision": 39264, "side": "B"},
            "base_equals_theirs_input": True,
        },
        "sheet": str(args.sheet),
        "read_only": True,
        "tk_created": False,
        "legacy_fallback_allowed": False,
        "before": {name: _provenance(path) for name, path in protected.items()},
    }
    try:
        for name in ("mine", "base", "theirs"):
            if not protected[name].is_file():
                raise FileNotFoundError(f"missing revision input: {protected[name]}")
        inputs = _snapshot_inputs(
            mine_path, base_path, theirs_path,
            sheet=str(args.sheet), generation=0,
        )
        alias_snapshots, alias_timings, alias_telemetry, alias_before, alias_streams = _run_reader(
            inputs, sheet=str(args.sheet), generation=0,
            deadline_check=deadline_check, disable_alias=False,
        )
        alias_after = sm._snapshot_alias_input_fingerprints(inputs)
        if tuple(alias_before) != tuple(alias_after):
            raise AssertionError("alias reader input freshness changed")
        base_alias = alias_snapshots["BASE"]
        theirs_alias = alias_snapshots["B"]
        alias_signatures = alias_telemetry.get("input_signatures") or {}
        base_signature = alias_signatures.get("BASE")
        source_signature = alias_signatures.get("B")
        expected_base_version = sm._selected_sheet_snapshot_version(
            inputs["BASE"]["value_path"], topology_generation=0, mutation_generation=0,
        )
        if (
            not bool(alias_telemetry.get("used"))
            or alias_telemetry.get("source") != "B"
            or alias_telemetry.get("saved_side") != "BASE"
            or alias_streams != {"A": 1, "B": 1}
            or base_alias is theirs_alias
            or base_alias.side != "BASE"
            or base_alias.version != expected_base_version
            or base_alias.rows is not theirs_alias.rows
            or base_alias.fields is not theirs_alias.fields
        ):
            raise AssertionError("Base=B alias contract failed")
        alias_result, alias_adapter, alias_normalized = _compare_and_adapt(
            str(args.sheet), alias_snapshots, deadline_check,
        )
        alias_facts = _hard_assert_exact(alias_result, alias_adapter, label="alias")
        payload["alias"] = {
            "reader_before_fingerprints": alias_before,
            "reader_after_fingerprints": alias_after,
            "stream_counts": alias_streams,
            "actual_stream_count": sum(alias_streams.values()),
            "timings_ms": alias_timings,
            "telemetry": alias_telemetry,
            "wrapper": {
                "is_distinct_from_B": base_alias is not theirs_alias,
                "wrapper_object_id": id(base_alias),
                "source_B_object_id": id(theirs_alias),
                "side": base_alias.side,
                "version": _version_facts(base_alias.version),
                "expected_base_version": _version_facts(expected_base_version),
                "base_value_path": str(Path(inputs["BASE"]["value_path"]).resolve()),
                "base_formula_path": str(Path(inputs["BASE"]["formula_path"]).resolve()),
                "base_input_signature": _signature_facts(base_signature),
                "source_B_input_signature": _signature_facts(source_signature),
                "source_payload_digest": alias_telemetry.get("source_payload_digest"),
                "wrapper_payload_digest": alias_telemetry.get("wrapper_payload_digest"),
                "shares_frozen_rows_with_B": base_alias.rows is theirs_alias.rows,
                "shares_frozen_fields_with_B": base_alias.fields is theirs_alias.fields,
            },
            "facts": alias_facts,
            "mapping_digest": _canonical_digest(alias_normalized),
        }

        baseline_snapshots, baseline_timings, baseline_telemetry, baseline_before, baseline_streams = _run_reader(
            inputs, sheet=str(args.sheet), generation=0,
            deadline_check=deadline_check, disable_alias=True,
        )
        baseline_after = sm._snapshot_alias_input_fingerprints(inputs)
        if tuple(baseline_before) != tuple(baseline_after):
            raise AssertionError("baseline reader input freshness changed")
        if baseline_streams != {"A": 1, "B": 1, "BASE": 1} or baseline_telemetry.get("used"):
            raise AssertionError("forced baseline did not stream all three sides")
        baseline_result, baseline_adapter, baseline_normalized = _compare_and_adapt(
            str(args.sheet), baseline_snapshots, deadline_check,
        )
        baseline_facts = _hard_assert_exact(baseline_result, baseline_adapter, label="three-stream baseline")
        parity = {
            "row_pairs": alias_normalized["row_pairs"] == baseline_normalized["row_pairs"],
            "base_rows_by_pair": alias_normalized["base_rows_by_pair"] == baseline_normalized["base_rows_by_pair"],
            "pair_diff_cols": alias_normalized["pair_diff_cols"] == baseline_normalized["pair_diff_cols"],
            "pair_base_diff_cols": alias_normalized["pair_base_diff_cols"] == baseline_normalized["pair_base_diff_cols"],
            "conflict_cols": alias_normalized["conflict_cols"] == baseline_normalized["conflict_cols"],
            "slot_map": alias_normalized["slot_map"] == baseline_normalized["slot_map"],
            "base_targets": alias_normalized["adapter_targets"] == baseline_normalized["adapter_targets"],
            # The same slot mapping is insufficient: an alias must preserve the
            # exact top-cache confidence/fallback proof as well.  Keep the
            # proof-triple comparison explicit in the report even though it is
            # also nested in top_cache, so a failed proof is immediately clear.
            "top_cache": alias_facts["top_cache"] == baseline_facts["top_cache"],
            "proof_triples": (
                alias_facts["top_cache"]["proof_triples"]
                == baseline_facts["top_cache"]["proof_triples"]
            ),
            "mapping_digest": _canonical_digest(alias_normalized) == _canonical_digest(baseline_normalized),
        }
        if not all(parity.values()):
            raise AssertionError(f"Base alias parity mismatch: {parity}")
        payload["three_stream_baseline"] = {
            "reader_before_fingerprints": baseline_before,
            "reader_after_fingerprints": baseline_after,
            "stream_counts": baseline_streams,
            "actual_stream_count": sum(baseline_streams.values()),
            "timings_ms": baseline_timings,
            "telemetry": baseline_telemetry,
            "facts": baseline_facts,
            "mapping_digest": _canonical_digest(baseline_normalized),
        }
        payload["parity"] = parity
    except Exception as exc:
        payload["error"] = {
            "type": type(exc).__name__,
            "message": str(exc),
            "traceback": traceback.format_exc(),
        }
    finally:
        payload["after"] = {name: _provenance(path) for name, path in protected.items()}
        payload["hashes_unchanged"] = payload["before"] == payload["after"]
        if not payload["hashes_unchanged"]:
            payload.setdefault("error", {
                "type": "AssertionError",
                "message": "source, diagnostic script, or revision input changed during read-only diagnostic",
            })
        payload["elapsed_ms"] = round((time.monotonic() - started) * 1000.0, 3)
        _atomic_json(output, payload)
        print(json.dumps({
            "output": str(output),
            "error": payload.get("error"),
            "elapsed_ms": payload["elapsed_ms"],
            "hashes_unchanged": payload["hashes_unchanged"],
        }, ensure_ascii=False, sort_keys=True))
    return 1 if "error" in payload else 0


if __name__ == "__main__":
    raise SystemExit(main())
