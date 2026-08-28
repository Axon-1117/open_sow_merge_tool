"""Read-only structured diagnostic for MonsterGroup 3-way duplicate proof.

This tool deliberately constructs no Tk application, editable workbook, or
comparison fallback.  It streams immutable snapshots from the revision exports
and records why Decision-14's duplicate proof is (or is not) admitted.  Its
only write is an atomic JSON report outside the revision inputs.
"""

from __future__ import annotations

import argparse
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


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
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
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, dict):
        return {str(key): _json(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set, frozenset)):
        return [_json(item) for item in value]
    return repr(value)


def _snapshot_alignment_summary(
    name: str,
    alignment: sm.SnapshotAlignment,
    left: sm.SheetSnapshot,
    right: sm.SheetSnapshot,
) -> dict[str, Any]:
    pairs = tuple(alignment.row_pairs)
    return {
        "name": name,
        "unresolved": bool(alignment.unresolved),
        "used_declared_keys": bool(alignment.used_declared_keys),
        "row_pair_count": len(pairs),
        "complete_keyed_pairs": bool(
            sm._snapshot_row_pairs_are_complete_and_keyed(left, right, pairs)
        ),
        "one_sided_pair_count": sum((left_row is None) != (right_row is None) for left_row, right_row in pairs),
        "none_pair_count": sum(left_row is None or right_row is None for left_row, right_row in pairs),
        "duplicate_field_proof_present": alignment.duplicate_field_proof is not None,
        "duplicate_field_proof_pair_count": (
            0 if alignment.duplicate_field_proof is None else len(alignment.duplicate_field_proof.pairs)
        ),
    }


def _column_alignment_summary(alignment: Any | None) -> dict[str, Any] | None:
    if alignment is None:
        return None
    return {
        "has_unresolved": bool(getattr(alignment, "has_unresolved", False)),
        "used_physical_fallback": bool(getattr(alignment, "used_physical_fallback", False)),
        "fallback_reason": str(getattr(alignment, "fallback_reason", "")),
        "fallback_slot_indices": [
            int(index) for index in getattr(alignment, "fallback_slot_indices", ())
        ],
        "anchor_pair_count": len(getattr(alignment, "anchor_pairs", ())),
    }


def _descriptor_evidence(snapshot: sm.SheetSnapshot) -> tuple[dict[str, Any], Any | None]:
    data = sm._snapshot_duplicate_run_descriptors(snapshot)
    if data is None:
        return {"ok": False, "reason": "duplicate-run-descriptor-returned-none"}, None
    groups, duplicate_identities, descriptors = data
    duplicates = {}
    for identity in sorted(duplicate_identities, key=repr):
        entries = [
            (key, payload)
            for key, payload in descriptors.items()
            if key[0] == identity
        ]
        entries.sort(key=lambda item: repr(item[0]))
        duplicates[repr(identity)] = {
            "field_count": len(groups.get(identity, ())),
            "physical_cols": [
                int(field.physical_col) for field in groups.get(identity, ())
            ],
            "descriptor_count": len(entries),
            "descriptor_keys": [repr(key) for key, _payload in entries],
            "full_digest_sha256": [
                hashlib.sha256(repr(payload[1]).encode("utf-8")).hexdigest()
                for _key, payload in entries
            ],
        }
    return {
        "ok": True,
        "duplicate_identity_count": len(duplicate_identities),
        "duplicates": duplicates,
    }, data


def _candidate_stage_evidence(
    mine: sm.SheetSnapshot,
    base: sm.SheetSnapshot,
    theirs: sm.SheetSnapshot,
    ab: sm.SnapshotAlignment,
    mb: sm.SnapshotAlignment,
    tb: sm.SnapshotAlignment,
) -> tuple[sm.SnapshotDuplicateFieldProof | None, dict[str, Any]]:
    side_data: dict[str, Any] = {}
    raw_data: dict[str, Any | None] = {}
    for name, snapshot in (("mine", mine), ("base", base), ("theirs", theirs)):
        side_data[name], raw_data[name] = _descriptor_evidence(snapshot)

    stages: dict[str, Any] = {
        "row_pairs_complete": {
            "AB": sm._snapshot_row_pairs_are_complete_and_keyed(mine, theirs, ab.row_pairs),
            "MB": sm._snapshot_row_pairs_are_complete_and_keyed(mine, base, mb.row_pairs),
            "TB": sm._snapshot_row_pairs_are_complete_and_keyed(theirs, base, tb.row_pairs),
        },
        "descriptors": side_data,
    }
    if any(raw_data[name] is None for name in raw_data):
        stages["first_failing_gate"] = "duplicate-run-descriptor-returned-none"
        return None, stages

    mine_groups, mine_duplicates, mine_descriptors = raw_data["mine"]
    base_groups, base_duplicates, base_descriptors = raw_data["base"]
    theirs_groups, theirs_duplicates, theirs_descriptors = raw_data["theirs"]
    identities = sorted(
        set(mine_duplicates) | set(base_duplicates) | set(theirs_duplicates),
        key=repr,
    )
    per_identity = {}
    for identity in identities:
        by_side = {
            "mine": {
                key: payload for key, payload in mine_descriptors.items() if key[0] == identity
            },
            "base": {
                key: payload for key, payload in base_descriptors.items() if key[0] == identity
            },
            "theirs": {
                key: payload for key, payload in theirs_descriptors.items() if key[0] == identity
            },
        }
        key_sets = {side: set(entries) for side, entries in by_side.items()}
        common_keys = set.intersection(*key_sets.values()) if key_sets else set()
        per_descriptor = {}
        for descriptor in sorted(common_keys, key=repr):
            digests = {
                side: hashlib.sha256(repr(entries[descriptor][1]).encode("utf-8")).hexdigest()
                for side, entries in by_side.items()
            }
            per_descriptor[repr(descriptor)] = {
                "digest_sha256": digests,
                "equal_all_sides": len(set(digests.values())) == 1,
                "physical_cols": {
                    side: int(entries[descriptor][0].physical_col)
                    for side, entries in by_side.items()
                },
            }
        per_identity[repr(identity)] = {
            "field_counts": {
                "mine": len(mine_groups.get(identity, ())),
                "base": len(base_groups.get(identity, ())),
                "theirs": len(theirs_groups.get(identity, ())),
            },
            "descriptor_keys_equal": key_sets["mine"] == key_sets["base"] == key_sets["theirs"],
            "common_descriptor_count": len(common_keys),
            "per_descriptor": per_descriptor,
        }
    stages["per_duplicate_identity"] = per_identity

    direct_error = None
    try:
        proof = sm._build_snapshot_duplicate_field_identity_proof(
            mine, theirs, ab, base, mb, tb
        )
    except Exception as exc:  # Diagnostic only: capture, do not relax.
        proof = None
        direct_error = {
            "type": type(exc).__name__,
            "message": str(exc),
            "traceback": traceback.format_exc(),
        }
    stages["builder"] = {
        "direct_exception": direct_error,
        "direct_returned": proof is not None,
        "safe_returned": sm._try_snapshot_duplicate_field_identity_proof(
            mine, theirs, ab, base, mb, tb
        ) is not None,
        "proof_occurrences": [] if proof is None else [
            {
                "descriptor": repr(occurrence.descriptor),
                "mine_col": int(occurrence.mine_col),
                "base_col": int(occurrence.base_col) if occurrence.base_col is not None else None,
                "theirs_col": int(occurrence.theirs_col),
                "mine_bounds": list(occurrence.mine_bounds),
                "base_bounds": list(occurrence.base_bounds or ()),
                "theirs_bounds": list(occurrence.theirs_bounds),
            }
            for occurrence in proof.occurrences
        ],
    }
    if proof is None:
        stages["first_failing_gate"] = "duplicate-field-candidate-builder-returned-none"
    return proof, stages


def _build_top_cache(
    mine: sm.SheetSnapshot,
    base: sm.SheetSnapshot,
    theirs: sm.SheetSnapshot,
    ab: sm.SnapshotAlignment,
    mb: sm.SnapshotAlignment,
    tb: sm.SnapshotAlignment,
) -> sm.LogicalColumnComparisonCache:
    row_pairs = tuple(ab.row_pairs)
    mine_payloads = tuple(sm._snapshot_row_payload(mine, row) for row, _ in row_pairs)
    theirs_payloads = tuple(sm._snapshot_row_payload(theirs, row) for _, row in row_pairs)
    mine_to_base = {
        mine_row: base_row
        for mine_row, base_row in mb.row_pairs
        if mine_row is not None and base_row is not None
    }
    theirs_to_base = {
        theirs_row: base_row
        for theirs_row, base_row in tb.row_pairs
        if theirs_row is not None and base_row is not None
    }
    base_rows = tuple(
        mine_to_base.get(mine_row, theirs_to_base.get(theirs_row))
        for mine_row, theirs_row in row_pairs
    )
    base_payloads = tuple(sm._snapshot_row_payload(base, row) for row in base_rows)
    cache_key = sm.ColumnModelCacheKey(
        mine.sheet,
        mine.version.topology_generation,
        mine.version.mutation_generation,
    )
    return sm.build_logical_column_comparison_cache_3way(
        cache_key,
        tuple(payload[0] for payload in mine_payloads),
        tuple(payload[0] for payload in base_payloads),
        tuple(payload[0] for payload in theirs_payloads),
        tuple(payload[1] for payload in mine_payloads),
        tuple(payload[1] for payload in base_payloads),
        tuple(payload[1] for payload in theirs_payloads),
        mine_max_col=mine.max_col,
        base_max_col=base.max_col,
        theirs_max_col=theirs.max_col,
    )


def _cache_summary(cache: sm.LogicalColumnComparisonCache) -> dict[str, Any]:
    alignment = cache.three_way_alignment
    slots = tuple(cache.model.slots)
    widths = {"mine_col": 31, "base_col": 31, "theirs_col": 31}
    return {
        "slot_count": len(slots),
        "model": {
            "ambiguous": bool(cache.model.confidence.ambiguous),
            "score": float(cache.model.confidence.score),
            "reason": cache.model.confidence.reason,
            "cause_codes": list(cache.model.confidence.cause_codes),
        },
        "top_alignment": {
            "has_unresolved": bool(getattr(alignment, "has_unresolved", False)),
            "used_physical_fallback": bool(getattr(alignment, "used_physical_fallback", False)),
            "fallback_reason": str(getattr(alignment, "fallback_reason", "")),
            "fallback_slot_indices": [
                int(index) for index in getattr(alignment, "fallback_slot_indices", ())
            ],
            "mine_base": _column_alignment_summary(
                getattr(alignment, "mine_to_base", None)
            ),
            "theirs_base": _column_alignment_summary(
                getattr(alignment, "theirs_to_base", None)
            ),
        },
        "unresolved_cols": sorted(int(value) for value in cache.unresolved_cols),
        "structural_diff_cols": sorted(int(value) for value in cache.structural_diff_cols),
        "all_retained": all(slot.state == "retained" for slot in slots),
        "all_nonambiguous": all(not slot.confidence.ambiguous for slot in slots),
        "all_score_at_least_min": all(
            float(slot.confidence.score) >= float(sm._COLUMN_ALIGNMENT_MIN_SCORE)
            for slot in slots
        ),
        "side_bijection": {
            side: (
                len(slots) == width
                and all(getattr(slot, side) is not None for slot in slots)
                and {int(getattr(slot, side)) for slot in slots} == set(range(1, width + 1))
            )
            for side, width in widths.items()
        },
        "slots": [
            {
                "logical_idx0": int(slot.logical_idx),
                "mine_col": slot.mine_col,
                "base_col": slot.base_col,
                "theirs_col": slot.theirs_col,
                "state": slot.state,
                "ambiguous": bool(slot.confidence.ambiguous),
                "score": float(slot.confidence.score),
                "reason": slot.confidence.reason,
                "cause_codes": list(slot.confidence.cause_codes),
            }
            for slot in slots
        ],
    }


def _replacement_by_proof(
    cache: sm.LogicalColumnComparisonCache,
    proof: sm.SnapshotDuplicateFieldProof | None,
) -> tuple[dict[tuple[int, int, int | None], int], str | None]:
    if proof is None:
        return {}, "candidate-builder-returned-none"
    replacement: dict[tuple[int, int, int | None], int] = {}
    for pair in proof.pairs:
        matches = [
            slot for slot in cache.model.slots
            if sm._snapshot_duplicate_proof_slot_matches(slot, pair, three_way=True)
        ]
        if len(matches) != 1:
            return replacement, f"proof-slot-match-count:{pair!r}:{len(matches)}"
        replacement[pair] = int(matches[0].logical_idx)
    if len(set(replacement.values())) != len(proof.pairs):
        return replacement, "proof-logical-slot-crosswire"
    return replacement, None


def _top_predicate_breakdown(
    cache: sm.LogicalColumnComparisonCache,
    proof: sm.SnapshotDuplicateFieldProof | None,
    replacement: dict[tuple[int, int, int | None], int],
    mine: sm.SheetSnapshot,
    base: sm.SheetSnapshot,
    theirs: sm.SheetSnapshot,
) -> dict[str, Any]:
    if proof is None:
        return {"evaluated": False, "reason": "candidate-builder-returned-none"}
    top = cache.three_way_alignment
    model = cache.model.confidence
    slots = tuple(cache.model.slots)
    proof_slot_matches = {}
    for pair, logical in replacement.items():
        matches = [slot for slot in slots if int(slot.logical_idx) == int(logical)]
        proof_slot_matches[repr(pair)] = {
            "logical_idx0": int(logical),
            "single_slot": len(matches) == 1,
            "same_physical_triple": bool(matches) and sm._snapshot_duplicate_proof_slot_matches(
                matches[0], pair, three_way=True
            ),
        }
    result = {
        "evaluated": True,
        "mode_matches": proof.three_way,
        "replacement_keys_equal_proof_pairs": set(replacement) == set(proof.pairs),
        "replacement_logicals_unique": len(set(replacement.values())) == len(proof.pairs),
        "top_alignment_present": top is not None,
        "top_no_physical_fallback": bool(
            top is not None
            and not getattr(top, "used_physical_fallback", False)
            and not getattr(top, "fallback_slot_indices", ())
        ),
        "model_score_at_least_min": float(model.score) >= float(sm._COLUMN_ALIGNMENT_MIN_SCORE),
        "model_no_low_confidence_cause": (
            sm.COLUMN_MAPPING_CAUSE_LOW_CONFIDENCE not in set(model.cause_codes)
        ),
        "model_not_ambiguous": not bool(model.ambiguous),
        "no_unresolved_cols": not bool(cache.unresolved_cols),
        "no_structural_diff_cols": not bool(cache.structural_diff_cols),
        "all_slots_retained_nonambiguous_scored": all(
            slot.state == "retained"
            and not slot.confidence.ambiguous
            and float(slot.confidence.score) >= float(sm._COLUMN_ALIGNMENT_MIN_SCORE)
            and sm.COLUMN_MAPPING_CAUSE_LOW_CONFIDENCE not in set(slot.confidence.cause_codes)
            for slot in slots
        ),
        "proof_slot_matches": proof_slot_matches,
        "all_side_bijection": sm._snapshot_duplicate_proof_cache_is_bijective(
            cache, mine, theirs, base
        ),
    }
    result["proof_slots_all_match"] = all(
        entry["single_slot"] and entry["same_physical_triple"]
        for entry in proof_slot_matches.values()
    )
    result["production_predicate"] = sm._snapshot_duplicate_proof_matches_exact_top_cache(
        cache, proof, replacement, mine, theirs, base
    )
    return result


def _atomic_json(path: Path, payload: dict[str, Any]) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_name(path.name + ".partial")
    encoded = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True)
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
    """Fail closed before any report write can approach real inputs/source."""
    raw = Path(raw_output)
    if not raw.is_absolute():
        raise ValueError("--output must be an absolute JSON path")
    if raw.suffix.casefold() != ".json":
        raise ValueError("--output must use a .json suffix")
    output = raw.resolve()
    partial = output.with_name(output.name + ".partial")
    protected = tuple(path.resolve() for path in protected_paths)
    if output in protected or partial in protected:
        raise ValueError("--output may not overwrite source, diagnostic, or revision input")
    forbidden_roots = {
        revision_dir.resolve(),
        *(path.parent.resolve() for path in protected_paths if path.suffix.casefold() in {".xlsx", ".xlsm"}),
    }
    if any(
        _is_within(candidate, root)
        for candidate in (output, partial)
        for root in forbidden_roots
    ):
        raise ValueError("--output may not be inside a revision export or input directory")
    workspace_reports = Path(__file__).resolve().parent / "benchmark_results"
    temp_root = Path(tempfile.gettempdir()).resolve()
    allowed_roots = (workspace_reports.resolve(), temp_root)
    if not any(_is_within(output, root) for root in allowed_roots):
        raise ValueError("--output must be under benchmark_results or the system temp directory")
    return output


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--revision-dir", type=Path, default=DEFAULT_REVISION_DIR)
    parser.add_argument("--mine", type=Path)
    parser.add_argument("--base", type=Path)
    parser.add_argument("--theirs", type=Path)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument("--output", type=Path)
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    mine_path = args.mine or args.revision_dir / "Dungeon-r39265.xlsx"
    base_path = args.base or args.revision_dir / "Dungeon-r39264.xlsx"
    theirs_path = args.theirs or args.revision_dir / "Dungeon-r39264.xlsx"
    started = time.monotonic()
    source_path = Path(__file__).with_name("sow_merge_tool.py")
    diagnostic_path = Path(__file__).resolve()
    inputs = {
        "source": source_path,
        "diagnostic_script": diagnostic_path,
        "mine": mine_path,
        "base": base_path,
        "theirs": theirs_path,
    }
    default_output = diagnostic_path.parent / "benchmark_results" / (
        "monster_3way_duplicate_proof_diagnostic_"
        + datetime.now().strftime("%Y%m%d_%H%M%S") + ".json"
    )
    output = _validate_output_path(
        args.output or default_output,
        revision_dir=args.revision_dir,
        protected_paths=tuple(inputs.values()),
    )
    payload: dict[str, Any] = {
        "schema": "monster-3way-duplicate-proof-diagnostic-v2",
        "mode": "3way",
        "sheet": args.sheet,
        "read_only": True,
        "before": {name: _provenance(path) for name, path in inputs.items()},
    }
    original_rebuild = sm._rebuild_snapshot_duplicate_proof_interval
    original_align = sm._align_selected_sheet_snapshots
    rebuild_calls: list[dict[str, int]] = []
    try:
        mine = sm._stream_selected_sheet_snapshot(str(mine_path), str(mine_path), args.sheet, "mine")
        base = sm._stream_selected_sheet_snapshot(str(base_path), str(base_path), args.sheet, "base")
        theirs = sm._stream_selected_sheet_snapshot(str(theirs_path), str(theirs_path), args.sheet, "theirs")
        ab = sm._align_selected_sheet_snapshots(mine, theirs)
        mb = sm._align_selected_sheet_snapshots(mine, base)
        tb = sm._align_selected_sheet_snapshots(theirs, base)
        payload["row_alignments"] = [
            _snapshot_alignment_summary("AB", ab, mine, theirs),
            _snapshot_alignment_summary("MB", mb, mine, base),
            _snapshot_alignment_summary("TB", tb, theirs, base),
        ]
        proof, candidate = _candidate_stage_evidence(mine, base, theirs, ab, mb, tb)
        payload["candidate"] = candidate
        cache = _build_top_cache(mine, base, theirs, ab, mb, tb)
        payload["top_cache"] = _cache_summary(cache)
        replacement, replacement_error = _replacement_by_proof(cache, proof)
        payload["replacement"] = {
            "by_pair": {repr(pair): logical for pair, logical in replacement.items()},
            "error": replacement_error,
        }
        payload["top_predicate"] = _top_predicate_breakdown(
            cache, proof, replacement, mine, base, theirs
        )

        def _rebuild_spy(*call_args: Any, **call_kwargs: Any):
            rebuild_calls.append({"ordinal": len(rebuild_calls) + 1})
            return original_rebuild(*call_args, **call_kwargs)

        sm._rebuild_snapshot_duplicate_proof_interval = _rebuild_spy
        applied = (
            sm._try_apply_snapshot_duplicate_field_proof_to_column_cache(
                cache, proof, mine, theirs, base
            )
            if proof is not None
            else None
        )
        payload["manual_apply"] = {
            "attempted": proof is not None,
            "returned": applied is not None,
            "returned_same_cache": applied is cache if applied is not None else False,
            "rebuild_calls": len(rebuild_calls),
        }
        rebuild_before_compare = len(rebuild_calls)
        final = sm._compare_selected_sheet_snapshots(mine, theirs, base)
        payload["final_compare"] = {
            "unresolved": bool(final.unresolved),
            "row_pair_count": len(final.row_pairs),
            "base_row_count": len(final.base_rows_by_pair),
            "rebuild_calls": len(rebuild_calls) - rebuild_before_compare,
            "final_cache": _cache_summary(final.column_cache),
        }

        adapter_align_calls: list[dict[str, str]] = []

        def _forbid_adapter_realign(left, right, *call_args: Any, **call_kwargs: Any):
            adapter_align_calls.append({
                "left": str(getattr(left, "side", "")),
                "right": str(getattr(right, "side", "")),
            })
            raise AssertionError("immutable adapter must consume result-owned Base mapping")

        sm._align_selected_sheet_snapshots = _forbid_adapter_realign
        try:
            adapter_cache = sm._snapshot_result_to_sheet_cache_immutable(
                args.sheet, final, mine, theirs, base, has_base=True,
            )
        finally:
            sm._align_selected_sheet_snapshots = original_align
        overrides = adapter_cache.get("pair_base_row_override", {}) or {}
        payload["adapter"] = {
            "prepared_complete": bool(adapter_cache.get("prepared_complete")),
            "unresolved_reason": adapter_cache.get("unresolved_reason"),
            "row_pair_count": len(adapter_cache.get("row_pairs") or ()),
            "mine_to_base_count": len(adapter_cache.get("mine_to_base_row") or {}),
            "theirs_to_base_count": len(adapter_cache.get("theirs_to_base_row") or {}),
            "pair_base_override_count": len(overrides),
            "first_base_target": overrides.get(0),
            "last_base_target": overrides.get(len(final.row_pairs) - 1),
            "child_realign_calls": adapter_align_calls,
        }
    except Exception as exc:  # Persist the partial facts for first-fail review.
        payload["error"] = {
            "type": type(exc).__name__,
            "message": str(exc),
            "traceback": traceback.format_exc(),
        }
    finally:
        sm._rebuild_snapshot_duplicate_proof_interval = original_rebuild
        sm._align_selected_sheet_snapshots = original_align
        payload["after"] = {name: _provenance(path) for name, path in inputs.items()}
        payload["hashes_unchanged"] = payload["before"] == payload["after"]
        if not payload["hashes_unchanged"]:
            payload.setdefault("error", {
                "type": "AssertionError",
                "message": "source, diagnostic script, or revision input changed during read-only diagnostic",
            })
        payload["elapsed_ms"] = round((time.monotonic() - started) * 1000.0, 3)
        _atomic_json(output, payload)
        print(json.dumps({
            "output": str(output.resolve()),
            "error": payload.get("error"),
            "elapsed_ms": payload["elapsed_ms"],
            "hashes_unchanged": payload["hashes_unchanged"],
        }, ensure_ascii=False, sort_keys=True))
    return 1 if "error" in payload else 0


if __name__ == "__main__":
    raise SystemExit(main())
