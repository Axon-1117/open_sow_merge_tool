"""Explain exact snapshot ambiguity for the real Language 3-way inputs."""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path

import sow_merge_tool as sm


SHEET = "default@design@na_TLanguageCn"


def snapshot(path: str, side: str):
    return sm._stream_selected_sheet_snapshot(path, path, SHEET, side)


def record_diagnostic(item):
    key_offsets = [
        field.physical_col - 1
        for field in item.fields
        if field.markers & {"id", "const"}
    ]
    identities = defaultdict(list)
    formula_without_cache = []
    leading_blank = []
    current_key = None
    for row in item.rows[2:]:
        parts = []
        for offset in key_offsets:
            cell = (
                row.cells[offset]
                if offset < len(row.cells)
                else sm._SNAPSHOT_BLANK_CELL
            )
            identity = sm._snapshot_declared_key_cell_identity(cell)
            if identity is None:
                formula_without_cache.append((row.physical_row, offset + 1))
            parts.append(identity)
        key = tuple(parts)
        blank = bool(parts) and all(value == "BLANK:" for value in parts)
        if blank:
            if current_key is None:
                leading_blank.append(row.physical_row)
            continue
        if any(value is None for value in parts):
            continue
        current_key = key
        identities[key].append(row.physical_row)
    duplicates = [
        {"key": repr(key), "rows": rows}
        for key, rows in identities.items()
        if len(rows) > 1
    ]
    return {
        "key_offsets": [offset + 1 for offset in key_offsets],
        "record_result": sm._snapshot_declared_records(item) is not None,
        "formula_without_cache": formula_without_cache[:20],
        "leading_blank": leading_blank[:20],
        "duplicate_count": len(duplicates),
        "duplicates": duplicates[:20],
    }


def field_diagnostic(item):
    identities = [sm._snapshot_field_identity(field) for field in item.fields]
    counts = Counter(identities)
    return {
        "max_row": item.max_row,
        "max_col": item.max_col,
        "fields": [
            {
                "col": field.physical_col,
                "declaration": field.declaration,
                "type": field.type_declaration,
                "markers": sorted(field.markers),
                "identity": repr(sm._snapshot_field_identity(field)),
            }
            for field in item.fields
        ],
        "duplicate_field_identities": [
            {"identity": repr(identity), "count": count}
            for identity, count in counts.items()
            if count > 1
        ],
    }


def alignment_diagnostic(left, right):
    alignment = sm._align_selected_sheet_snapshots(left, right)
    return {
        "unresolved": alignment.unresolved,
        "declared": alignment.used_declared_keys,
        "row_pairs": len(alignment.row_pairs),
        "duplicate_proof": alignment.duplicate_field_proof is not None,
        "field_pairs": alignment.field_pairs,
    }


def duplicate_digest_diagnostic(items):
    streams = {
        side: sm._snapshot_duplicate_run_descriptors(item, keyed_payload=True)
        for side, item in items.items()
    }
    descriptors = {
        side: data[2] if data is not None else {}
        for side, data in streams.items()
    }
    base_entries = descriptors["base"]
    rows = []
    for descriptor, base_payload in sorted(base_entries.items(), key=lambda item: repr(item[0])):
        base_digest = base_payload[1]
        entry = {
            "descriptor": repr(descriptor),
            "base_col": base_payload[0].physical_col,
            "base_digest": base_digest[:16],
        }
        for side in ("mine", "theirs"):
            payload = descriptors[side].get(descriptor)
            if payload is None:
                entry[side] = None
                continue
            digest = payload[1]
            cross = [
                other_payload[0].physical_col
                for other_descriptor, other_payload in base_entries.items()
                if other_descriptor != descriptor and other_payload[1] == digest
            ]
            entry[side] = {
                "col": payload[0].physical_col,
                "digest": digest[:16],
                "equals_base": digest == base_digest,
                "cross_base_cols": cross,
            }
        rows.append(entry)
    return rows


def main():
    parser = argparse.ArgumentParser()
    root = Path(
        "D:/Tools/sow_merge_tool_proj/tmp/"
        "language_3way_unresolved_repro_20260828_183244"
    )
    parser.add_argument("--base", default=str(root / "base.xlsx"))
    parser.add_argument("--mine", default=str(root / "mine.xlsx"))
    parser.add_argument("--theirs", default=str(root / "theirs.xlsx"))
    args = parser.parse_args()
    items = {
        "base": snapshot(str(Path(args.base).resolve()), "BASE"),
        "mine": snapshot(str(Path(args.mine).resolve()), "A"),
        "theirs": snapshot(str(Path(args.theirs).resolve()), "B"),
    }
    result = {
        "sides": {
            side: {
                "field": field_diagnostic(item),
                "record": record_diagnostic(item),
            }
            for side, item in items.items()
        },
        "alignments": {
            "mine_theirs": alignment_diagnostic(items["mine"], items["theirs"]),
            "mine_base": alignment_diagnostic(items["mine"], items["base"]),
            "theirs_base": alignment_diagnostic(items["theirs"], items["base"]),
        },
        "duplicate_digests": duplicate_digest_diagnostic(items),
    }
    comparison = sm._compare_selected_sheet_snapshots(
        items["mine"], items["theirs"], items["base"]
    )
    result["comparison"] = {
        "unresolved": comparison.unresolved,
        "row_pairs": len(comparison.row_pairs),
        "unresolved_columns": sorted(comparison.column_cache.unresolved_cols),
        "structural_columns": sorted(comparison.column_cache.structural_diff_cols),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=repr))


if __name__ == "__main__":
    main()
