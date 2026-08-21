"""Normalized comparator for frozen legacy and future snapshot manifests.

This intentionally has no Tk or worksheet dependency.  A new comparison
engine can emit its prepared immutable result as JSON, then use this module to
prove parity with a manifest captured by ``_large_sheet_legacy_oracle.py``.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any


_TOP_LEVEL = ("schema", "sheet", "three_way", "columns", "only_diff_rows", "records")


def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def load_manifest(path: str | os.PathLike[str]) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as source:
        return json.load(source)


def normalize_manifest(raw: dict[str, Any]) -> dict[str, Any]:
    """Return stable ordering while rejecting incomplete Oracle manifests."""
    missing = [key for key in _TOP_LEVEL if key not in raw]
    if missing:
        raise ValueError("manifest missing required fields: " + ", ".join(missing))
    records = []
    for record in raw["records"]:
        cells = {str(key): record["cells"][key] for key in sorted(record.get("cells", {}), key=lambda item: int(item))}
        records.append({
            "pair": int(record["pair"]),
            "mine_row": record.get("mine_row"),
            "theirs_row": record.get("theirs_row"),
            "base_row": record.get("base_row"),
            "row_structure": bool(record.get("row_structure", False)),
            "diff_cols": sorted(int(value) for value in record.get("diff_cols", ())),
            "base_diff_cols": sorted(int(value) for value in record.get("base_diff_cols", ())),
            "conflicts": sorted(int(value) for value in record.get("conflicts", ())),
            "cells": cells,
        })
    return {
        "schema": str(raw["schema"]),
        "sheet": str(raw["sheet"]),
        "three_way": bool(raw["three_way"]),
        "columns": sorted(raw["columns"], key=lambda slot: int(slot["logical"])),
        "only_diff_rows": sorted(int(value) for value in raw["only_diff_rows"]),
        "records": sorted(records, key=lambda item: item["pair"]),
    }


def compare_manifests(legacy: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    """Report exact parity; only explicit candidate unresolved blocks are conservative."""
    expected, actual = normalize_manifest(legacy), normalize_manifest(candidate)
    mismatches: list[dict[str, Any]] = []
    for key in ("sheet", "three_way", "columns", "only_diff_rows"):
        if expected[key] != actual[key]:
            mismatches.append({"kind": key, "expected": expected[key], "actual": actual[key]})
    by_pair = {item["pair"]: item for item in expected["records"]}
    candidate_pairs = {item["pair"]: item for item in actual["records"]}
    for pair in sorted(set(by_pair) | set(candidate_pairs)):
        left, right = by_pair.get(pair), candidate_pairs.get(pair)
        if left != right:
            mismatches.append({"kind": "record", "pair": pair, "expected": left, "actual": right})
    return {"exact": not mismatches, "mismatches": mismatches, "legacy_digest": _json(expected), "candidate_digest": _json(actual)}


def capture_legacy(args: argparse.Namespace) -> dict[str, Any]:
    tool = Path(__file__).with_name("_large_sheet_legacy_oracle.py")
    with tempfile.TemporaryDirectory(prefix="sow_large_sheet_legacy_manifest_") as temporary:
        output = Path(temporary) / "legacy.json"
        configured_timeout = float(args.timeout)
        absolute_deadline = getattr(args, "absolute_deadline", None)
        worker_timeout = configured_timeout
        parent_timeout = max(5.0, configured_timeout + 10.0)
        if absolute_deadline is not None:
            remaining = float(absolute_deadline) - time.monotonic()
            if remaining <= 0:
                raise TimeoutError(
                    f"frozen legacy Oracle deadline expired before launch for {args.sheet}"
                )
            # The parent and its Tk worker share the caller's single deadline;
            # unlike the legacy default there is no extra parent-side grace.
            worker_timeout = min(configured_timeout, remaining)
            parent_timeout = remaining
        command = [sys.executable, str(tool), "--mine", args.mine, "--theirs", args.theirs, "--sheet", args.sheet, "--out", str(output), "--timeout", str(worker_timeout)]
        if args.base:
            command.extend(("--base", args.base))
        # The legacy worker owns a Tk loop.  A hard parent-side timeout is
        # required even though it also receives an internal READY deadline:
        # malformed workbooks or modal regressions must not strand a corpus
        # acceptance process.  ``run`` kills and waits for this direct worker.
        try:
            subprocess.run(command, check=True, timeout=parent_timeout)
        except subprocess.TimeoutExpired as exc:
            raise TimeoutError(
                f"frozen legacy Oracle exceeded {worker_timeout:.1f}s for {args.sheet}"
            ) from exc
        if absolute_deadline is not None and time.monotonic() >= float(absolute_deadline):
            raise TimeoutError(f"frozen legacy Oracle deadline expired after worker for {args.sheet}")
        return normalize_manifest(load_manifest(output))


def main() -> None:
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--compare", action="store_true")
    mode.add_argument("--capture-legacy", action="store_true")
    parser.add_argument("--legacy")
    parser.add_argument("--candidate")
    parser.add_argument("--mine")
    parser.add_argument("--theirs")
    parser.add_argument("--base")
    parser.add_argument("--sheet")
    parser.add_argument("--timeout", type=float, default=90.0)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    if args.capture_legacy:
        if not all((args.mine, args.theirs, args.sheet)):
            parser.error("--capture-legacy requires --mine, --theirs, and --sheet")
        result: dict[str, Any] = capture_legacy(args)
    else:
        if not args.legacy or not args.candidate:
            parser.error("--compare requires --legacy and --candidate")
        result = compare_manifests(load_manifest(args.legacy), load_manifest(args.candidate))
    with open(args.out, "w", encoding="utf-8", newline="\n") as destination:
        json.dump(result, destination, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    if args.compare and not result["exact"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
