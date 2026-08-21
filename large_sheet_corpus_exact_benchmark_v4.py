"""Production launcher for the corpus exact harness with Windows-safe evidence.

This is the supported entry point for v3's benchmark workflow.  It retains
the v3 comparison implementation but makes its evidence layer resilient to
read locks: the primary JSON is atomically replaced only after a retry window,
while JSONL and a separate atomic heartbeat expose progress without readers
opening the active primary checkpoint.
"""
from __future__ import annotations

import os
import time
from pathlib import Path

import large_sheet_corpus_exact_benchmark as benchmark


def _atomic_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    next_path = path.with_name(path.name + ".next")
    next_path.write_text(benchmark.json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    last_error = None
    for _ in range(150):
        try:
            os.replace(next_path, path)
            return
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.2)
    raise RuntimeError(f"atomic checkpoint remained locked for 30 seconds: {path}") from last_error


_append_jsonl = benchmark.append_jsonl


def _append_with_heartbeat(jsonl_path, value):
    _append_jsonl(jsonl_path, value)
    # A distinct file prevents monitor reads from locking the primary JSON
    # during its atomic replacement.  It has no source-workbook content.
    heartbeat = Path(jsonl_path).with_suffix(Path(jsonl_path).suffix + ".status.json")
    _atomic_json(heartbeat, {"schema": "large-sheet-corpus-heartbeat-v1", "updated_at_epoch": time.time(), "last_event": value})


benchmark.write_json = _atomic_json
benchmark.append_jsonl = _append_with_heartbeat


if __name__ == "__main__":
    benchmark.main()
