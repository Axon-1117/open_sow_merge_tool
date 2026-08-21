"""Lock-tolerant launcher for the exact corpus benchmark.

Windows readers can transiently lock a JSON checkpoint while an operator is
inspecting it.  Retrying the final atomic replacement preserves the original
file until the full replacement succeeds; the underlying benchmark's JSONL
continues to provide per-Sheet durable evidence in the meantime.
"""
from __future__ import annotations

import os
import time
from pathlib import Path

import large_sheet_corpus_exact_benchmark as benchmark


def lock_tolerant_write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".next")
    temporary.write_text(benchmark.json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    last_error = None
    for _ in range(150):  # 30 seconds: checkpoint is retained, never partially written.
        try:
            os.replace(temporary, path)
            return
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.2)
    raise RuntimeError(f"atomic checkpoint remained locked for 30 seconds: {path}") from last_error


benchmark.write_json = lock_tolerant_write_json

if __name__ == "__main__":
    benchmark.main()
