"""Pure regression for generation-zero snapshot-alias telemetry collection.

This test imports the real changed-revision harness but supplies only a tiny
immutable fake app.  It therefore exercises the production harness collector
and shared alias-payload validator without starting Tk, a worker process, or
an Excel reader.
"""

from __future__ import annotations

import hashlib
import tempfile
from pathlib import Path
from types import SimpleNamespace

import _gui_real_changed_revision_benchmark as benchmark
import sow_merge_tool as sm


def _fingerprint(path: Path) -> tuple[str, int, int, str]:
    stat = path.stat()
    return (
        str(path.resolve()),
        int(stat.st_size),
        int(stat.st_mtime_ns),
        hashlib.sha256(path.read_bytes()).hexdigest().upper(),
    )


class _FakeApp:
    """Only the attributes consumed by collector/validator are modelled."""

    def __init__(self, path: Path, *, metrics_generation) -> None:
        self.has_base = True
        self._file_a_val_path = path
        self._file_b_val_path = path
        self._file_base_val_path = path
        fingerprint = _fingerprint(path)
        signature = ((fingerprint[1], fingerprint[3]),) * 2
        versions = {
            side: sm._selected_sheet_snapshot_version(
                str(path), topology_generation=0, mutation_generation=0
            )
            for side in ("A", "B", "BASE")
        }
        self.sheet_views = {}
        for index, sheet in enumerate(benchmark.SHEETS, start=1):
            token = f"generation-zero-{index}"
            raw = {
                "sheet": sheet,
                "generation": 0,
                "cache_applied_generation": 0,
                "request_token": token,
                "input_fingerprints_before": (fingerprint, fingerprint, fingerprint),
                "input_fingerprints_after": (fingerprint, fingerprint, fingerprint),
                "snapshot_versions": versions,
                "base_alias": {
                    "used": True,
                    "source": "B",
                    "saved_side": "BASE",
                    "reason": "test-alias",
                    "input_signatures": {"B": signature, "BASE": signature},
                    "source_payload_digest": "payload",
                    "wrapper_payload_digest": "payload",
                    "ingest_ms": {"A": 1.0, "B": 1.0, "BASE": 0.0},
                },
                "base_wrapper": {
                    "side": "BASE",
                    "source_side": "B",
                    "generation": 0,
                    "request_token": token,
                    "is_distinct_from_source": True,
                    "shares_frozen_rows_with_source": True,
                    "shares_frozen_fields_with_source": True,
                },
                "stream_counts": {"A": 1, "B": 1, "BASE": 0},
            }
            self.sheet_views[sheet] = SimpleNamespace(
                _snapshot_alias_telemetry=raw,
                _snapshot_child_metrics={
                    "base_alias_telemetry": raw,
                    "pid": index,
                    "generation": metrics_generation,
                    "request_token": token,
                    "peak_rss_bytes": 1,
                    "last_cpu_ms": 0.0,
                },
                _snapshot_metrics_ms={
                    "mine_ingest": 1.0,
                    "theirs_ingest": 1.0,
                    "base_ingest": 0.0,
                    "compare": 1.0,
                    "adapter": 1.0,
                    "total": 1.0,
                },
            )

    def _sheet_exact_entry(self, sheet: str) -> dict:
        assert sheet in self.sheet_views, sheet
        return {
            "generation": 0,
            "state": "EXACT_CHANGED",
            "full_detail_terminal": True,
        }


def _assert_collector_rejects(path: Path, value, label: str) -> None:
    try:
        benchmark._collect_current_snapshot_alias_telemetry(
            _FakeApp(path, metrics_generation=value)
        )
    except AssertionError:
        return
    raise AssertionError(f"collector accepted {label}: {value!r}")


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="sow_alias_generation_") as root:
        value_path = Path(root) / "immutable-input.bin"
        value_path.write_bytes(b"immutable snapshot alias generation regression\n")

        app = _FakeApp(value_path, metrics_generation=0)
        records = benchmark._collect_current_snapshot_alias_telemetry(app)
        benchmark._validate_snapshot_alias_telemetry_payload(
            records,
            mode="3way",
            final_states={
                sheet: app._sheet_exact_entry(sheet) for sheet in benchmark.SHEETS
            },
        )
        assert {record["child_metrics"]["generation"] for record in records.values()} == {0}

        _assert_collector_rejects(value_path, None, "None")
        _assert_collector_rejects(value_path, False, "bool")
        _assert_collector_rejects(value_path, -1, "negative")
        _assert_collector_rejects(value_path, 1, "mismatch")

        missing = _FakeApp(value_path, metrics_generation=0)
        for view in missing.sheet_views.values():
            del view._snapshot_child_metrics["generation"]
        try:
            benchmark._collect_current_snapshot_alias_telemetry(missing)
        except AssertionError:
            pass
        else:
            raise AssertionError("collector accepted missing generation")

    print("PASS: snapshot alias telemetry preserves generation zero and fails closed")


if __name__ == "__main__":
    main()
