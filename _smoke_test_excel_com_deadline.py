"""Focused no-COM deadline and owned-root gates for real Excel fidelity."""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import tempfile
import time
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as sm
import _large_sheet_excel_fidelity_gate as gate
import _large_sheet_snapshot_oracle as snapshot_oracle


_CASE = "excel-com-deadline"


def _result(returncode: int, stderr: str = "") -> SimpleNamespace:
    return SimpleNamespace(returncode=returncode, stderr=stderr)


def _assert_default_retry_contract() -> None:
    calls: list[float] = []
    sleeps: list[float] = []
    results = iter((
        _result(1, "0x80070520"),
        _result(1, "0x80070520"),
        _result(1, "0x80070520"),
        _result(0),
    ))

    def fake_run(*_args, **kwargs):
        calls.append(kwargs["timeout"])
        return next(results)

    with patch.object(sm.subprocess, "run", fake_run), patch.object(sm.time, "sleep", sleeps.append):
        result = sm._run_excel_powershell_with_transient_retry("exit 0", timeout=120)
    assert result.returncode == 0
    assert calls == [120, 120, 120, 120], calls
    assert sleeps == [1.0, 2.0, 3.0], sleeps


def _assert_deadline_stops_retry() -> None:
    calls: list[float] = []
    sleeps: list[float] = []
    moments = iter((10.0, 10.0, 10.25))

    def fake_run(*_args, **kwargs):
        calls.append(kwargs["timeout"])
        return _result(1, "0x80070520")

    with (
        patch.object(sm.subprocess, "run", fake_run),
        patch.object(sm.time, "sleep", sleeps.append),
        patch.object(sm.time, "monotonic", lambda: next(moments)),
    ):
        try:
            sm._run_excel_powershell_with_transient_retry(
                "exit 0", timeout=120, absolute_deadline=10.25
            )
        except sm._ExcelComDeadlineExceeded:
            pass
        else:
            raise AssertionError("deadline-truncated retry unexpectedly continued")
    assert calls == [0.25], calls
    assert sleeps == [0.25], sleeps

    with patch.object(
        sm.subprocess,
        "run",
        side_effect=subprocess.TimeoutExpired(cmd="powershell", timeout=1),
    ) as run:
        try:
            sm._run_excel_powershell_with_transient_retry(
                "exit 0", timeout=120, absolute_deadline=time.monotonic() + 5.0
            )
        except sm._ExcelComDeadlineExceeded:
            pass
        else:
            raise AssertionError("TimeoutExpired did not become deadline failure")
    assert run.call_count == 1


def _assert_reopen_propagates_deadline() -> None:
    captured: list[dict] = []

    def legacy_mock(_script: str, *, timeout: int):
        captured.append({"timeout": timeout})
        return _result(0)

    with (
        patch.object(sm, "_workbook_package_ready", lambda _path: True),
        patch.object(sm, "_run_excel_powershell_with_transient_retry", legacy_mock),
    ):
        assert sm._excel_reopen_validate("C:/owned/default.xlsx") is True
    assert captured == [{"timeout": 120}], captured

    with (
        patch.object(sm, "_workbook_package_ready", lambda _path: True),
        patch.object(
            sm,
            "_run_excel_powershell_with_transient_retry",
            side_effect=sm._ExcelComDeadlineExceeded("expired"),
        ),
    ):
        try:
            sm._excel_reopen_validate(
                "C:/owned/deadline.xlsx", absolute_deadline=time.monotonic() + 5.0
            )
        except sm._ExcelComDeadlineExceeded:
            pass
        else:
            raise AssertionError("reopen converted explicit deadline failure to False")


def _assert_gate_and_legacy_share_deadline() -> None:
    manifest = {"sheet": "S", "three_way": True, "columns": [], "records": [], "only_diff_rows": []}
    deadline = time.monotonic() + 5.0
    received: list[SimpleNamespace] = []

    def fake_capture(args):
        received.append(args)
        return manifest

    with (
        patch.object(gate, "capture_legacy", fake_capture),
        patch.object(gate, "_snapshot_manifest", lambda *_args, **_kwargs: (manifest, False)),
    ):
        gate._assert_frozen_three_way_parity(
            Path("mine.xlsx"),
            Path("theirs.xlsx"),
            Path("base.xlsx"),
            "S",
            "deadline-gate",
            timeout=90.0,
            absolute_deadline=deadline,
        )
    assert len(received) == 1
    assert received[0].absolute_deadline == deadline
    assert 0.0 < received[0].timeout <= 5.0, received[0].timeout

    legacy_calls: list[tuple[list[str], float]] = []
    args = SimpleNamespace(
        mine="mine.xlsx",
        theirs="theirs.xlsx",
        base="base.xlsx",
        sheet="S",
        timeout=90.0,
        absolute_deadline=time.monotonic() + 4.0,
    )

    def fake_worker(command, **kwargs):
        legacy_calls.append((command, kwargs["timeout"]))
        return _result(0)

    with (
        patch.object(snapshot_oracle.subprocess, "run", fake_worker),
        patch.object(snapshot_oracle, "load_manifest", lambda _path: manifest),
        patch.object(snapshot_oracle, "normalize_manifest", lambda value: value),
    ):
        assert snapshot_oracle.capture_legacy(args) == manifest
    command, parent_timeout = legacy_calls[0]
    worker_timeout = float(command[command.index("--timeout") + 1])
    assert 0.0 < worker_timeout <= 4.0
    assert 0.0 < parent_timeout <= 4.0


def _assert_package_probe_forwards_absolute_deadline() -> None:
    with tempfile.TemporaryDirectory(prefix="sow_excel_deadline_probe_") as temporary:
        path = Path(temporary) / "probe.xlsx"
        workbook = Workbook()
        try:
            workbook.active.title = "S"
            workbook.active["A1"] = "id@id"
            workbook.active["A2"] = "string"
            workbook.active["A3"] = "row"
            workbook.save(path)
        finally:
            workbook.close()
        deadline = time.monotonic() + 5.0
        received: list[tuple[str, float]] = []

        def fake_reopen(path_text: str, *, absolute_deadline: float) -> bool:
            received.append((path_text, absolute_deadline))
            return True

        with patch.object(sm, "_excel_reopen_validate", fake_reopen):
            probe = gate._package_probe(
                path, real_excel=True, absolute_deadline=deadline
            )
        assert probe["excel_reopen"] == "ok"
        assert received == [(str(path), deadline)]


def _assert_validation_failure_main_report_and_cleanup() -> None:
    configured_parent = Path(tempfile.gettempdir()).resolve()
    root = Path(tempfile.mkdtemp(prefix="sow_large_sheet_excel_fidelity_forced_"))
    handle, report_name = tempfile.mkstemp(prefix="sow_excel_deadline_report_", suffix=".json")
    os.close(handle)
    report_path = Path(report_name)
    report_path.unlink()
    primary = None
    try:
        with (
            patch.dict(os.environ, {"SOW_TEST_TMPDIR": str(configured_parent)}),
            patch.object(gate, "make_temp_dir", lambda _prefix: str(root)),
            patch.object(
                gate,
                "_validate_owned_case_root",
                side_effect=AssertionError("forced validation failure"),
            ),
        ):
            try:
                gate.main(["--case", "special-formulas", "--out", str(report_path)])
            except AssertionError as exc:
                primary = exc
        assert primary is not None
        assert str(primary) == "forced validation failure"
        report = json.loads(report_path.read_text(encoding="utf-8"))
        assert report["status"] == "FAILED"
        assert report["error"] == "AssertionError: forced validation failure"
        assert report["inputs_before"] == report["inputs_after"] == []
        assert report["runtime_sources_before"] == report["runtime_sources_after"]
        assert report["cleanup"] == {
            "disposable_root_deleted": True,
            "disposable_root_absent": True,
        }
        assert report["disposable_root"] == str(root)
        assert not os.path.lexists(str(root))

        source_root = configured_parent / "sow_large_sheet_excel_fidelity_source_guard"
        source_root.mkdir()
        try:
            with patch.object(gate, "REAL_SOURCE_ROOT", configured_parent):
                try:
                    gate._cleanup_owned_case_root(source_root)
                except AssertionError:
                    pass
                else:
                    raise AssertionError("cleanup accepted a root under REAL_SOURCE_ROOT")
            assert source_root.is_dir()
        finally:
            source_root.rmdir()
    finally:
        if report_path.exists():
            report_path.unlink()
        # The gate must have consumed its exact root; this is a test-side audit,
        # not a fallback deletion path.
        assert not os.path.lexists(str(root))


def run_case() -> None:
    started = time.monotonic()
    _assert_default_retry_contract()
    _assert_deadline_stops_retry()
    _assert_reopen_propagates_deadline()
    _assert_gate_and_legacy_share_deadline()
    _assert_package_probe_forwards_absolute_deadline()
    _assert_validation_failure_main_report_and_cleanup()
    assert time.monotonic() - started < 30.0


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case", choices=(_CASE,))
    parser.add_argument("--list-cases", action="store_true")
    args = parser.parse_args()
    if args.list_cases:
        print(_CASE)
        return
    run_case()
    print("PASS " + (args.case or _CASE))


if __name__ == "__main__":
    main()
