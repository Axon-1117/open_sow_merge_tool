"""Pure contract for bounded fidelity/Section 10/11 test dispatch.

This test deliberately uses parser/plan functions and synthetic call spies only:
it creates no workbooks, opens no Excel/COM application, and creates no Tk root.
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import tempfile
import uuid
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import _gui_self_test_openspec_section10 as section10
import _gui_self_test_openspec_section11 as section11
import _large_sheet_excel_fidelity_gate as fidelity


class _ExpectedFailure(RuntimeError):
    pass


def _capture_stdout(callback, *args):
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        result = callback(*args)
    return output.getvalue(), result


def _assert_raises(expected, callback, *args, **kwargs) -> BaseException:
    try:
        callback(*args, **kwargs)
    except expected as exc:
        return exc
    raise AssertionError(f"expected {expected.__name__}")


def _assert_gui_dispatch_contract(module, expected_ids: tuple[str, ...]) -> None:
    listed, _ = _capture_stdout(module.main, ["--list-cases"])
    assert tuple(line for line in listed.splitlines() if line) == expected_ids

    original_cases = module._CASES
    calls: list[str] = []

    def _first() -> None:
        calls.append("first")

    def _selected() -> None:
        calls.append("selected")

    def _later() -> None:
        calls.append("later")

    def _boom() -> None:
        calls.append("boom")
        raise _ExpectedFailure("synthetic dispatch failure")

    try:
        module._CASES = (("first", _first), ("selected", _selected), ("later", _later))
        _capture_stdout(module.main, ["--case", "selected"])
        assert calls == ["selected"], calls

        calls.clear()
        module._CASES = (("boom", _boom), ("later", _later))
        _assert_raises(_ExpectedFailure, _capture_stdout, module.main, [])
        assert calls == ["boom"], calls
    finally:
        module._CASES = original_cases


def test_fidelity_plans_exactly_one_natural_case_and_rejects_aggregate_flags() -> None:
    original_sources = fidelity._corpus_sources
    synthetic_sources = (
        Path(r"D:\synthetic\one.xlsx"),
        Path(r"D:\synthetic\two.xlsx"),
    )
    fidelity._corpus_sources = lambda: synthetic_sources
    try:
        plans = [fidelity._plan_case("special-formulas", None)]
        plans.extend(
            fidelity._plan_case("fixture", f"{fixture.name}:{variant}")
            for fixture in fidelity.REAL_FIXTURES
            for variant in fidelity._FIXTURE_VARIANTS
        )
        plans.extend(
            fidelity._plan_case("corpus-shard", f"{ordinal}/{len(synthetic_sources)}")
            for ordinal in range(1, len(synthetic_sources) + 1)
        )
        assert all(len(plan.work_items) == 1 for plan in plans), plans
        assert all(plan.work_items == ((plan.case, plan.variant),) for plan in plans)
        assert len(plans) == 1 + len(fidelity.REAL_FIXTURES) * len(fidelity._FIXTURE_VARIANTS) + len(synthetic_sources)

        listed, _ = _capture_stdout(fidelity.main, ["--list-cases"])
        catalog = json.loads(listed)
        assert catalog["case_budget_seconds"] == 90.0
        assert catalog["outer_timeout_recommendation_seconds"] == 115.0
        assert catalog["cases"]["corpus-shard"] == ["1/2", "2/2"]
        assert len(catalog["cases"]["fixture"]) == len(fidelity.REAL_FIXTURES) * len(fidelity._FIXTURE_VARIANTS)

        for aggregate_args in (("--real",), ("--real", "Skill"), ("--no-op-corpus",), ("--no-fixtures",)):
            with contextlib.redirect_stderr(io.StringIO()):
                _assert_raises(SystemExit, fidelity.main, list(aggregate_args))

        list_conflicts = (
            ("--case", "fixture"),
            ("--case", "corpus-shard"),
            ("--variant", "Skill:value"),
            ("--real",),
            ("--real-excel",),
            ("--timeout", "90"),
            ("--out", str(Path(tempfile.gettempdir()) / "list-cases.json")),
            ("--no-op-corpus",),
            ("--no-fixtures",),
        )
        for execution_args in list_conflicts:
            with contextlib.redirect_stderr(io.StringIO()):
                _assert_raises(
                    SystemExit,
                    fidelity.main,
                    ["--list-cases", *execution_args],
                )

        assert fidelity._validate_case_timeout(90.0) == 90.0
        for invalid in (0.0, -1.0, 90.001):
            _assert_raises(ValueError, fidelity._validate_case_timeout, invalid)
    finally:
        fidelity._corpus_sources = original_sources


def _fresh_temp_json(label: str) -> Path:
    path = Path(tempfile.gettempdir()) / f"sow_dispatch_{label}_{uuid.uuid4().hex}.json"
    assert not path.exists(), path
    return path


def test_fidelity_output_guard_rejects_aliases_and_overwrites() -> None:
    allowed = _fresh_temp_json("allowed")
    approved = fidelity._approve_report_destination(allowed, forbidden_paths=())
    assert approved.path == allowed.resolve(strict=False)
    assert approved.partial_path.name == f"{allowed.name}.{os.getpid()}.partial"

    _assert_raises(ValueError, fidelity._approve_report_destination, Path("relative.json"), forbidden_paths=())
    _assert_raises(ValueError, fidelity._approve_report_destination, allowed.with_suffix(".txt"), forbidden_paths=())
    input_path = _fresh_temp_json("input")
    runtime_path = _fresh_temp_json("runtime")
    _assert_raises(ValueError, fidelity._approve_report_destination, input_path, forbidden_paths=(input_path,))
    _assert_raises(ValueError, fidelity._approve_report_destination, runtime_path, forbidden_paths=(runtime_path,))

    existing = _fresh_temp_json("existing")
    partial_target = _fresh_temp_json("partial")
    partial = partial_target.with_name(f"{partial_target.name}.{os.getpid()}.partial")
    try:
        existing.write_text("{}", encoding="utf-8")
        _assert_raises(FileExistsError, fidelity._approve_report_destination, existing, forbidden_paths=())
        partial.write_text("partial", encoding="utf-8")
        _assert_raises(FileExistsError, fidelity._approve_report_destination, partial_target, forbidden_paths=())
    finally:
        for path in (existing, partial):
            if path.exists():
                path.unlink()

    class _JunctionPath:
        def exists(self) -> bool:
            return True

        def is_symlink(self) -> bool:
            return False

        def is_junction(self) -> bool:
            return True

    class _ReparsePath:
        def exists(self) -> bool:
            return True

        def is_symlink(self) -> bool:
            return False

        def is_junction(self) -> bool:
            return False

        def lstat(self):
            return SimpleNamespace(st_file_attributes=0x0400)

    assert fidelity._component_is_link_or_reparse(_JunctionPath())
    assert fidelity._component_is_link_or_reparse(_ReparsePath())
    with patch.object(fidelity, "_component_is_link_or_reparse", return_value=True):
        _assert_raises(
            ValueError,
            fidelity._approve_report_destination,
            _fresh_temp_json("mocked-reparse"),
            forbidden_paths=(),
        )


def test_section_case_ids_are_stable_and_dispatch_is_fail_fast() -> None:
    _assert_gui_dispatch_contract(
        section10,
        (
            "centered-navigation",
            "global-mode-only-diff-undo",
            "three-way-global-base-layout",
            "global-merge-conflict-atomicity",
            "global-structural-zero-write",
            "global-conflict-block-batching",
            "sheet-wide-excel-headers",
            "real-replay-readonly-hash",
        ),
    )
    _assert_gui_dispatch_contract(
        section11,
        (
            "gunships-header-pane-pixels",
            "gunships-combined-action-row",
            "gunships-root-utility-layout",
            "excel-column-guidance",
        ),
    )


def main() -> None:
    test_fidelity_plans_exactly_one_natural_case_and_rejects_aggregate_flags()
    test_fidelity_output_guard_rejects_aliases_and_overwrites()
    test_section_case_ids_are_stable_and_dispatch_is_fail_fast()
    print("PASS: fidelity case-dispatch contract")


if __name__ == "__main__":
    main()
