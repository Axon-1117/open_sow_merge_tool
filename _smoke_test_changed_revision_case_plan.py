"""Pure contract for changed-revision harness parent case selection.

This test intentionally imports no Tk controls and never invokes a worker,
subprocess, workbook, or revision input.  It keeps the outer CLI plan distinct
from a selected worker's established two-Sheet readiness exercise.
"""

from __future__ import annotations

import _gui_real_changed_revision_benchmark as bench


def _assert_case_plan_selectors() -> None:
    assert bench._changed_revision_case_plan(None, None) == (
        ("2way", "Dungeon@design"),
        ("2way", "MonsterGroup@design"),
        ("3way", "Dungeon@design"),
        ("3way", "MonsterGroup@design"),
    )
    assert bench._changed_revision_case_plan("3way", "Dungeon@design") == (
        ("3way", "Dungeon@design"),
    )
    assert bench._changed_revision_case_plan("2way", None) == (
        ("2way", "Dungeon@design"),
        ("2way", "MonsterGroup@design"),
    )
    assert bench._changed_revision_case_plan(None, "MonsterGroup@design") == (
        ("2way", "MonsterGroup@design"),
        ("3way", "MonsterGroup@design"),
    )


def _assert_default_fail_fast() -> None:
    plan = bench._changed_revision_case_plan(None, None)
    calls: list[tuple[str, str]] = []

    def run_case(mode: str, sheet: str) -> dict:
        calls.append((mode, sheet))
        return {
            "schema": "real-changed-revision-gui-v2",
            "mode": mode,
            "sheet": sheet,
            "status": "FAILED" if len(calls) == 2 else "PASS",
            "error": "synthetic failure" if len(calls) == 2 else None,
        }

    results, stopped_after = bench._execute_changed_revision_case_plan(
        plan, run_case, continue_after_failure=False,
    )
    assert calls == list(plan[:2]), calls
    assert [item["status"] for item in results] == [
        "PASS", "FAILED", "NOT_RUN_AFTER_FAILURE", "NOT_RUN_AFTER_FAILURE",
    ], results
    assert stopped_after == {
        "case_index": 1,
        "mode": "2way",
        "sheet": "MonsterGroup@design",
        "status": "FAILED",
        "error": "synthetic failure",
    }, stopped_after
    assert all(
        item["stopped_after"] == stopped_after for item in results[2:]
    ), results


def _assert_explicit_continue_after_failure() -> None:
    plan = bench._changed_revision_case_plan(None, None)
    calls: list[tuple[str, str]] = []

    def run_case(mode: str, sheet: str) -> dict:
        calls.append((mode, sheet))
        return {
            "mode": mode,
            "sheet": sheet,
            "status": "FAILED" if len(calls) == 2 else "PASS",
            "error": "synthetic failure" if len(calls) == 2 else None,
        }

    results, stopped_after = bench._execute_changed_revision_case_plan(
        plan, run_case, continue_after_failure=True,
    )
    assert calls == list(plan), calls
    assert [item["status"] for item in results] == [
        "PASS", "FAILED", "PASS", "PASS",
    ], results
    assert stopped_after is None, stopped_after


def _assert_launch_exception_is_fail_fast() -> None:
    plan = bench._changed_revision_case_plan("3way", None)
    calls: list[tuple[str, str]] = []

    def run_case(mode: str, sheet: str) -> dict:
        calls.append((mode, sheet))
        raise OSError("synthetic launch failure")

    results, stopped_after = bench._execute_changed_revision_case_plan(
        plan, run_case, continue_after_failure=False,
    )
    assert calls == [plan[0]], calls
    assert [item["status"] for item in results] == [
        "WORKER_LAUNCH_FAILURE", "NOT_RUN_AFTER_FAILURE",
    ], results
    assert stopped_after == {
        "case_index": 0,
        "mode": "3way",
        "sheet": "Dungeon@design",
        "status": "WORKER_LAUNCH_FAILURE",
        "error": "OSError: synthetic launch failure",
    }, stopped_after


def _assert_single_anchor_never_expands_outer_workers() -> None:
    plan = bench._changed_revision_case_plan("3way", "Dungeon@design")
    calls: list[tuple[str, str]] = []

    def run_case(mode: str, sheet: str) -> dict:
        calls.append((mode, sheet))
        return {"mode": mode, "sheet": sheet, "status": "PASS"}

    results, stopped_after = bench._execute_changed_revision_case_plan(
        plan, run_case, continue_after_failure=False,
    )
    assert calls == [("3way", "Dungeon@design")], calls
    assert [item["status"] for item in results] == ["PASS"], results
    assert stopped_after is None, stopped_after
    assert bench._case_plan_descriptors(plan) == [{
        "case_index": 0,
        "mode": "3way",
        "sheet": "Dungeon@design",
    }]


def main() -> None:
    _assert_case_plan_selectors()
    _assert_default_fail_fast()
    _assert_explicit_continue_after_failure()
    _assert_launch_exception_is_fail_fast()
    _assert_single_anchor_never_expands_outer_workers()
    print("changed revision case plan: PASS")


if __name__ == "__main__":
    main()
