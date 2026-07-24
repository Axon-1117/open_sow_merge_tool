"""Headless regressions for unresolved-conflict save navigation."""

from __future__ import annotations

from types import SimpleNamespace

import sow_merge_tool as smt


def _headless_app():
    app = object.__new__(smt.SowMergeApp)
    app.display_sheets = ["Second", "First", "Third"]
    app.merge_conflict_cells_by_sheet = {
        "First": {8: {4, 2}, 3: {7}},
        "Second": {12: {5}, 2: {9, 1}},
    }
    app.initial_conflict_cell_count = 6
    return app


def test_first_conflict_uses_visible_sheet_then_row_then_column_order():
    app = _headless_app()
    assert app._first_unresolved_conflict_cell() == ("Second", 2, 1)
    app.display_sheets = ["First", "Second"]
    assert app._first_unresolved_conflict_cell() == ("First", 3, 7)
    assert app._conflict_cell_location_text(("Data", 14, 14)) == (
        "Sheet：Data\n位置：N14\n行号：14　列号：14　逻辑列：L14"
    )


def test_save_goto_branch_navigates_and_never_enters_save_pipeline():
    app = _headless_app()
    app.merged_path = "unused.xlsx"
    app.merge_mode = True
    app._guard_save_readiness = lambda *_args: True
    app._ensure_live_column_mappings_current = lambda *_args: None
    app._ensure_column_replay_available = lambda *_args: None
    dialog_calls = []
    app._show_unresolved_conflict_save_dialog = (
        lambda unresolved, first: dialog_calls.append((unresolved, first)) or "goto"
    )
    navigation_calls = []
    app._navigate_to_conflict_cell = (
        lambda *location: navigation_calls.append(location) or True
    )
    app._with_progress = lambda *_args, **_kwargs: (_ for _ in ()).throw(
        AssertionError("goto must stop before save")
    )

    app.save_merged_and_exit(auto=False)
    assert dialog_calls == [(6, ("Second", 2, 1))]
    assert navigation_calls == [("Second", 2, 1)]


def test_navigation_switches_sheet_and_focuses_requested_logical_cell():
    app = object.__new__(smt.SowMergeApp)
    selected = []
    focused = []
    statuses = []

    class _ImmediateRoot:
        def after(self, _delay, callback):
            callback()
            return "after-id"

    app.root = _ImmediateRoot()
    app.nb = SimpleNamespace(select=lambda container: selected.append(container))
    app._sheet_containers = {"Data": "data-container"}
    app.sheet_views = {
        "Data": SimpleNamespace(
            _data_ready=True,
            focus_logical_cell=lambda row, col: focused.append((row, col)) or True,
        )
    }
    app._is_closing = False
    app._set_task_status = lambda message, **kwargs: statuses.append(
        (message, kwargs)
    )

    assert app._navigate_to_conflict_cell("Data", 14, 14)
    assert selected == ["data-container"]
    assert focused == [(14, 14)]
    assert statuses[-1][0] == "已定位首个冲突：Data!N14"


def main():
    tests = (
        test_first_conflict_uses_visible_sheet_then_row_then_column_order,
        test_save_goto_branch_navigates_and_never_enters_save_pipeline,
        test_navigation_switches_sheet_and_focuses_requested_logical_cell,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"SMOKE_TEST_CONFLICT_NAVIGATION_OK ({len(tests)} tests)")


if __name__ == "__main__":
    main()
