"""Headless regressions for unresolved-conflict save navigation."""

from __future__ import annotations

from types import SimpleNamespace

import sow_merge_tool as smt


def _headless_app():
    app = object.__new__(smt.SowMergeApp)
    app.display_sheets = ["Second", "First", "Third"]
    app._sheet_containers = {
        "Second": "second-container",
        "First": "first-container",
        "Third": "third-container",
    }
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
        "Sheet：Data\n位置：N14\n行号：14　列号：N"
    )


def test_first_conflict_skips_pseudo_and_non_navigable_sheets():
    app = _headless_app()
    app.merge_conflict_cells_by_sheet = {
        "<workbook>": {1: {1}},
        "Hidden": {1: {1}},
        "First": {4: {2}},
    }
    assert app._first_unresolved_conflict_cell() == ("First", 4, 2)

    app.merge_conflict_cells_by_sheet = {
        "<workbook>": {1: {1}},
        "Hidden": {1: {1}},
    }
    assert app._first_unresolved_conflict_cell() is None


def test_startup_primary_uses_manual_review_when_only_pseudo_markers_remain():
    app = _headless_app()
    app.merge_conflict_cells_by_sheet = {"<workbook>": {1: {1}}}
    navigations = []
    reviews = []
    app._navigate_to_conflict_cell = lambda *location: navigations.append(location)
    app._focus_full_three_way_manual_review = lambda review_outcome=None: reviews.append(
        review_outcome
    )
    outcome = smt.StartupMergeOutcome(
        automatic_action="manual-review",
        unresolved_count=1,
    )

    assert app._activate_startup_outcome_primary_action(outcome) == "manual-review"
    assert navigations == []
    assert reviews == [outcome]


def test_source_delta_manual_review_maps_source_before_part_to_real_sheet():
    app = _headless_app()
    app.display_sheets = ["GunshipsModify@design"]
    app._sheet_containers = {"GunshipsModify@design": "modify-container"}
    app.launch_context = SimpleNamespace(
        identity_for=lambda role: (
            SimpleNamespace(effective_path="source-before.xlsx")
            if role == "base" else None
        )
    )
    outcome = smt.StartupMergeOutcome(
        automatic_action="manual-review",
        unresolved_count=1,
        fallback_reasons=[
            "source-delta projection safely declined: "
            "Unsupported Source OOXML representation change: xl/worksheets/sheet2.xml"
        ],
    )
    original_fingerprint = smt._ooxml_package_fingerprint
    original_map = smt._ooxml_sheet_part_map
    try:
        smt._ooxml_package_fingerprint = lambda _path: SimpleNamespace(
            ready=True,
            payloads={},
        )
        smt._ooxml_sheet_part_map = lambda _payloads: {
            "GunshipsModify@design": "xl/worksheets/sheet2.xml",
        }
        assert app._source_delta_review_sheet_from_fallback(outcome) == (
            "GunshipsModify@design"
        )
    finally:
        smt._ooxml_package_fingerprint = original_fingerprint
        smt._ooxml_sheet_part_map = original_map


def test_manual_review_prioritizes_resolved_source_delta_sheet():
    app = _headless_app()
    selected_tabs = []
    expanded = []
    focused = []

    class _ImmediateRoot:
        def after(self, _delay, callback):
            callback()
            return "after-id"

        def focus_set(self):
            focused.append(True)

    class _Var:
        def set(self, value):
            assert value == 1

    review_view = SimpleNamespace(
        _is_three_way_enabled=lambda: True,
        _is_three_way_expanded=lambda: bool(expanded),
        _derive_lifecycle_state=lambda: "READY",
        three_way_var=_Var(),
        _toggle_three_way_view=lambda: expanded.append(True),
    )
    app.root = _ImmediateRoot()
    app._root_after_ids = set()
    app._is_closing = False
    app.nb = SimpleNamespace(select=lambda container: selected_tabs.append(container))
    app._sheet_containers = {
        "Second": "second-container",
        "GunshipsModify@design": "modify-container",
    }
    app.sheet_views = {"GunshipsModify@design": review_view}
    app.selected_sheet = "Second"
    app._set_task_status = lambda *_args, **_kwargs: None
    app._source_delta_review_sheet_from_fallback = (
        lambda _outcome: "GunshipsModify@design"
    )

    app._focus_full_three_way_manual_review(smt.StartupMergeOutcome())

    assert selected_tabs == ["modify-container"]
    assert app.selected_sheet == "GunshipsModify@design"
    assert expanded == [True]
    assert focused


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


def test_batch_conflict_resolution_defers_repaint_and_can_propagate_errors():
    calls = []
    view = object.__new__(smt.SheetView)
    view.sheet = "Data"
    view.app = SimpleNamespace(
        resolve_conflict_row=lambda sheet, row, cols: calls.append(
            (sheet, row, set(cols))
        ) or True,
    )
    view.refresh = lambda **kwargs: calls.append(("refresh", kwargs))
    view._update_cursor_lines = lambda: calls.append(("cursor",))

    view._resolve_conflict_row(4, {2}, refresh=False)
    assert calls == [("Data", 4, {2})]

    view._resolve_conflict_row(5, {3})
    assert calls[-3:] == [
        ("Data", 5, {3}),
        ("refresh", {"row_only": None, "rescan": False}),
        ("cursor",),
    ]

    view.app.resolve_conflict_row = lambda *_args: (_ for _ in ()).throw(
        RuntimeError("conflict-map-write-failed")
    )
    try:
        view._resolve_conflict_row(6, {4}, refresh=False, raise_on_error=True)
    except RuntimeError as exc:
        assert "conflict-map-write-failed" in str(exc)
    else:
        raise AssertionError("batch conflict resolution must propagate requested errors")


def main():
    tests = (
        test_first_conflict_uses_visible_sheet_then_row_then_column_order,
        test_first_conflict_skips_pseudo_and_non_navigable_sheets,
        test_startup_primary_uses_manual_review_when_only_pseudo_markers_remain,
        test_source_delta_manual_review_maps_source_before_part_to_real_sheet,
        test_manual_review_prioritizes_resolved_source_delta_sheet,
        test_save_goto_branch_navigates_and_never_enters_save_pipeline,
        test_navigation_switches_sheet_and_focuses_requested_logical_cell,
        test_batch_conflict_resolution_defers_repaint_and_can_propagate_errors,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"SMOKE_TEST_CONFLICT_NAVIGATION_OK ({len(tests)} tests)")


if __name__ == "__main__":
    main()
