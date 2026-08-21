"""Focused GUI acceptance for the compact three-way merge workspace.

This suite deliberately exercises presentation and navigation only.  It never
applies a row/column action or saves a workbook.  The real Gunships fixture is
used when it is available in the local GM15 working copy.
"""

from __future__ import annotations

import hashlib
import os
import time
from contextlib import ExitStack, contextmanager
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


GUNSHIPS_TARGET = (
    r"C:\GM15\design\sheets\develop\Gunships护山神兽.xlsx"
)
GUNSHIPS_SOURCE_BEFORE = (
    r"C:\GM15\design\sheets\develop"
    r"\Gunships护山神兽.xlsx.merge-left.r37347"
)
GUNSHIPS_SOURCE_AFTER = (
    r"C:\GM15\design\sheets\develop"
    r"\Gunships护山神兽.xlsx.merge-right.r37348"
)
GUNSHIPS_SHEET = "GunshipsModify@design"


def _pump(root, seconds: float = 0.05) -> None:
    deadline = time.time() + seconds
    while time.time() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.005)


def _wait_until(root, predicate, message: str, timeout: float = 30.0) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        _pump(root, 0.025)
        if predicate():
            return
    raise AssertionError(message)


@contextmanager
def _quiet_dialogs():
    """Keep an unattended GUI regression from waiting on a modal dialog."""
    with ExitStack() as stack:
        for name in ("showinfo", "showwarning", "showerror"):
            stack.enter_context(
                patch.object(smt.messagebox, name, lambda *_args, **_kwargs: None)
            )
        stack.enter_context(
            patch.object(
                smt.messagebox,
                "askyesno",
                lambda *_args, **_kwargs: False,
            )
        )
        yield


def _make_single_sheet_book(path: str, marker: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["id", "name", "value", "comment"])
    for row in range(2, 80):
        sheet.append([row - 1, f"name-{row}", f"{marker}-{row}", f"note-{row}"])
    workbook.save(path)
    workbook.close()


def _set_test_identity(
    context: smt.MergeLaunchContext,
    role: str,
    *,
    author: str,
    repository_identity: str,
) -> None:
    identity = context.identity_for(role)
    assert identity is not None
    identity.author = author
    identity.author_status = "resolved"
    identity.author_source = "focused-gui-fixture"
    identity.repository_identity = repository_identity


def _construct_app(
    mine: str,
    theirs: str,
    *,
    base: str,
    merged: str,
    context: smt.MergeLaunchContext,
    outcome: smt.StartupMergeOutcome | None = None,
    conflict_map: dict | None = None,
    startup_owned_paths: set[str] | None = None,
    startup_inputs_prepared: bool = False,
    initial_sheet: str | None = None,
) -> smt.SowMergeApp:
    owned_paths = startup_owned_paths if startup_owned_paths is not None else set()
    app = None
    try:
        for role in ("base", "mine", "theirs"):
            identity = context.identity_for(role)
            if identity is not None and identity.path and not identity.stable_path:
                identity.stable_path = smt._ensure_xlsx_copy(
                    identity.path,
                    owned_paths=owned_paths,
                )
        base_identity = context.identity_for("base")
        mine_identity = context.identity_for("mine")
        theirs_identity = context.identity_for("theirs")
        assert base_identity and mine_identity and theirs_identity
        ui_mine = (
            outcome.candidate_path
            if outcome is not None and outcome.candidate_path
            else mine_identity.effective_path
        )
        app = smt.SowMergeApp(
            str(ui_mine),
            str(theirs_identity.effective_path),
            merge_mode=True,
            merged_path=merged,
            base_path=str(base_identity.effective_path),
            merge_conflict_cells_by_sheet=conflict_map or {},
            # Production main() deliberately keeps every real Sheet available.
            merge_conflict_mode=False,
            raw_base=base,
            raw_mine=mine,
            raw_theirs=theirs,
            launch_context=context,
            startup_outcome=outcome,
            startup_owned_paths=owned_paths,
            startup_inputs_prepared=startup_inputs_prepared,
            initial_sheet=initial_sheet,
        )
        app.root.deiconify()
        _pump(app.root, 0.15)
        app.root.state("normal")
        app.root.geometry("1450x860")
        _pump(app.root, 0.15)
        return app
    except BaseException as primary:
        cleanup_errors = []
        expected_owned = {
            os.path.normcase(os.path.abspath(path)) for path in owned_paths
        }
        if app is not None:
            try:
                app._shutdown_root()
            except BaseException as exc:
                cleanup_errors.append(exc)
            try:
                registry = set(getattr(app, "_owned_startup_temp_paths", ()))
                evidence = list(
                    getattr(app, "_owned_startup_temp_cleanup_evidence", ())
                )
                evidence_paths = {
                    os.path.normcase(os.path.abspath(item.get("path", "")))
                    for item in evidence
                }
                invalid = [
                    item
                    for item in evidence
                    if not item.get("removed")
                    or item.get("exists_after")
                    or item.get("error")
                ]
                if registry or evidence_paths != expected_owned or invalid:
                    cleanup_errors.append(
                        AssertionError(
                            "focused acceptance app cleanup evidence invalid: "
                            + repr(
                                {
                                    "registry": sorted(registry),
                                    "expected": sorted(expected_owned),
                                    "actual": sorted(evidence_paths),
                                    "invalid": invalid,
                                }
                            )
                        )
                    )
            except BaseException as exc:
                cleanup_errors.append(exc)
        if owned_paths:
            try:
                remaining_expected = {
                    os.path.normcase(os.path.abspath(path)) for path in owned_paths
                }
                evidence = smt._cleanup_unclaimed_startup_temp_paths(
                    owned_paths,
                    reason="focused acceptance app initialization failed",
                )
                evidence_paths = {
                    os.path.normcase(os.path.abspath(item.get("path", "")))
                    for item in evidence
                }
                invalid = [
                    item
                    for item in evidence
                    if not item.get("removed")
                    or item.get("exists_after")
                    or item.get("error")
                ]
                if evidence_paths != remaining_expected or invalid:
                    cleanup_errors.append(
                        AssertionError(
                            "focused acceptance unclaimed cleanup evidence invalid: "
                            + repr(
                                {
                                    "expected": sorted(remaining_expected),
                                    "actual": sorted(evidence_paths),
                                    "invalid": invalid,
                                }
                            )
                        )
                    )
            except BaseException as exc:
                cleanup_errors.append(exc)
        for exc in cleanup_errors:
            primary.add_note(f"focused acceptance cleanup failure: {exc!r}")
        raise


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def test_single_visible_sheet_navigation_compact_author_and_grid_height() -> None:
    root_dir = make_temp_dir("sow_focused_compact_workspace_")
    base = os.path.join(root_dir, "Design.xlsx.merge-left.r100")
    mine = os.path.join(root_dir, "Design.xlsx")
    theirs = os.path.join(root_dir, "Design.xlsx.merge-right.r101")
    merged = os.path.join(root_dir, "merged.xlsx")
    _make_single_sheet_book(base, "source-before")
    _make_single_sheet_book(mine, "target-working")
    _make_single_sheet_book(theirs, "source-after")

    context = smt.build_merge_launch_context(base, mine, theirs, merged)
    _set_test_identity(
        context,
        "base",
        author="before.author",
        repository_identity="sheets/release/Design.xlsx",
    )
    _set_test_identity(
        context,
        "mine",
        author="target.author",
        repository_identity="sheets/develop/Design.xlsx",
    )
    _set_test_identity(
        context,
        "theirs",
        author="after.author",
        repository_identity="sheets/release/Design.xlsx",
    )

    with _quiet_dialogs():
        app = _construct_app(
            mine,
            theirs,
            base=base,
            merged=merged,
            context=context,
        )
        try:
            assert app.display_sheets == ["Data"]
            assert app.merge_conflict_mode is False
            app.nb.select(app._sheet_containers["Data"])
            _wait_until(
                app.root,
                lambda: (
                    app.sheet_views.get("Data") is not None
                    and bool(getattr(app.sheet_views["Data"], "_data_ready", False))
                    and app.sheet_views["Data"]._derive_lifecycle_state() == "READY"
                ),
                "single visible Data Sheet did not reach READY",
            )
            view = app.sheet_views["Data"]
            _pump(app.root, 0.2)

            # The lower strip is the sole Sheet navigator; duplicate top tabs
            # are hidden while the Notebook remains the lazy-loading host.
            layout = app.root.tk.call("ttk::style", "layout", "SheetHost.TNotebook")
            assert "Notebook.tab" not in str(layout), layout
            nav_buttons = list(app.nav_inner.winfo_children())
            assert len(nav_buttons) == 1, nav_buttons
            assert nav_buttons[0].cget("text") == "Data"
            assert str(nav_buttons[0].cget("state")) != "disabled"

            # A real one-Sheet conflict route must still switch/focus normally.
            app.merge_conflict_cells_by_sheet = {"Data": {2: {2}}}
            assert app._first_unresolved_conflict_cell() == ("Data", 2, 2)
            assert app._navigate_to_conflict_cell("Data", 2, 2)
            _wait_until(
                app.root,
                lambda: "Data!B2" in str(app.task_status_var.get()),
                "single visible Sheet navigation did not focus Data!B2",
            )
            assert app.selected_sheet == "Data"

            expected_authors = (
                (view.path_label_a, "target.author", mine),
                (view.path_label_base, "before.author", base),
                (view.path_label_b, "after.author", theirs),
            )
            for label, author, full_path in expected_authors:
                compact = str(label.cget("text"))
                detail = str(getattr(label, "_identity_detail_text", ""))
                assert f"Author = {author}" in compact, compact
                assert compact.index(f"Author = {author}") < compact.index(
                    "Design.xlsx"
                ), compact
                assert "SVN路径" not in compact and full_path not in compact, compact
                assert full_path in detail, detail
                assert f"Author = {author}" in detail, detail
                assert label.winfo_reqwidth() <= label.winfo_width(), (
                    compact,
                    label.winfo_reqwidth(),
                    label.winfo_width(),
                )

            selected_container = app._sheet_containers["Data"]
            grid_height = int(view._main_paned.winfo_height())
            page_height = int(selected_container.winfo_height())
            lower_height = int(view.lower_area.winfo_height())
            assert grid_height >= 280, (grid_height, page_height, lower_height)
            assert grid_height >= int(page_height * 0.42), (
                grid_height,
                page_height,
                lower_height,
            )
            assert grid_height > lower_height, (
                "main grid must receive more height than compact lower panels",
                grid_height,
                lower_height,
            )
        finally:
            app._shutdown_root()


def test_pseudo_only_conflict_uses_manual_review_without_navigation() -> None:
    app = object.__new__(smt.SowMergeApp)
    app.display_sheets = ["OnlyVisibleSheet"]
    app._sheet_containers = {"OnlyVisibleSheet": "only-container"}
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

    assert app._first_unresolved_conflict_cell() is None
    assert app._activate_startup_outcome_primary_action(outcome) == "manual-review"
    assert navigations == []
    assert reviews == [outcome]


def test_real_gunships_modify_auto_selects_l14_without_touching_inputs() -> None:
    paths = (
        GUNSHIPS_TARGET,
        GUNSHIPS_SOURCE_BEFORE,
        GUNSHIPS_SOURCE_AFTER,
    )
    missing = [path for path in paths if not os.path.isfile(path)]
    if missing:
        print(f"SKIP: real Gunships r37348 fixture unavailable: {missing}")
        return
    hashes_before = {path: _sha256(path) for path in paths}
    context = smt.build_merge_launch_context(
        GUNSHIPS_SOURCE_BEFORE,
        GUNSHIPS_TARGET,
        GUNSHIPS_SOURCE_AFTER,
        GUNSHIPS_TARGET,
    )
    analysis = smt.run_startup_merge_analysis(context)
    outcome = analysis.outcome
    assert context.scenario == smt.MergeScenario.CROSS_BRANCH_MERGE
    assert outcome.automatic_action == "manual-review", outcome
    assert outcome.unresolved_count == 277, outcome
    assert tuple(analysis.conflict_cells_by_sheet) == ("<workbook>",)

    candidate_path = outcome.candidate_path
    app = None
    try:
        with _quiet_dialogs():
            app = _construct_app(
                GUNSHIPS_TARGET,
                GUNSHIPS_SOURCE_AFTER,
                base=GUNSHIPS_SOURCE_BEFORE,
                merged=GUNSHIPS_TARGET,
                context=context,
                outcome=outcome,
                conflict_map=analysis.conflict_cells_by_sheet,
            )
            assert app.merge_conflict_mode is False
            assert GUNSHIPS_SHEET in app.display_sheets
            app.nb.select(app._sheet_containers[GUNSHIPS_SHEET])
            _wait_until(
                app.root,
                lambda: (
                    app.sheet_views.get(GUNSHIPS_SHEET) is not None
                    and bool(
                        getattr(
                            app.sheet_views[GUNSHIPS_SHEET],
                            "_data_ready",
                            False,
                        )
                    )
                    and app.sheet_views[
                        GUNSHIPS_SHEET
                    ]._derive_lifecycle_state()
                    == "READY"
                    and app.sheet_views[
                        GUNSHIPS_SHEET
                    ].selected_column_logical_range
                    is not None
                ),
                "real GunshipsModify did not reach READY with an automatic selection",
                timeout=120.0,
            )
            view = app.sheet_views[GUNSHIPS_SHEET]
            _pump(app.root, 0.2)

            assert view.selected_column_logical_range == (14, 14)
            assert view.selected_column_source_side == "LOGICAL"
            block = view._selected_column_block()
            assert block is not None and block.state == "mine-deleted", block
            assert 14 in view.column_comparison_cache.structural_diff_cols
            assert 20 in view.column_comparison_cache.structural_diff_cols
            assert not view.column_comparison_cache.unresolved_cols
            action_status = str(view.column_action_status_var.get())
            assert "N" in action_status, action_status
            assert "L14" not in action_status, action_status
            assert all(
                str(button.cget("state")) == "normal"
                for button in (
                    view.use_mine_col_btn,
                    view.use_base_col_btn,
                    view.use_theirs_col_btn,
                )
            )

            retain = view._plan_selected_column_block_action(
                "A",
                "A",
                action_id="acceptance-retain-l14",
            )
            source_before = view._plan_selected_column_block_action(
                "BASE",
                "A",
                action_id="acceptance-before-l14",
            )
            source_after = view._plan_selected_column_block_action(
                "B",
                "A",
                action_id="acceptance-after-l14",
            )
            assert retain.action_kind == "retain"
            for plan in (source_before, source_after):
                assert plan.action_kind == "insert_copy", plan
                assert plan.logical_start == plan.logical_end == 14
                assert plan.target_physical_anchor == 14
                assert plan.count == 1
                assert plan.target_physical_cols == ()
                assert plan.source_physical_cols == (14,)
    finally:
        if app is not None:
            app._shutdown_root()
        if (
            candidate_path
            and candidate_path not in paths
            and os.path.isfile(candidate_path)
        ):
            os.remove(candidate_path)
    assert {path: _sha256(path) for path in paths} == hashes_before


def main() -> None:
    tests = (
        test_single_visible_sheet_navigation_compact_author_and_grid_height,
        test_pseudo_only_conflict_uses_manual_review_without_navigation,
        test_real_gunships_modify_auto_selects_l14_without_touching_inputs,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"FOCUSED_MERGE_GUI_ACCEPTANCE_OK ({len(tests)} tests)")


if __name__ == "__main__":
    main()
