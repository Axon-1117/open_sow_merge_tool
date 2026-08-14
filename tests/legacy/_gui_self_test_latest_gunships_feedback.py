"""Regressions for the latest Gunships column-workflow and SVN Author feedback.

The real workbooks are treated as immutable inputs.  Column actions affect only
the startup candidate held by the GUI and are never saved.
"""

from __future__ import annotations

import os
import sqlite3
import traceback
from contextlib import contextmanager

import sow_merge_tool as smt
from _gui_self_test_focused_merge_acceptance import (
    GUNSHIPS_SHEET,
    GUNSHIPS_SOURCE_AFTER,
    GUNSHIPS_SOURCE_BEFORE,
    GUNSHIPS_TARGET,
    _construct_app,
    _pump,
    _quiet_dialogs,
    _sha256,
    _wait_until,
)


REAL_INPUTS = (
    GUNSHIPS_TARGET,
    GUNSHIPS_SOURCE_BEFORE,
    GUNSHIPS_SOURCE_AFTER,
)


def _require_real_inputs() -> None:
    missing = [path for path in REAL_INPUTS if not os.path.isfile(path)]
    assert not missing, f"real Gunships r37348 fixture unavailable: {missing}"


def _all_column_buttons(view):
    return (
        view.use_mine_col_btn,
        view.use_base_col_btn,
        view.use_theirs_col_btn,
    )


def _assert_column_buttons_fully_visible(view) -> None:
    bar = view.column_action_bar
    bar_left = int(bar.winfo_rootx())
    bar_right = bar_left + int(bar.winfo_width())
    buttons = _all_column_buttons(view)
    assert all(bool(button.winfo_ismapped()) for button in buttons)
    for button in buttons:
        left = int(button.winfo_rootx())
        width = int(button.winfo_width())
        right = left + width
        assert width >= int(button.winfo_reqwidth()), (
            button.cget("text"),
            width,
            button.winfo_reqwidth(),
        )
        assert bar_left <= left < right <= bar_right, (
            button.cget("text"),
            (left, right),
            (bar_left, bar_right),
        )
    ordered = sorted(
        (
            int(button.winfo_rootx()),
            int(button.winfo_rootx()) + int(button.winfo_width()),
            str(button.cget("text")),
        )
        for button in buttons
    )
    for previous, current in zip(ordered, ordered[1:]):
        assert previous[1] <= current[0], (previous, current)


def _mapped_toolbar_controls(view):
    candidates = (
        ("only_diff", view.only_diff_cb),
        ("force_align", view.force_align_cb),
        ("grid", view.grid_overlay_cb),
        ("three_way", view.three_way_cb),
        ("use_left", view.use_left_group),
        ("use_base", view.use_base_btn),
        ("use_right", view.use_right_group),
        ("undo", view.undo_btn),
        ("manual_rescan", view.manual_rescan_btn),
        ("load_all", view._load_all_btn),
    )
    return [
        (name, widget)
        for name, widget in candidates
        if widget is not None and bool(widget.winfo_ismapped())
    ]


def _widget_rect(widget) -> tuple[int, int, int, int]:
    left = int(widget.winfo_rootx())
    top = int(widget.winfo_rooty())
    return (
        left,
        top,
        left + int(widget.winfo_width()),
        top + int(widget.winfo_height()),
    )


def _toolbar_overflow_names(view) -> list[str]:
    left, top, right, bottom = _widget_rect(view._toolbar)
    overflow = []
    for name, widget in _mapped_toolbar_controls(view):
        widget_left, widget_top, widget_right, widget_bottom = _widget_rect(widget)
        if not bool(widget.winfo_ismapped()) or not (
            left <= widget_left < widget_right <= right
            and top <= widget_top < widget_bottom <= bottom
        ):
            overflow.append(name)
    return overflow


def _second_row_overflow_names(view) -> list[str]:
    controls = (
        ("keep_target_column", view.use_mine_col_btn, view.column_action_bar),
        ("source_before_column", view.use_base_col_btn, view.column_action_bar),
        ("source_after_column", view.use_theirs_col_btn, view.column_action_bar),
        # Native Tk button borders can exceed the shared action surface by one
        # theme pixel, so child visibility is owned by the grouped controls.
        # The section-11 test separately proves that the navigation and column
        # groups share one collision-safe row at both required widths.
        ("prev_diff", view.prev_diff_btn, view.diff_nav_group),
        ("diff_status", view.diff_block_status, view.diff_nav_group),
        ("next_diff", view.next_diff_btn, view.diff_nav_group),
    )
    overflow = []
    for name, widget, owner in controls:
        left, top, right, bottom = _widget_rect(owner)
        widget_left, widget_top, widget_right, widget_bottom = _widget_rect(widget)
        if not bool(widget.winfo_ismapped()) or not (
            left <= widget_left < widget_right <= right
            and top <= widget_top < widget_bottom <= bottom
        ):
            overflow.append(name)
    return overflow


def _assert_second_row_groups_do_not_overlap(view) -> None:
    diff_rect = _widget_rect(view.diff_nav_group)
    column_rect = _widget_rect(view.column_action_button_group)
    horizontal_overlap = min(diff_rect[2], column_rect[2]) > max(
        diff_rect[0], column_rect[0]
    )
    vertical_overlap = min(diff_rect[3], column_rect[3]) > max(
        diff_rect[1], column_rect[1]
    )
    assert not (horizontal_overlap and vertical_overlap), (diff_rect, column_rect)


def _assert_toolbar_peers_do_not_overlap(view) -> None:
    controls = [
        (name, _widget_rect(widget))
        for name, widget in _mapped_toolbar_controls(view)
    ]
    for index, (left_name, left_rect) in enumerate(controls):
        for right_name, right_rect in controls[index + 1 :]:
            horizontal_overlap = min(left_rect[2], right_rect[2]) > max(
                left_rect[0], right_rect[0]
            )
            vertical_overlap = min(left_rect[3], right_rect[3]) > max(
                left_rect[1], right_rect[1]
            )
            assert not (horizontal_overlap and vertical_overlap), (
                left_name,
                left_rect,
                right_name,
                right_rect,
            )


def _toolbar_geometry_snapshot(view):
    return {
        "action_group": _widget_rect(view.toolbar_action_group),
        "controls": {
            name: _widget_rect(widget)
            for name, widget in _mapped_toolbar_controls(view)
        },
        "second_controls": {
            "column_group": _widget_rect(view.column_action_button_group),
            "diff_group": _widget_rect(view.diff_nav_group),
            "keep_target_column": _widget_rect(view.use_mine_col_btn),
            "source_before_column": _widget_rect(view.use_base_col_btn),
            "source_after_column": _widget_rect(view.use_theirs_col_btn),
            "prev_diff": _widget_rect(view.prev_diff_btn),
            "diff_status": _widget_rect(view.diff_block_status),
            "next_diff": _widget_rect(view.next_diff_btn),
        },
    }


def _assert_split_and_diff_children_visible(view) -> None:
    parent_children = (
        (
            "use_left",
            view.use_left_group,
            (view.use_left_btn, view.use_left_menu_btn),
        ),
        (
            "use_right",
            view.use_right_group,
            (view.use_right_btn, view.use_right_menu_btn),
        ),
        (
            "diff_nav",
            view.diff_nav_group,
            (view.prev_diff_btn, view.diff_block_status, view.next_diff_btn),
        ),
    )
    for name, parent, children in parent_children:
        parent_left, parent_top, parent_right, parent_bottom = _widget_rect(parent)
        for child in children:
            assert bool(child.winfo_ismapped()), (name, child)
            left, top, right, bottom = _widget_rect(child)
            assert (
                parent_left <= left < right <= parent_right
                and parent_top <= top < bottom <= parent_bottom
            ), (name, _widget_rect(parent), child, _widget_rect(child))


@contextmanager
def _real_gunships_app():
    _require_real_inputs()
    hashes_before = {path: _sha256(path) for path in REAL_INPUTS}
    context = smt.build_merge_launch_context(
        GUNSHIPS_SOURCE_BEFORE,
        GUNSHIPS_TARGET,
        GUNSHIPS_SOURCE_AFTER,
        GUNSHIPS_TARGET,
    )
    analysis = smt.run_startup_merge_analysis(context)
    candidate_path = analysis.outcome.candidate_path
    app = None
    try:
        with _quiet_dialogs():
            app = _construct_app(
                GUNSHIPS_TARGET,
                GUNSHIPS_SOURCE_AFTER,
                base=GUNSHIPS_SOURCE_BEFORE,
                merged=GUNSHIPS_TARGET,
                context=context,
                outcome=analysis.outcome,
                conflict_map=analysis.conflict_cells_by_sheet,
            )
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
                    == (14, 14)
                ),
                "real GunshipsModify did not reach READY at automatic L14",
                timeout=120.0,
            )
            yield app, app.sheet_views[GUNSHIPS_SHEET], analysis
    finally:
        if app is not None:
            app._shutdown_root()
        if (
            candidate_path
            and candidate_path not in REAL_INPUTS
            and os.path.isfile(candidate_path)
        ):
            os.remove(candidate_path)
        assert {path: _sha256(path) for path in REAL_INPUTS} == hashes_before


def test_column_buttons_win_over_long_status_at_1450_and_narrow_width() -> None:
    with _real_gunships_app() as (app, view, _analysis):
        long_status = (
            "整列差异：GunshipsModify@design 中检测到完整列样式、公式、批注、"
            "数据验证、条件格式和合并单元格变化；"
        ) * 12
        for geometry in ("1450x860", "1024x760"):
            app.root.state("normal")
            app.root.geometry(geometry)
            _pump(app.root, 0.25)
            view.info.configure(text="")
            view.column_action_status_var.set("")
            _pump(app.root, 0.15)
            toolbar_before_dynamic_info = _toolbar_geometry_snapshot(view)

            view.info.configure(text=long_status)
            view.column_action_status_var.set(long_status)
            _pump(app.root, 0.15)
            assert _toolbar_geometry_snapshot(view) == toolbar_before_dynamic_info, (
                geometry,
                toolbar_before_dynamic_info,
                _toolbar_geometry_snapshot(view),
            )
            assert all(
                str(button.cget("state")) == "normal"
                for button in _all_column_buttons(view)
            ), geometry
            _assert_column_buttons_fully_visible(view)
            _assert_toolbar_peers_do_not_overlap(view)
            _assert_second_row_groups_do_not_overlap(view)
            _assert_split_and_diff_children_visible(view)
            top_overflow = _toolbar_overflow_names(view)
            second_overflow = _second_row_overflow_names(view)
            assert not top_overflow, (geometry, top_overflow)
            assert not second_overflow, (geometry, second_overflow)
            print(
                f"ACTION_LAYOUT {geometry} "
                f"top_req={view.toolbar_action_group.winfo_reqwidth()} "
                f"top_width={view._toolbar.winfo_width()} "
                f"second_req="
                f"{view.column_action_button_group.winfo_reqwidth() + view.diff_nav_group.winfo_reqwidth()} "
                f"second_width={view.column_action_bar.winfo_width()} "
                f"overflow=[]"
            )


def test_real_gunships_advances_l14_to_l20_then_disables_buttons() -> None:
    with _real_gunships_app() as (app, view, _analysis):
        assert view.selected_column_logical_range == (14, 14)
        first = view._apply_selected_column_block("BASE", "A")
        assert first.action_kind == "insert_copy"
        assert first.logical_start == first.logical_end == 14
        _wait_until(
            app.root,
            lambda: (
                view._derive_lifecycle_state() == "READY"
                and view.selected_column_logical_range == (20, 20)
            ),
            (
                "after applying L14 the workflow must automatically select L20; "
                f"selection={view.selected_column_logical_range!r}, "
                f"states={[button.cget('state') for button in _all_column_buttons(view)]!r}, "
                f"structural={sorted(view.column_comparison_cache.structural_diff_cols)!r}"
            ),
            timeout=30.0,
        )
        assert all(
            str(button.cget("state")) == "normal"
            for button in _all_column_buttons(view)
        )
        first_status = view.column_action_status_var.get()
        assert "待处理 T 已自动选｜可执行" in first_status, first_status
        assert "L20" not in first_status, first_status

        second = view._apply_selected_column_block("BASE", "A")
        assert second.action_kind == "insert_copy"
        assert second.logical_start == second.logical_end == 20
        _wait_until(
            app.root,
            lambda: (
                view._derive_lifecycle_state() == "READY"
                and not view.column_comparison_cache.structural_diff_cols
                and not view.column_comparison_cache.unresolved_cols
            ),
            "after applying L20 no structural column difference should remain",
            timeout=30.0,
        )
        assert view.selected_column_logical_range is None
        assert all(
            str(button.cget("state")) == "disabled"
            for button in _all_column_buttons(view)
        )
        assert "列结构处理完成" in view.column_action_status_var.get()


def _release_wc_revision_and_author() -> tuple[int, str | None]:
    wc_root = smt._find_svn_wc_root_for_path(GUNSHIPS_TARGET)
    assert wc_root
    db_path = os.path.join(wc_root, ".svn", "wc.db")
    with sqlite3.connect(f"file:{db_path}?mode=ro", uri=True) as connection:
        row = connection.execute(
            """
            select changed_revision, changed_author
            from NODES
            where local_relpath = ? and op_depth = 0 and kind = 'file'
              and presence = 'normal'
            limit 1
            """,
            ("sheets/release/Gunships护山神兽.xlsx",),
        ).fetchone()
    assert row
    return int(row[0]), row[1]


def test_real_gunships_authors_survive_release_wc_advancing_past_r37348() -> None:
    _require_real_inputs()
    hashes_before = {path: _sha256(path) for path in REAL_INPUTS}
    try:
        release_revision, _release_author = _release_wc_revision_and_author()
        assert release_revision > 37348, release_revision

        context = smt.build_merge_launch_context(
            GUNSHIPS_SOURCE_BEFORE,
            GUNSHIPS_TARGET,
            GUNSHIPS_SOURCE_AFTER,
            GUNSHIPS_TARGET,
        )
        smt.resolve_cross_branch_source_metadata(context)
        identities = smt.resolve_svn_author_metadata(context)
        source_before = identities["base"]
        source_after = identities["theirs"]

        assert source_after.revision == 37348
        assert source_after.author_status == "resolved", source_after
        assert source_after.author == "rongheng.xue", source_after
        assert source_before.revision == 37347
        assert source_before.author_status == "resolved", source_before
        assert source_before.author == "cheng.zhu2", source_before
        assert source_before.author_source.startswith(
            ("tortoise-svn-revprop-author:", "svn-author-memory-cache:")
        ), source_before.author_source
        assert source_after.author_source.startswith(
            ("tortoise-svn-revprop-author:", "svn-author-memory-cache:")
        ), source_after.author_source
    finally:
        assert {path: _sha256(path) for path in REAL_INPUTS} == hashes_before


def main() -> None:
    tests = (
        test_column_buttons_win_over_long_status_at_1450_and_narrow_width,
        test_real_gunships_advances_l14_to_l20_then_disables_buttons,
        test_real_gunships_authors_survive_release_wc_advancing_past_r37348,
    )
    failures = []
    for test in tests:
        try:
            test()
        except Exception:
            failures.append(test.__name__)
            print(f"FAIL: {test.__name__}")
            traceback.print_exc()
        else:
            print(f"PASS: {test.__name__}")
    if failures:
        raise SystemExit(f"LATEST_GUNSHIPS_FEEDBACK_FAILED: {failures}")
    print(f"LATEST_GUNSHIPS_FEEDBACK_OK ({len(tests)} tests)")


if __name__ == "__main__":
    main()
