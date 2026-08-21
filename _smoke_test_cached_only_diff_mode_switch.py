"""Pure contracts for the prepared-cache only-difference mode switch.

The public GUI regression keeps the actual Checkbutton route under workbook and
Tk sentinels.  These contracts isolate the generation gate and scheduler so a
stale callback cannot fall back to ``refresh`` (and worksheet I/O) while tests
remain deterministic and headless.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import sow_merge_tool as sm


class _Var:
    def __init__(self, value: int):
        self.value = int(value)

    def get(self) -> int:
        return int(self.value)

    def set(self, value: int) -> None:
        self.value = int(value)


class _Info:
    def __init__(self):
        self.values = []

    def configure(self, **kwargs) -> None:
        self.values.append(dict(kwargs))


class _Widget:
    def __init__(self, state: str = "disabled"):
        self.options = {"state": str(state)}

    def configure(self, **kwargs) -> None:
        self.options.update(kwargs)

    def cget(self, name: str):
        return self.options.get(str(name))


class _Event:
    def __init__(self):
        self.value = False

    def is_set(self) -> bool:
        return bool(self.value)


@dataclass
class _Frame:
    callbacks: list[tuple[int, object]] = field(default_factory=list)
    canceled: set[int] = field(default_factory=set)
    next_id: int = 1

    def after(self, _delay: int, callback):
        token = self.next_id
        self.next_id += 1
        self.callbacks.append((token, callback))
        return token

    def after_cancel(self, token) -> None:
        self.canceled.add(int(token))

    def run_all_including_canceled(self) -> None:
        callbacks = list(self.callbacks)
        self.callbacks.clear()
        for _token, callback in callbacks:
            callback()


class _App:
    def __init__(self):
        self.selected_sheet = "S1"
        self._is_closing = False
        self.has_base = False
        self.releases = []
        self.forbidden_hits = []
        self.exact = True
        self.exact_state = sm._SHEET_EXACT_CHANGED
        self.modified_sheets_a = set()
        self.modified_sheets_b = set()
        self.merge_conflict_mode = False
        self._interactive_action_event = _Event()
        self.recalc_btn = _Widget()
        self.exact_updates = []
        self.progress_outcomes = []
        self.status_updates = 0
        self.nav_updates = 0
        self.enqueues = []
        self.kicks = 0
        self.priority_claims = []
        self.priority_submissions = []
        self._only_diff_progress_owner = None
        self._sheet_compute_generation = {"S1": 0}

    def _is_sheet_exact_current(self, sheet: str) -> bool:
        return self.exact and sheet == "S1"

    def _sheet_exact_entry(self, sheet: str) -> dict:
        return {
            "sheet": sheet,
            "state": self.exact_state,
            "generation": int(self._sheet_compute_generation.get(sheet, -1)),
            "full_detail_terminal": True,
        }

    def _edit_workbooks_ready(self) -> bool:
        # The view-only positive deliberately leaves editable workbooks absent.
        return False

    def _set_sheet_exact_state(self, sheet, state, **kwargs) -> bool:
        self.exact = state in (sm._SHEET_EXACT_SAME, sm._SHEET_EXACT_CHANGED)
        self.exact_state = str(state)
        self.exact_updates.append((str(sheet), str(state), dict(kwargs)))
        return True

    def _update_exact_status_ui(self) -> None:
        self.status_updates += 1

    def refresh_sheet_nav(self) -> None:
        self.nav_updates += 1

    def _finish_only_diff_progress(self, view, build_seq, *, outcome: str) -> None:
        token = (view, int(build_seq))
        if self._only_diff_progress_owner == token:
            self._only_diff_progress_owner = None
        self.progress_outcomes.append((view.sheet, int(build_seq), str(outcome)))

    def _claim_priority_exact(self, view, build_seq: int) -> None:
        self.priority_claims.append((view, int(build_seq)))

    def _submit_priority_exact(self, view, build_seq: int, worker):
        self.priority_submissions.append((view, int(build_seq), worker))
        return object()

    def _enqueue_sheet(self, *args, **kwargs) -> None:
        self.enqueues.append((tuple(args), dict(kwargs)))

    def _kick_worker(self) -> None:
        self.kicks += 1

    def _reserve_ui_transition_window(self, _milliseconds: int):
        released = False

        def _release():
            nonlocal released
            assert not released, "transition released twice"
            released = True
            self.releases.append("released")

        return _release

    # The cache-only gate/publisher must never touch any worksheet or edit
    # accessor.  Keep concrete traps instead of merely omitting these names.
    def ws_a_val(self, *_args, **_kwargs):
        self.forbidden_hits.append("ws_a_val")
        raise AssertionError("worksheet access")

    def ws_b_val(self, *_args, **_kwargs):
        self.forbidden_hits.append("ws_b_val")
        raise AssertionError("worksheet access")

    def ws_base_val(self, *_args, **_kwargs):
        self.forbidden_hits.append("ws_base_val")
        raise AssertionError("worksheet access")

    def _request_edit_preload(self, *_args, **_kwargs):
        self.forbidden_hits.append("_request_edit_preload")
        raise AssertionError("edit preload")


def _reference_window_start(rows, anchor_pair: int) -> tuple[int, int]:
    """Independent logical-anchor reference; never call the production helper."""
    rows = [int(pair_idx) for pair_idx in rows]
    assert rows
    try:
        target_index = rows.index(int(anchor_pair))
        target_pair = int(anchor_pair)
    except ValueError:
        target_index = next(
            (
                offset
                for offset, pair_idx in enumerate(rows)
                if int(pair_idx) >= int(anchor_pair)
            ),
            len(rows) - 1,
        )
        target_pair = int(rows[target_index])
    cap = min(int(sm._VIRTUAL_VIEWPORT_MAX_ROWS), len(rows))
    return (
        max(0, min(target_index, max(0, len(rows) - cap))),
        target_pair,
    )


def _make_view(*, row_count: int = 400):
    app = _App()
    frame = _Frame()
    view = sm.SheetView.__new__(sm.SheetView)
    view.app = app
    view.frame = frame
    view.info = _Info()
    view.sheet = "S1"
    view.only_diff_var = _Var(0)
    view.max_col = 3
    view.row_pairs = [(index + 1, index + 1) for index in range(row_count)]
    view.row_a_to_pair_idx = {row_a: pair_idx for pair_idx, (row_a, _row_b) in enumerate(view.row_pairs)}
    view.row_b_to_pair_idx = {row_b: pair_idx for pair_idx, (_row_a, row_b) in enumerate(view.row_pairs)}
    view.mine_to_base_row = {}
    view.theirs_to_base_row = {}
    view.pair_base_row_override = {}
    view._align_rows_enabled = True
    view.pair_raw_parts_a = {index: ("a", index) for index in range(row_count)}
    view.pair_raw_parts_b = {index: ("b", index) for index in range(row_count)}
    view.pair_raw_parts_base = {}
    view._full_display_rows = list(range(row_count))
    view.display_rows = list(range(row_count - 100, row_count))
    view.row_to_line = {
        pair_idx: line
        for line, pair_idx in enumerate(view.display_rows, start=1)
    }
    view._virtual_window_start = row_count - 100
    view._virtual_column_window_start = 7
    view._mode_switch_pending = False
    view._mode_switch_seq = 0
    view._mode_switch_after_id = None
    view._mode_switch_release_transition = None
    view._mode_switch_requested_value = None
    view._mode_switch_origin_value = None
    view._mode_switch_completion = None
    view._last_only_diff_value = 0
    view._pending_exact_render = False
    view._only_diff_async_building = False
    view._only_diff_async_build_key = None
    view._only_diff_async_build_seq = 41
    view._only_diff_async_requested_value = 0
    view._only_diff_async_thread = None
    view._only_diff_request_origin_value = 0
    view._virtual_publishing = False
    view._prepared_cache_publish_active = False
    view._lifecycle_error = None
    view._lifecycle_canceled = False
    view._data_ready = True
    view._prepared_complete = True
    view._row_model_exact = True
    view._cache_formula_aware = True
    view._pair_diff_full_exact = True
    view._base_diff_full_exact = True
    view._only_diff_rows_exact = True
    view._only_diff_rows_cache = list(range(0, row_count, 5))
    view.pair_diff_cols = {5: {2}, 10: {3}}
    view.pair_base_diff_cols = {}
    view._sheet_structural_diff = False
    view._data_version = 0
    view._current_only_diff_cache_key = lambda: ("S1", 0, 0)
    view._full_render = False
    view._render_limit = sm._LARGE_SHEET_INITIAL_ROWS
    view._is_large_sheet = True
    view._only_diff_preview_full = False
    view._lifecycle_state = "READY"
    view.touched_rows = set()
    view._terminal_surfaces = []
    view._published_rows = []
    view._restored_selection = []
    view._cleared_selection = 0
    view._cleared_hover = 0
    view._refresh_states = []
    view._is_three_way_enabled = lambda: False
    view._logical_slot_count = lambda: 20
    view._column_mapping_is_current = lambda: True
    view._has_valid_only_diff_snapshot_cache = lambda: True
    view._only_diff_rows_with_touched = lambda rows: list(rows)
    view._active_column_comparison_cache = lambda: type(
        "_Cache", (), {"structural_diff_cols": set(), "unresolved_cols": set()}
    )()
    view._cache_only_diff_rows_snapshot = lambda rows, exact=False: (
        setattr(view, "_only_diff_rows_cache", sorted({int(row) for row in rows})),
        setattr(view, "_only_diff_rows_exact", bool(exact) or view._only_diff_rows_exact),
    )
    view._pending_pair_parts_cache = None

    def _stage_cached_parts(a, b, base, widths, *, replace=True):
        incoming_a = dict(a or {})
        incoming_b = dict(b or {})
        incoming_base = dict(base or {})
        incoming_widths = dict(widths or {})
        pending = view._pending_pair_parts_cache
        if not replace and pending is not None:
            old_a, old_b, old_base, old_widths, _old_replace = pending
            incoming_a = {**dict(old_a or {}), **incoming_a}
            incoming_b = {**dict(old_b or {}), **incoming_b}
            incoming_base = {**dict(old_base or {}), **incoming_base}
            incoming_widths = {**dict(old_widths or {}), **incoming_widths}
        view._pending_pair_parts_cache = (
            incoming_a,
            incoming_b,
            incoming_base,
            incoming_widths,
            bool(replace),
        )
        view._staged_async_parts = (
            dict(incoming_a),
            dict(incoming_b),
            dict(incoming_base),
            dict(incoming_widths),
            {"replace": bool(replace)},
        )

    def _materialize_staged_parts(*_args, **_kwargs):
        pending = view._pending_pair_parts_cache
        if pending is None:
            return
        parts_a, parts_b, parts_base, _widths, replace = pending
        view._pending_pair_parts_cache = None
        if replace:
            view.pair_raw_parts_a = dict(parts_a)
            view.pair_raw_parts_b = dict(parts_b)
            view.pair_raw_parts_base = dict(parts_base)
        else:
            view.pair_raw_parts_a.update(dict(parts_a))
            view.pair_raw_parts_b.update(dict(parts_b))
            view.pair_raw_parts_base.update(dict(parts_base))
        view._materialized_async_parts = True

    view._stage_cached_pair_parts = _stage_cached_parts
    view._materialize_staged_pair_parts = _materialize_staged_parts
    view._invalidate_render_cache = lambda: setattr(
        view, "_invalidated_async_render", getattr(view, "_invalidated_async_render", 0) + 1
    )
    view._install_exact_diff_map_cache = lambda rows: setattr(
        view, "_installed_async_diff_rows", tuple(int(row) for row in rows)
    )
    view._hide_loading = lambda: setattr(view, "_hid_async_loading", True)
    view._show_loading = lambda _message: setattr(view, "_showed_async_loading", True)
    view._persist_only_diff_setting_debounced = lambda: setattr(
        view, "_persisted_async_setting", getattr(view, "_persisted_async_setting", 0) + 1
    )
    view._snapshot_explicit_selection_state = lambda: {"pair_idx": row_count - 70, "main_col": 9}
    view._pair_idx_from_selection_snapshot = lambda snapshot: snapshot.get("pair_idx") if snapshot else None
    view._clear_selection_visuals = lambda: setattr(
        view, "_cleared_selection", view._cleared_selection + 1
    )
    view._clear_hover_state = lambda **_kwargs: setattr(
        view, "_cleared_hover", view._cleared_hover + 1
    )

    def _publish(*, on_staged_complete=None, prepared_rows=None):
        assert on_staged_complete is None
        assert prepared_rows is not None
        rows = list(prepared_rows)
        view._published_rows.append(rows)
        view._full_display_rows = rows
        start = max(0, int(view._virtual_window_start))
        cap = min(sm._VIRTUAL_VIEWPORT_MAX_ROWS, len(rows))
        view.display_rows = rows[start : start + cap]
        view.row_to_line = {
            pair_idx: line
            for line, pair_idx in enumerate(view.display_rows, start=1)
        }
        return True

    def _restore(snapshot):
        pair_idx = snapshot.get("pair_idx") if snapshot else None
        if pair_idx not in view.row_to_line:
            return False
        view._restored_selection.append((pair_idx, snapshot.get("main_col")))
        return True

    view._publish_prepared_cache_surface = _publish
    view._restore_explicit_selection_state = _restore
    view.clear_explicit_cell_selection = lambda: view._restored_selection.append((None, None))
    view._update_cursor_lines = lambda: None
    view._refresh_diff_block_ui = lambda: None
    view._update_diff_nav_state = lambda: None
    view._show_exact_unavailable = lambda message: view._terminal_surfaces.append(str(message))
    view._refresh_mode_switch_preserving_selection = lambda *, rescan: view._legacy_refreshes.append(bool(rescan))
    view._legacy_refreshes = []

    def _refresh_gate():
        if not app.exact and app.exact_state == sm._SHEET_EXACT_UNRESOLVED:
            state = "UNRESOLVED"
        elif view._lifecycle_error:
            state = "FAILED"
        elif view._mode_switch_pending or view._pending_exact_render:
            state = "DIFFING"
        else:
            state = "READY"
        view._lifecycle_state = state
        view._refresh_states.append(state)

    view._refresh_interaction_gate = _refresh_gate
    return view, app, frame


def _assert_immutable_view_ready_gate_is_independent_of_edit_backend() -> None:
    """Only-diff may use exact immutable data, never lazy edit readiness."""
    view, app, _frame = _make_view()
    assert app._edit_workbooks_ready() is False
    assert sm.SheetView._is_exact_immutable_view_ready(view) is True
    assert sm.SheetView._derive_lifecycle_state(view) == "EDIT_DEFERRED"
    assert not app.forbidden_hits, app.forbidden_hits

    # Use the actual gate projection for the positive: only-diff is enabled
    # from the exact immutable surface, while edit-gated controls stay locked
    # (ordinary operation buttons remain explainably clickable by design).
    view.only_diff_cb = _Widget()
    view.force_align_cb = _Widget()
    view.use_left_btn = _Widget()
    view.save_a_btn = _Widget()
    view.manual_rescan_btn = _Widget()
    view._lifecycle_generation = 0
    view._update_sheet_role_labels = lambda: None
    view._refresh_column_action_buttons = lambda: None
    sm.SheetView._refresh_interaction_gate(view)
    assert view._lifecycle_state == "EDIT_DEFERRED"
    assert view.only_diff_cb.cget("state") == "normal"
    assert view.force_align_cb.cget("state") == "disabled"
    assert app.recalc_btn.cget("state") == "disabled"
    assert view.manual_rescan_btn.cget("state") == "disabled"
    assert view.use_left_btn.cget("state") == "normal"
    assert view.save_a_btn.cget("state") == "normal"
    assert not app.forbidden_hits, app.forbidden_hits

    blockers = (
        ("not-current", lambda view, app: setattr(app, "exact", False)),
        ("unresolved", lambda view, app: setattr(app, "exact_state", sm._SHEET_EXACT_UNRESOLVED)),
        ("closing", lambda view, app: setattr(app, "_is_closing", True)),
        ("hidden", lambda view, app: setattr(app, "selected_sheet", "other")),
        ("lifecycle-error", lambda view, app: setattr(view, "_lifecycle_error", "failed")),
        ("canceled", lambda view, app: setattr(view, "_lifecycle_canceled", True)),
        ("interactive", lambda view, app: setattr(app._interactive_action_event, "value", True)),
        ("mode-switch", lambda view, app: setattr(view, "_mode_switch_pending", True)),
        ("pending-render", lambda view, app: setattr(view, "_pending_exact_render", True)),
        ("only-diff-build", lambda view, app: setattr(view, "_only_diff_async_building", True)),
        ("virtual-publishing", lambda view, app: setattr(view, "_virtual_publishing", True)),
        ("prepared-publish", lambda view, app: setattr(view, "_prepared_cache_publish_active", True)),
        ("not-data-ready", lambda view, app: setattr(view, "_data_ready", False)),
        ("not-prepared", lambda view, app: setattr(view, "_prepared_complete", False)),
        ("row-model", lambda view, app: setattr(view, "_row_model_exact", False)),
        ("formula-cache", lambda view, app: setattr(view, "_cache_formula_aware", False)),
        ("pair-diff", lambda view, app: setattr(view, "_pair_diff_full_exact", False)),
        ("column-map", lambda view, app: setattr(view, "_column_mapping_is_current", lambda: False)),
        (
            "requested-only-diff-without-exact-cache",
            lambda view, app: (
                view.only_diff_var.set(1),
                setattr(view, "_only_diff_rows_exact", False),
            ),
        ),
    )
    for name, mutate in blockers:
        blocked_view, blocked_app, _blocked_frame = _make_view()
        mutate(blocked_view, blocked_app)
        assert sm.SheetView._is_exact_immutable_view_ready(blocked_view) is False, name
        assert not blocked_app.forbidden_hits, (name, blocked_app.forbidden_hits)

    base_view, base_app, _base_frame = _make_view()
    base_app.has_base = True
    base_view._is_three_way_enabled = lambda: True
    base_view._base_diff_full_exact = False
    assert sm.SheetView._is_exact_immutable_view_ready(base_view) is False


def _assert_cache_only_publisher_preserves_logical_state() -> None:
    view, app, _frame = _make_view()
    view.only_diff_var.set(0)
    view._mode_switch_pending = True
    view._mode_switch_seq = 17

    # Guard against accidental reintroduction of any producer-path call.
    originals = {
        name: getattr(sm, name)
        for name in (
            "_stream_selected_sheet_snapshot",
            "_align_selected_sheet_snapshots",
            "_compare_selected_sheet_snapshots",
        )
    }

    def _forbidden(*_args, **_kwargs):
        raise AssertionError("snapshot producer invoked by cache-only publisher")

    try:
        for name in originals:
            setattr(sm, name, _forbidden)
        published, reason = view._publish_cached_only_diff_mode_switch(0, 17)
    finally:
        for name, original in originals.items():
            setattr(sm, name, original)

    assert published and reason == "published", reason
    expected_rows = list(range(400))
    expected_start, expected_anchor = _reference_window_start(expected_rows, 330)
    assert view._published_rows == [expected_rows]
    # Pair 330 remains selected and visible after its logical index is mapped
    # from the old 100-row full window to the bounded full result window.
    assert expected_anchor == 330
    assert view._virtual_window_start == expected_start, view._virtual_window_start
    assert view.display_rows == expected_rows[
        expected_start : expected_start + sm._VIRTUAL_VIEWPORT_MAX_ROWS
    ]
    assert 330 in view.display_rows
    assert view._restored_selection == [(330, 9)], view._restored_selection
    assert view._virtual_column_window_start == 7
    assert view._cleared_selection == 1 and view._cleared_hover == 1
    assert not app.forbidden_hits, app.forbidden_hits


def _assert_missing_anchor_uses_independent_target_fallback() -> None:
    view, app, _frame = _make_view()
    target_rows = list(range(0, 400, 5))
    view.only_diff_var.set(1)
    view._mode_switch_pending = True
    view._mode_switch_seq = 23
    view._snapshot_explicit_selection_state = lambda: {"pair_idx": 331, "main_col": 9}
    expected_start, expected_anchor = _reference_window_start(target_rows, 331)
    assert expected_anchor == 335  # first prepared only-diff row >= missing anchor
    published, reason = view._publish_cached_only_diff_mode_switch(1, 23)
    assert published and reason == "published"
    assert view._published_rows == [target_rows]
    assert view._virtual_window_start == expected_start
    assert view.display_rows == target_rows[
        expected_start : expected_start + sm._VIRTUAL_VIEWPORT_MAX_ROWS
    ]
    assert expected_anchor in view.display_rows
    # The original selected pair is not a target row, so it must not be
    # remapped to an unrelated line; the cached publisher clears it instead.
    assert view._restored_selection == [(None, None)]
    last_start, last_anchor = _reference_window_start(target_rows, 999)
    assert last_anchor == target_rows[-1]
    assert last_start == max(0, len(target_rows) - sm._VIRTUAL_VIEWPORT_MAX_ROWS)
    assert not app.forbidden_hits

    tail_view, tail_app, _tail_frame = _make_view()
    tail_view.only_diff_var.set(1)
    tail_view._mode_switch_pending = True
    tail_view._mode_switch_seq = 24
    tail_view._snapshot_explicit_selection_state = lambda: {"pair_idx": 999, "main_col": 9}
    tail_published, tail_reason = tail_view._publish_cached_only_diff_mode_switch(1, 24)
    assert tail_published and tail_reason == "published"
    assert tail_view._virtual_window_start == last_start
    assert tail_view.display_rows == target_rows[
        last_start : last_start + sm._VIRTUAL_VIEWPORT_MAX_ROWS
    ]
    assert last_anchor in tail_view.display_rows
    assert tail_view._restored_selection == [(None, None)]
    assert not tail_app.forbidden_hits


def _assert_transient_terminal_and_failure_gates() -> None:
    transient_cases = (
        ("pending", lambda view, app: setattr(view, "_pending_exact_render", True)),
        ("viewport-publishing", lambda view, app: setattr(view, "_virtual_publishing", True)),
        ("tab-away", lambda view, app: setattr(app, "selected_sheet", "other")),
        ("new-generation", lambda view, app: setattr(app, "exact", False)),
        ("closing", lambda view, app: setattr(app, "_is_closing", True)),
    )
    for name, mutate in transient_cases:
        view, app, frame = _make_view()
        view.only_diff_var.set(0)
        view._last_only_diff_value = 1
        assert view._schedule_cached_only_diff_mode_switch(0)
        mutate(view, app)
        frame.run_all_including_canceled()
        assert not view._published_rows, (name, view._published_rows)
        assert view._lifecycle_error is None, (name, view._lifecycle_error)
        assert not view._terminal_surfaces, (name, view._terminal_surfaces)
        assert not app.forbidden_hits, (name, app.forbidden_hits)
        assert app.releases == ["released"], (name, app.releases)
        if name != "closing":
            assert view._pending_exact_render is True, name
            assert view._lifecycle_state == "DIFFING", (name, view._refresh_states)

    terminal_view, terminal_app, terminal_frame = _make_view()
    terminal_view.only_diff_var.set(0)
    terminal_app.exact = False
    terminal_app.exact_state = sm._SHEET_EXACT_UNRESOLVED
    assert terminal_view._schedule_cached_only_diff_mode_switch(0)
    terminal_frame.run_all_including_canceled()
    assert not terminal_view._published_rows
    assert terminal_view._lifecycle_error is None
    assert terminal_view._terminal_surfaces
    assert terminal_view._lifecycle_state == "UNRESOLVED"
    assert not terminal_app.forbidden_hits

    failure_cases = (
        ("raw-incomplete", lambda view: view.pair_raw_parts_a.pop(12)),
        ("publisher-error", lambda view: setattr(
            view,
            "_publish_prepared_cache_surface",
            lambda **_kwargs: (_ for _ in ()).throw(RuntimeError("publisher")),
        )),
    )
    for name, mutate in failure_cases:
        view, app, frame = _make_view()
        view.only_diff_var.set(0)
        view._last_only_diff_value = 1
        mutate(view)
        assert view._schedule_cached_only_diff_mode_switch(0)
        frame.run_all_including_canceled()
        assert not view._published_rows, (name, view._published_rows)
        assert view._lifecycle_state == "FAILED", (name, view._refresh_states)
        assert view._lifecycle_error and "未发布" in view._lifecycle_error
        assert view.only_diff_var.get() == 1, (name, view.only_diff_var.get())
        assert view._terminal_surfaces, (name, view._terminal_surfaces)
        assert not app.forbidden_hits, (name, app.forbidden_hits)
        assert app.releases == ["released"], (name, app.releases)


def _assert_edited_sheet_bypasses_raw_cache_publisher() -> None:
    view, app, _frame = _make_view()
    view.only_diff_var.set(0)
    view._last_only_diff_value = 1
    view.touched_rows.add(101)
    view._rendered_semantics = "edited-overlay"

    def _legacy_refresh(*, rescan: bool) -> None:
        assert rescan is False
        view._legacy_refreshes.append(rescan)
        # The established edit-ready path owns the current overlay value.  A
        # stale raw cache publisher would replace this marker with "old-raw".
        assert view._rendered_semantics == "edited-overlay"

    view._refresh_mode_switch_preserving_selection = _legacy_refresh
    view._derive_lifecycle_state = lambda: "READY"
    view._persist_only_diff_setting_debounced = lambda: None
    sm.SheetView._toggle_only_diff(view)
    assert view._legacy_refreshes == [False]
    assert not view._published_rows
    assert view._rendered_semantics == "edited-overlay"
    assert view._last_only_diff_value == 0
    assert not app.forbidden_hits


def _assert_scheduler_coalesces_and_stale_token_cannot_publish() -> None:
    view, app, frame = _make_view()
    view.only_diff_var.set(0)
    view._last_only_diff_value = 1
    assert view._schedule_cached_only_diff_mode_switch(0)
    first_seq = view._mode_switch_seq
    view.only_diff_var.set(1)
    assert view._schedule_cached_only_diff_mode_switch(1)
    assert view._mode_switch_seq > first_seq
    # Deliberately execute the canceled callback too: its sequence guard must
    # make it a no-op, then the newest callback alone publishes the new mode.
    frame.run_all_including_canceled()
    assert len(view._published_rows) == 1, view._published_rows
    assert view._published_rows[0] == list(range(0, 400, 5))
    assert view._last_only_diff_value == 1
    assert view._lifecycle_state == "READY", view._refresh_states
    assert app.releases == ["released", "released"], app.releases

    stale_view, stale_app, stale_frame = _make_view()
    stale_view.only_diff_var.set(0)
    assert stale_view._schedule_cached_only_diff_mode_switch(0)
    stale_view._mode_switch_seq += 1  # a newer owner invalidated the ticket
    stale_frame.run_all_including_canceled()
    assert not stale_view._published_rows
    assert not stale_app.forbidden_hits



def _async_result_payload(view, *, build_seq: int, error: str | None = None) -> dict:
    """Deterministic immutable worker payload; no workbook or comparator input."""
    rows = [5, 10]
    return {
        "build_seq": int(build_seq),
        "build_key": view._current_only_diff_cache_key(),
        "sheet": view.sheet,
        "error": error,
        "diff_pair_indices": list(rows),
        "pair_diff_cols": {5: {2}, 10: {3}},
        "pair_base_diff_cols": {},
        "pair_parts_a": {pair_idx: view.pair_raw_parts_a[pair_idx] for pair_idx in rows},
        "pair_parts_b": {pair_idx: view.pair_raw_parts_b[pair_idx] for pair_idx in rows},
        "pair_parts_base": {},
    }


def _assert_async_result_apply_uses_cache_only_scheduler() -> None:
    """Worker completion must publish exact rows without refresh or worksheet I/O."""
    view, app, frame = _make_view()
    view.only_diff_var.set(1)
    view._last_only_diff_value = 0
    view._only_diff_async_building = True
    view._only_diff_async_build_key = view._current_only_diff_cache_key()
    original_legacy = view._refresh_mode_switch_preserving_selection
    original_refresh = getattr(view, "refresh")

    def _forbidden_legacy(*_args, **_kwargs):
        raise AssertionError("async only-diff completion entered legacy refresh")

    try:
        view._refresh_mode_switch_preserving_selection = _forbidden_legacy
        view.refresh = _forbidden_legacy
        result = sm.SheetView._apply_async_only_diff_result(
            view,
            _async_result_payload(view, build_seq=41),
            build_seq=41,
            has_base=False,
        )
        assert result == "publish-scheduled", result
        assert view._mode_switch_pending is True
        assert view._published_rows == []
        frame.run_all_including_canceled()
    finally:
        view._refresh_mode_switch_preserving_selection = original_legacy
        view.refresh = original_refresh

    assert view._published_rows == [[5, 10]], view._published_rows
    assert view._installed_async_diff_rows == (5, 10)
    assert view._staged_async_parts[0] == {
        5: ("a", 5), 10: ("a", 10),
    }
    assert view._materialized_async_parts is True
    assert view._invalidated_async_render == 1
    assert view._last_only_diff_value == 1
    assert view._hid_async_loading is True
    assert app.exact_updates[-1][1] == sm._SHEET_EXACT_CHANGED
    assert app.progress_outcomes[-1] == ("S1", 41, "success")
    assert app.status_updates == 1 and app.nav_updates == 1
    assert not app.forbidden_hits, app.forbidden_hits
    assert view._legacy_refreshes == []


def _assert_async_result_apply_rejects_stale_hidden_and_failure_paths() -> None:
    # Old worker sequence cannot alter cache, viewport, or progress ownership.
    stale_view, stale_app, stale_frame = _make_view()
    stale_view.only_diff_var.set(1)
    stale_view._only_diff_async_building = True
    stale = sm.SheetView._apply_async_only_diff_result(
        stale_view,
        _async_result_payload(stale_view, build_seq=40),
        build_seq=40,
        has_base=False,
    )
    assert stale == "stale-build-seq"
    assert not stale_view._published_rows and not stale_frame.callbacks
    assert not stale_app.progress_outcomes and not stale_app.forbidden_hits

    # A selected-sheet tab handoff rejects the queued publisher rather than
    # publishing a result into the wrong tab or falling back to refresh.
    tab_view, tab_app, tab_frame = _make_view()
    tab_view.only_diff_var.set(1)
    tab_view._only_diff_async_building = True
    result = sm.SheetView._apply_async_only_diff_result(
        tab_view,
        _async_result_payload(tab_view, build_seq=41),
        build_seq=41,
        has_base=False,
    )
    assert result == "publish-scheduled"
    tab_app.selected_sheet = "other"
    tab_frame.run_all_including_canceled()
    assert not tab_view._published_rows
    assert tab_view._pending_exact_render is True
    assert tab_app.progress_outcomes[-1] == ("S1", 41, "publish-cancel")
    assert not tab_app.forbidden_hits

    # A hidden exact result is retained as cache-only pending render; no Tk
    # publisher, worksheet read, or fake success is permitted.
    hidden_view, hidden_app, hidden_frame = _make_view()
    hidden_view.only_diff_var.set(1)
    hidden_view._only_diff_async_building = True
    hidden_app.selected_sheet = "other"
    hidden = sm.SheetView._apply_async_only_diff_result(
        hidden_view,
        _async_result_payload(hidden_view, build_seq=41),
        build_seq=41,
        has_base=False,
    )
    assert hidden == "cached-hidden"
    assert hidden_view._pending_exact_render is True
    assert not hidden_view._published_rows and not hidden_frame.callbacks
    assert hidden_app.progress_outcomes[-1] == ("S1", 41, "cached-hidden")
    assert not hidden_app.forbidden_hits

    # Result errors and publisher rejection are terminal/fail-closed; neither
    # may substitute a worksheet refresh or a success outcome.
    error_view, error_app, error_frame = _make_view()
    error_view.only_diff_var.set(1)
    error_view._only_diff_async_building = True
    error_before = {
        "pair_diff_cols": {
            int(pair_idx): set(cols)
            for pair_idx, cols in error_view.pair_diff_cols.items()
        },
        "pair_base_diff_cols": {
            int(pair_idx): set(cols)
            for pair_idx, cols in error_view.pair_base_diff_cols.items()
        },
        "pair_raw_parts_a": dict(error_view.pair_raw_parts_a),
        "pair_raw_parts_b": dict(error_view.pair_raw_parts_b),
        "pair_raw_parts_base": dict(error_view.pair_raw_parts_base),
        "only_diff_rows_cache": list(error_view._only_diff_rows_cache),
        "only_diff_rows_exact": bool(error_view._only_diff_rows_exact),
        "published_rows": list(error_view._published_rows),
    }
    worker_error = sm.SheetView._apply_async_only_diff_result(
        error_view,
        _async_result_payload(error_view, build_seq=41, error="fixture-error"),
        build_seq=41,
        has_base=False,
    )
    assert worker_error == "worker-error"
    assert not error_view._published_rows and not error_frame.callbacks
    assert error_app._sheet_exact_entry("S1")["state"] == sm._SHEET_EXACT_FAILED
    assert error_app.exact_state == sm._SHEET_EXACT_FAILED
    assert error_view._only_diff_async_building is False
    assert error_view._only_diff_async_build_key is None
    assert error_view._only_diff_preview_full is False
    assert error_view._pending_exact_render is False
    assert error_view._lifecycle_error == "fixture-error"
    assert error_view.pair_diff_cols == error_before["pair_diff_cols"]
    assert error_view.pair_base_diff_cols == error_before["pair_base_diff_cols"]
    assert error_view.pair_raw_parts_a == error_before["pair_raw_parts_a"]
    assert error_view.pair_raw_parts_b == error_before["pair_raw_parts_b"]
    assert error_view.pair_raw_parts_base == error_before["pair_raw_parts_base"]
    assert error_view._only_diff_rows_cache == error_before["only_diff_rows_cache"]
    assert error_view._only_diff_rows_exact is error_before["only_diff_rows_exact"]
    assert error_view._published_rows == error_before["published_rows"]
    assert error_app.progress_outcomes == [("S1", 41, "worker-error")]
    assert not error_app.forbidden_hits

    failed_view, failed_app, failed_frame = _make_view()
    failed_view.only_diff_var.set(1)
    failed_view._only_diff_async_building = True
    failed_view._publish_prepared_cache_surface = lambda **_kwargs: False
    failed = sm.SheetView._apply_async_only_diff_result(
        failed_view,
        _async_result_payload(failed_view, build_seq=41),
        build_seq=41,
        has_base=False,
    )
    assert failed == "publish-scheduled"
    failed_frame.run_all_including_canceled()
    assert not failed_view._published_rows
    assert failed_app.exact_state == sm._SHEET_EXACT_FAILED
    assert failed_app.progress_outcomes[-1] == ("S1", 41, "publish-failure")
    assert failed_view._terminal_surfaces
    assert not failed_app.forbidden_hits

def _assert_async_invalid_and_completion_failures_release_owner() -> None:
    """Every result-completion failure releases build/progress ownership once."""
    invalid_view, invalid_app, _invalid_frame = _make_view()
    invalid_view.only_diff_var.set(1)
    invalid_view._last_only_diff_value = 0
    invalid_view._only_diff_request_origin_value = 0
    invalid_view._only_diff_async_building = True
    invalid_view._only_diff_async_build_key = invalid_view._current_only_diff_cache_key()
    invalid_view._only_diff_preview_full = True
    invalid_view._pending_exact_render = True
    invalid_app._only_diff_progress_owner = (invalid_view, 41)
    invalid = sm.SheetView._apply_async_only_diff_result(
        invalid_view,
        object(),
        build_seq=41,
        has_base=False,
    )
    assert invalid == "invalid-payload"
    assert invalid_app.exact_state == sm._SHEET_EXACT_FAILED
    assert invalid_app.progress_outcomes == [("S1", 41, "invalid-payload")]
    assert invalid_app._only_diff_progress_owner is None
    assert invalid_view._only_diff_async_building is False
    assert invalid_view._only_diff_async_build_key is None
    assert invalid_view._only_diff_preview_full is False
    assert invalid_view._pending_exact_render is False
    assert invalid_view.only_diff_var.get() == 0
    assert not invalid_app.forbidden_hits

    # The actual production starter must be eligible again after the invalid
    # payload has released its old build flag; submit is a pure broker spy.
    invalid_view._has_valid_only_diff_snapshot_cache = lambda: False
    invalid_view._set_only_diff_pending_info = lambda: None
    retry_before_seq = invalid_view._only_diff_async_build_seq
    retry = sm.SheetView._start_async_large_only_diff_build(
        invalid_view,
        user_initiated=False,
    )
    assert retry is True
    assert retry_before_seq == 41
    assert invalid_view._only_diff_async_build_seq == retry_before_seq + 1 == 42
    assert invalid_app.exact_updates[-1][0:2] == (
        "S1",
        sm._SHEET_EXACT_CALCULATING,
    )
    assert invalid_app._sheet_exact_entry("S1")["generation"] == 0
    assert invalid_view._only_diff_async_building is True
    assert invalid_view._only_diff_async_build_key == invalid_view._current_only_diff_cache_key()
    assert invalid_view._only_diff_async_build_key
    assert getattr(invalid_view, "_only_diff_async_prior_exact", None) is None
    assert invalid_app.priority_claims == [(invalid_view, 42)]
    assert len(invalid_app.priority_submissions) == 1
    assert invalid_app.priority_submissions[0][0:2] == (invalid_view, 42)

    # Completion runs under a transaction: each individual UI owner may fail,
    # but the current cache result must become FAILED and release progress once.
    for name, owner_name in (
        ("hide", "view"),
        ("status", "app"),
        ("sheet-nav", "app"),
        ("persist", "view"),
    ):
        view, app, frame = _make_view()
        view.only_diff_var.set(1)
        view._only_diff_async_building = True
        view._only_diff_async_build_key = view._current_only_diff_cache_key()
        app._only_diff_progress_owner = (view, 41)

        def _raise_completion_failure(*_args, **_kwargs):
            raise RuntimeError(name)

        if owner_name == "view":
            attr = {
                "hide": "_hide_loading",
                "persist": "_persist_only_diff_setting_debounced",
            }[name]
            previous = getattr(view, attr)
            setattr(view, attr, _raise_completion_failure)
        else:
            attr = {"status": "_update_exact_status_ui", "sheet-nav": "refresh_sheet_nav"}[name]
            previous = getattr(app, attr)
            setattr(app, attr, _raise_completion_failure)
        try:
            scheduled = sm.SheetView._apply_async_only_diff_result(
                view,
                _async_result_payload(view, build_seq=41),
                build_seq=41,
                has_base=False,
            )
            assert scheduled == "publish-scheduled", (name, scheduled)
            frame.run_all_including_canceled()
        finally:
            if owner_name == "view":
                setattr(view, attr, previous)
            else:
                setattr(app, attr, previous)
        assert app.exact_state == sm._SHEET_EXACT_FAILED, name
        assert view._only_diff_async_building is False, name
        assert view._only_diff_async_build_key is None, name
        assert view._only_diff_preview_full is False, name
        assert view._pending_exact_render is False, name
        assert app._only_diff_progress_owner is None, name
        assert app.progress_outcomes == [("S1", 41, "publish-completion-failed")], (
            name,
            app.progress_outcomes,
        )
        assert view._legacy_refreshes == [], name
        assert not app.forbidden_hits, (name, app.forbidden_hits)

    # The scheduler itself must not silently discard a foreign completion
    # exception either.  With no specialized error owner, it records FAILED.
    callback_view, callback_app, callback_frame = _make_view()
    callback_view.only_diff_var.set(0)

    def _raising_callback(*_args, **_kwargs):
        raise RuntimeError("callback")

    assert callback_view._schedule_cached_only_diff_mode_switch(
        0,
        on_complete=_raising_callback,
    )
    callback_frame.run_all_including_canceled()
    assert callback_view._mode_switch_completion is None
    assert callback_app.exact_state == sm._SHEET_EXACT_FAILED
    assert callback_view._lifecycle_error and "回调异常" in callback_view._lifecycle_error
    assert not callback_app.forbidden_hits


def _assert_async_three_way_base_payload_contract() -> None:
    """Three-way async publication retains Base fragments or rejects partial data."""
    view, app, frame = _make_view()
    app.has_base = True
    view._is_three_way_enabled = lambda: True
    view.only_diff_var.set(1)
    view._only_diff_async_building = True
    view._only_diff_async_build_key = view._current_only_diff_cache_key()
    payload = _async_result_payload(view, build_seq=41)
    payload["pair_base_diff_cols"] = {5: {2}, 10: set()}
    payload["pair_parts_base"] = {5: ("base", 5), 10: ("base", 10)}
    published = sm.SheetView._apply_async_only_diff_result(
        view,
        payload,
        build_seq=41,
        has_base=True,
    )
    assert published == "publish-scheduled"
    frame.run_all_including_canceled()
    assert view._base_diff_full_exact is True
    assert view._staged_async_parts[2] == payload["pair_parts_base"]
    assert view.pair_raw_parts_base[5] == ("base", 5)
    assert view.pair_raw_parts_base[10] == ("base", 10)
    assert view.pair_base_diff_cols[5] == {2}
    assert app.progress_outcomes[-1] == ("S1", 41, "success")
    assert not app.forbidden_hits

    missing_view, missing_app, missing_frame = _make_view()
    missing_app.has_base = True
    missing_view._is_three_way_enabled = lambda: True
    missing_view.only_diff_var.set(1)
    # The incomplete payload must not partially replace an earlier immutable
    # cache/map before its Base completeness gate rejects publication.
    missing_view.pair_diff_cols = {99: {1}}
    missing_view.pair_base_diff_cols = {99: {2}}
    missing_view.pair_raw_parts_base = {99: ("old-base", 99)}
    missing_view._only_diff_async_building = True
    missing_view._only_diff_async_build_key = missing_view._current_only_diff_cache_key()
    missing_payload = _async_result_payload(missing_view, build_seq=41)
    missing_payload["pair_base_diff_cols"] = {5: {2}, 10: set()}
    missing_payload["pair_parts_base"] = {5: ("base", 5)}
    missing = sm.SheetView._apply_async_only_diff_result(
        missing_view,
        missing_payload,
        build_seq=41,
        has_base=True,
    )
    assert missing == "base-parts-incomplete"
    assert not missing_frame.callbacks and not missing_view._published_rows
    assert missing_app.exact_state == sm._SHEET_EXACT_FAILED
    assert missing_app.progress_outcomes == [("S1", 41, "base-parts-incomplete")]
    assert missing_view._only_diff_async_building is False
    assert missing_view._only_diff_async_build_key is None
    assert missing_view.pair_diff_cols == {99: {1}}
    assert missing_view.pair_base_diff_cols == {99: {2}}
    assert missing_view.pair_raw_parts_base == {99: ("old-base", 99)}
    assert missing_view._pending_pair_parts_cache is None
    assert not hasattr(missing_view, "_staged_async_parts")
    assert not missing_app.forbidden_hits


def main() -> None:
    _assert_immutable_view_ready_gate_is_independent_of_edit_backend()
    _assert_cache_only_publisher_preserves_logical_state()
    _assert_missing_anchor_uses_independent_target_fallback()
    _assert_transient_terminal_and_failure_gates()
    _assert_edited_sheet_bypasses_raw_cache_publisher()
    _assert_scheduler_coalesces_and_stale_token_cannot_publish()
    _assert_async_result_apply_uses_cache_only_scheduler()
    _assert_async_result_apply_rejects_stale_hidden_and_failure_paths()
    _assert_async_invalid_and_completion_failures_release_owner()
    _assert_async_three_way_base_payload_contract()
    print("SMOKE_CACHED_ONLY_DIFF_MODE_SWITCH_OK", flush=True)


if __name__ == "__main__":
    main()
