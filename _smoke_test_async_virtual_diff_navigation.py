"""Headless contracts for requestized offscreen virtual diff navigation.

No workbook, Tk root, or parser is constructed.  The test drives the real
viewport request/scheduler/navigation helpers with a fake ``after(0)`` owner
and a bounded publisher seam, so stale and failure terminals cannot be hidden
by a model-only coordinator.
"""

from __future__ import annotations

from collections import deque

import sow_merge_tool as sm


class _Button:
    def __init__(self, state: str = "disabled") -> None:
        self.state = state

    def cget(self, name: str) -> str:
        assert name == "state"
        return self.state


class _Text:
    """Minimum Text seam for the real strict deferred finalizer."""

    def __init__(self) -> None:
        self._x = 0.0
        self.fail_late_tag = False
        self.fail_xview_moveto = False
        self.state = "disabled"
        self.tag_add_calls = 0
        self.xview_moveto_calls = 0

    def mark_set(self, _mark: str, _index: str) -> None:
        return

    def see(self, _index: str) -> None:
        return

    def xview(self):
        return (self._x, 1.0)

    def xview_moveto(self, fraction: float) -> None:
        self.xview_moveto_calls += 1
        if self.fail_xview_moveto:
            raise RuntimeError("injected-late-c-x-restore")
        self._x = float(fraction)

    def tag_remove(self, *_args) -> None:
        return

    def tag_add(self, *_args) -> None:
        self.tag_add_calls += 1
        if self.fail_late_tag:
            raise RuntimeError("injected-late-c-tag")
        return

    def configure(self, **options) -> None:
        self.state = str(options.get("state", self.state))

    def delete(self, *_args) -> None:
        return

    def insert(self, *_args) -> None:
        return


class _Scrollbar:
    def __init__(self) -> None:
        self.values = None
        self.set_calls = 0
        self.fail_set = False

    def set(self, first, last) -> None:
        self.set_calls += 1
        if self.fail_set:
            raise RuntimeError("injected-main-scrollbar-set")
        self.values = (float(first), float(last))
        return


class _Frame:
    def __init__(self, *, fail_after: bool = False) -> None:
        self.fail_after = bool(fail_after)
        self.callbacks = {}
        self.delays = []
        self.cancelled = []
        self._seq = 0

    def after(self, delay, callback):
        self.delays.append(int(delay))
        if self.fail_after:
            raise RuntimeError("injected-after-install-failure")
        self._seq += 1
        token = f"after-{self._seq}"
        self.callbacks[token] = callback
        return token

    def after_cancel(self, token) -> None:
        self.cancelled.append(token)
        self.callbacks.pop(token, None)

    def fire_one(self) -> None:
        assert len(self.callbacks) == 1, self.callbacks
        _token, callback = self.callbacks.popitem()
        callback()


class _App:
    def __init__(self) -> None:
        self.selected_sheet = "S1"
        self.generation = 0
        self.ws_calls = 0
        self.activity = []

    def _sheet_exact_entry(self, sheet: str) -> dict:
        assert sheet == "S1"
        return {"generation": self.generation}

    def _note_ui_activity(self, reason: str) -> None:
        self.activity.append(str(reason))

    def __getattr__(self, name: str):
        if name.startswith("ws_") or name.startswith("_request_edit_preload"):
            self.ws_calls += 1
            raise AssertionError(f"async navigation must not access {name}")
        raise AttributeError(name)


def _blocks():
    return [
        sm._DiffBlock(ordinal=1, pair_indices=(1, 2), start_pair_idx=1, end_pair_idx=2, pending=True),
        sm._DiffBlock(ordinal=2, pair_indices=(15, 16), start_pair_idx=15, end_pair_idx=16, pending=True),
    ]


def _make_view(*, fail_after: bool = False, virtual: bool = True):
    app = _App()
    frame = _Frame(fail_after=fail_after)
    view = sm.SheetView.__new__(sm.SheetView)
    view.app = app
    view.frame = frame
    view.sheet = "S1"
    view._viewport_request_seq = 0
    view._viewport_request_active = None
    view._viewport_request_samples_ms = deque(maxlen=128)
    view._viewport_request_completed = deque(maxlen=128)
    view._viewport_request_terminal = deque(maxlen=128)
    view._viewport_request_superseded = 0
    view._virtual_publish_after_id = None
    view._virtual_column_publish_after_id = None
    view._virtual_publish_token = 0
    view._virtual_pending_start = None
    view._virtual_pending_column_start = None
    view._staged_virtual_surface_active = False
    view._virtual_window_start = 0
    view._virtual_column_window_start = 0
    view._full_display_rows = list(range(40))
    view.display_rows = [0, 1, 2]
    view.row_to_line = {0: 1, 1: 2, 2: 3}
    view.selected_pair_idx = 1
    view._last_selected_line = None
    view._last_cursor_cmp_pair_idx = None
    view._c_area_last_render_key = None
    view.left = _Text()
    view.base = _Text()
    view.right = _Text()
    view.left_colhdr = _Text()
    view.base_colhdr = _Text()
    view.right_colhdr = _Text()
    view.cursor_cmp = _Text()
    view.cursor_cmp_colhdr = _Text()
    view.cell_cmp_text = _Text()
    view.cursor_hsb = _Scrollbar()
    view.cell_cmp_hsb = _Scrollbar()
    view.hsb_left = _Scrollbar()
    view.hsb_mid = _Scrollbar()
    view.hsb_right = _Scrollbar()
    view._xsyncing = False
    view.prev_diff_btn = _Button("disabled")
    view.next_diff_btn = _Button("normal")
    view._diff_navigation_telemetry = deque(maxlen=128)
    view._last_diff_navigation_telemetry = {}
    view._last_virtual_render_phases_ms = {}
    view._last_virtual_publication_telemetry = {}
    view._last_c_area_render_ms = 0.0
    view._last_strict_c_final_restore_ms = 0.0
    view._c_replace_count = 0
    view._c_area_render_samples_ms = deque(maxlen=128)
    view._c_area_same_row_skips = 0
    view._suppress_c_xsync = False
    view._data_version = 0
    view._column_projection_generation = 0
    view._virtual_column_window_generation = 0
    view._enable_c_cell = False
    view.row_pairs = [(pair_idx, pair_idx) for pair_idx in range(40)]
    view.pair_text_a = {}
    view.pair_text_b = {}
    view.pair_text_base = {}
    view._virtual_mode_active = lambda: bool(virtual)
    view._wide_column_virtual_active = lambda: False
    view._rendered_logical_columns = lambda: (1,)
    view._apply_pending_virtual_column_window = lambda: None
    view._ensure_full_diff_blocks = _blocks
    view._active_full_diff_block_index = lambda: (
        1 if int(getattr(view, "selected_pair_idx", -1)) in (15, 16) else 0
    )
    view._refresh_diff_block_ui = lambda: None
    view._normalize_pair_idx = lambda value: (
        int(value) if value is not None and 0 <= int(value) < 40 else None
    )
    view._virtual_window_rows = lambda start=None: list(
        range(
            int(view._virtual_window_start if start is None else start),
            min(
                len(view._full_display_rows),
                int(view._virtual_window_start if start is None else start)
                + min(sm._VIRTUAL_VIEWPORT_MAX_ROWS, len(view._full_display_rows)),
            ),
        )
    )
    view._is_three_way_enabled = lambda: False
    order = []

    view._logical_horizontal_first = lambda: float(view.left.xview()[0])

    def _map_x(_left, target, fraction):
        source = float(fraction)
        if target in (view.cursor_cmp, view.cursor_cmp_colhdr):
            return source * 0.5 + 0.1
        if target is view.cell_cmp_text:
            return source * 0.25 + 0.35
        return source

    view._map_xfirst_between_widgets = _map_x
    view._main_sync_calls = 0

    def _sync_main(fraction) -> None:
        view._main_sync_calls += 1
        first = float(fraction)
        panes = [
            (view.left, view.left_colhdr, view.hsb_left),
            (view.right, view.right_colhdr, view.hsb_right),
        ]
        if view._is_three_way_enabled():
            panes.insert(1, (view.base, view.base_colhdr, view.hsb_mid))
        for pane, header, _scrollbar in panes:
            pane.xview_moveto(first)
            if header is not None:
                header.xview_moveto(first)
        for pane, _header, scrollbar in panes:
            try:
                scrollbar.set(*pane.xview())
            except Exception:
                pass

    view._sync_main_x_to_frac = _sync_main
    view._sync_c_x_calls = 0

    def _sync_c_x(fraction):
        view._sync_c_x_calls += 1
        source = float(fraction)
        cursor_x = view._map_xfirst_between_widgets(view.left, view.cursor_cmp, source)
        cell_x = view._map_xfirst_between_widgets(view.left, view.cell_cmp_text, source)
        view.cursor_cmp.xview_moveto(cursor_x)
        view.cursor_cmp_colhdr.xview_moveto(cursor_x)
        view.cell_cmp_text.xview_moveto(cell_x)

    view._sync_c_x_to_frac = _sync_c_x
    view._pair_idx_for_line = lambda line: next(
        (pair for pair, mapped_line in view.row_to_line.items() if mapped_line == int(line)),
        None,
    )
    view._pair_for_line = lambda line: (view._pair_idx_for_line(line),) * 2
    view._row_for_side = lambda pair, _side: None if pair is None else pair[0]
    view.resolved_pair_idx_for_c_area = lambda: view.selected_pair_idx
    view._all_logical_diff_cols_for_pair = lambda _pair_idx: set()
    view._render_cursor_row_headers = lambda _pair, _three_way, *, strict=False: (
        True if strict else None
    )
    view._spans_for_line = lambda _text: {}
    view._apply_main_selected_cell_highlight = lambda: None

    def _replace_text_document(_widget, _text) -> None:
        view._c_replace_count += 1
        order.append("C")

    def _update_nav_state() -> None:
        block_idx = view._active_full_diff_block_index()
        blocks = view._ensure_full_diff_blocks()
        view.prev_diff_btn.state = "normal" if block_idx and block_idx > 0 else "disabled"
        view.next_diff_btn.state = (
            "normal"
            if block_idx is not None and int(block_idx) + 1 < len(blocks)
            else "disabled"
        )

    view._replace_text_document = _replace_text_document
    view._update_diff_nav_state = _update_nav_state

    def _highlight(line: int) -> None:
        order.append("selection")
        view._last_selected_line = int(line)

    view._highlight_selected_line = _highlight

    def _select(line, *, navigation_phases=None):
        order.append("selection")
        assert int(line) >= 1
        target = next(pair for pair, mapped_line in view.row_to_line.items() if mapped_line == int(line))
        view.selected_pair_idx = target
        view.prev_diff_btn.state = "normal"
        view.next_diff_btn.state = "disabled"
        navigation_phases.update({"selection": 1.0, "restore": 2.0, "ui": 3.0})
        order.append("C")

    view._goto_block_start = _select
    return view, app, frame, order


def _install_bounded_publisher(
    view,
    order,
    *,
    missing_line: bool = False,
    raise_error: bool = False,
    wrong_window: bool = False,
):
    def _publish(start: int):
        order.append("publish")
        if raise_error:
            raise RuntimeError("injected-publish-failure")
        view._virtual_window_start = int(start) + (1 if wrong_window else 0)
        rows = view._virtual_window_rows(view._virtual_window_start)
        if missing_line:
            rows = [pair_idx for pair_idx in rows if pair_idx != 15]
        view.display_rows = list(rows)
        view.row_to_line = {pair_idx: line for line, pair_idx in enumerate(rows, start=1)}
        result = sm.SheetView._finish_deferred_diff_navigation_after_publish(view)
        if result is True:
            sm.SheetView._complete_viewport_request_if_current(view)
            order.append("complete")
        return result

    view._publish_virtual_window = _publish


def _terminal(view):
    assert view._viewport_request_terminal
    return dict(view._viewport_request_terminal[-1])


def _strict_x_restore_context(view, source_x: float | None = None) -> dict:
    source = float(view.left.xview()[0] if source_x is None else source_x)
    return {
        "source_x": source,
        "cursor_x": float(
            view._map_xfirst_between_widgets(view.left, view.cursor_cmp, source)
        ),
        "cell_x": float(
            view._map_xfirst_between_widgets(view.left, view.cell_cmp_text, source)
        ),
    }


def _assert_offscreen_callback_is_after0_and_terminal_is_exact() -> None:
    view, app, frame, order = _make_view()
    view.left._x = 0.37
    _install_bounded_publisher(view, order)
    sm.SheetView._goto_full_diff_block(view, 1)
    assert order == []
    active = dict(view._viewport_request_active)
    assert active["reason"] == "diff-block"
    assert active["kind"] == "diff-navigation"
    assert active["navigation"]["target_pair_idx"] == 15
    assert active["navigation"]["block_idx"] == 1
    assert frame.delays == [0]
    assert len(frame.callbacks) == 1
    assert app.ws_calls == 0
    frame.fire_one()
    terminal = _terminal(view)
    assert terminal["status"] == "complete"
    assert terminal["counted"] is True
    assert terminal["surface_changed"] is True
    navigation = terminal["navigation"]
    assert navigation["status"] == "selected"
    assert navigation["actual_pair_idx"] == 15
    assert navigation["actual_block_idx"] == 1
    assert navigation["actual_button_states"] == {"prev": "normal", "next": "disabled"}
    assert view._c_area_last_render_key[0] == 15
    assert view._c_area_last_completed_render_key[0] == 15
    assert view._last_cursor_cmp_pair_idx == 15
    assert {
        "block_lookup", "materialize_publish", "selection", "restore",
        "restore_main", "c_final_restore", "ui", "total",
    } <= set(navigation["phase_ms"])
    expected_x = _strict_x_restore_context(view)
    assert view._sync_c_x_calls == 0
    assert abs(view.cursor_cmp.xview()[0] - expected_x["cursor_x"]) <= 0.02
    assert abs(view.cursor_cmp_colhdr.xview()[0] - expected_x["cursor_x"]) <= 0.02
    assert abs(view.cell_cmp_text.xview()[0] - expected_x["cell_x"]) <= 0.02
    assert view.cursor_cmp.xview_moveto_calls == 1
    assert view.cursor_cmp_colhdr.xview_moveto_calls == 1
    assert view.cell_cmp_text.xview_moveto_calls == 1
    assert order == ["publish", "selection", "C", "complete"]
    # A completed same-pair key is not enough to skip a newly supplied strict
    # source-x context: its C document must be redrawn and mapped again rather
    # than accepting the prior horizontal position as terminal evidence.
    replace_before = view._c_replace_count
    cursor_moves_before = view.cursor_cmp.xview_moveto_calls
    view.left._x = 0.51
    remapped = _strict_x_restore_context(view)
    assert sm.SheetView._update_cursor_lines(
        view,
        strict=True,
        strict_x_restore=remapped,
    ) is True
    assert view._c_replace_count == replace_before + 1
    assert view.cursor_cmp.xview_moveto_calls == cursor_moves_before + 1
    assert abs(view.cursor_cmp.xview()[0] - remapped["cursor_x"]) <= 0.02
    assert abs(view.cell_cmp_text.xview()[0] - remapped["cell_x"]) <= 0.02
    assert app.ws_calls == 0


def _assert_visible_and_nonvirtual_remain_synchronous() -> None:
    visible, _app, frame, order = _make_view()
    visible.display_rows = [0, 15, 16]
    visible.row_to_line = {0: 1, 15: 2, 16: 3}
    visible._materialize_pair_for_navigation = lambda pair_idx: int(pair_idx) == 15
    sm.SheetView._goto_full_diff_block(visible, 1)
    assert frame.callbacks == {}
    assert order == ["selection", "C"]
    assert visible._viewport_request_active is None

    nonvirtual, _app, frame, order = _make_view(virtual=False)
    nonvirtual._materialize_pair_for_navigation = lambda pair_idx: (
        nonvirtual.row_to_line.update({15: 1}) or int(pair_idx) == 15
    )
    sm.SheetView._goto_full_diff_block(nonvirtual, 1)
    assert frame.callbacks == {}
    assert order == ["selection", "C"]
    assert nonvirtual._viewport_request_active is None


def _assert_tail_target_uses_the_canonical_window_start() -> None:
    tail, _app, frame, order = _make_view()
    tail._ensure_full_diff_blocks = lambda: [
        sm._DiffBlock(
            ordinal=1,
            pair_indices=(39,),
            start_pair_idx=39,
            end_pair_idx=39,
            pending=True,
        )
    ]
    _install_bounded_publisher(tail, order)
    sm.SheetView._goto_full_diff_block(tail, 0)
    active = dict(tail._viewport_request_active)
    assert active["row_start"] == 20
    assert active["navigation"]["expected_row_start"] == 20
    frame.fire_one()
    terminal = _terminal(tail)
    assert terminal["status"] == "complete"
    assert terminal["actual_row_start"] == 20
    assert terminal["navigation"]["actual_pair_idx"] == 39
    assert terminal["navigation"]["actual_block_idx"] == 0
    assert terminal["navigation"]["actual_button_states"] == {
        "prev": "disabled",
        "next": "disabled",
    }
    assert order == ["publish", "selection", "C", "complete"]


def _assert_newer_requests_and_stale_inputs_never_select_old_target() -> None:
    view, app, frame, order = _make_view()
    _install_bounded_publisher(view, order)
    sm.SheetView._goto_full_diff_block(view, 1)
    sm.SheetView._begin_viewport_request(view, "wheel", row_start=0, column_start=0)
    first_terminal = _terminal(view)
    assert first_terminal["status"] == "superseded"
    assert first_terminal["navigation"]["status"] == "superseded"
    frame.fire_one()
    assert "selection" not in order

    stale, app, frame, order = _make_view()
    _install_bounded_publisher(stale, order)
    sm.SheetView._goto_full_diff_block(stale, 1)
    app.generation = 1
    frame.fire_one()
    assert _terminal(stale)["status"] == "superseded"
    assert "selection" not in order

    tab, app, frame, order = _make_view()
    _install_bounded_publisher(tab, order)
    sm.SheetView._goto_full_diff_block(tab, 1)
    app.selected_sheet = "Other"
    frame.fire_one()
    assert _terminal(tab)["status"] == "superseded"
    assert "selection" not in order

    closing, _app, frame, order = _make_view()
    _install_bounded_publisher(closing, order)
    sm.SheetView._goto_full_diff_block(closing, 1)
    sm.SheetView._cancel_virtual_2d_publish(closing, "close")
    assert _terminal(closing)["status"] == "superseded"
    assert frame.callbacks == {}
    assert "selection" not in order


def _assert_missing_publish_and_after_fail_closed() -> None:
    missing, _app, frame, order = _make_view()
    _install_bounded_publisher(missing, order, missing_line=True)
    sm.SheetView._goto_full_diff_block(missing, 1)
    frame.fire_one()
    failed = _terminal(missing)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-target-line-missing"
    assert "selection" not in order

    broken_publish, _app, frame, order = _make_view()
    _install_bounded_publisher(broken_publish, order, raise_error=True)
    sm.SheetView._goto_full_diff_block(broken_publish, 1)
    frame.fire_one()
    failed = _terminal(broken_publish)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "publish-exception"
    assert "selection" not in order

    broken_after, _app, _frame, order = _make_view(fail_after=True)
    _install_bounded_publisher(broken_after, order)
    sm.SheetView._goto_full_diff_block(broken_after, 1)
    failed = _terminal(broken_after)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "after-install-failed"
    assert "publish" not in order
    assert "selection" not in order

    wrong_window, _app, frame, order = _make_view()
    _install_bounded_publisher(wrong_window, order, wrong_window=True)
    sm.SheetView._goto_full_diff_block(wrong_window, 1)
    frame.fire_one()
    failed = _terminal(wrong_window)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-window-mismatch"
    assert "selection" not in order

    returned_false, _app, frame, order = _make_view()
    returned_false._publish_virtual_window = lambda _start: False
    sm.SheetView._goto_full_diff_block(returned_false, 1)
    frame.fire_one()
    failed = _terminal(returned_false)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "publish-returned-false"
    assert "selection" not in order


def _assert_staged_and_strict_selection_fail_closed() -> None:
    staged, _app, frame, order = _make_view()
    staged._staged_virtual_surface_active = True
    sm.SheetView._goto_full_diff_block(staged, 1)
    failed = _terminal(staged)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-staged-surface"
    assert frame.callbacks == {}
    assert "selection" not in order

    restore_failure, _app, frame, order = _make_view()
    _install_bounded_publisher(restore_failure, order)
    restore_failure.left._x = 0.37
    restore_failure.cursor_cmp.fail_xview_moveto = True
    sm.SheetView._goto_full_diff_block(restore_failure, 1)
    frame.fire_one()
    failed = _terminal(restore_failure)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-selection-restore-failed"
    assert "deferred navigation C surface did not finalize" in failed["failure_error"]
    assert "complete" not in order
    assert restore_failure._sync_c_x_calls == 0
    assert restore_failure._c_area_last_completed_render_key is None
    restore_failure.cursor_cmp.fail_xview_moveto = False
    cursor_moves_before = restore_failure.cursor_cmp.xview_moveto_calls
    assert sm.SheetView._update_cursor_lines(
        restore_failure,
        strict=True,
        strict_x_restore=_strict_x_restore_context(restore_failure),
    ) is True
    assert restore_failure.cursor_cmp.xview_moveto_calls == cursor_moves_before + 1
    assert restore_failure._c_area_last_completed_render_key[0] == 15

    c_failure, _app, frame, order = _make_view()
    _install_bounded_publisher(c_failure, order)
    # Exercise the real strict C renderer: the pair/key bookkeeping happens
    # before its Text tags, so this late tag failure proves the final success
    # signal is not inferred from those early bookkeeping fields.
    c_failure.cursor_cmp.fail_late_tag = True
    sm.SheetView._goto_full_diff_block(c_failure, 1)
    frame.fire_one()
    failed = _terminal(c_failure)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-selection-restore-failed"
    assert "deferred navigation C surface did not finalize" in failed["failure_error"]
    assert "complete" not in order
    assert c_failure._c_area_last_render_key is None
    assert c_failure._c_area_last_completed_render_key is None
    assert c_failure._last_cursor_cmp_pair_idx is None
    replace_before = c_failure._c_replace_count
    tag_before = c_failure.cursor_cmp.tag_add_calls
    c_failure.cursor_cmp.fail_late_tag = False
    assert sm.SheetView._update_cursor_lines(c_failure, strict=True) is True
    assert c_failure._c_replace_count == replace_before + 1
    assert c_failure.cursor_cmp.tag_add_calls > tag_before
    assert c_failure._c_area_last_completed_render_key[0] == 15

    x_failure, _app, frame, order = _make_view()
    _install_bounded_publisher(x_failure, order)
    x_failure.left.xview = lambda: (_ for _ in ()).throw(
        RuntimeError("injected-initial-main-x-read")
    )
    sm.SheetView._goto_full_diff_block(x_failure, 1)
    frame.fire_one()
    failed = _terminal(x_failure)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-selection-restore-failed"
    assert "injected-initial-main-x-read" in failed["failure_error"]
    assert "complete" not in order


def _assert_strict_main_x_fast_path_is_conservative() -> None:
    aligned, _app, frame, order = _make_view()
    _install_bounded_publisher(aligned, order)
    sm.SheetView._goto_full_diff_block(aligned, 1)
    frame.fire_one()
    terminal = _terminal(aligned)
    navigation = terminal["navigation"]
    assert terminal["status"] == "complete"
    assert navigation["main_x_fast_path"] is True
    assert navigation["main_x_observe_ms"] >= 0.0
    assert navigation["main_x_fallback_ms"] == 0.0
    assert aligned._main_sync_calls == 0
    for widget in (
        aligned.left,
        aligned.right,
        aligned.left_colhdr,
        aligned.right_colhdr,
    ):
        assert widget.xview_moveto_calls == 0
    assert aligned.hsb_left.values == aligned.left.xview()
    assert aligned.hsb_right.values == aligned.right.xview()

    text_mismatch, _app, frame, order = _make_view()
    text_mismatch.right._x = 0.25
    _install_bounded_publisher(text_mismatch, order)
    sm.SheetView._goto_full_diff_block(text_mismatch, 1)
    frame.fire_one()
    terminal = _terminal(text_mismatch)
    navigation = terminal["navigation"]
    assert terminal["status"] == "complete"
    assert navigation["main_x_fast_path"] is False
    assert navigation["main_x_fallback_ms"] >= 0.0
    assert text_mismatch._main_sync_calls == 1
    assert text_mismatch.right.xview() == text_mismatch.left.xview()
    assert text_mismatch.right_colhdr.xview()[0] == text_mismatch.right.xview()[0]

    header_mismatch, _app, frame, order = _make_view()
    header_mismatch.right_colhdr._x = 0.25
    _install_bounded_publisher(header_mismatch, order)
    sm.SheetView._goto_full_diff_block(header_mismatch, 1)
    frame.fire_one()
    assert _terminal(header_mismatch)["status"] == "complete"
    assert header_mismatch._main_sync_calls == 1

    persistent, _app, frame, order = _make_view()
    persistent.right.xview = lambda: (0.25, 1.0)
    _install_bounded_publisher(persistent, order)
    sm.SheetView._goto_full_diff_block(persistent, 1)
    frame.fire_one()
    failed = _terminal(persistent)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-selection-restore-failed"
    assert persistent._main_sync_calls == 1
    assert "complete" not in order

    bar_failure, _app, frame, order = _make_view()
    bar_failure.hsb_right.fail_set = True
    _install_bounded_publisher(bar_failure, order)
    sm.SheetView._goto_full_diff_block(bar_failure, 1)
    frame.fire_one()
    failed = _terminal(bar_failure)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-selection-restore-failed"
    assert bar_failure._main_sync_calls == 1
    assert "complete" not in order

    for label, header_attr, three_way in (
        ("mine", "left_colhdr", False),
        ("theirs", "right_colhdr", False),
        ("base", "base_colhdr", True),
    ):
        missing_header, app, frame, order = _make_view()
        if three_way:
            app.has_base = True
            missing_header._is_three_way_enabled = lambda: True
            missing_header._base_row_for_pair = (
                lambda pair_idx, _pair=None: int(pair_idx)
            )
        setattr(missing_header, header_attr, None)
        _install_bounded_publisher(missing_header, order)
        sm.SheetView._goto_full_diff_block(missing_header, 1)
        frame.fire_one()
        failed = _terminal(missing_header)
        assert failed["status"] == "failed", label
        assert failed["failure_reason"] == "navigation-selection-restore-failed", label
        assert missing_header._main_sync_calls == 1, label
        assert "complete" not in order, label

    zero_width, _app, frame, order = _make_view()
    zero_width.right.xview = lambda: (0.0, 0.0)
    _install_bounded_publisher(zero_width, order)
    sm.SheetView._goto_full_diff_block(zero_width, 1)
    frame.fire_one()
    failed = _terminal(zero_width)
    assert failed["status"] == "failed"
    assert failed["failure_reason"] == "navigation-selection-restore-failed"
    assert zero_width._main_sync_calls == 1
    assert "complete" not in order


def _assert_three_way_and_wide_strict_paths_remain_complete() -> None:
    three_way, app, frame, order = _make_view()
    app.has_base = True
    three_way.left._x = 0.37
    three_way._is_three_way_enabled = lambda: True
    three_way._base_row_for_pair = lambda pair_idx, _pair=None: int(pair_idx)
    _install_bounded_publisher(three_way, order)
    sm.SheetView._goto_full_diff_block(three_way, 1)
    frame.fire_one()
    terminal = _terminal(three_way)
    assert terminal["status"] == "complete"
    assert terminal["navigation"]["actual_pair_idx"] == 15
    assert three_way._sync_c_x_calls == 0
    assert three_way.cursor_cmp.xview_moveto_calls == 1
    assert three_way.cursor_cmp_colhdr.xview_moveto_calls == 1
    assert three_way.cell_cmp_text.xview_moveto_calls == 1

    three_fast, app, frame, order = _make_view()
    app.has_base = True
    three_fast._is_three_way_enabled = lambda: True
    three_fast._base_row_for_pair = lambda pair_idx, _pair=None: int(pair_idx)
    _install_bounded_publisher(three_fast, order)
    sm.SheetView._goto_full_diff_block(three_fast, 1)
    frame.fire_one()
    terminal = _terminal(three_fast)
    assert terminal["status"] == "complete"
    assert terminal["navigation"]["main_x_fast_path"] is True
    assert three_fast._main_sync_calls == 0
    assert three_fast.hsb_mid.values == three_fast.base.xview()

    wide, _app, frame, order = _make_view()
    wide._wide_column_virtual_active = lambda: True
    wide._set_wide_column_scrollbars_calls = 0
    wide._set_wide_column_scrollbars = lambda: setattr(
        wide,
        "_set_wide_column_scrollbars_calls",
        wide._set_wide_column_scrollbars_calls + 1,
    )
    wide._wide_main_sync_calls = 0
    wide._sync_main_x_to_frac = lambda _fraction: setattr(
        wide, "_wide_main_sync_calls", wide._wide_main_sync_calls + 1
    )
    _install_bounded_publisher(wide, order)
    sm.SheetView._goto_full_diff_block(wide, 1)
    frame.fire_one()
    terminal = _terminal(wide)
    assert terminal["status"] == "complete"
    assert terminal["navigation"]["actual_pair_idx"] == 15
    assert terminal["navigation"]["main_x_fast_path"] is False
    assert wide._wide_main_sync_calls == 1
    assert wide._set_wide_column_scrollbars_calls >= 1
    assert wide._sync_c_x_calls == 0
    assert "complete" in order


def main() -> None:
    tests = (
        _assert_offscreen_callback_is_after0_and_terminal_is_exact,
        _assert_visible_and_nonvirtual_remain_synchronous,
        _assert_tail_target_uses_the_canonical_window_start,
        _assert_newer_requests_and_stale_inputs_never_select_old_target,
        _assert_missing_publish_and_after_fail_closed,
        _assert_staged_and_strict_selection_fail_closed,
        _assert_strict_main_x_fast_path_is_conservative,
        _assert_three_way_and_wide_strict_paths_remain_complete,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}", flush=True)
    print(f"SMOKE_ASYNC_VIRTUAL_DIFF_NAVIGATION_OK ({len(tests)} tests)", flush=True)


if __name__ == "__main__":
    main()
