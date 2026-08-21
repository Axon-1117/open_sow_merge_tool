"""Pure contract checks for token-bound only-diff progress recovery."""

from __future__ import annotations

import copy
import threading
import types
from collections import deque

import sow_merge_tool as sm


class _IntVar:
    """Tiny IntVar stand-in for the public only-diff control."""

    def __init__(self, value=0):
        self.value = int(value)

    def get(self):
        return self.value

    def set(self, value):
        self.value = int(value)


class _StringVar:
    """StringVar-faithful progress text stand-in; never coerces stage text."""

    def __init__(self, value=""):
        self.value = str(value)

    def get(self):
        return self.value

    def set(self, value):
        self.value = str(value)


class _Widget:
    def __init__(self):
        self.options = {}

    def configure(self, **kwargs):
        self.options.update(kwargs)

    def cget(self, name):
        return self.options.get(name, "")


class _Root:
    def __init__(self, *, fail_after_calls=()):
        self.fail_after_calls = set(fail_after_calls)
        self.after_calls = 0
        self.callbacks = {}
        self.canceled = []

    def after(self, delay_ms, callback):
        self.after_calls += 1
        if self.after_calls in self.fail_after_calls:
            return None
        after_id = f"after-{self.after_calls}"
        self.callbacks[after_id] = (int(delay_ms), callback)
        return after_id

    def after_cancel(self, after_id):
        self.canceled.append(after_id)
        self.callbacks.pop(after_id, None)

    def run_delay(self, delay_ms):
        for after_id, (delay, callback) in tuple(self.callbacks.items()):
            if delay == int(delay_ms):
                self.callbacks.pop(after_id, None)
                callback()
                return after_id
        raise AssertionError(("missing delayed callback", delay_ms, self.callbacks))

    def winfo_rootx(self):
        return 0

    def winfo_rooty(self):
        return 0

    def winfo_width(self):
        return 900

    def winfo_height(self):
        return 700

    def state(self):
        return "normal"

    def winfo_ismapped(self):
        return True

    def winfo_viewable(self):
        return True


class _ProgressWindow:
    def __init__(
        self, *, raise_show=False, raise_grab=False, mapped=True, viewable=True
    ):
        self.raise_show = bool(raise_show)
        self.raise_grab = bool(raise_grab)
        self.mapped = bool(mapped)
        self.viewable = bool(viewable)
        self.normal = False
        self.grabbed = False
        self.lift_calls = 0
        self.grab_calls = 0
        self.focus_calls = 0
        self.on_grab = None

    def winfo_exists(self):
        return True

    def title(self, _value):
        return None

    def deiconify(self):
        if self.raise_show:
            raise RuntimeError("injected deiconify failure")
        self.normal = True

    def withdraw(self):
        self.normal = False
        self.grabbed = False

    def winfo_reqwidth(self):
        return 300

    def winfo_reqheight(self):
        return 130

    def geometry(self, _value):
        return None

    def lift(self):
        self.lift_calls += 1

    def update_idletasks(self):
        return None

    def grab_set(self):
        self.grab_calls += 1
        if self.raise_grab:
            raise RuntimeError("injected grab failure")
        self.grabbed = True
        if callable(self.on_grab):
            self.on_grab()

    def grab_release(self):
        self.grabbed = False

    def grab_current(self):
        return self if self.grabbed else None

    def focus_set(self):
        self.focus_calls += 1
        return None

    def winfo_ismapped(self):
        return self.mapped and self.normal

    def winfo_viewable(self):
        return self.viewable and self.normal

    def state(self):
        return "normal" if self.normal else "withdrawn"


class _Info:
    def configure(self, **_kwargs):
        return None


def _bind(app, *names):
    for name in names:
        setattr(app, name, types.MethodType(getattr(sm.SowMergeApp, name), app))


def _exact_entry(*, generation=0, state=sm._SHEET_EXACT_CHANGED):
    return {
        "generation": int(generation),
        "state": str(state),
        "stage": "prior exact terminal",
        "processed": 1,
        "total": 1,
        "reason": "proved before only-diff",
        "request_started_at": 1.0,
        "request_start_state": sm._SHEET_EXACT_CALCULATING,
        "full_detail_terminal": True,
        "full_detail_terminal_at": 2.0,
        "request_to_full_detail_ms": 1000.0,
    }


def _build(
    *, fail_after_calls=(), raise_show=False, raise_grab=False, mapped=True, viewable=True
):
    root = _Root(fail_after_calls=fail_after_calls)
    window = _ProgressWindow(
        raise_show=raise_show,
        raise_grab=raise_grab,
        mapped=mapped,
        viewable=viewable,
    )
    app = types.SimpleNamespace(
        root=root,
        _root_after_ids=set(),
        _is_closing=False,
        _only_diff_progress_owner=None,
        _only_diff_progress_win=window,
        _only_diff_progress_stage_var=_StringVar(),
        _only_diff_progress_detail_var=_StringVar(),
        _only_diff_progress_bar=_Widget(),
        _only_diff_progress_cancel_btn=_Widget(),
        _only_diff_progress_started=0.0,
        _only_diff_progress_show_after_id=None,
        _only_diff_progress_show_token=None,
        _only_diff_progress_watchdog_after_id=None,
        _only_diff_progress_watchdog_token=None,
        _only_diff_progress_confirm_after_id=None,
        _only_diff_progress_confirm_token=None,
        _only_diff_progress_visible_token=None,
        _only_diff_progress_visibility_attempts=deque(maxlen=32),
        _only_diff_progress_visibility_attempt_count=0,
        _only_diff_tab_states={},
        _priority_diff_lock=threading.Lock(),
        _priority_exact_owner=None,
        _exact_broker_lock=threading.Lock(),
        _exact_broker_pending=None,
        _only_diff_failure_handoff_lock=threading.Lock(),
        _only_diff_failure_handoffs=[],
        _sheet_compute_generation={},
        _sheet_exact_states={},
        _update_exact_status_ui=lambda: None,
        refresh_sheet_nav=lambda: None,
        _foreground_resume_on_exact_state=lambda *_args: None,
        sheet_views={},
        _queue_ui_task=lambda _callback: True,
    )
    _bind(
        app,
        "_safe_root_after",
        "_clear_only_diff_progress_schedule",
        "_fail_only_diff_progress_show",
        "_finish_only_diff_progress",
        "_begin_only_diff_progress",
        "_release_priority_exact",
        "_claim_priority_exact",
        "_enqueue_only_diff_failure_handoff",
        "_drain_only_diff_failure_handoffs",
        "_clear_only_diff_shutdown_state",
    )
    app._ensure_only_diff_progress_dialog = lambda: window

    def _sheet_exact_entry(sheet):
        entry = dict(app._sheet_exact_states.get(str(sheet)) or {})
        generation = int(app._sheet_compute_generation.get(str(sheet), -1))
        if int(entry.get("generation", -2)) != generation:
            return {"generation": generation, "state": sm._SHEET_EXACT_PENDING}
        return entry

    def _set_sheet_exact_state(
        sheet,
        state,
        *,
        stage="",
        reason="",
        only_diff_reopen_capability=None,
        **_kwargs,
    ):
        sheet = str(sheet)
        entry = dict(app._sheet_exact_states.get(sheet) or {})
        generation = int(app._sheet_compute_generation.get(sheet, -1))
        if (
            int(entry.get("generation", -1)) == generation
            and str(entry.get("state") or "") in sm._SHEET_EXACT_TERMINAL
            and str(state) == sm._SHEET_EXACT_CALCULATING
        ):
            ready_view = app.sheet_views.get(sheet)
            if (
                ready_view is not None
                and bool(getattr(ready_view, "_prepared_complete", False))
                and bool(getattr(ready_view, "_data_ready", False))
                and not bool(getattr(ready_view, "_pending_exact_render", False))
            ):
                try:
                    cap_view, cap_seq, cap_key, cap_generation, cap_prior = (
                        only_diff_reopen_capability
                    )
                    record = ready_view._only_diff_async_prior_exact
                    allowed = bool(
                        cap_view is ready_view
                        and str(getattr(cap_view, "sheet", "")) == sheet
                        and int(cap_generation) == generation
                        and int(ready_view._only_diff_async_build_seq) == int(cap_seq)
                        and bool(ready_view._only_diff_async_building)
                        and bool(cap_key)
                        and ready_view._only_diff_async_build_key == cap_key
                        and isinstance(record, dict)
                        and int(record.get("build_seq", -1)) == int(cap_seq)
                        and int(record.get("generation", -1)) == generation
                        and dict(record.get("entry") or {}) == entry
                        and dict(cap_prior or {}) == entry
                    )
                except Exception:
                    allowed = False
                if not allowed:
                    return False
        entry.update(
            {
                "generation": generation,
                "state": str(state),
                "stage": str(stage),
                "reason": str(reason),
            }
        )
        if str(state) not in sm._SHEET_EXACT_TERMINAL:
            entry["full_detail_terminal"] = False
        app._sheet_exact_states[sheet] = entry
        return True

    app._sheet_exact_entry = _sheet_exact_entry
    app._set_sheet_exact_state = _set_sheet_exact_state

    def _view(name):
        app._sheet_compute_generation[name] = 0
        app._sheet_exact_states[name] = _exact_entry()
        view = types.SimpleNamespace(
            app=app,
            sheet=name,
            _only_diff_async_build_seq=1,
            _only_diff_async_building=True,
            _only_diff_async_build_key=(name, "key"),
            _only_diff_async_prior_exact=None,
            _only_diff_async_thread=object(),
            _only_diff_source_version=0,
            _only_diff_rows_cache=None,
            _only_diff_rows_cache_key=None,
            _only_diff_rows_exact=False,
            _only_diff_preview_full=True,
            _pending_exact_render=False,
            _prepared_complete=True,
            _data_ready=True,
            _lifecycle_error=None,
            _lifecycle_canceled=False,
            _only_diff_request_origin_value=0,
            _last_only_diff_value=1,
            _prefer_only_diff_when_ready=True,
            only_diff_var=_IntVar(1),
            row_pairs=[(3, 3)],
            max_row=3,
            info=_Info(),
            _refresh_diff_block_ui=lambda: None,
            _refresh_interaction_gate=lambda: None,
            _persist_only_diff_setting_debounced=lambda: None,
            _set_only_diff_pending_info=lambda: None,
            _invalidate_diff_block_model=lambda: None,
        )
        for method in (
            "_begin_only_diff_exact_transition",
            "_clear_only_diff_prior_exact",
            "_restore_only_diff_prior_exact",
            "_fail_only_diff_exact_transition",
            "_queue_only_diff_ui_or_failure",
            "_invalidate_only_diff_snapshot_cache",
            "_cancel_only_diff_calculation",
        ):
            setattr(view, method, types.MethodType(getattr(sm.SheetView, method), view))
        app.sheet_views[name] = view
        return view

    return app, root, window, _view


def _assert_progress_released(app, root, view):
    """Assert only the responsibilities owned by progress finish/cancel."""
    assert app._only_diff_progress_owner is None
    assert app._only_diff_progress_visible_token is None
    assert app._only_diff_progress_show_after_id is None
    assert app._only_diff_progress_show_token is None
    assert app._only_diff_progress_watchdog_after_id is None
    assert app._only_diff_progress_watchdog_token is None
    assert app._only_diff_progress_confirm_after_id is None
    assert app._only_diff_progress_confirm_token is None
    assert len(app._only_diff_progress_visibility_attempts) <= 32
    assert view._only_diff_async_prior_exact is None
    assert not root.callbacks, root.callbacks
    assert not app._root_after_ids, app._root_after_ids
    window = app._only_diff_progress_win
    assert window.grab_current() is not window
    assert window.state() == "withdrawn"
    assert not window.winfo_ismapped() and not window.winfo_viewable()


def _assert_reverted_to_full(app, root, view):
    """Cancel/show-failure must additionally restore a full, retryable view."""
    _assert_progress_released(app, root, view)
    assert not view._only_diff_async_building
    assert view._only_diff_async_build_key is None
    assert not view._only_diff_preview_full
    assert not view._pending_exact_render
    assert view.only_diff_var.get() == 0


def _assert_failed_disposition(app, root, view, *, stage: str):
    entry = app._sheet_exact_entry(view.sheet)
    assert entry["state"] == sm._SHEET_EXACT_FAILED, (stage, entry)
    assert not view._only_diff_async_building, stage
    assert view._only_diff_async_build_key is None, stage
    assert not view._only_diff_preview_full, stage
    assert not view._pending_exact_render, stage
    assert view._only_diff_async_prior_exact is None, stage
    assert app._only_diff_progress_owner is None, stage
    assert app._priority_exact_owner is None, stage
    assert not root.callbacks, (stage, root.callbacks)

def _begin_exact_transition(app, view, seq=1):
    prior = dict(app._sheet_exact_entry(view.sheet))
    view._only_diff_async_build_seq = int(seq)
    assert view._begin_only_diff_exact_transition(int(seq), (view.sheet, "key"))
    calculating = app._sheet_exact_entry(view.sheet)
    assert calculating["state"] == sm._SHEET_EXACT_CALCULATING
    assert calculating["generation"] == prior["generation"]
    assert calculating["full_detail_terminal"] is False
    assert view._only_diff_async_prior_exact["entry"] == prior
    return prior


def _assert_only_diff_reopen_capability_gate():
    app, _root, _window, make_view = _build()
    view = make_view("S")
    prior = dict(app._sheet_exact_entry("S"))

    # A normal terminal-to-CALCULATING update remains a strict no-op once the
    # terminal prepared surface exists.
    assert not app._set_sheet_exact_state("S", sm._SHEET_EXACT_CALCULATING)
    assert app._sheet_exact_entry("S") == prior

    def _cap(*, seq=7, key=("S", "cap"), generation=0, record_entry=None, view_ref=None):
        view._only_diff_async_build_seq = int(seq)
        view._only_diff_async_building = True
        view._only_diff_async_build_key = key
        view._only_diff_async_prior_exact = {
            "build_seq": int(seq),
            "generation": int(generation),
            "entry": copy.deepcopy(
                prior if record_entry is None else record_entry
            ),
        }
        return (
            view if view_ref is None else view_ref,
            int(seq),
            key,
            int(generation),
            copy.deepcopy(prior),
        )

    valid = _cap()
    assert app._set_sheet_exact_state(
        "S", sm._SHEET_EXACT_CALCULATING, only_diff_reopen_capability=valid
    )
    assert app._sheet_exact_entry("S")["state"] == sm._SHEET_EXACT_CALCULATING

    # Restore the original terminal and independently corrupt every authority
    # component; none may reopen a prepared terminal surface.
    for mutate in (
        lambda cap: (cap[0], cap[1] + 1, cap[2], cap[3], cap[4]),
        lambda cap: (cap[0], cap[1], ("S", "wrong"), cap[3], cap[4]),
        lambda cap: (cap[0], cap[1], cap[2], cap[3] + 1, cap[4]),
        lambda cap: (cap[0], cap[1], cap[2], cap[3], {**cap[4], "stage": "wrong"}),
    ):
        app._sheet_exact_states["S"] = copy.deepcopy(prior)
        valid = _cap()
        assert not app._set_sheet_exact_state(
            "S",
            sm._SHEET_EXACT_CALCULATING,
            only_diff_reopen_capability=mutate(valid),
        )
        assert app._sheet_exact_entry("S") == prior

    foreign = make_view("other")
    app._sheet_exact_states["S"] = copy.deepcopy(prior)
    valid = _cap()
    wrong_view = (foreign, valid[1], valid[2], valid[3], valid[4])
    assert not app._set_sheet_exact_state(
        "S", sm._SHEET_EXACT_CALCULATING, only_diff_reopen_capability=wrong_view
    )
    assert app._sheet_exact_entry("S") == prior

    # If the state publisher declines even a build that has claimed its local
    # primitives, begin itself releases every primitive and never reaches modal
    # or worker launch.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    prior = dict(app._sheet_exact_entry("S"))
    original_set = app._set_sheet_exact_state
    app._set_sheet_exact_state = lambda *_args, **_kwargs: False
    assert not view._begin_only_diff_exact_transition(1, ("S", "key"))
    assert app._sheet_exact_entry("S") == prior
    assert not view._only_diff_async_building
    assert view._only_diff_async_build_key is None
    assert not view._only_diff_preview_full
    assert not view._pending_exact_render
    assert view._only_diff_async_prior_exact is None
    assert app._priority_exact_owner is None
    app._set_sheet_exact_state = original_set


def _assert_success_and_finish_cleanup():
    app, root, window, make_view = _build()
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    assert app._only_diff_progress_owner == (view, 1)
    assert root.run_delay(0)
    assert app._only_diff_progress_visible_token == (view, 1)
    assert window.state() == "normal"
    assert window.winfo_ismapped() and window.winfo_viewable()
    assert window.grab_current() is window
    assert not root.callbacks, root.callbacks
    # Model the post-publish state immediately before progress finish.  The real
    # immutable apply/scheduler owns these flags; finish must leave them intact.
    app._sheet_exact_states["S"] = _exact_entry(state=sm._SHEET_EXACT_SAME)
    view._only_diff_async_building = False
    view._only_diff_async_build_key = None
    view._only_diff_preview_full = False
    view._pending_exact_render = False
    view._last_only_diff_value = 1
    view.only_diff_var.set(1)
    post_publish = (
        dict(app._sheet_exact_entry("S")),
        view._only_diff_async_building,
        view._only_diff_async_build_key,
        view._only_diff_preview_full,
        view._pending_exact_render,
        view._last_only_diff_value,
        view.only_diff_var.get(),
    )
    assert app._finish_only_diff_progress(view, 1, outcome="test-success")
    assert (
        dict(app._sheet_exact_entry("S")),
        view._only_diff_async_building,
        view._only_diff_async_build_key,
        view._only_diff_preview_full,
        view._pending_exact_render,
        view._last_only_diff_value,
        view.only_diff_var.get(),
    ) == post_publish
    assert app._sheet_exact_entry("S") != prior
    _assert_progress_released(app, root, view)


def _assert_cancel_and_show_fail_restore_prior_exact():
    for kwargs in (
        {"raise_show": True},
        {"raise_grab": True},
    ):
        app, root, _window, make_view = _build(**kwargs)
        view = make_view("S")
        prior = _begin_exact_transition(app, view)
        app._begin_only_diff_progress(view, 1)
        root.run_delay(0)
        assert app._sheet_exact_entry("S") == prior
        _assert_reverted_to_full(app, root, view)

    # Mapping is no longer an immediate show failure: the dialog retains the
    # current token and retries at 10 ms until the original 100 ms watchdog.
    # Exercise both independent WM facts without conflating them with a Tk
    # exception or weakening the existing fail-closed restoration contract.
    for kwargs in ({"mapped": False}, {"viewable": False}):
        app, root, _window, make_view = _build(**kwargs)
        view = make_view("S")
        prior = _begin_exact_transition(app, view)
        token = (view, 1)
        app._begin_only_diff_progress(view, 1)
        root.run_delay(0)
        assert app._only_diff_progress_owner == token
        assert app._sheet_exact_entry("S")["state"] == sm._SHEET_EXACT_CALCULATING
        assert app._only_diff_progress_visible_token is None
        assert app._only_diff_progress_confirm_after_id is not None
        assert app._only_diff_progress_confirm_token == token
        assert app._only_diff_progress_watchdog_after_id is not None
        assert app._only_diff_progress_watchdog_token == token
        root.run_delay(100)
        assert app._sheet_exact_entry("S") == prior
        _assert_reverted_to_full(app, root, view)

    app, root, window, make_view = _build()
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    assert window.grab_current() is window
    view._cancel_only_diff_calculation(1, outcome="test-explicit-cancel")
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)


def _assert_scheduler_none_fails_closed_and_restores_prior_exact():
    app, root, _window, make_view = _build(fail_after_calls=(1,))
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)

    # The show callback has been scheduled successfully, but the watchdog
    # scheduling itself fails. This must release the same token and restore the
    # terminal exact entry rather than leave a disabled build behind.
    app, root, _window, make_view = _build(fail_after_calls=(2,))
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    # The successfully scheduled show timer is cancelled immediately when the
    # watchdog schedule rejects.  No detached callback may survive that
    # fail-closed restoration path.
    assert root.after_calls == 2, root.after_calls
    assert root.callbacks == {}, root.callbacks
    assert root.canceled == ["after-1"], root.canceled
    assert app._root_after_ids == set(), app._root_after_ids
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)


def _assert_visibility_confirmation_lifecycle():
    # A real WM can need one event-loop turn after deiconify.  The confirm
    # callback must commit only after the same token has actually mapped.
    app, root, window, make_view = _build(mapped=False, viewable=True)
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    assert app._only_diff_progress_owner == (view, 1)
    assert app._only_diff_progress_visible_token is None
    assert app._only_diff_progress_confirm_token == (view, 1)
    assert app._only_diff_progress_confirm_after_id is not None
    assert window.grab_calls == 0
    assert window.focus_calls == 0
    window.mapped = True
    root.run_delay(10)
    assert app._only_diff_progress_visible_token == (view, 1)
    assert app._only_diff_progress_confirm_after_id is None
    assert app._only_diff_progress_confirm_token is None
    assert window.grab_current() is window
    assert window.grab_calls == 1
    assert window.focus_calls == 1
    assert app._only_diff_progress_visibility_attempts
    assert len(app._only_diff_progress_visibility_attempts) <= 32
    assert {
        "attempt", "elapsed_ms", "child_state", "child_mapped",
        "child_viewable", "child_grabbed", "root_state", "root_mapped",
        "root_viewable",
    } <= set(app._only_diff_progress_visibility_attempts[-1])
    app._sheet_exact_states["S"] = _exact_entry(state=sm._SHEET_EXACT_SAME)
    assert app._finish_only_diff_progress(view, 1, outcome="confirm-success")
    _assert_progress_released(app, root, view)

    # A dialog that never maps remains uncommitted until the existing watchdog
    # restores the same exact terminal entry and cancels its confirm callback.
    app, root, _window, make_view = _build(mapped=False, viewable=True)
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    assert app._only_diff_progress_confirm_after_id is not None
    assert _window.grab_calls == 0
    assert _window.focus_calls == 0
    root.run_delay(10)
    root.run_delay(10)
    assert len(app._only_diff_progress_visibility_attempts) <= 32
    root.run_delay(100)
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)

    # A mapped child that throws during its late grab fails closed and restores
    # the prior terminal without ever committing visible state.
    app, root, window, make_view = _build(raise_grab=True)
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    assert window.grab_calls == 1
    assert window.focus_calls == 0
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)

    # A newer token that wins between map and grab cannot receive a visible
    # commit from the old callback.
    app, root, window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    window.on_grab = lambda: app._finish_only_diff_progress(
        view, 1, outcome="stale-during-grab"
    )
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    assert window.grab_calls == 1
    assert window.focus_calls == 0
    assert app._only_diff_progress_visible_token is None
    assert app._only_diff_progress_owner is None
    assert not root.callbacks

    # If installing the confirmation itself fails, use the same immediate
    # fail-closed restoration instead of waiting for the watchdog.
    app, root, _window, make_view = _build(
        fail_after_calls=(3,), mapped=False, viewable=True
    )
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    assert root.after_calls == 3
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)

    # Supersede/cancel/close must revoke a pending confirmation without giving
    # an old token any later opportunity to claim visible state.
    app, root, _window, make_view = _build(mapped=False, viewable=True)
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    view._only_diff_async_build_seq = 2
    assert app._finish_only_diff_progress(view, 1, outcome="stale-confirm")
    assert app._only_diff_progress_confirm_after_id is None
    assert app._only_diff_progress_confirm_token is None
    assert root.callbacks == {}

    app, root, _window, make_view = _build(mapped=False, viewable=True)
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    view._cancel_only_diff_calculation(1, outcome="confirm-cancel")
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)

    app, root, _window, make_view = _build(mapped=False, viewable=True)
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(0)
    app._is_closing = True
    assert app._finish_only_diff_progress(view, 1, outcome="confirm-close")
    assert app._only_diff_progress_confirm_after_id is None
    assert app._only_diff_progress_confirm_token is None
    assert root.callbacks == {}


def _assert_watchdog_stale_and_supersede_cleanup():
    app, root, _window, make_view = _build()
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    app._begin_only_diff_progress(view, 1)
    root.run_delay(100)
    assert app._sheet_exact_entry("S") == prior
    _assert_reverted_to_full(app, root, view)

    # The independent transition token rejects a same-generation old sequence
    # once a newer progress owner/build has claimed the view.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    view._only_diff_async_build_seq = 2
    app._only_diff_progress_owner = (view, 2)
    assert not view._restore_only_diff_prior_exact(1)
    assert app._sheet_exact_entry("S")["state"] == sm._SHEET_EXACT_CALCULATING
    assert not app._finish_only_diff_progress(view, 1, outcome="stale-seq")
    assert view._only_diff_async_prior_exact is None

    # A newer generation must never be overwritten by an old cancellation.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._only_diff_progress_owner = (view, 1)
    newer = _exact_entry(generation=1, state=sm._SHEET_EXACT_SAME)
    newer["stage"] = "new generation exact"
    app._sheet_compute_generation["S"] = 1
    app._sheet_exact_states["S"] = dict(newer)
    assert not view._restore_only_diff_prior_exact(1)
    assert app._sheet_exact_entry("S") == newer
    assert app._finish_only_diff_progress(view, 1, outcome="stale-worker-error")
    assert view._only_diff_async_prior_exact is None
    assert app._sheet_exact_entry("S") == newer

    # Superseding a token restores its current prior terminal before advancing
    # its sequence, then clears the old record without touching the new owner.
    app, _root, _window, make_view = _build()
    old_view = make_view("old")
    new_view = make_view("new")
    old_prior = _begin_exact_transition(app, old_view)
    app._begin_only_diff_progress(old_view, 1)
    new_view._only_diff_async_build_seq = 2
    assert new_view._begin_only_diff_exact_transition(2, ("new", "key"))
    assert app._sheet_exact_entry("old") == old_prior
    assert old_view._only_diff_async_prior_exact is None
    assert app._only_diff_progress_owner is None

    # Rescan/invalidate follows the same current-token disposition: restore
    # first, then revoke the sequence and discard the stored record.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    prior = _begin_exact_transition(app, view)
    view._invalidate_only_diff_snapshot_cache()
    assert app._sheet_exact_entry("S") == prior
    assert view._only_diff_async_build_seq == 2
    assert view._only_diff_async_prior_exact is None


def _assert_hidden_nonuser_failure_dispositions():
    for outcome, stage in (
        ("open-failed", "open failure"),
        ("submit-failed", "submit failure"),
    ):
        app, root, _window, make_view = _build()
        app.selected_sheet = "other"
        view = make_view("hidden")
        _begin_exact_transition(app, view)
        assert view._only_diff_async_prior_exact is not None
        assert view._fail_only_diff_exact_transition(
            1,
            stage=f"{stage} terminal",
            reason=stage,
            outcome=outcome,
        )
        _assert_failed_disposition(app, root, view, stage=stage)


def _assert_queue_rejections_handoff_on_tk_only():
    for source in ("open", "result"):
        for queue_mode in ("false", "none", "raise"):
            app, root, _window, make_view = _build()
            app.selected_sheet = "hidden-other-tab"
            view = make_view("S")
            _begin_exact_transition(app, view)
            callback_calls = []
            if queue_mode == "false":
                app._queue_ui_task = lambda _callback: False
            elif queue_mode == "none":
                app._queue_ui_task = lambda _callback: None
            else:
                def _raise_queue(_callback):
                    raise RuntimeError("injected queue rejection")
                app._queue_ui_task = _raise_queue
            assert not view._queue_only_diff_ui_or_failure(
                lambda: callback_calls.append(source),
                build_seq=1,
                generation=0,
                stage=f"{source} queue failure",
                reason=f"{source} queue rejected",
                outcome=f"{source}-queue-failed",
            )
            # The worker has only recorded immutable facts: no callback and no
            # lifecycle change occur until the main-thread drain consumes them.
            assert callback_calls == []
            assert app._sheet_exact_entry("S")["state"] == sm._SHEET_EXACT_CALCULATING
            assert len(app._only_diff_failure_handoffs) == 1
            assert app._drain_only_diff_failure_handoffs()
            assert callback_calls == []
            _assert_failed_disposition(
                app, root, view, stage=f"{source}/{queue_mode} current handoff"
            )
            assert app._only_diff_failure_handoffs == []


def _assert_handoff_stale_generation_closing_and_shutdown():
    # A stale token never overwrites the newer terminal entry.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._queue_ui_task = lambda _callback: False
    assert not view._queue_only_diff_ui_or_failure(
        lambda: None,
        build_seq=1,
        generation=0,
        stage="stale queue",
        reason="stale",
        outcome="stale",
    )
    newer = _exact_entry(generation=0, state=sm._SHEET_EXACT_SAME)
    view._only_diff_async_build_seq = 2
    app._sheet_exact_states["S"] = dict(newer)
    assert app._drain_only_diff_failure_handoffs()
    assert app._sheet_exact_entry("S") == newer
    assert view._only_diff_async_prior_exact is None

    # A newer generation is equally protected from an old queue rejection.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._queue_ui_task = lambda _callback: False
    assert not view._queue_only_diff_ui_or_failure(
        lambda: None,
        build_seq=1,
        generation=0,
        stage="generation queue",
        reason="new generation",
        outcome="new-generation",
    )
    app._sheet_compute_generation["S"] = 1
    newer = _exact_entry(generation=1, state=sm._SHEET_EXACT_SAME)
    app._sheet_exact_states["S"] = dict(newer)
    assert app._drain_only_diff_failure_handoffs()
    assert app._sheet_exact_entry("S") == newer
    assert view._only_diff_async_prior_exact is None

    # Close drops pending handoffs without writing a terminal over shutdown.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._queue_ui_task = lambda _callback: False
    assert not view._queue_only_diff_ui_or_failure(
        lambda: None,
        build_seq=1,
        generation=0,
        stage="closing queue",
        reason="closing",
        outcome="closing",
    )
    app._is_closing = True
    assert app._drain_only_diff_failure_handoffs()
    assert app._only_diff_failure_handoffs == []
    assert view._only_diff_async_prior_exact is None

    # The shutdown helper clears modal/timer records, exact ownership, queued
    # handoffs, and every view build primitive without restoring ownerless work.
    app, _root, _window, make_view = _build()
    view = make_view("S")
    _begin_exact_transition(app, view)
    app._only_diff_progress_owner = (view, 1)
    app._only_diff_progress_visible_token = (view, 1)
    app._only_diff_progress_show_after_id = "show"
    app._only_diff_progress_show_token = (view, 1)
    app._only_diff_progress_watchdog_after_id = "watchdog"
    app._only_diff_progress_watchdog_token = (view, 1)
    app._exact_broker_pending = (view, 1, object())
    app._queue_ui_task = lambda _callback: False
    assert not view._queue_only_diff_ui_or_failure(
        lambda: None,
        build_seq=1,
        generation=0,
        stage="shutdown queue",
        reason="shutdown",
        outcome="shutdown",
    )
    app._clear_only_diff_shutdown_state()
    assert app._only_diff_failure_handoffs == []
    assert app._exact_broker_pending is None
    assert app._priority_exact_owner is None
    assert app._only_diff_progress_owner is None
    assert app._only_diff_progress_visible_token is None
    assert app._only_diff_progress_show_after_id is None
    assert app._only_diff_progress_watchdog_after_id is None
    assert view._only_diff_async_build_seq == 2
    assert not view._only_diff_async_building
    assert view._only_diff_async_build_key is None
    assert not view._only_diff_preview_full
    assert not view._pending_exact_render
    assert view._only_diff_async_thread is None
    assert view._only_diff_async_prior_exact is None


def main():
    _assert_only_diff_reopen_capability_gate()
    _assert_success_and_finish_cleanup()
    _assert_cancel_and_show_fail_restore_prior_exact()
    _assert_scheduler_none_fails_closed_and_restores_prior_exact()
    _assert_visibility_confirmation_lifecycle()
    _assert_watchdog_stale_and_supersede_cleanup()
    _assert_hidden_nonuser_failure_dispositions()
    _assert_queue_rejections_handoff_on_tk_only()
    _assert_handoff_stale_generation_closing_and_shutdown()
    print("ONLY_DIFF_PROGRESS_DIALOG_CONTRACT_OK")


if __name__ == "__main__":
    main()