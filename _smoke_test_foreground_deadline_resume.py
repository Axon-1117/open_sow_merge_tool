"""Deterministic protocol test for foreground deadline resume.

This imports the production coordinator directly.  It starts no Tk interpreter,
workbook, parser child, GUI worker, or revision input; ``_FakeRoot`` supplies
only the ``after``/``after_cancel`` contract used by the production scheduler.
"""

from __future__ import annotations

import sow_merge_tool as sm


def _entry(
    generation: int,
    state: str = sm._SHEET_EXACT_CALCULATING,
    *,
    started: float | None = 1.0,
    full: bool = False,
) -> dict:
    return {
        "generation": int(generation),
        "state": state,
        "request_started_at": started,
        "full_detail_terminal": bool(full),
    }


class _FakeRoot:
    def __init__(self):
        self._next = 0
        self.pending: dict[str, tuple[int, object]] = {}
        self.cancelled: list[str] = []

    def after(self, delay_ms: int, callback):
        self._next += 1
        after_id = f"after-{self._next}"
        self.pending[after_id] = (int(delay_ms), callback)
        return after_id

    def after_cancel(self, after_id: str):
        self.cancelled.append(str(after_id))
        self.pending.pop(str(after_id), None)

    def fire(self, after_id: str):
        _delay, callback = self.pending.pop(after_id)
        callback()

    def capture(self, after_id: str):
        return self.pending[after_id][1]


class _Runtime:
    def __init__(self):
        self.root = _FakeRoot()
        self.entries: dict[str, dict] = {}
        self.generations: dict[str, int] = {}
        self.selected_sheet = ""
        self.closing = False
        self.enqueued: list[str] = []
        self.kicks: list[str] = []
        self.events: list[tuple[str, dict]] = []

    def snapshot(self):
        return (
            {sheet: dict(entry) for sheet, entry in self.entries.items()},
            dict(self.generations),
        )

    def selected_is_full(self):
        entry = self.entries.get(self.selected_sheet) or {}
        return bool(
            entry.get("full_detail_terminal", False)
            and entry.get("state") in sm._SHEET_EXACT_TERMINAL
        )

    def enqueue_front(self, sheet: str):
        self.enqueued.append(str(sheet))

    def kick_worker(self):
        self.kicks.append("kick")

    def emit(self, event: str, **details):
        self.events.append((str(event), dict(details)))

    def request_priority(self, coordinator: sm._ForegroundResumeCoordinator):
        return coordinator.request_priority_after_terminal(
            self.selected_sheet,
            self.entries[self.selected_sheet],
            snapshot=self.snapshot,
            closing=lambda: self.closing,
            selected_is_full=self.selected_is_full,
            after=self.root.after,
            after_cancel=self.root.after_cancel,
            enqueue_front=self.enqueue_front,
            kick_worker=self.kick_worker,
            emit=self.emit,
        )


def _confirm(
    coordinator: sm._ForegroundResumeCoordinator,
    runtime: _Runtime,
    sheet: str,
    generation: int,
    entry: dict,
):
    runtime.entries[sheet] = dict(entry)
    runtime.generations[sheet] = int(generation)
    runtime.selected_sheet = str(sheet)
    return coordinator.confirm_selection(
        sheet,
        generation,
        runtime.entries[sheet],
        after_cancel=runtime.root.after_cancel,
        confirmed_at=100.0 + len(runtime.events),
    )


def _assert_confirmed_pending_visit_becomes_away_request_then_after_zero():
    coordinator = sm._ForegroundResumeCoordinator()
    runtime = _Runtime()
    # Monster is visibly selected while PENDING, before the worker publishes a
    # request clock.  The visit must survive selecting Dungeon.
    monster_visit = _confirm(
        coordinator, runtime, "Monster", 2, _entry(2, sm._SHEET_EXACT_PENDING, started=None)
    )
    assert monster_visit is not None
    assert not coordinator.ledger.entries
    _confirm(coordinator, runtime, "Dungeon", 1, _entry(1, started=30.0))
    # The first Monster CALCULATING state arrives while Dungeon is selected.
    runtime.entries["Monster"] = _entry(2, started=11.0)
    monster_request = coordinator.note_request_started(
        "Monster", 2, runtime.entries["Monster"]
    )
    assert monster_request is not None
    assert monster_request.request_started_at == 11.0
    assert monster_request.tab_seq == monster_visit.tab_seq
    assert runtime.selected_sheet == "Dungeon"

    runtime.entries["Dungeon"] = _entry(
        1, sm._SHEET_EXACT_CHANGED, started=30.0, full=True
    )
    ticket = runtime.request_priority(coordinator)
    assert ticket is not None and ticket.request.sheet == "Monster"
    after_id = coordinator.priority_after_id
    assert after_id is not None and runtime.root.pending[after_id][0] == 0
    runtime.root.fire(after_id)
    assert runtime.enqueued == ["Monster"], runtime.enqueued
    assert runtime.kicks == ["kick"], runtime.kicks
    assert coordinator.ticket_is_active(
        "Monster", 2, *runtime.snapshot(), closing=False
    )


def _assert_deadline_order_tie_and_single_ticket():
    coordinator = sm._ForegroundResumeCoordinator()
    runtime = _Runtime()
    _confirm(coordinator, runtime, "Monster", 2, _entry(2, started=50.0))
    coordinator.note_request_started("Monster", 2, runtime.entries["Monster"])
    _confirm(coordinator, runtime, "Chapter", 3, _entry(3, started=50.0))
    chapter = coordinator.note_request_started("Chapter", 3, runtime.entries["Chapter"])
    _confirm(coordinator, runtime, "Dungeon", 1, _entry(1, started=90.0))
    runtime.entries["Dungeon"] = _entry(
        1, sm._SHEET_EXACT_CHANGED, started=90.0, full=True
    )
    ticket = runtime.request_priority(coordinator)
    assert ticket is not None and ticket.request is chapter
    # Repeating the terminal hand-off while its after(0) is owned must not
    # create a second child/after callback.
    assert runtime.request_priority(coordinator) is None
    assert len(runtime.root.pending) == 1


def _assert_never_selected_hidden_uses_quiet_and_stale_quiet_cannot_fire():
    coordinator = sm._ForegroundResumeCoordinator()
    runtime = _Runtime()
    runtime.entries["Hidden"] = _entry(7, started=5.0)
    runtime.generations["Hidden"] = 7
    # No confirmed visit => no ledger entry; the hidden worker remains an
    # opportunistic quiet-timer case rather than being deadline-prioritized.
    token = coordinator.claim_quiet_install()
    assert token is not None
    after_id = coordinator.install_quiet_timer(
        token,
        1200,
        snapshot=runtime.snapshot,
        closing=lambda: runtime.closing,
        after=runtime.root.after,
        after_cancel=runtime.root.after_cancel,
        kick_worker=runtime.kick_worker,
        emit=runtime.emit,
    )
    assert after_id is not None and not coordinator.ledger.entries
    runtime.root.fire(after_id)
    assert runtime.kicks == ["kick"]

    token = coordinator.claim_quiet_install()
    after_id = coordinator.install_quiet_timer(
        token,
        1200,
        snapshot=runtime.snapshot,
        closing=lambda: runtime.closing,
        after=runtime.root.after,
        after_cancel=runtime.root.after_cancel,
        kick_worker=runtime.kick_worker,
        emit=runtime.emit,
    )
    stale_callback = runtime.root.capture(after_id)
    # A subsequent real confirmation bumps epoch and cancels the quiet owner.
    _confirm(coordinator, runtime, "Dungeon", 1, _entry(1, sm._SHEET_EXACT_PENDING, started=None))
    assert after_id in runtime.root.cancelled
    stale_callback()  # Simulate the already-dispatched Tk callback race.
    assert runtime.kicks == ["kick"], runtime.kicks


def _assert_stale_ticket_releases_and_cache_pending_defer_is_preserved():
    coordinator = sm._ForegroundResumeCoordinator()
    runtime = _Runtime()
    _confirm(coordinator, runtime, "Monster", 2, _entry(2, started=4.0))
    coordinator.note_request_started("Monster", 2, runtime.entries["Monster"])
    _confirm(coordinator, runtime, "Dungeon", 1, _entry(1, started=9.0))
    runtime.entries["Dungeon"] = _entry(
        1, sm._SHEET_EXACT_CHANGED, started=9.0, full=True
    )
    ticket = runtime.request_priority(coordinator)
    assert ticket is not None
    runtime.generations["Monster"] = 3
    assert not coordinator.ticket_is_active(
        "Monster", 2, *runtime.snapshot(), closing=False
    )
    assert coordinator.ledger._priority_ticket is None
    # A new generation is not blocked by the released stale ticket.
    _confirm(coordinator, runtime, "Monster", 3, _entry(3, started=8.0))
    request = coordinator.note_request_started("Monster", 3, runtime.entries["Monster"])
    assert request is not None
    runtime.selected_sheet = "Dungeon"
    fresh_ticket = runtime.request_priority(coordinator)
    assert fresh_ticket is not None and fresh_ticket.request == request
    runtime.root.fire(coordinator.priority_after_id)
    assert coordinator.ticket_is_active(
        "Monster", 3, *runtime.snapshot(), closing=False
    )

    # UI-cache pending is intentionally not runnable even if its registry
    # remains CALCULATING.  The coordinator owns no second worker here.
    assert sm._hidden_interrupt_must_defer(sm._SHEET_EXACT_CALCULATING, False)
    assert not sm._selected_sheet_is_runnable_queue_front("Monster", ("Dungeon",))
    coordinator.release_worker_turn("Monster", 3, emit=runtime.emit)
    assert not coordinator.ticket_is_active(
        "Monster", 3, *runtime.snapshot(), closing=False
    )


def _assert_terminal_failed_and_close_clear_visit_and_request():
    coordinator = sm._ForegroundResumeCoordinator()
    runtime = _Runtime()
    _confirm(coordinator, runtime, "Monster", 2, _entry(2, started=2.0))
    coordinator.note_request_started("Monster", 2, runtime.entries["Monster"])
    assert coordinator.ledger.visits and coordinator.ledger.entries
    runtime.entries["Monster"] = _entry(2, sm._SHEET_EXACT_FAILED, started=2.0)
    coordinator.ledger.discard_stale(*runtime.snapshot())
    assert not coordinator.ledger.visits and not coordinator.ledger.entries
    _confirm(coordinator, runtime, "Monster", 3, _entry(3, started=3.0))
    coordinator.note_request_started("Monster", 3, runtime.entries["Monster"])
    coordinator.clear(after_cancel=runtime.root.after_cancel)
    assert not coordinator.ledger.visits and not coordinator.ledger.entries


def main() -> None:
    _assert_confirmed_pending_visit_becomes_away_request_then_after_zero()
    _assert_deadline_order_tie_and_single_ticket()
    _assert_never_selected_hidden_uses_quiet_and_stale_quiet_cannot_fire()
    _assert_stale_ticket_releases_and_cache_pending_defer_is_preserved()
    _assert_terminal_failed_and_close_clear_visit_and_request()
    print("foreground deadline resume: PASS")


if __name__ == "__main__":
    main()
