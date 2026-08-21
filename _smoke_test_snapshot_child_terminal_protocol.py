"""Deterministic single-owner protocol tests for isolated snapshot children.

This intentionally exercises no Tk root, workbook, Excel process, or spawned
child.  It proves the lifecycle helpers used by the real runner cannot emit a
second ``finished`` event after a cancellation has already claimed the owner.
"""

from __future__ import annotations

from collections import deque
import tempfile
import threading
import time
from pathlib import Path

import sow_merge_tool as sm


class _FakeProcess:
    def __init__(self, pid: int, *, exitcode: int | None, alive: bool = False):
        self.pid = int(pid)
        self.exitcode = exitcode
        self._alive = bool(alive)
        self.closed = False
        self.terminate_calls = 0
        self.kill_calls = 0
        self.join_calls = 0
        self.close_calls = 0

    def is_alive(self):
        return self._alive

    def terminate(self):
        self.terminate_calls += 1
        self._alive = False
        self.exitcode = -15

    def kill(self):
        self.kill_calls += 1
        self._alive = False
        self.exitcode = -9

    def join(self, timeout=None):
        del timeout
        self.join_calls += 1

    def close(self):
        self.closed = True
        self.close_calls += 1


class _FailingEventSink:
    """A deterministic event writer failure; replacement must retain proof."""

    def __iter__(self):
        return iter(())

    def append(self, _record):
        raise OSError("injected terminal event sink failure")


def _app_shell():
    app = sm.SowMergeApp.__new__(sm.SowMergeApp)
    app._snapshot_child_lock = threading.RLock()
    app._snapshot_child_owner = None
    app._snapshot_child_temp_paths = set()
    app._snapshot_child_events = deque(maxlen=32)
    return app


def _owner(app, root: Path, token: str, *, exitcode: int | None, alive=False):
    result = root / f"{token}.result.pickle"
    partial = root / f"{token}.partial"
    result.write_bytes(b"snapshot-result")
    partial.write_bytes(b"partial")
    owner = {
        "token": token,
        "sheet": "SnapshotProtocol",
        "generation": 7,
        "process": _FakeProcess(1000 + len(token), exitcode=exitcode, alive=alive),
        "result_path": str(result),
        "partial_path": str(partial),
        "last_resources": {"tree_pids": (1000 + len(token),)},
        "terminal_record_lock": threading.Lock(),
        "terminal_recorded_event": threading.Event(),
    }
    with app._snapshot_child_lock:
        app._snapshot_child_owner = owner
        app._snapshot_child_temp_paths.add(str(result))
    return owner, result, partial


def _assert_terminated_claim_wins(root: Path):
    app = _app_shell()
    owner, result, partial = _owner(app, root, "cancel", exitcode=None, alive=True)
    assert app._terminate_snapshot_child("cancel-or-preempt")
    assert app._finalize_snapshot_child_runner(
        owner,
        normal_result_verified=False,
        result_decoded=False,
        failure_reason="runner unwound after cancellation",
    ) is None
    events = list(app._snapshot_child_events)
    assert len(events) == 1, events
    event = events[0]
    assert event["event"] == "terminated", event
    assert event["reason"] == "cancel-or-preempt", event
    assert event["exitcode"] == -15, event
    assert event["result_existed_before_cleanup"] is True, event
    assert event["partial_existed_before_cleanup"] is True, event
    assert event["result_exists_after_cleanup"] is False, event
    assert event["partial_exists_after_cleanup"] is False, event
    assert not result.exists() and not partial.exists()
    assert owner["terminal_recorded_event"].is_set()
    assert owner["process"].close_calls == 1


def _assert_verified_normal_is_finished_once(root: Path):
    app = _app_shell()
    owner, result, partial = _owner(app, root, "normal", exitcode=0)
    event = app._finalize_snapshot_child_runner(
        owner,
        normal_result_verified=True,
        result_decoded=True,
    )
    assert event is not None
    events = list(app._snapshot_child_events)
    assert len(events) == 1, events
    event = events[0]
    assert event["event"] == "finished", event
    assert event["reason"] == "normal-result-verified", event
    assert event["exitcode"] == 0 and event["result_decoded"] is True, event
    assert event["result_existed_before_cleanup"] is True, event
    assert event["result_exists_after_cleanup"] is False, event
    assert not result.exists() and not partial.exists()
    assert owner["terminal_recorded_event"].is_set()
    assert owner["process"].close_calls == 1


def _assert_exception_is_failed_once(root: Path):
    app = _app_shell()
    owner, result, partial = _owner(app, root, "failure", exitcode=1)
    event = app._finalize_snapshot_child_runner(
        owner,
        normal_result_verified=False,
        result_decoded=False,
        failure_reason="OSError: injected protocol failure",
    )
    assert event is not None
    events = list(app._snapshot_child_events)
    assert len(events) == 1, events
    event = events[0]
    assert event["event"] == "failed", event
    assert event["exitcode"] == 1, event
    assert "injected protocol failure" in event["exception"], event
    assert event["result_exists_after_cleanup"] is False, event
    assert event["partial_exists_after_cleanup"] is False, event
    assert not result.exists() and not partial.exists()
    assert owner["terminal_recorded_event"].is_set()
    assert owner["process"].close_calls == 1


def _assert_stale_runner_cannot_pollute_latest_owner(root: Path):
    app = _app_shell()
    old_owner, old_result, old_partial = _owner(app, root, "old", exitcode=1)
    latest_owner, latest_result, latest_partial = _owner(app, root, "latest", exitcode=0)
    assert app._finalize_snapshot_child_runner(
        old_owner,
        normal_result_verified=False,
        result_decoded=False,
        failure_reason="old runner lost ownership",
    ) is None
    assert app._snapshot_child_owner is latest_owner
    assert not list(app._snapshot_child_events)
    assert not old_result.exists() and not old_partial.exists()
    event = app._finalize_snapshot_child_runner(
        latest_owner,
        normal_result_verified=True,
        result_decoded=True,
    )
    assert event is not None
    events = list(app._snapshot_child_events)
    assert len(events) == 1 and events[0]["token"] == "latest", events
    assert events[0]["event"] == "finished", events
    assert not latest_result.exists() and not latest_partial.exists()


def _assert_claimed_terminator_excludes_runner_until_fence(root: Path):
    """Pause after terminate claims ownership and race the real runner finalizer."""
    app = _app_shell()
    owner, result, partial = _owner(app, root, "barrier", exitcode=None, alive=True)
    process = owner["process"]
    record_entered = threading.Event()
    release_record = threading.Event()
    runner_started = threading.Event()
    runner_done = threading.Event()
    cleanup_threads = []
    runner_results = []
    original_record = app._record_snapshot_child_terminal
    original_cleanup = app._cleanup_snapshot_child_terminal_paths

    def paused_record(*args, **kwargs):
        record_entered.set()
        assert release_record.wait(timeout=2.0), "test did not release terminal record"
        return original_record(*args, **kwargs)

    def spy_cleanup(*args, **kwargs):
        cleanup_threads.append(threading.current_thread().name)
        return original_cleanup(*args, **kwargs)

    app._record_snapshot_child_terminal = paused_record
    app._cleanup_snapshot_child_terminal_paths = spy_cleanup
    terminator = threading.Thread(
        target=lambda: app._terminate_snapshot_child("cancel-or-preempt"),
        name="terminator",
    )

    def _runner_finally():
        runner_started.set()
        runner_results.append(app._finalize_snapshot_child_runner(
            owner,
            normal_result_verified=False,
            result_decoded=False,
            failure_reason="runner unwound after claimed termination",
        ))
        runner_done.set()

    terminator.start()
    assert record_entered.wait(timeout=1.0), "terminator did not claim before record"
    assert owner["terminal_event"] == "terminated"
    runner = threading.Thread(target=_runner_finally, name="stale-runner")
    runner.start()
    assert runner_started.wait(timeout=1.0)
    time.sleep(0.05)
    # The stale runner waits on the private record fence; it may not remove
    # IPC or close the process while the terminator owns the terminal record.
    assert not runner_done.is_set()
    assert cleanup_threads == [], cleanup_threads
    assert process.close_calls == 0
    release_record.set()
    terminator.join(timeout=2.0)
    runner.join(timeout=2.0)
    assert not terminator.is_alive() and not runner.is_alive()
    assert runner_results == [None], runner_results
    assert owner["terminal_recorded_event"].is_set()
    assert cleanup_threads == ["terminator"], cleanup_threads
    assert process.close_calls == 1
    events = list(app._snapshot_child_events)
    assert len(events) == 1 and events[0]["event"] == "terminated", events
    assert not result.exists() and not partial.exists()


def _assert_cleanup_and_record_errors_still_emit_failed(root: Path):
    app = _app_shell()
    owner, result, partial = _owner(app, root, "cleanup-error", exitcode=0)

    def _broken_cleanup(_owner):
        raise OSError("injected cleanup helper failure")

    app._cleanup_snapshot_child_terminal_paths = _broken_cleanup
    event = app._finalize_snapshot_child_runner(
        owner,
        normal_result_verified=True,
        result_decoded=True,
    )
    assert event is not None and event["event"] == "failed", event
    assert event["original_event"] == "finished", event
    assert "cleanup:OSError" in event["terminal_record_error"], event
    assert owner["terminal_recorded_event"].is_set()
    assert owner["process"].close_calls == 1
    assert not result.exists() and not partial.exists()

    app = _app_shell()
    owner, result, partial = _owner(app, root, "record-error", exitcode=0)
    app._snapshot_child_events = _FailingEventSink()
    event = app._finalize_snapshot_child_runner(
        owner,
        normal_result_verified=True,
        result_decoded=True,
    )
    assert event is not None and event["event"] == "failed", event
    assert "event-append:OSError" in event["terminal_record_error"], event
    events = list(app._snapshot_child_events)
    assert len(events) == 1 and events[0]["event"] == "failed", events
    assert owner["terminal_recorded_event"].is_set()
    assert owner["process"].close_calls == 1
    assert not result.exists() and not partial.exists()


def main():
    with tempfile.TemporaryDirectory(prefix="sow_snapshot_terminal_protocol_") as temp:
        root = Path(temp)
        _assert_terminated_claim_wins(root)
        _assert_verified_normal_is_finished_once(root)
        _assert_exception_is_failed_once(root)
        _assert_stale_runner_cannot_pollute_latest_owner(root)
        _assert_claimed_terminator_excludes_runner_until_fence(root)
        _assert_cleanup_and_record_errors_still_emit_failed(root)
    print("SNAPSHOT_CHILD_TERMINAL_PROTOCOL PASS cases=6")


if __name__ == "__main__":
    main()
