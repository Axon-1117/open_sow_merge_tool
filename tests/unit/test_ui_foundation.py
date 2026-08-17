from __future__ import annotations

import time
from threading import Event

from sow_merge_tool.ui_foundation import THEME, UiTaskRunner, UiTrace, batched


def test_theme_tokens_are_neutral_and_branded() -> None:
    assert THEME.window_bg.startswith("#")
    assert THEME.panel_bg == "#FFFFFF"
    assert THEME.accent == "#0F6CBD"


def test_trace_marks_are_monotonic() -> None:
    trace = UiTrace()
    trace.mark("first")
    time.sleep(0.001)
    trace.mark("second")
    durations = trace.durations()
    assert list(durations) == ["first", "second"]
    assert durations["first"] >= 0
    assert durations["second"] >= 0


def test_batched_keeps_order_and_tail() -> None:
    assert list(batched(range(5), 2)) == [[0, 1], [2, 3], [4]]


class _FakeRoot:
    def __init__(self) -> None:
        self.callbacks = []

    def after(self, _delay, callback):
        self.callbacks.append(callback)
        return len(self.callbacks)

    def drain(self) -> None:
        callbacks, self.callbacks = self.callbacks, []
        for callback in callbacks:
            callback()


def test_task_runner_returns_result_on_ui_callback() -> None:
    root = _FakeRoot()
    runner = UiTaskRunner(root, poll_ms=20)
    received = []
    runner.submit(lambda _cancel: "ready", lambda value, error, _generation: received.append((value, error)))
    for _ in range(100):
        root.drain()
        if received:
            break
        time.sleep(0.002)
    runner.close()
    assert received == [("ready", None)]


def test_task_runner_discards_stale_generation() -> None:
    root = _FakeRoot()
    runner = UiTaskRunner(root, poll_ms=20)
    release = Event()
    received = []
    runner.submit(lambda _cancel: (release.wait(0.05) and "old"), lambda value, error, _generation: received.append(value))
    runner.submit(lambda _cancel: "new", lambda value, error, _generation: received.append(value))
    for _ in range(100):
        root.drain()
        if received:
            break
        time.sleep(0.002)
    release.set()
    for _ in range(30):
        root.drain()
        time.sleep(0.002)
    runner.close()
    assert received == ["new"]
