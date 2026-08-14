"""Shared Windows-style Tk UI primitives.

The application has two substantial Tk surfaces.  This module intentionally
contains only reusable chrome, scheduling and tracing helpers; domain logic
stays in the Excel and SVN modules.
"""

from __future__ import annotations

import queue
import threading
import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field


@dataclass(frozen=True)
class UiTheme:
    """Small set of tokens shared by both windows."""

    window_bg: str = "#F5F6F7"
    panel_bg: str = "#FFFFFF"
    border: str = "#D3D6DA"
    text: str = "#202124"
    secondary_text: str = "#5F6368"
    accent: str = "#0F6CBD"
    accent_active: str = "#0B5CAD"
    success: str = "#107C10"
    warning: str = "#986F0B"
    error: str = "#C42B1C"
    disabled: str = "#8A8F98"
    row_alt: str = "#F8F9FA"
    font_family: str = "Segoe UI"
    fallback_font_family: str = "Microsoft YaHei UI"


THEME = UiTheme()


def configure_ttk_style(root, *, theme: UiTheme = THEME):
    """Apply the common style without assuming a particular Tk theme."""

    from tkinter import TclError

    style = root.tk.call("ttk::style", "theme", "use")
    # Vista/Windows themes ignore a few background options.  Configure them
    # anyway so the same code remains usable on a developer workstation with
    # clam or a minimal Tk installation.
    try:
        root.tk.call("ttk::style", "theme", "use", "vista")
    except TclError:
        pass
    style_obj = root.tk.call("ttk::style", "theme", "use") or style
    del style_obj
    ttk = root.ttk if hasattr(root, "ttk") else None
    if ttk is None:
        # Tk roots do not expose ttk; import lazily to keep the module cheap.
        from tkinter import ttk as ttk_module

        ttk = ttk_module
    style = ttk.Style(root)
    style.configure("App.TFrame", background=theme.window_bg)
    style.configure("Panel.TFrame", background=theme.panel_bg)
    style.configure("App.TLabel", background=theme.window_bg, foreground=theme.text, font=(theme.font_family, 9))
    style.configure("Muted.App.TLabel", background=theme.window_bg, foreground=theme.secondary_text, font=(theme.font_family, 9))
    style.configure("Title.App.TLabel", background=theme.window_bg, foreground=theme.text, font=(theme.font_family, 10, "bold"))
    style.configure("Primary.TButton", padding=(12, 5), font=(theme.font_family, 9, "bold"))
    style.configure("App.TButton", padding=(10, 5), font=(theme.font_family, 9))
    style.configure("Treeview", rowheight=25, font=(theme.font_family, 9), background=theme.panel_bg, fieldbackground=theme.panel_bg, foreground=theme.text)
    style.configure("Treeview.Heading", font=(theme.font_family, 9, "bold"))
    style.map("Treeview", background=[("selected", "#DCEBFA")], foreground=[("selected", theme.text)])
    style.configure("App.Horizontal.TProgressbar", troughcolor="#E5E7EA", background=theme.accent)
    return style


@dataclass
class UiTrace:
    """In-memory startup trace that can be emitted without leaking file data."""

    started_at: float = field(default_factory=time.perf_counter)
    marks: list[tuple[str, float]] = field(default_factory=list)

    def mark(self, name: str) -> None:
        self.marks.append((str(name), time.perf_counter()))

    def durations(self) -> dict[str, float]:
        previous = self.started_at
        result: dict[str, float] = {}
        for name, stamp in self.marks:
            result[name] = max(0.0, stamp - previous)
            previous = stamp
        return result


class UiTaskRunner:
    """Run cancellable work off the Tk thread and marshal results safely."""

    def __init__(self, root, *, poll_ms: int = 40):
        self.root = root
        self.poll_ms = max(20, int(poll_ms))
        self._queue: queue.Queue = queue.Queue()
        self._generation = 0
        self._closed = False
        self._polling = False
        self._cancel_events: dict[int, threading.Event] = {}
        self._callbacks: dict[int, Callable[[object, Exception | None, int], None]] = {}

    def close(self) -> None:
        self._closed = True
        for event in self._cancel_events.values():
            event.set()
        self._callbacks.clear()

    def cancel(self, generation: int | None = None) -> None:
        if generation is None:
            generation = self._generation
        event = self._cancel_events.get(generation)
        if event is not None:
            event.set()

    def submit(
        self,
        worker: Callable[[threading.Event], object],
        on_done: Callable[[object, Exception | None, int], None],
    ) -> int:
        self._generation += 1
        generation = self._generation
        cancel_event = threading.Event()
        self._cancel_events[generation] = cancel_event
        self._callbacks[generation] = on_done

        def run() -> None:
            value = None
            error: Exception | None = None
            try:
                value = worker(cancel_event)
            except Exception as exc:  # noqa: BLE001  # worker failures are returned to Tk
                error = exc
            self._queue.put((generation, value, error))

        threading.Thread(target=run, name=f"sow-ui-task-{generation}", daemon=True).start()
        self._ensure_polling()
        return generation

    def _ensure_polling(self) -> None:
        if self._polling or self._closed:
            return
        self._polling = True
        self.root.after(self.poll_ms, self._poll)

    def _poll(self) -> None:
        if self._closed:
            self._polling = False
            return
        try:
            while True:
                generation, value, error = self._queue.get_nowait()
                self._cancel_events.pop(generation, None)
                callback = self._callbacks.pop(generation, None)
                if generation == self._generation and callback is not None:
                    # All callbacks are invoked on Tk's event thread.
                    callback(value, error, generation)
        except queue.Empty:
            pass
        self.root.after(self.poll_ms, self._poll)

def batched(values: Iterable[object], size: int) -> Iterable[list[object]]:
    """Yield small lists so a Tk render can return to its event loop."""

    batch: list[object] = []
    for value in values:
        batch.append(value)
        if len(batch) >= size:
            yield batch
            batch = []
    if batch:
        yield batch
