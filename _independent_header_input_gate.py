"""Disposable independent gate for editable-looking row-header Text widgets.

This test is intentionally synthetic and does not mutate production workbooks.
It asserts that keyboard, clipboard, middle-click, and virtual edit gestures
cannot alter a row-header document or create an undo/redo action, while its
click/hover dispatch remains live.
"""

from __future__ import annotations

import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


def _make_book(path: Path, changed: bool) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Gate"
    ws.append(["id@id", "value"])
    for row in range(1, 12):
        value = f"v-{row}" + ("-changed" if changed and row == 6 else "")
        ws.append([row, value])
    wb.save(path)
    wb.close()


def _pump(root, seconds: float = 0.05) -> None:
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.002)


def _wait(root, predicate, timeout: float, label: str) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        _pump(root, 0.02)
        if predicate():
            return
    raise AssertionError(f"timeout: {label}")


def _header_snapshot(view) -> tuple:
    return tuple(widget.get("1.0", "end-1c") for widget in (view.left_ln, view.right_ln))


def _emit_edit_gestures(root, widget) -> None:
    root.clipboard_clear()
    root.clipboard_append("HEADER-MUTATION-MUST-NOT-APPEAR")
    widget.focus_force()
    _pump(root)
    # Character insertion, destructive keys, clipboard shortcuts/virtual
    # events, middle-paste and composition-like key input all traverse Tk's
    # ordinary event dispatch.  Any unblocked Text class binding will change
    # the header document and be caught below.
    for sequence, options in (
        ("<KeyPress>", {"keysym": "x"}),
        ("<KeyPress>", {"keysym": "Multi_key"}),
        ("<BackSpace>", {}),
        ("<Delete>", {}),
        ("<Control-v>", {}),
        ("<Control-x>", {}),
        ("<<Paste>>", {}),
        ("<<Cut>>", {}),
        ("<<Clear>>", {}),
        ("<Button-2>", {"x": 1, "y": 1}),
        ("<B2-Motion>", {"x": 2, "y": 1}),
        ("<ButtonRelease-2>", {"x": 2, "y": 1}),
    ):
        widget.event_generate(sequence, **options)
        _pump(root)


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="sow_header_gate_") as tmp:
        mine = Path(tmp) / "mine.xlsx"
        theirs = Path(tmp) / "theirs.xlsx"
        _make_book(mine, changed=False)
        _make_book(theirs, changed=True)
        app = sm.SowMergeApp(str(mine), str(theirs), initial_sheet="Gate")
        try:
            _wait(
                app.root,
                lambda: (
                    app.sheet_views.get("Gate") is not None
                    and app._is_sheet_exact_current("Gate")
                ),
                12.0,
                "exact Gate",
            )
            view = app.sheet_views["Gate"]
            _pump(app.root)
            assert str(view.left_ln.cget("state")) == "normal"
            before = _header_snapshot(view)
            undo_before = len(app.undo_stack)
            redo_before = len(app.redo_stack)
            for header in (view.left_ln, view.right_ln):
                _emit_edit_gestures(app.root, header)
                assert _header_snapshot(view) == before, (header, before, _header_snapshot(view))
                assert len(app.undo_stack) == undo_before
                assert len(app.redo_stack) == redo_before

            dispatched: list[tuple[str, str]] = []
            original_click = view._on_row_header_click
            original_hover = view._on_row_header_hover
            view._on_row_header_click = lambda _w, _e, direction: dispatched.append(("click", direction)) or "break"
            view._on_row_header_hover = lambda _w, _e, direction: dispatched.append(("hover", direction)) or "break"
            try:
                view.left_ln.event_generate("<Motion>", x=1, y=1)
                view.left_ln.event_generate("<Button-1>", x=1, y=1)
                _pump(app.root)
            finally:
                view._on_row_header_click = original_click
                view._on_row_header_hover = original_hover
            assert ("hover", "A2B") in dispatched, dispatched
            assert ("click", "A2B") in dispatched, dispatched
            assert _header_snapshot(view) == before
            assert len(app.undo_stack) == undo_before
            assert len(app.redo_stack) == redo_before
            print("INDEPENDENT_HEADER_INPUT_GATE_PASS")
        finally:
            app._shutdown_root()


if __name__ == "__main__":
    main()
