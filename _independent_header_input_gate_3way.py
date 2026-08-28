"""Three-way coverage for all normal-state row header Text widgets."""

from __future__ import annotations

import tempfile
from pathlib import Path

import _independent_header_input_gate as gate
import sow_merge_tool as sm


def _snapshot(view) -> tuple[str, ...]:
    return tuple(
        widget.get("1.0", "end-1c")
        for widget in (view.left_ln, view.base_ln, view.right_ln, view.cursor_cmp_ln)
    )


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="sow_header_gate_3way_") as tmp:
        mine = Path(tmp) / "mine.xlsx"
        theirs = Path(tmp) / "theirs.xlsx"
        base = Path(tmp) / "base.xlsx"
        gate._make_book(mine, changed=False)
        gate._make_book(theirs, changed=True)
        gate._make_book(base, changed=False)
        app = sm.SowMergeApp(
            str(mine),
            str(theirs),
            merge_mode=True,
            base_path=str(base),
            initial_sheet="Gate",
        )
        try:
            gate._wait(
                app.root,
                lambda: (
                    app.sheet_views.get("Gate") is not None
                    and app._is_sheet_exact_current("Gate")
                ),
                12.0,
                "three-way exact Gate",
            )
            view = app.sheet_views["Gate"]
            view._update_cursor_lines()
            gate._pump(app.root)
            headers = (view.left_ln, view.base_ln, view.right_ln, view.cursor_cmp_ln)
            assert all(str(header.cget("state")) == "normal" for header in headers)
            before = _snapshot(view)
            undo_before, redo_before = len(app.undo_stack), len(app.redo_stack)
            for header in headers:
                gate._emit_edit_gestures(app.root, header)
                assert _snapshot(view) == before
                assert (len(app.undo_stack), len(app.redo_stack)) == (undo_before, redo_before)

            dispatched: list[tuple[str, str]] = []
            original_click, original_hover = view._on_row_header_click, view._on_row_header_hover
            view._on_row_header_click = lambda _w, _e, direction: dispatched.append(("click", direction)) or "break"
            view._on_row_header_hover = lambda _w, _e, direction: dispatched.append(("hover", direction)) or "break"
            try:
                for header in (view.left_ln, view.base_ln, view.right_ln):
                    header.event_generate("<Motion>", x=1, y=1)
                    header.event_generate("<Button-1>", x=1, y=1)
                gate._pump(app.root)
            finally:
                view._on_row_header_click, view._on_row_header_hover = original_click, original_hover
            for direction in ("A2B", "BASE2A", "B2A"):
                assert ("hover", direction) in dispatched, dispatched
                assert ("click", direction) in dispatched, dispatched
            assert _snapshot(view) == before
            assert (len(app.undo_stack), len(app.redo_stack)) == (undo_before, redo_before)
            print("INDEPENDENT_HEADER_INPUT_GATE_3WAY_PASS")
        finally:
            app._shutdown_root()


if __name__ == "__main__":
    main()
