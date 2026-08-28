"""Retry wrapper: replace an unavailable synthetic IME key with Unicode paste."""

import _independent_header_input_gate as gate


def _emit(root, widget) -> None:
    root.clipboard_clear()
    root.clipboard_append("HEADER-输入法提交-MUST-NOT-APPEAR")
    widget.focus_force()
    gate._pump(root)
    for sequence, options in (
        ("<KeyPress>", {"keysym": "x"}),
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
        gate._pump(root)


gate._emit_edit_gestures = _emit
gate.main()
