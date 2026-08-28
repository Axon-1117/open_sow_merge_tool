"""Final one-shot wrapper guarded for the snapshot child spawn process."""

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


def main() -> None:
    gate._emit_edit_gestures = _emit
    gate.main()


if __name__ == "__main__":
    main()
