"""Prototype: true viewport virtualization for the side-by-side Text view.

This is an evaluation prototype for the "perf-virtual-eval" task. It is NOT wired
into sow_merge_tool.py; it demonstrates, in isolation, the core mechanism a real
virtualized renderer would need, and surfaces the main architectural challenge.

Key idea (vs current "append-prefix" pseudo-virtualization):
  * The tk.Text only ever holds VIEWPORT + buffer rows (a fixed, small number),
    regardless of the dataset size (here 200,000 rows).
  * The Text's own yview/scrollbar cannot represent the full range (it only knows
    about the lines it holds), so a SEPARATE "proxy" scrollbar drives a virtual
    top-row index. On scroll we replace the Text content with the new window.
  * Two synchronized panes (mimicking left/right) share the same virtual offset.

Run:
    python _proto_virtual_text_render.py
Then drag the right-hand scrollbar / use the mouse wheel. The status line shows
the virtual row window; note that render cost is constant (window-sized), not
O(total rows), and memory stays flat.

What this prototype deliberately does NOT solve (the hard parts for the real app):
  * line<->pair identity is replaced by (virtual_top + widget_line); EVERY feature
    that currently keys off the Text line number (selection, hover, diff-nav jump,
    manual-merge row clicks, see()/anchor restore, per-line diff/sel tags, row
    header gutter) must be rewritten to translate through the window offset.
  * tag coloring must be re-applied to the window on every swap.
  * the 3 data panes + 3 row-header gutters + C-area must all swap in lock-step.
"""
import tkinter as tk
from tkinter import ttk

TOTAL_ROWS = 200_000
COLS = 8
BUFFER = 4  # extra rows above/below the viewport


def make_row(i: int) -> str:
    cells = [f"r{i:06d}c{c}" for c in range(COLS)]
    return "   ".join(cells)


class VirtualPane:
    def __init__(self, parent):
        self.text = tk.Text(parent, wrap="none", font=("Consolas", 10),
                            width=80, height=30, cursor="arrow")
        self.text.pack(side="left", fill="both", expand=True)
        self.text.configure(state="disabled")

    def render_window(self, top: int, count: int):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        lines = [make_row(top + k) for k in range(count) if top + k < TOTAL_ROWS]
        self.text.insert("1.0", "\n".join(lines))
        self.text.configure(state="disabled")


class App:
    def __init__(self, root):
        self.root = root
        root.title("Virtualized Text render prototype")
        top = ttk.Frame(root)
        top.pack(fill="both", expand=True)

        self.left = VirtualPane(top)
        self.right = VirtualPane(top)

        self.sb = ttk.Scrollbar(top, orient="vertical", command=self._on_scrollbar)
        self.sb.pack(side="right", fill="y")

        self.status = ttk.Label(root, anchor="w")
        self.status.pack(fill="x")

        self.visible_rows = 30
        self.virtual_top = 0

        # Mouse wheel on either pane drives the virtual offset.
        for pane in (self.left, self.right):
            pane.text.bind("<MouseWheel>", self._on_wheel)

        self.root.after(50, self._refresh)

    def _max_top(self) -> int:
        return max(0, TOTAL_ROWS - self.visible_rows)

    def _on_scrollbar(self, *args):
        if not args:
            return
        op = args[0]
        if op == "moveto":
            frac = float(args[1])
            self.virtual_top = int(frac * self._max_top())
        elif op == "scroll":
            amount = int(args[1])
            unit = args[2]
            step = self.visible_rows if unit == "pages" else 1
            self.virtual_top += amount * step
        self.virtual_top = max(0, min(self.virtual_top, self._max_top()))
        self._refresh()

    def _on_wheel(self, event):
        delta = -1 if event.delta > 0 else 1
        self.virtual_top = max(0, min(self.virtual_top + delta * 3, self._max_top()))
        self._refresh()
        return "break"

    def _refresh(self):
        top = max(0, self.virtual_top - BUFFER)
        count = self.visible_rows + 2 * BUFFER
        self.left.render_window(top, count)
        self.right.render_window(top, count)
        # Proxy scrollbar reflects the FULL virtual range.
        denom = float(TOTAL_ROWS)
        first = self.virtual_top / denom
        last = (self.virtual_top + self.visible_rows) / denom
        self.sb.set(first, last)
        self.status.configure(
            text=(f"virtual rows {self.virtual_top}..{self.virtual_top + self.visible_rows} "
                  f"of {TOTAL_ROWS}  |  Text holds {count} lines (constant)")
        )


def main():
    root = tk.Tk()
    root.geometry("1100x650")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
