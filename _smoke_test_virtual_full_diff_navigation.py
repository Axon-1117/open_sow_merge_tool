"""Headless contracts for full-logical diff navigation in virtual panes.

The tests exercise the real SheetView navigation methods with a bounded fake
Text document.  They intentionally provide no worksheet or parser access:
any attempt to consult one raises immediately.
"""

from __future__ import annotations

import math
from types import SimpleNamespace

import sow_merge_tool as sm


class _Var:
    def __init__(self, value: int) -> None:
        self._value = int(value)

    def get(self) -> int:
        return int(self._value)


class _Button:
    def __init__(self) -> None:
        self.state = "disabled"

    def configure(self, **kwargs) -> None:
        if "state" in kwargs:
            self.state = str(kwargs["state"])

    def cget(self, name: str) -> str:
        assert name == "state"
        return self.state


class _Text:
    def __init__(self, *, top_line: int = 1, insert_line: int = 2) -> None:
        self.top_line = int(top_line)
        self.insert_line = int(insert_line)

    def index(self, value: str) -> str:
        if value == "@0,0":
            return f"{self.top_line}.0"
        if value == "insert":
            return f"{self.insert_line}.0"
        raise AssertionError(f"unexpected Text index query: {value!r}")


class _ForbiddenApp:
    def __getattr__(self, name: str):
        raise AssertionError(f"navigation must not access app.{name}")


def _forbidden(name: str):
    def _raise(*_args, **_kwargs):
        raise AssertionError(f"navigation must not access {name}")

    return _raise


def _make_view(*, only_diff: bool, virtual: bool, diff_pairs: set[int], display_rows):
    view = sm.SheetView.__new__(sm.SheetView)
    view.app = _ForbiddenApp()
    view.only_diff_var = _Var(int(only_diff))
    view._virtual_mode_active = lambda: bool(virtual)
    view._data_ready = True
    view._only_diff_async_building = False
    view._only_diff_source_version = 1
    view._data_version = 1
    view.row_pairs = [(index + 1, index + 1) for index in range(7)]
    # A touched, currently equal row must not bridge two logical diff blocks.
    view.touched_rows = {4}
    view._full_display_rows = list(range(7))
    view.display_rows = list(display_rows)
    view.row_to_line = {
        int(pair_idx): line
        for line, pair_idx in enumerate(view.display_rows, start=1)
    }
    view._full_diff_blocks_cache_key = None
    view._full_diff_blocks = []
    view._pair_to_full_diff_block = {}
    view._diff_blocks_cache = None
    view.selected_pair_idx = None
    view._last_selected_line = None
    view.has_explicit_cell_selection = lambda: False
    view.left = _Text()
    view.prev_diff_btn = _Button()
    view.next_diff_btn = _Button()
    view._visual_diff_cols_for_pair = lambda pair_idx: (
        {1} if int(pair_idx) in diff_pairs else set()
    )
    view._update_diff_block_indicator = lambda: None
    view._get_row_values = _forbidden("SheetView._get_row_values")
    view._materialize_pair_for_navigation = _forbidden(
        "SheetView._materialize_pair_for_navigation"
    )
    view._goto_block_start = _forbidden("SheetView._goto_block_start")
    navigated = []

    def _record_global_goto(block_idx: int) -> None:
        blocks = sm.SheetView._ensure_full_diff_blocks(view)
        block = blocks[int(block_idx)]
        navigated.append((int(block_idx), int(block.start_pair_idx)))

    view._goto_full_diff_block = _record_global_goto
    return view, navigated


def _assert_virtual_full_uses_complete_blocks() -> None:
    view, navigated = _make_view(
        only_diff=False,
        virtual=True,
        diff_pairs={1, 2, 5, 6},
        display_rows=[0, 1, 2],
    )
    sm.SheetView._update_diff_nav_state(view)
    blocks = sm.SheetView._ensure_full_diff_blocks(view)
    assert [block.pair_indices for block in blocks] == [(1, 2), (5, 6)]
    assert 4 in view.touched_rows
    assert view.prev_diff_btn.cget("state") == "disabled"
    assert view.next_diff_btn.cget("state") == "normal"
    sm.SheetView._goto_next_diff_block(view)
    assert navigated == [(1, 5)]

    # A recycled document at the second block must navigate back through the
    # same logical model, not a current-window scan.
    view.display_rows = [5, 6]
    view.row_to_line = {5: 1, 6: 2}
    view.left = _Text(top_line=1, insert_line=1)
    sm.SheetView._update_diff_nav_state(view)
    assert view.prev_diff_btn.cget("state") == "normal"
    assert view.next_diff_btn.cget("state") == "disabled"
    sm.SheetView._goto_prev_diff_block(view)
    assert navigated == [(1, 5), (0, 1)]


def _assert_virtual_single_block_stays_disabled() -> None:
    view, navigated = _make_view(
        only_diff=False,
        virtual=True,
        diff_pairs={1, 2},
        display_rows=[0, 1, 2],
    )
    sm.SheetView._update_diff_nav_state(view)
    assert [block.pair_indices for block in sm.SheetView._ensure_full_diff_blocks(view)] == [
        (1, 2)
    ]
    assert view.prev_diff_btn.cget("state") == "disabled"
    assert view.next_diff_btn.cget("state") == "disabled"
    sm.SheetView._goto_next_diff_block(view)
    sm.SheetView._goto_prev_diff_block(view)
    assert navigated == []


def _assert_nonvirtual_full_keeps_visible_window_semantics() -> None:
    view, navigated = _make_view(
        only_diff=False,
        virtual=False,
        diff_pairs={1, 2, 5, 6},
        display_rows=[0, 1, 2],
    )
    sm.SheetView._update_diff_nav_state(view)
    # The second logical block is off the non-virtual current document, so
    # the historic visible-window navigator remains disabled.
    assert view.next_diff_btn.cget("state") == "disabled"
    sm.SheetView._goto_next_diff_block(view)
    assert navigated == []


def _assert_only_diff_still_uses_complete_blocks() -> None:
    view, navigated = _make_view(
        only_diff=True,
        virtual=False,
        diff_pairs={1, 2, 5, 6},
        display_rows=[1, 2],
    )
    sm.SheetView._update_diff_nav_state(view)
    assert view.next_diff_btn.cget("state") == "normal"
    sm.SheetView._goto_next_diff_block(view)
    assert navigated == [(1, 5)]


def _assert_full_navigation_outer_phase_telemetry_is_bounded() -> None:
    """Use the real outer navigator while keeping all UI work headless."""
    view = sm.SheetView.__new__(sm.SheetView)
    view.app = SimpleNamespace(selected_sheet="S1")
    view.sheet = "S1"
    view._viewport_request_generation = lambda: 9
    view._ensure_full_diff_blocks = lambda: [
        sm._DiffBlock(ordinal=0, pair_indices=(4, 5), start_pair_idx=4, end_pair_idx=5, pending=False)
    ]
    view._materialize_pair_for_navigation = lambda pair_idx: int(pair_idx) == 4
    view.row_to_line = {4: 3}
    captured = {}

    def _goto(line, *, navigation_phases=None):
        assert int(line) == 3
        captured["line"] = int(line)
        navigation_phases.update({"selection": 1.0, "restore": 2.0, "ui": 3.0})

    view._goto_block_start = _goto
    sm.SheetView._goto_full_diff_block(view, 0)
    assert captured == {"line": 3}
    record = dict(view._last_diff_navigation_telemetry)
    assert record["action"] == "full-diff-block"
    assert record["block_idx"] == 0
    assert record["pair_idx"] == 4
    assert record["outcome"] == "complete"
    assert record["generation"] == 9
    assert record["selected_sheet"] == "S1"
    assert set(record["phase_ms"]) == {
        "block_lookup",
        "materialize_publish",
        "selection",
        "restore",
        "main_x_observe",
        "main_x_fallback",
        "ui",
        "total",
    }
    assert record["phase_ms"]["selection"] == 1.0
    assert record["phase_ms"]["restore"] == 2.0
    assert record["phase_ms"]["main_x_observe"] == 0.0
    assert record["phase_ms"]["main_x_fallback"] == 0.0
    assert record["phase_ms"]["ui"] == 3.0
    assert all(
        math.isfinite(value) and value >= 0.0
        for value in record["phase_ms"].values()
    )
    assert len(view._diff_navigation_telemetry) == 1
    assert view._diff_navigation_telemetry.maxlen == 128


def main() -> None:
    tests = (
        _assert_virtual_full_uses_complete_blocks,
        _assert_virtual_single_block_stays_disabled,
        _assert_nonvirtual_full_keeps_visible_window_semantics,
        _assert_only_diff_still_uses_complete_blocks,
        _assert_full_navigation_outer_phase_telemetry_is_bounded,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}", flush=True)
    print(f"SMOKE_VIRTUAL_FULL_DIFF_NAVIGATION_OK ({len(tests)} tests)", flush=True)


if __name__ == "__main__":
    main()
