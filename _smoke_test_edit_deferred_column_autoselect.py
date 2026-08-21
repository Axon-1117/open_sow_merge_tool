"""Headless contract for structural auto-selection before edit preload."""
from __future__ import annotations

import argparse
from types import SimpleNamespace

import sow_merge_tool as sm


_CASE = "edit-deferred-column-autoselect"


class _FakeView:
    def __init__(self, state: str, *, structural: bool = True):
        self._lifecycle_state = state
        self._column_projection_generation = 7
        self._auto_structural_selection_generation = None
        self.selected_column_logical_range = None
        self.selected_column_block_ordinal = None
        self._selected = None
        self.select_calls = []
        self.status_calls = []
        self.preload_calls = []
        self.manual_ops = []
        self.undo_stack = []
        self.refresh_calls = []
        self._cache = SimpleNamespace(
            structural_diff_cols=frozenset({14}) if structural else frozenset(),
            unresolved_cols=frozenset(),
        )
        self._block = SimpleNamespace(
            ordinal=1,
            start_slot_idx=13,
            end_slot_idx=13,
            slot_indices=(13,),
        )
        self._projection = SimpleNamespace(
            model=SimpleNamespace(blocks=(self._block,))
        )
        self.app = SimpleNamespace(
            _request_edit_preload=lambda *args, **kwargs: self.preload_calls.append(
                (args, kwargs)
            )
        )

    def _ensure_column_projection_current(self, reason):
        assert reason == "自动选择结构列块"
        return self._projection

    def _selected_column_block(self):
        return self._selected

    def _active_column_comparison_cache(self):
        return self._cache

    def _column_block_is_structural(self, block):
        assert block is self._block
        return bool(self._cache.structural_diff_cols or self._cache.unresolved_cols)

    def _select_column_block_by_logical_col(self, logical_col, source_side):
        self.select_calls.append((logical_col, source_side))
        assert logical_col == 14 and source_side == "LOGICAL"
        self._selected = self._block
        self.selected_column_block_ordinal = self._block.ordinal
        self.selected_column_logical_range = (14, 14)
        return self._block

    def _column_block_cause_text(self, block):
        assert block is self._block
        return "Target Working缺列"

    def _set_column_action_status(self, compact, detail):
        self.status_calls.append((compact, detail))


def _run_auto(view: _FakeView):
    return sm.SheetView._auto_select_first_structural_column_block_if_ready(view)


def _assert_no_write(view: _FakeView) -> None:
    assert view.preload_calls == []
    assert view.manual_ops == []
    assert view.undo_stack == []
    assert view.refresh_calls == []


def run_case() -> None:
    for state in ("READY", "EDIT_DEFERRED"):
        view = _FakeView(state)
        selected = _run_auto(view)
        assert selected is view._block
        assert view.select_calls == [(14, "LOGICAL")]
        assert view.selected_column_logical_range == (14, 14)
        assert view._auto_structural_selection_generation == 7
        assert view.status_calls == [
            (
                "待处理 N 已自动选｜可执行",
                "已自动选择待处理列块 N｜原因：Target Working缺列",
            )
        ]
        _assert_no_write(view)

    for state in ("EDIT_LOADING", "DIFFING", "FAILED", "UNRESOLVED", "CLOSING"):
        view = _FakeView(state)
        assert _run_auto(view) is None
        assert view.select_calls == []
        assert view._auto_structural_selection_generation is None
        _assert_no_write(view)

    no_structure = _FakeView("EDIT_DEFERRED", structural=False)
    assert _run_auto(no_structure) is None
    assert no_structure.select_calls == []
    assert no_structure._auto_structural_selection_generation == 7
    _assert_no_write(no_structure)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case", choices=(_CASE,))
    parser.add_argument("--list-cases", action="store_true")
    args = parser.parse_args()
    if args.list_cases:
        print(_CASE)
        return
    run_case()
    print("PASS " + (args.case or _CASE))


if __name__ == "__main__":
    main()
