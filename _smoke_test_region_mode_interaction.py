"""Headless guards for region-mode fallback selection and visible feedback."""

from __future__ import annotations

from types import SimpleNamespace

import sow_merge_tool as smt


class _Var:
    def __init__(self, value=0):
        self.value = value

    def get(self):
        return self.value


class _Info:
    def __init__(self):
        self.texts = []

    def configure(self, **kwargs):
        if "text" in kwargs:
            self.texts.append(str(kwargs["text"]))


class _Root:
    def __init__(self):
        self.bells = 0

    def bell(self):
        self.bells += 1

    def configure(self, **_kwargs):
        return None

    def update_idletasks(self):
        return None


class _Projection:
    def slot(self, logical_col):
        return SimpleNamespace(
            logical_idx=int(logical_col) - 1,
            state="retained",
            confidence=SimpleNamespace(ambiguous=False),
        )

    def physical_col(self, _side, logical_col):
        return int(logical_col)


def _base_view():
    view = object.__new__(smt.SheetView)
    view.sheet = "Data"
    view.row_pairs = [(idx + 1, idx + 1) for idx in range(10)]
    view.display_rows = list(range(10))
    view.pair_diff_cols = {1: {1}, 2: {1}, 7: {1}}
    view.pair_base_diff_cols = {7: {1}}
    view.only_diff_var = _Var(0)
    view.app = SimpleNamespace(
        has_base=True,
        ws_a_val=lambda *_args: (_ for _ in ()).throw(AssertionError("worksheet read")),
        ws_b_val=lambda *_args: (_ for _ in ()).throw(AssertionError("worksheet read")),
        ws_base_val=lambda *_args: (_ for _ in ()).throw(AssertionError("worksheet read")),
    )
    view._pair_has_visual_diff = lambda pair_idx: bool(
        view.pair_diff_cols.get(pair_idx) or view.pair_base_diff_cols.get(pair_idx)
    )
    view._active_column_projection = lambda: _Projection()
    view._base_row_for_pair = lambda pair_idx, _pair: pair_idx + 1
    return view


def test_nearest_first_and_direction_filtering_are_deterministic():
    view = _base_view()
    block, relocated = view._resolve_region_action_target("B2A", 1)
    assert (block.start_pair_idx, block.end_pair_idx, relocated) == (1, 2, False)

    block, relocated = view._resolve_region_action_target("B2A", 5)
    assert (block.start_pair_idx, relocated) == (7, True)

    # Pair 4 is equally far from [1,2] and [7,7]; earlier wins.
    block, relocated = view._resolve_region_action_target("B2A", 4)
    assert (block.start_pair_idx, relocated) == (1, True)

    block, relocated = view._resolve_region_action_target("B2A", None)
    assert (block.start_pair_idx, relocated) == (1, True)

    # Direction-specific source rows prevent landing on a one-sided block that
    # the clicked button cannot apply.
    view.pair_diff_cols = {1: {-1}, 7: {-1}}
    view.row_pairs[1] = (2, None)
    view.row_pairs[7] = (None, 8)
    left_target = view._resolve_region_action_target("A2B", 7)
    right_target = view._resolve_region_action_target("B2A", 1)
    assert left_target is not None and left_target[0].start_pair_idx == 1
    assert right_target is not None and right_target[0].start_pair_idx == 7

    base_target = view._resolve_region_action_target("BASE2A", 1)
    assert base_target is None  # pair 7 has no Mine destination row


def test_locator_reuses_navigation_and_selects_block_start():
    view = _base_view()
    block = smt._DiffBlock(1, (7,), 7, 7, True)
    calls = []
    view.row_to_line = {7: 3}
    view.selected_pair_idx = None
    view._materialize_pair_for_navigation = lambda pair_idx: calls.append(
        ("materialize", pair_idx)
    ) or True

    def _goto(line):
        calls.append(("goto", line))
        view.selected_pair_idx = 7

    view._goto_block_start = _goto
    assert view._locate_region_action_block(block)
    assert calls == [("materialize", 7), ("goto", 3)]


def test_region_button_fallback_locates_then_requires_second_click_headlessly():
    view = _base_view()
    view.pair_diff_cols = {7: {1}}
    view.pair_base_diff_cols = {}
    view._suppress_bg_apply = False
    view._only_diff_async_building = False
    view._formula_copy_skips_pending = 0
    view.selected_pair_idx = 4
    view._last_selected_line = 1
    view.snapshot_only_diff = False
    view.info = _Info()
    view.root = _Root()
    view.app = SimpleNamespace(
        has_base=False,
        _begin_interactive_action=None,
        _end_interactive_action=None,
        push_undo=lambda _action: None,
    )
    view._current_line = lambda: 5
    view.has_explicit_cell_selection = lambda: True
    view.resolved_pair_idx_for_c_area = lambda: 4
    view._preflight_region_formula_copy = lambda *_args: None
    view._capture_view_anchor = lambda: ("anchor",)
    located = []

    def _locate(block):
        located.append(block.start_pair_idx)
        view.selected_pair_idx = block.start_pair_idx
        return True

    view._locate_region_action_block = _locate
    applied = []

    def _copy_row(_direction, **kwargs):
        applied.append(kwargs["override_pair_idx"])
        return True

    view._copy_selected_row = _copy_row
    view._mark_nonstructural_cell_edit = lambda *_args, **_kwargs: None
    view._invalidate_render_cache = lambda: None
    view._refresh_pair_indices_exact = lambda pairs: len(tuple(pairs))
    view.refresh = lambda **_kwargs: None
    view._refresh_diff_block_ui = lambda: None
    view._restore_view_anchor = lambda _anchor: None
    view._update_cursor_lines = lambda: None
    view._show_formula_copy_skip_notice = lambda _count: None

    view._copy_selected_region("A2B")
    assert located == [7]
    assert applied == []
    assert view.selected_pair_idx == 7
    assert view.root.bells == 0
    assert view.info.texts[-1] == (
        "已定位到可使用的左侧差异区域，请再次点击“使用左侧区域”。"
    )

    view._copy_selected_region("A2B")
    assert applied == [7]
    assert view.info.texts[-1].startswith("已采用左侧区域：")


def test_no_applicable_region_has_visible_non_modal_feedback():
    view = _base_view()
    view.pair_diff_cols = {}
    view.pair_base_diff_cols = {}
    view._suppress_bg_apply = False
    view._only_diff_async_building = False
    view._formula_copy_skips_pending = 0
    view.selected_pair_idx = None
    view._last_selected_line = None
    view.info = _Info()
    view.root = _Root()
    view.app = SimpleNamespace(
        has_base=False,
        _begin_interactive_action=None,
        _end_interactive_action=None,
    )
    view._current_line = lambda: 1
    view.has_explicit_cell_selection = lambda: False
    view.resolved_pair_idx_for_c_area = lambda: None

    view._copy_selected_region("B2A")
    assert view.info.texts[-1] == "当前 Sheet 没有可使用的右侧差异区域。"
    assert view.root.bells == 0


def test_explicit_wrong_direction_block_does_not_jump_or_write():
    view = _base_view()
    view.pair_diff_cols = {1: {-1}, 7: {-1}}
    view.row_pairs[1] = (2, None)  # left-only: A2B applies, B2A cannot
    view.row_pairs[7] = (None, 8)  # right-only: B2A applies
    view._suppress_bg_apply = False
    view._only_diff_async_building = False
    view._formula_copy_skips_pending = 0
    view.selected_pair_idx = 1
    view._last_selected_line = 2
    view.info = _Info()
    view.root = _Root()
    view.app = SimpleNamespace(
        has_base=False,
        _begin_interactive_action=None,
        _end_interactive_action=None,
    )
    view._current_line = lambda: 2
    view.has_explicit_cell_selection = lambda: False
    view.resolved_pair_idx_for_c_area = lambda: 1
    located = []
    view._locate_region_action_block = lambda block: located.append(
        block.start_pair_idx
    ) or True

    view._copy_selected_region("B2A")
    assert located == []
    assert view.selected_pair_idx == 1
    assert view.info.texts[-1] == (
        "当前差异区域不能使用右侧内容，请选择另一侧或其他差异区域。"
    )
    assert view.root.bells == 0


def test_processed_snapshot_block_falls_back_to_next_pending_block():
    view = _base_view()
    view.only_diff_var = _Var(1)
    view.display_rows = [1, 7]
    view.pair_diff_cols = {7: {1}}
    view.pair_base_diff_cols = {}
    processed = smt._DiffBlock(1, (1,), 1, 1, False)
    pending = smt._DiffBlock(2, (7,), 7, 7, True)
    view._diff_block_model_ready = lambda: True
    view._ensure_full_diff_blocks = lambda: [processed, pending]
    view._suppress_bg_apply = False
    view._only_diff_async_building = False
    view._formula_copy_skips_pending = 0
    view.selected_pair_idx = 1
    view._last_selected_line = 1
    view.snapshot_only_diff = True
    view.info = _Info()
    view.root = _Root()
    view.app = SimpleNamespace(
        has_base=False,
        _begin_interactive_action=None,
        _end_interactive_action=None,
    )
    view._current_line = lambda: 1
    view.has_explicit_cell_selection = lambda: False
    view.resolved_pair_idx_for_c_area = lambda: 1
    located = []

    def _locate(block):
        located.append(block.start_pair_idx)
        view.selected_pair_idx = block.start_pair_idx
        return True

    view._locate_region_action_block = _locate
    view._copy_selected_region("B2A")
    assert located == [7]
    assert view.selected_pair_idx == 7
    assert view.info.texts[-1] == (
        "已定位到可使用的右侧差异区域，请再次点击“使用右侧区域”。"
    )
    assert view.root.bells == 0


def test_empty_direction_map_short_circuits_without_block_scan():
    view = _base_view()
    view.pair_base_diff_cols = {}

    def _forbidden_visual_scan():
        raise AssertionError("empty direction map scanned visual blocks")

    view._region_action_visual_blocks = _forbidden_visual_scan
    assert view._resolve_region_action_target("BASE2A", None) is None


def main():
    tests = (
        test_nearest_first_and_direction_filtering_are_deterministic,
        test_locator_reuses_navigation_and_selects_block_start,
        test_region_button_fallback_locates_then_requires_second_click_headlessly,
        test_no_applicable_region_has_visible_non_modal_feedback,
        test_explicit_wrong_direction_block_does_not_jump_or_write,
        test_processed_snapshot_block_falls_back_to_next_pending_block,
        test_empty_direction_map_short_circuits_without_block_scan,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"SMOKE_TEST_REGION_MODE_INTERACTION_OK ({len(tests)} tests)")


if __name__ == "__main__":
    main()
