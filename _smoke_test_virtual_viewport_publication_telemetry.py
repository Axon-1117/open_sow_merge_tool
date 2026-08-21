"""Pure contract for bounded virtual-publication telemetry.

This test has no Tk root or workbook.  It exercises the production publisher
with fake Text widgets, a deterministic clock, and a fake formatter so phase
boundaries plus A/B/Base cache hit/miss counts remain observable without
changing render output or interaction semantics.
"""

from __future__ import annotations

from collections import OrderedDict, deque

import sow_merge_tool as sm


class _Var:
    def __init__(self, value: int = 0) -> None:
        self.value = int(value)

    def get(self) -> int:
        return int(self.value)


class _Widget:
    def __init__(self) -> None:
        self._w = "fake-text"
        self.tk = self
        self.text = ""
        self.calls = []

    def call(self, *_args) -> None:
        self.text = str(_args[-1])

    def tag_add(self, *args) -> None:
        self.calls.append(tuple(args))

    def configure(self, **_kwargs) -> None:
        return None


class _Clock:
    def __init__(self) -> None:
        self.value = 100.0

    def perf_counter(self) -> float:
        self.value += 0.0005
        return self.value


class _App:
    selected_sheet = "S1"
    has_base = False

    @staticmethod
    def _is_sheet_exact_current(sheet: str) -> bool:
        return sheet == "S1"

    @staticmethod
    def _sheet_exact_entry(sheet: str) -> dict:
        assert sheet == "S1"
        return {"generation": 3}


def _make_view():
    view = sm.SheetView.__new__(sm.SheetView)
    view.app = _App()
    view.sheet = "S1"
    view.only_diff_var = _Var(0)
    view.row_pairs = [(1, 1), (2, 2)]
    view.pair_raw_parts_a = {0: ("a0",), 1: ("a1",)}
    view.pair_raw_parts_b = {0: ("b0",), 1: ("b1",)}
    view.pair_raw_parts_base = {}
    # Pair zero is a pre-existing formatted cache entry; pair one is a miss.
    view.pair_text_a = {0: "A:a0"}
    view.pair_text_b = {0: "B:b0"}
    view.pair_text_base = {}
    view._prepared_text_lru = OrderedDict(((0, None),))
    view._virtual_window_start = 0
    view._virtual_column_window_start = 0
    view._virtual_scroll_publications = 0
    # The production publisher records a timestamp only after the complete
    # bounded publication and end-to-end completion callback have run.
    view._virtual_publication_timestamps = deque(maxlen=128)
    view._viewport_render_samples_ms = deque(maxlen=128)
    view._virtual_publication_telemetry = deque(maxlen=128)
    view._last_virtual_publication_telemetry = {}
    view._viewport_request_active = {"id": 17, "status": "pending"}
    view.left = _Widget()
    view.base = _Widget()
    view.right = _Widget()
    view.left_ln = _Widget()
    view.base_ln = _Widget()
    view.right_ln = _Widget()
    view.info = _Widget()
    formatter_calls = []

    def _formatter(raw_parts, side):
        formatter_calls.append((str(side), tuple(raw_parts)))
        return f"{side}:{','.join(str(value) for value in raw_parts)}"

    view._render_line_from_raw_parts = _formatter
    view._virtual_mode_active = lambda: True
    view._apply_pending_virtual_column_window = lambda: False
    view._virtual_window_rows = lambda _start: [0, 1]
    view._build_visual_diff_surface_context = lambda _rows: None
    view._visual_diff_cols_for_pair = lambda _pair_idx, **_kwargs: {1}
    view._rendered_logical_columns = lambda: (1,)
    view._build_diffcell_surface_context = lambda _cols: object()
    view._set_virtual_flat_diff_background = lambda _enabled: None
    view._is_three_way_enabled = lambda: False
    view._diffcell_tag_args_for_line = lambda *_args, **_kwargs: ([], [], [], [], [], [])
    view._padding_column_tag_args = lambda *_args: []
    view._render_col_headers = lambda: None
    view._render_row_headers_full = lambda: None
    view._normal_sheet_info_text = lambda: "ready"
    view._refresh_diff_block_ui = lambda: None
    view._update_diff_nav_state = lambda: None
    view._update_cursor_lines = lambda: None
    view._virtual_scroll_fractions = lambda: (0.0, 1.0)
    view._yscroll_all = lambda *_args: None
    view._set_wide_column_scrollbars = lambda: None
    view._trim_prepared_text_cache = lambda: None
    view._viewport_request_generation = lambda: 3
    view._complete_viewport_request_if_current = lambda: None
    return view, formatter_calls


def _assert_publication_telemetry_is_bounded_and_exact() -> None:
    view, formatter_calls = _make_view()
    clock = _Clock()
    original_clock = sm.time.perf_counter
    try:
        sm.time.perf_counter = clock.perf_counter
        assert sm.SheetView._publish_virtual_window(view, 0) is True
    finally:
        sm.time.perf_counter = original_clock

    # Existing text cache output is retained; only the missing pair is passed
    # through the fake formatter, once per populated A/B side.
    assert view.pair_text_a == {0: "A:a0", 1: "A:a1"}
    assert view.pair_text_b == {0: "B:b0", 1: "B:b1"}
    assert formatter_calls == [("A", ("a1",)), ("B", ("b1",))]
    assert len(view._virtual_publication_telemetry) == 1
    assert view._virtual_publication_telemetry.maxlen == 128
    assert len(view._virtual_publication_timestamps) == 1
    assert view._virtual_publication_timestamps.maxlen == 128
    assert float(view._virtual_publication_timestamps[-1]) > 100.0
    telemetry = dict(view._virtual_publication_telemetry[-1])
    assert telemetry["publication_seq"] == 1
    assert telemetry["request_id"] == 17
    assert telemetry["generation"] == 3
    assert telemetry["window"] == {
        "row_start": 0,
        "row_count": 2,
        "column_start": 0,
        "rendered_logical_columns": (1,),
    }
    assert set(telemetry["phase_ms"]) == {
        "bounded_row_selection",
        "visual_context",
        "visual_map_projection",
        "diffcell_context",
        "line_render_context",
        "flat_bg",
        "materialize",
        "row_map",
        "tk_replace",
        "headers",
        "tags",
        "cursor",
        "finalize",
        "render_total",
    }
    assert all(float(value) >= 0.0 for value in telemetry["phase_ms"].values())
    sides = telemetry["materialize"]["sides"]
    assert sides["A"] == {"hit": 1, "miss": 1, "absent": 0}
    assert sides["B"] == {"hit": 1, "miss": 1, "absent": 0}
    assert sides["BASE"] == {"hit": 0, "miss": 0, "absent": 2}
    assert telemetry["materialize"]["rows"] == 2
    context = dict(telemetry["render_context"])
    # This stripped fake has no immutable column projection, so the publisher
    # must conservatively retain the fake legacy formatter for its complete
    # bounded surface rather than inventing a projection.
    assert context["enabled"] is False
    assert context["fallback"] is True
    assert context["readiness_scan_count"] == 0
    assert context["formatted"] == {
        "context": 0,
        "legacy": 2,
        "base_empty_legacy": 0,
    }
    assert view._last_virtual_render_phases_ms["detail_ms"] == telemetry["phase_ms"]
    assert view._last_virtual_publication_telemetry == telemetry


def main() -> None:
    _assert_publication_telemetry_is_bounded_and_exact()
    print("SMOKE_VIRTUAL_VIEWPORT_PUBLICATION_TELEMETRY_OK", flush=True)


if __name__ == "__main__":
    main()
