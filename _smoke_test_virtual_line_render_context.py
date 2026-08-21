"""Pure parity contract for frozen virtual line-render contexts.

The test deliberately uses immutable ``ColumnSlot``/``ColumnModel`` records
and raw fragments only.  It creates no Tk root, workbook, worksheet, or
editable backend.  Every optimized line must remain byte-identical to the
historical formatter for the same one-publication projection.
"""

from __future__ import annotations

from collections import OrderedDict

import sow_merge_tool as sm


class _Var:
    def __init__(self, value: int) -> None:
        self.value = int(value)

    def get(self) -> int:
        return int(self.value)


def _projection(slots: tuple[sm.ColumnSlot, ...]) -> sm.LogicalColumnProjection:
    model = sm.ColumnModel.from_slots(
        sm.ColumnModelCacheKey("render-context", 1, 1, 0, 0, 0),
        slots,
    )
    return sm.LogicalColumnProjection.from_model(model)


def _make_view(
    slots: tuple[sm.ColumnSlot, ...],
    *,
    grid: bool,
    widths: tuple[int, ...],
) -> sm.SheetView:
    view = sm.SheetView.__new__(sm.SheetView)
    projection = _projection(slots)
    view.column_projection = projection
    view._active_column_projection = lambda: projection
    view.grid_overlay_var = _Var(int(grid))
    view.col_char_widths = {
        index: int(width)
        for index, width in enumerate(widths, start=1)
    }
    view._column_projection_generation = 7
    view._col_widths_version = 11
    view._virtual_column_window_generation = 13
    view._virtual_column_window_start = 0
    view._data_version = 17
    view.pair_raw_parts_a = {}
    view.pair_raw_parts_b = {}
    view.pair_raw_parts_base = {}
    view.pair_text_a = {}
    view.pair_text_b = {}
    view.pair_text_base = {}
    view._prepared_text_lru = OrderedDict()
    # A renderer context must never need completeness scanning.  If a future
    # change does, this pure contract fails rather than hiding an O(full rows)
    # publication path behind an optimized line formatter.
    view._has_complete_prepared_rows = lambda: (_ for _ in ()).throw(
        AssertionError("render context scanned full prepared rows")
    )
    return view


def _context(view: sm.SheetView, columns: tuple[int, ...]):
    context, telemetry = sm.SheetView._build_virtual_line_render_context(
        view, columns
    )
    assert context is not None, telemetry
    assert telemetry["enabled"] is True, telemetry
    assert telemetry["fallback"] is False, telemetry
    assert telemetry["readiness_scan_count"] == 0, telemetry
    return context, telemetry


def _legacy_and_context_equal(view, raw_parts, side: str, context) -> None:
    legacy = sm.SheetView._render_line_from_raw_parts(view, list(raw_parts), side)
    optimized = sm.SheetView._render_line_from_raw_parts(
        view,
        list(raw_parts),
        side,
        render_context=context,
    )
    assert optimized.encode("utf-8") == legacy.encode("utf-8"), (
        side,
        raw_parts,
        legacy,
        optimized,
    )


def _two_way_reorder_missing_placeholder_parity() -> None:
    slots = (
        sm.ColumnSlot(0, mine_col=2, theirs_col=1),
        sm.ColumnSlot(1, mine_col=1, theirs_col=None),
        sm.ColumnSlot(2, mine_col=3, theirs_col=3),
    )
    view = _make_view(slots, grid=False, widths=(5, 4, 7))
    context, telemetry = _context(view, (1, 2, 3))
    _legacy_and_context_equal(view, ("A-one", "A-two", "A-three"), "A", context)
    _legacy_and_context_equal(view, ("B-one", "B-two", "B-three"), "B", context)
    # The missing second theirs slot is a structural placeholder, not a short
    # raw-row fallback.  Preserve the placeholder byte-for-byte.
    rendered = sm.SheetView._render_line_from_raw_parts(
        view, ["B-one", "B-two", "B-three"], "B", render_context=context
    )
    assert sm._LOGICAL_COLUMN_PLACEHOLDER in rendered
    assert telemetry["offset_counts"] == {"A": 3, "BASE": 0, "B": 2}


def _three_way_grid_wide_unicode_and_empty_base_parity() -> None:
    slots = (
        sm.ColumnSlot(0, mine_col=2, base_col=1, theirs_col=3),
        sm.ColumnSlot(1, mine_col=1, base_col=2, theirs_col=1),
        sm.ColumnSlot(2, mine_col=3, base_col=None, theirs_col=2),
    )
    view = _make_view(slots, grid=True, widths=(6, 5, 4))
    raw_a = ("é\u200b", "表", "tail")
    raw_b = ("B-one", "B-two", "B-three")
    raw_base = ("base-one", "base-two")
    # First/middle/last-style bounded column windows use distinct contexts but
    # must always equal the historical projection/render path.
    window_outputs = []
    for columns in ((1, 2), (2, 3), (1, 2, 3)):
        # The optimized context receives the publication's bounded columns.
        # Give the legacy comparison the exact same publication window while
        # keeping each loop isolated from a prior window's fake override.
        view = _make_view(slots, grid=True, widths=(6, 5, 4))
        view._rendered_logical_columns = lambda expected=tuple(columns): expected
        context, telemetry = _context(view, columns)
        _legacy_and_context_equal(view, raw_a, "A", context)
        _legacy_and_context_equal(view, raw_b, "B", context)
        _legacy_and_context_equal(view, raw_base, "BASE", context)
        base_line = sm.SheetView._render_line_from_raw_parts(
            view, list(raw_base), "BASE", render_context=context
        )
        if 3 in columns:
            assert sm._LOGICAL_COLUMN_PLACEHOLDER in base_line, base_line
        window_outputs.append(
            sm.SheetView._render_line_from_raw_parts(
                view, list(raw_a), "A", render_context=context
            )
        )
        # ``side=None`` has logical raw semantics and deliberately bypasses
        # the publication context even when one is supplied.
        legacy_logical = sm.SheetView._render_line_from_raw_parts(
            view, ["logical", "raw", "parts"], None
        )
        optimized_logical = sm.SheetView._render_line_from_raw_parts(
            view,
            ["logical", "raw", "parts"],
            None,
            render_context=context,
        )
        assert optimized_logical.encode("utf-8") == legacy_logical.encode("utf-8")
        assert telemetry["versions"]["window_generation"] == 13

    assert len(set(window_outputs)) == 3, window_outputs
    view = _make_view(slots, grid=True, widths=(6, 5, 4))
    view._rendered_logical_columns = lambda: (1, 2, 3)
    context, _telemetry = _context(view, (1, 2, 3))
    view.pair_raw_parts_a = {0: raw_a}
    view.pair_raw_parts_b = {0: raw_b}
    view.pair_raw_parts_base = {0: ()}
    stats = sm.SheetView._materialize_prepared_pair_text(
        view, [0], render_context=context
    )
    assert view.pair_text_base[0] == ""
    assert stats["formatter"]["base_empty_legacy"] == 1
    assert stats["formatter"]["context_formatted"] == 2
    assert stats["formatter"]["legacy_formatted"] == 0


def _short_raw_and_exception_fallback_are_publication_consistent() -> None:
    slots = (
        sm.ColumnSlot(0, mine_col=1, theirs_col=1),
        sm.ColumnSlot(1, mine_col=3, theirs_col=2),
        sm.ColumnSlot(2, mine_col=2, theirs_col=3),
    )
    view = _make_view(slots, grid=False, widths=(4, 4, 4))
    context, _telemetry = _context(view, (1, 2, 3))
    # A short raw side is legal and must retain the legacy empty-cell result.
    _legacy_and_context_equal(view, ("short",), "A", context)

    raw_a = ("A1", "A2", "A3")
    raw_b = ("B1", "B2", "B3")
    expected_a = sm.SheetView._render_line_from_raw_parts(view, list(raw_a), "A")
    expected_b = sm.SheetView._render_line_from_raw_parts(view, list(raw_b), "B")
    view.pair_raw_parts_a = {0: raw_a}
    view.pair_raw_parts_b = {0: raw_b}
    view.pair_raw_parts_base = {}
    original = view._render_line_from_raw_parts
    context_calls = 0

    def _raise_once(raw_parts, side=None, *, render_context=None):
        nonlocal context_calls
        if render_context is not None:
            context_calls += 1
            if context_calls == 2:
                raise RuntimeError("injected-context-format-failure")
        return original(raw_parts, side, render_context=render_context)

    view._render_line_from_raw_parts = _raise_once
    stats = sm.SheetView._materialize_prepared_pair_text(
        view, [0], render_context=context
    )
    assert stats["formatter"] == {
        "context_enabled": True,
        "fallback": True,
        "fallback_reason": "format:RuntimeError",
        "context_formatted": 0,
        "legacy_formatted": 2,
        "base_empty_legacy": 0,
    }
    assert view.pair_text_a[0].encode("utf-8") == expected_a.encode("utf-8")
    assert view.pair_text_b[0].encode("utf-8") == expected_b.encode("utf-8")


def main() -> None:
    _two_way_reorder_missing_placeholder_parity()
    _three_way_grid_wide_unicode_and_empty_base_parity()
    _short_raw_and_exception_fallback_are_publication_consistent()
    print("SMOKE_VIRTUAL_LINE_RENDER_CONTEXT_OK", flush=True)


if __name__ == "__main__":
    main()
