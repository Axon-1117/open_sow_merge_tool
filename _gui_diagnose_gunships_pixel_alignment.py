"""Read-only Tk pixel-geometry diagnostic for real GunshipsModify@design.

The diagnostic intentionally records evidence before choosing an acceptance
tolerance.  It never saves the startup candidate or any SVN input.

Run:
  python _gui_diagnose_gunships_pixel_alignment.py
"""

from __future__ import annotations

import json
import re
import sys
from collections import defaultdict

import tkinter as tk

from _gui_self_test_latest_gunships_feedback import (
    _pump,
    _real_gunships_app,
    _wait_until,
)


_L_TOKEN_RE = re.compile(r"\bL\d+(?::L\d+)?\b")
_CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
_ASCII_WORD_RE = re.compile(r"[A-Za-z]")
_ROOT_TOP_BUTTON_TEXTS = (
    "重算并刷新",
    "导出诊断包",
    "复制反馈信息",
    "检查更新",
)


def _rect(widget) -> tuple[int, int, int, int]:
    left = int(widget.winfo_rootx())
    top = int(widget.winfo_rooty())
    return (
        left,
        top,
        left + int(widget.winfo_width()),
        top + int(widget.winfo_height()),
    )


def _language(value) -> str:
    text = "" if value is None else str(value)
    has_cjk = bool(_CJK_RE.search(text))
    has_ascii = bool(_ASCII_WORD_RE.search(text))
    if has_cjk and has_ascii:
        return "mixed"
    if has_cjk:
        return "zh"
    if has_ascii:
        return "en"
    if text:
        return "numeric/symbol"
    return "empty"


def _widget_text(widget) -> str:
    if isinstance(widget, tk.Text):
        return str(widget.get("1.0", "end-1c"))
    try:
        text = str(widget.cget("text"))
        if text:
            return text
    except Exception:
        pass
    try:
        variable = str(widget.cget("textvariable"))
        return str(widget.getvar(variable)) if variable else ""
    except Exception:
        return ""


def _descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from _descendants(child)


def _document_xpixels(widget: tk.Text, line_no: int, char_index: int) -> int:
    result = widget.count(
        f"{line_no}.0",
        f"{line_no}.{int(char_index)}",
        "xpixels",
    )
    assert result is not None and len(result) == 1, (
        widget,
        line_no,
        char_index,
        result,
    )
    return int(result[0])


def _physical_row_for_side(view, pair_idx: int, side: str) -> int | None:
    pair = view.row_pairs[pair_idx]
    if side == "A":
        return view._row_for_side(pair, "A")
    if side == "B":
        return view._row_for_side(pair, "B")
    return view._base_row_for_pair(pair_idx, pair)


def _worksheet_for_side(app, sheet: str, side: str):
    if side == "A":
        return app.ws_a_val(sheet)
    if side == "B":
        return app.ws_b_val(sheet)
    return app.ws_base_val(sheet)


def _pixel_boundary_records(app, view):
    projection = view._active_column_projection()
    side_widgets = (
        ("A", view.left),
        ("BASE", view.base),
        ("B", view.right),
    )
    records = []
    for mine_row in range(1, 6):
        pair_idx = view.row_a_to_pair_idx.get(mine_row)
        assert pair_idx is not None, f"Gunships Mine row A{mine_row} is not aligned"
        line_no = view.row_to_line.get(pair_idx)
        assert line_no is not None, (
            mine_row,
            pair_idx,
            view.display_rows[:10],
        )
        for side, widget in side_widgets:
            physical_row = _physical_row_for_side(view, pair_idx, side)
            assert physical_row is not None, (side, mine_row, pair_idx)
            line_text = widget.get(f"{line_no}.0", f"{line_no}.end")
            spans = view._spans_for_line(line_text)
            boundaries = {}
            cells = {}
            worksheet = _worksheet_for_side(app, view.sheet, side)
            for logical_col in range(1, 6):
                assert logical_col in spans, (side, mine_row, spans)
                physical_col = projection.physical_col(side, logical_col)
                assert physical_col is not None, (side, logical_col)
                _start, end = spans[logical_col]
                boundaries[chr(64 + logical_col)] = _document_xpixels(
                    widget,
                    line_no,
                    end,
                )
                value = worksheet.cell(
                    row=physical_row,
                    column=physical_col,
                ).value
                cells[chr(64 + logical_col)] = {
                    "value": value,
                    "language": _language(value),
                }
            records.append(
                {
                    "side": side,
                    "mine_row": mine_row,
                    "physical_row": physical_row,
                    "pair_idx": pair_idx,
                    "display_line": line_no,
                    "a_right_px": boundaries["A"],
                    "boundaries_px": boundaries,
                    "cells": cells,
                }
            )
    return records


def _drift_summary(records):
    grouped = defaultdict(list)
    for record in records:
        side = record["side"]
        for column, pixels in record["boundaries_px"].items():
            grouped[(side, column)].append(int(pixels))
    summary = []
    for (side, column), pixels in sorted(grouped.items()):
        summary.append(
            {
                "side": side,
                "column": column,
                "pixels": pixels,
                "min": min(pixels),
                "max": max(pixels),
                "drift_px": max(pixels) - min(pixels),
            }
        )
    return summary


def _merged_intervals(intervals):
    merged = []
    for start, end in sorted(intervals):
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return merged


def _bar_whitespace(bar, children):
    bar_left, bar_top, bar_right, bar_bottom = _rect(bar)
    width = bar_right - bar_left
    height = bar_bottom - bar_top
    mapped = []
    intervals = []
    for name, widget in children:
        if widget is None or not bool(widget.winfo_ismapped()):
            continue
        left, top, right, bottom = _rect(widget)
        relative = (left - bar_left, top - bar_top, right - bar_left, bottom - bar_top)
        mapped.append(
            {
                "name": name,
                "text": _widget_text(widget),
                "rect": relative,
                "requested": (
                    int(widget.winfo_reqwidth()),
                    int(widget.winfo_reqheight()),
                ),
            }
        )
        intervals.append((max(0, relative[0]), min(width, relative[2])))
    merged = _merged_intervals(intervals)
    gaps = []
    cursor = 0
    for start, end in merged:
        if start > cursor:
            gaps.append((cursor, start, start - cursor))
        cursor = max(cursor, end)
    if cursor < width:
        gaps.append((cursor, width, width - cursor))
    return {
        "bar_rect": (bar_left, bar_top, bar_right, bar_bottom),
        "size": (width, height),
        "children": mapped,
        "horizontal_blank_gaps": gaps,
        "largest_horizontal_blank_px": max(
            (gap[2] for gap in gaps),
            default=0,
        ),
    }


def _top_button_inventory(view):
    core = (
        ("undo", view.undo_btn),
        ("use_left", view.use_left_group),
        ("keep_mine", view.use_base_btn),
        ("use_right", view.use_right_group),
    )
    view_modes = (
        ("only_diff", view.only_diff_cb),
        ("force_align", view.force_align_cb),
        ("grid_overlay", view.grid_overlay_cb),
        ("three_way", view.three_way_cb),
    )
    toolbar_left, toolbar_top, _right, _bottom = _rect(view._toolbar)

    def _items(items):
        result = []
        for name, widget in items:
            assert widget is not None and bool(widget.winfo_ismapped()), name
            left, top, right, bottom = _rect(widget)
            result.append(
                {
                    "name": name,
                    "text": _widget_text(widget),
                    "rect": (
                        left - toolbar_left,
                        top - toolbar_top,
                        right - toolbar_left,
                        bottom - toolbar_top,
                    ),
                    "root_rect": (left, top, right, bottom),
                }
            )
        return sorted(result, key=lambda item: item["rect"][0])

    return {
        "core_four_left_to_right": _items(core),
        "view_mode_four_left_to_right": _items(view_modes),
    }


def _root_top_button_inventory(app):
    by_text = {
        _widget_text(widget): widget
        for widget in _descendants(app.root)
        if bool(widget.winfo_ismapped())
        and _widget_text(widget) in _ROOT_TOP_BUTTON_TEXTS
    }
    assert tuple(text for text in _ROOT_TOP_BUTTON_TEXTS if text in by_text) == (
        _ROOT_TOP_BUTTON_TEXTS
    ), by_text
    buttons = [by_text[text] for text in _ROOT_TOP_BUTTON_TEXTS]
    owners = {str(button.master) for button in buttons}
    assert len(owners) == 1, owners
    top = buttons[0].master
    top_left, top_y, top_right, top_bottom = _rect(top)
    records = []
    for expected_ordinal, (text, button) in enumerate(
        zip(_ROOT_TOP_BUTTON_TEXTS, buttons),
        start=1,
    ):
        left, y, right, bottom = _rect(button)
        records.append(
            {
                "ordinal": expected_ordinal,
                "text": text,
                "rect": (
                    left - top_left,
                    y - top_y,
                    right - top_left,
                    bottom - top_y,
                ),
                "root_rect": (left, y, right, bottom),
            }
        )
    visual_order = [
        item["text"]
        for item in sorted(records, key=lambda item: item["rect"][0])
    ]
    group_left = min(item["rect"][0] for item in records)
    group_right = max(item["rect"][2] for item in records)
    return {
        "top_rect": (top_left, top_y, top_right, top_bottom),
        "top_size": (top_right - top_left, top_bottom - top_y),
        "expected_order": list(_ROOT_TOP_BUTTON_TEXTS),
        "visual_order": visual_order,
        "order_matches": visual_order == list(_ROOT_TOP_BUTTON_TEXTS),
        "buttons": records,
        "first_left_gap_px": group_left,
        "last_right_gap_px": (top_right - top_left) - group_right,
        "group_width_px": group_right - group_left,
    }


def _column_action_hints(view):
    sources = {
        "status_var": view.column_action_status_var.get(),
        "prefix_var": view.column_action_status_prefix_var.get(),
        "token_var": view.column_action_selection_var.get(),
        "suffix_var": view.column_action_status_suffix_var.get(),
        "structure_summary": view._column_structure_summary(),
        "status_detail": str(
            getattr(view, "_column_action_status_detail", "")
        ),
    }
    widgets = (
        ("prefix_label", view.column_action_status_prefix_label),
        ("token_label", view.column_action_selection_label),
        ("suffix_label", view.column_action_status_suffix_label),
        ("legacy_status_label", view.column_action_status_label),
        ("sheet_info", view.info),
    )
    for name, widget in widgets:
        text = _widget_text(widget)
        detail = str(getattr(widget, "_identity_detail_text", ""))
        sources[f"widget:{name}"] = text
        sources[f"detail:{name}"] = detail
    matching = []
    for source, text in sources.items():
        tokens = _L_TOKEN_RE.findall(str(text))
        if tokens:
            matching.append(
                {
                    "source": source,
                    "text": str(text),
                    "tokens": tokens,
                }
            )
    return {
        "selected_range": view.selected_column_logical_range,
        "matching_sources": matching,
        "unique_tokens": sorted(
            {
                token
                for item in matching
                for token in item["tokens"]
            },
            key=lambda token: tuple(
                int(value)
                for value in re.findall(r"\d+", token)
            ),
        ),
    }


def _geometry_evidence(app, view):
    return {
        "diff_nav": _bar_whitespace(
            view.diff_nav_bar,
            (("diff_nav_group", view.diff_nav_group),),
        ),
        "column_action": _bar_whitespace(
            view.column_action_bar,
            (
                ("info", view.info),
                ("rich_status", view.column_action_status_group),
                ("column_buttons", view.column_action_button_group),
            ),
        ),
        "top_buttons": _top_button_inventory(view),
        "root_top_four": _root_top_button_inventory(app),
    }


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    with _real_gunships_app() as (app, view, _analysis):
        evidence = {
            "sheet": view.sheet,
            "editor_font": view.editor_font,
            "logical_widths": dict(view.col_char_widths),
            "pixel_boundaries": _pixel_boundary_records(app, view),
        }
        evidence["drift_summary"] = _drift_summary(evidence["pixel_boundaries"])
        evidence["layout_by_geometry"] = {}
        for geometry in ("1450x860", "1024x760"):
            app.root.geometry(geometry)
            _pump(app.root, 0.2)
            evidence["layout_by_geometry"][geometry] = _geometry_evidence(
                app,
                view,
            )

        evidence["column_hints"] = {
            "initial_l14": _column_action_hints(view),
        }
        result_l14 = view._apply_selected_column_block("BASE", "A")
        assert result_l14.logical_start == result_l14.logical_end == 14
        _wait_until(
            app.root,
            lambda: (
                view._derive_lifecycle_state() == "READY"
                and view.selected_column_logical_range == (20, 20)
            ),
            "Gunships did not advance to L20",
            timeout=35.0,
        )
        evidence["column_hints"]["after_l14_l20_pending"] = _column_action_hints(
            view
        )
        result_l20 = view._apply_selected_column_block("BASE", "A")
        assert result_l20.logical_start == result_l20.logical_end == 20
        _wait_until(
            app.root,
            lambda: (
                view._derive_lifecycle_state() == "READY"
                and not view._active_column_comparison_cache().structural_diff_cols
                and not view._active_column_comparison_cache().unresolved_cols
            ),
            "Gunships did not finish L20",
            timeout=35.0,
        )
        evidence["column_hints"]["after_l20_complete"] = _column_action_hints(
            view
        )
        evidence["root_top_after_l14_l20"] = {}
        for geometry in ("1450x860", "1024x760"):
            app.root.geometry(geometry)
            _pump(app.root, 0.2)
            evidence["root_top_after_l14_l20"][geometry] = (
                _root_top_button_inventory(app)
            )
        print(json.dumps(evidence, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
