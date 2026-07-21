"""GUI self-test: hover-driven C area, explicit selection lock, and right-click clear."""

import os
from types import SimpleNamespace

from openpyxl import Workbook

import sow_merge_tool as mod
from _test_temp_utils import make_temp_dir


def _make_xlsx(path: str, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = value
    wb.save(path)
    wb.close()


def _ensure_view(app: mod.SowMergeApp):
    sheet = app.common_sheets[0]
    view = app.sheet_views.get(sheet)
    if view is None:
        app.nb.select(app._sheet_containers[sheet])
        app.root.update_idletasks()
        app.root.update()
        view = app.sheet_views[sheet]
    return view


def _motion_event_for_cell(text_widget, line_no: int, char_pos: int):
    box = text_widget.bbox(f"{line_no}.{max(0, int(char_pos))}")
    assert box is not None, f"bbox is None for line={line_no}, char={char_pos}"
    x, y, w, h = box
    px = int(x + max(1, w // 2))
    py = int(y + max(1, h // 2))
    return SimpleNamespace(
        x=px,
        y=py,
        x_root=int(text_widget.winfo_rootx() + px),
        y_root=int(text_widget.winfo_rooty() + py),
    )


def _char_pos_for_col(view, col_idx: int, line_text: str | None = None) -> int:
    spans = view._spans_for_line(line_text) if line_text is not None else view._spans_for_line()
    assert col_idx in spans, f"Column {col_idx} not found in spans: {spans}"
    s, e = spans[col_idx]
    return s + 1 if (e - s) > 1 else s


def _drive_main_hover(view, line_no: int, col_idx: int, side: str = "A"):
    widget = view.left if side == "A" else (view.base if side == "BASE" else view.right)
    char_pos = _char_pos_for_col(view, col_idx)
    ev = _motion_event_for_cell(widget, line_no=line_no, char_pos=char_pos)
    view._on_cell_hover_tooltip(widget, ev, side)
    view.app.root.update_idletasks()
    view.app.root.update()
    return ev


def _select_main_cell(view, line_no: int, col_idx: int, side: str = "A"):
    view._highlight_selected_line(line_no)
    view.selected_pair_idx = view._pair_idx_for_line(line_no)
    pair = view._pair_for_line(line_no)
    view.selected_excel_row_a = view._row_for_side(pair, "A")
    view.selected_excel_row_b = view._row_for_side(pair, "B")
    view.selected_excel_row = view.selected_excel_row_a or view.selected_excel_row_b
    view._set_main_selected_cell(line_no, col_idx)
    if view._is_three_way_enabled():
        c_line = 1 if side == "BASE" else (2 if side == "A" else 3)
    else:
        c_line = 1 if side in ("A", "BASE") else 2
    view._cursor_cmp_sel_col = int(col_idx)
    view._cursor_cmp_sel_line = int(c_line)
    view.hover_pair_idx = view.selected_pair_idx
    view.hover_col_idx = int(col_idx)
    view.hover_side = side
    view._update_cursor_lines()
    view.app.root.update_idletasks()
    view.app.root.update()
    return None


def _click_cursor_cmp(view, line_no: int, col_idx: int):
    line_text = view.cursor_cmp.get(f"{line_no}.0", f"{line_no}.end")
    char_pos = _char_pos_for_col(view, col_idx, line_text=line_text)
    ev = _motion_event_for_cell(view.cursor_cmp, line_no=line_no, char_pos=char_pos)
    view._on_cursor_cmp_click(ev)
    view.app.root.update_idletasks()
    view.app.root.update()
    return ev


def _cursor_lines(view):
    return [
        view.cursor_cmp.get("1.0", "1.end"),
        view.cursor_cmp.get("2.0", "2.end"),
    ]


def _panel_text(view) -> str:
    return str(view.hover_cmp_text.get("1.0", "end-1c"))


def _assert_main_line_has_no_diff_bg(view, line_no: int, col_idx: int):
    char_pos = _char_pos_for_col(view, col_idx)
    tags = set(view.left.tag_names(f"{line_no}.{char_pos}"))
    assert "diffrow" not in tags, f"unexpected diffrow at left {line_no}.{char_pos}: {tags}"
    assert "diffcell" not in tags, f"unexpected diffcell at left {line_no}.{char_pos}: {tags}"
    tags = set(view.right.tag_names(f"{line_no}.{char_pos}"))
    assert "diffrow" not in tags, f"unexpected diffrow at right {line_no}.{char_pos}: {tags}"
    assert "diffcell" not in tags, f"unexpected diffcell at right {line_no}.{char_pos}: {tags}"


def main():
    td_a = make_temp_dir(prefix="sow_hover_state_a_")
    td_b = make_temp_dir(prefix="sow_hover_state_b_")
    file_a = os.path.join(td_a, "same.xlsx")
    file_b = os.path.join(td_b, "same.xlsx")
    rows_a = [
        ["id", "value", "note"],
        ["row2_key", "A2", "same"],
        ["row3_key", "A3", "same"],
    ]
    rows_b = [
        ["id", "value", "note"],
        ["row2_key", "B2", "same"],
        ["row3_key", "B3", "same"],
    ]
    _make_xlsx(file_a, rows_a)
    _make_xlsx(file_b, rows_b)

    app = mod.SowMergeApp(file_a, file_b)
    try:
        view = _ensure_view(app)
        view.only_diff_var.set(0)
        view.refresh(row_only=None, rescan=True)
        app.root.update_idletasks()
        app.root.update()

        _assert_main_line_has_no_diff_bg(view, line_no=1, col_idx=1)
        assert not view.has_explicit_cell_selection()

        _drive_main_hover(view, line_no=2, col_idx=2, side="A")
        lines = _cursor_lines(view)
        assert "row2_key" in lines[0] and "A2" in lines[0], lines
        assert "row2_key" in lines[1] and "B2" in lines[1], lines
        panel = _panel_text(view)
        assert "base[2]: A2" in panel and "mine[2]: B2" in panel, panel
        assert view._cursor_cmp_tooltip_payload(_char_pos_for_col(view, 2), force_panel=True) is not None

        _select_main_cell(view, line_no=3, col_idx=2, side="A")
        assert view.has_explicit_cell_selection()
        assert view.selected_pair_idx == view._pair_idx_for_line(3)
        lines = _cursor_lines(view)
        assert "row3_key" in lines[0] and "A3" in lines[0], lines
        assert "row3_key" in lines[1] and "B3" in lines[1], lines

        _drive_main_hover(view, line_no=2, col_idx=2, side="A")
        lines = _cursor_lines(view)
        assert "row3_key" in lines[0] and "A3" in lines[0], lines
        assert "row3_key" in lines[1] and "B3" in lines[1], lines
        panel = _panel_text(view)
        assert "base[2]: A2" in panel and "mine[2]: B2" in panel, panel

        pinned_panel = panel
        view.hover_cmp_pin_var.set(1)
        view._on_hover_compare_pin_toggle()
        _drive_main_hover(view, line_no=3, col_idx=2, side="A")
        assert _panel_text(view) == pinned_panel, _panel_text(view)
        view.hover_cmp_pin_var.set(0)
        view._on_hover_compare_pin_toggle()

        ev = _drive_main_hover(view, line_no=2, col_idx=2, side="A")
        view._on_main_pane_right_click(view.left, ev, "A")
        assert not view.has_explicit_cell_selection()
        assert view.selected_pair_idx is None
        lines = _cursor_lines(view)
        assert "row2_key" in lines[0] and "A2" in lines[0], lines
        assert "row2_key" in lines[1] and "B2" in lines[1], lines

        _drive_main_hover(view, line_no=3, col_idx=2, side="A")
        _click_cursor_cmp(view, line_no=1, col_idx=2)
        assert view.has_explicit_cell_selection()
        assert view.selected_pair_idx == view._pair_idx_for_line(3)

        _drive_main_hover(view, line_no=2, col_idx=2, side="A")
        lines = _cursor_lines(view)
        assert "row3_key" in lines[0] and "A3" in lines[0], lines
        assert "row3_key" in lines[1] and "B3" in lines[1], lines

        view._on_cursor_cmp_right_click(SimpleNamespace(x=1, y=1))
        assert not view.has_explicit_cell_selection()
        assert view.selected_pair_idx is None

        _drive_main_hover(view, line_no=2, col_idx=2, side="A")
        lines = _cursor_lines(view)
        assert "row2_key" in lines[0] and "A2" in lines[0], lines
        assert "row2_key" in lines[1] and "B2" in lines[1], lines
    finally:
        try:
            app._shutdown_root()
        except Exception:
            pass

    print("GUI_SELF_TEST_HOVER_SELECTION_STATE_OK")


if __name__ == "__main__":
    main()
