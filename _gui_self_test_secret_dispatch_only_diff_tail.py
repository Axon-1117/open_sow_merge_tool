"""Real-file regression for the SecretDispatch 2-way only-diff tail."""

from __future__ import annotations

import argparse
import json
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


SHEET = "Dispatch_tasks@design"


def write_fixture(path: Path, *, changed: bool) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET
    sheet.append(["id@id"] + [f"field_{column}" for column in range(2, 35)])
    sheet.append(["int32"] + ["string"] * 33)
    # Physical rows 49..97 differ, reproducing the 49-row only-diff surface.
    for physical_row in range(3, 98):
        values = [physical_row] + [
            f"value_{physical_row}_{column}" for column in range(2, 35)
        ]
        if changed and physical_row >= 49:
            values[-1] += "_changed"
        sheet.append(values)
    workbook.save(path)
    workbook.close()


def pump(app) -> None:
    app.root.update_idletasks()
    app.root.update()


def wait(app, predicate, label: str, timeout: float = 45.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        pump(app)
        if predicate():
            return
        time.sleep(0.01)
    raise AssertionError(f"timeout waiting for {label}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base")
    parser.add_argument("--mine")
    args = parser.parse_args()
    fixture_dir = None
    if bool(args.base) != bool(args.mine):
        parser.error("--base and --mine must be supplied together")
    if args.base:
        base_path = str(Path(args.base).resolve())
        mine_path = str(Path(args.mine).resolve())
    else:
        fixture_dir = tempfile.TemporaryDirectory(
            prefix="sow-secret-dispatch-tail-"
        )
        base_fixture = Path(fixture_dir.name) / "base.xlsx"
        mine_fixture = Path(fixture_dir.name) / "mine.xlsx"
        write_fixture(base_fixture, changed=False)
        write_fixture(mine_fixture, changed=True)
        base_path = str(base_fixture)
        mine_path = str(mine_fixture)
    app = None
    try:
        app = sm.SowMergeApp(
            base_path,
            mine_path,
            merge_mode=False,
            initial_sheet=SHEET,
        )

        def ready():
            view = app.sheet_views.get(SHEET)
            return bool(
                view is not None
                and view._data_ready
                and view._prepared_complete
                and app._is_sheet_exact_current(SHEET)
            )

        wait(app, ready, "exact sheet")
        view = app.sheet_views[SHEET]
        if not bool(view.only_diff_var.get()):
            view.only_diff_var.set(1)
            view._toggle_only_diff()
        wait(
            app,
            lambda: (
                bool(view.only_diff_var.get())
                and not bool(view._mode_switch_pending)
                and not bool(view._only_diff_async_building)
            ),
            "only diff",
        )
        pump(app)
        pre_last_line = len(view.display_rows)
        pre_tail = {
            "capacity": view._virtual_viewport_row_capacity(),
            "window_start": view._virtual_window_start,
            "display_rows": pre_last_line,
            "last_dline": view.right.dlineinfo(f"{pre_last_line}.0"),
            "after_last_dline": view.right.dlineinfo(f"{pre_last_line + 1}.0"),
            "text_end": view.right.index("end"),
        }
        view._yview_both("moveto", "1.0")
        wait(
            app,
            lambda: (
                getattr(view, "_virtual_pending_start", None) is None
                and getattr(view, "_virtual_2d_after_id", None) is None
            ),
            "tail publication",
        )
        pump(app)
        last_line = len(view.display_rows)
        last_pair_idx = view.display_rows[-1]
        last_pair = view.row_pairs[last_pair_idx]
        height = int(view.right.winfo_height())
        evidence = {
            "pre_tail": pre_tail,
            "full_rows": len(view._full_display_rows),
            "capacity": view._virtual_viewport_row_capacity(),
            "window_start": view._virtual_window_start,
            "display_rows": len(view.display_rows),
            "last_pair": last_pair,
            "last_right_header": view._row_label_for_pair_idx(last_pair_idx, "B"),
            "widget_height": height,
            "first_dline": view.right.dlineinfo("1.0"),
            "penultimate_dline": view.right.dlineinfo(f"{max(1, last_line - 1)}.0"),
            "last_dline": view.right.dlineinfo(f"{last_line}.0"),
            "after_last_dline": view.right.dlineinfo(f"{last_line + 1}.0"),
            "bottom_index": view.right.index(f"@0,{max(0, height - 1)}"),
            "text_yview": view.right.yview(),
            "logical_scroll": view._virtual_scroll_fractions(),
            "text_end": view.right.index("end"),
        }
        print("SECRET_DISPATCH_ONLY_DIFF_TAIL_EVIDENCE", json.dumps(evidence, ensure_ascii=False))
        assert str(evidence["last_right_header"]).strip() == "97", evidence
        assert evidence["last_dline"] is not None, evidence
        last_bottom = int(evidence["last_dline"][1]) + int(evidence["last_dline"][3])
        assert last_bottom <= height, evidence
        assert evidence["after_last_dline"] is not None, evidence
        assert int(evidence["after_last_dline"][1]) >= last_bottom, evidence
        blank_bottom = (
            int(evidence["after_last_dline"][1])
            + int(evidence["after_last_dline"][3])
        )
        assert blank_bottom <= height, evidence
        assert int(evidence["after_last_dline"][3]) >= min(
            int(evidence["penultimate_dline"][3]),
            int(evidence["last_dline"][3]),
        ), evidence

        # A recycled non-tail document must not destabilize the measured
        # capacity. Return to the top and then to 100% once more.
        view._yview_both("moveto", "0.0")
        wait(
            app,
            lambda: (
                view._virtual_window_start == 0
                and getattr(view, "_virtual_pending_start", None) is None
                and getattr(view, "_virtual_2d_after_id", None) is None
            ),
            "top publication",
        )
        view._yview_both("moveto", "1.0")
        wait(
            app,
            lambda: (
                view._virtual_scroll_fractions()[1] == 1.0
                and getattr(view, "_virtual_pending_start", None) is None
                and getattr(view, "_virtual_2d_after_id", None) is None
            ),
            "repeated tail publication",
        )
        pump(app)
        repeated_last_line = len(view.display_rows)
        repeated_pair_idx = view.display_rows[-1]
        repeated_last = view.right.dlineinfo(f"{repeated_last_line}.0")
        repeated_blank = view.right.dlineinfo(f"{repeated_last_line + 1}.0")
        assert view._row_label_for_pair_idx(repeated_pair_idx, "B").strip() == "97"
        assert repeated_last is not None and repeated_blank is not None
        assert int(repeated_blank[3]) >= int(repeated_last[3])
        assert int(repeated_blank[1]) + int(repeated_blank[3]) <= int(
            view.right.winfo_height()
        )
        print("GUI_SELF_TEST_SECRET_DISPATCH_ONLY_DIFF_TAIL_OK")
    finally:
        if app is not None:
            app._shutdown_root()
        if fixture_dir is not None:
            fixture_dir.cleanup()


if __name__ == "__main__":
    main()
