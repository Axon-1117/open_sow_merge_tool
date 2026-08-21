"""B3 GUI regression: current exact C-area hover is snapshot-only and deduplicated."""

import argparse
from contextlib import contextmanager
import hashlib
import json
import os
import tempfile
import time
from pathlib import Path
from types import SimpleNamespace

from openpyxl import Workbook

import sow_merge_tool as sm


_CASES = ("two-way", "three-way")


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _path_snapshot(path: Path) -> tuple[bool, bytes | None]:
    return (True, path.read_bytes()) if path.exists() else (False, None)


def _canonical(value):
    if value is None or isinstance(value, (bool, int, float, str, bytes)):
        return value
    if isinstance(value, dict):
        return tuple(
            sorted(
                ((_canonical(key), _canonical(item)) for key, item in value.items()),
                key=repr,
            )
        )
    if isinstance(value, (tuple, list, set, frozenset)):
        values = tuple(_canonical(item) for item in value)
        return tuple(sorted(values, key=repr)) if isinstance(value, (set, frozenset)) else values
    return repr(value)


def _fingerprint(value) -> str:
    return hashlib.sha256(repr(_canonical(value)).encode("utf-8")).hexdigest()


def _hard_mutation_snapshot(app, view, sheet: str, input_paths) -> dict:
    overlays = getattr(app, "sheet_operation_overlays", {}) or {}
    overlay = overlays.get(sheet)
    edit_handles = tuple(
        (
            name,
            id(getattr(app, name, None)),
            type(getattr(app, name, None)).__name__,
            getattr(getattr(app, name, None), "read_only", None),
        )
        for name in ("_wb_a_edit", "_wb_b_edit", "_wb_base_edit")
    )
    return {
        "input_hashes": tuple(sorted((name, _sha256(path)) for name, path in input_paths.items())),
        "manual": _fingerprint(
            {
                name: getattr(app, name, None)
                for name in (
                    "manual_a_cell_ops", "manual_b_cell_ops",
                    "manual_a_formula_cache_ops", "manual_b_formula_cache_ops",
                    "manual_a_row_ops", "manual_b_row_ops",
                    "manual_a_column_ops", "manual_b_column_ops",
                    "manual_sheet_ops", "auto_sheet_ops",
                )
            }
        ),
        "undo_redo": _fingerprint((getattr(app, "undo_stack", ()), getattr(app, "redo_stack", ()))),
        "modified": _canonical(
            (
                getattr(app, "modified_a", None), getattr(app, "modified_b", None),
                getattr(app, "modified_sheets_a", None), getattr(app, "modified_sheets_b", None),
                getattr(app, "user_touched_conflicts", None), getattr(view, "touched_rows", None),
            )
        ),
        "overlay": _fingerprint(
            {
                name: (
                    getattr(item, "topology_generation", None),
                    getattr(item, "mutation_generation", None),
                    getattr(item, "cells", None),
                )
                for name, item in overlays.items()
            }
        ),
        "prepared": _fingerprint(
            {
                "raw_a": getattr(view, "pair_raw_parts_a", None),
                "raw_b": getattr(view, "pair_raw_parts_b", None),
                "raw_base": getattr(view, "pair_raw_parts_base", None),
                "row_pairs": getattr(view, "row_pairs", None),
                "maps": (
                    getattr(view, "row_a_to_pair_idx", None),
                    getattr(view, "row_b_to_pair_idx", None),
                    getattr(view, "mine_to_base_row", None),
                    getattr(view, "theirs_to_base_row", None),
                    getattr(view, "pair_base_row_override", None),
                ),
                "diffs": (
                    getattr(view, "pair_diff_cols", None),
                    getattr(view, "pair_base_diff_cols", None),
                ),
            }
        ),
        "generations": (
            getattr(app, "_sheet_compute_generation", {}).get(sheet),
            getattr(view, "_row_model_version", None),
            getattr(view, "_column_model_version", None),
            getattr(view, "_row_model_exact", None),
            getattr(view, "_column_projection_generation", None),
            getattr(view, "_virtual_column_window_generation", None),
            getattr(overlay, "topology_generation", None),
            getattr(overlay, "mutation_generation", None),
        ),
        "edit_handles": edit_handles,
    }


def _render_cache_version(view) -> int:
    return int(getattr(view, "_data_version", -1))


def _snapshot_value_summary(value) -> dict:
    canonical = _canonical(value)
    preview = repr(canonical)
    try:
        length = len(value)
    except Exception:
        length = None
    return {
        "type": type(value).__name__,
        "len": length,
        "hash": _fingerprint(value),
        "preview": preview[:240] + ("..." if len(preview) > 240 else ""),
    }


def _recursive_field_diffs(before, after, path: str = "hard") -> list[dict]:
    if before == after:
        return []
    if isinstance(before, dict) and isinstance(after, dict):
        diffs = []
        keys = sorted(set(before) | set(after), key=repr)
        for key in keys:
            child_path = f"{path}[{key!r}]"
            if key not in before or key not in after:
                diffs.append(
                    {
                        "path": child_path,
                        "before": _snapshot_value_summary(before.get(key)),
                        "after": _snapshot_value_summary(after.get(key)),
                    }
                )
            else:
                diffs.extend(_recursive_field_diffs(before[key], after[key], child_path))
        return diffs
    if isinstance(before, (tuple, list)) and isinstance(after, (tuple, list)):
        diffs = []
        if len(before) != len(after):
            diffs.append(
                {
                    "path": f"{path}.len",
                    "before": _snapshot_value_summary(before),
                    "after": _snapshot_value_summary(after),
                }
            )
        for index, (left, right) in enumerate(zip(before, after)):
            diffs.extend(_recursive_field_diffs(left, right, f"{path}[{index}]"))
        return diffs
    return [
        {
            "path": path,
            "before": _snapshot_value_summary(before),
            "after": _snapshot_value_summary(after),
        }
    ]


def _assert_hard_mutation_unchanged(before, after, *, action: str) -> None:
    diffs = _recursive_field_diffs(before, after)
    assert not diffs, f"{action}: forbidden hard mutation drift: {json.dumps(diffs, ensure_ascii=False, sort_keys=True)}"


def _make_book(path: str, *, side: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["id@id", "value"])
    ws.append(["int32", "string"])
    long_value = f"{side}_FULL_" + (side.lower()[0] * 90)
    ws.append([1, long_value])
    ws.append([2, f"{side}_NEXT_" + (side.lower()[-1] * 20)])
    wb.save(path)
    wb.close()


def _pump(app, deadline: float) -> None:
    while time.monotonic() < deadline:
        app.root.update_idletasks()
        app.root.update()
        time.sleep(0.005)


def _wait(app, predicate, *, deadline: float, stage: str) -> None:
    while time.monotonic() < deadline:
        _pump(app, min(deadline, time.monotonic() + 0.05))
        if predicate():
            return
        view = app.sheet_views.get("S")
        if view is not None and view._derive_lifecycle_state() in {
            "FAILED",
            "UNRESOLVED",
            "CANCELED",
            "CLOSING",
        }:
            raise AssertionError(
                f"{stage} failed: {view._derive_lifecycle_state()} "
                f"{getattr(view, '_lifecycle_error', None)!r}"
            )
    raise AssertionError(
        f"timeout {stage}: {app._sheet_exact_entry('S')!r}"
    )


def _shutdown(app) -> None:
    if app is None:
        return
    for view in tuple(getattr(app, "sheet_views", {}).values()):
        if view is None:
            continue
        for attr in ("_settings_save_id", "_hover_debounce_id"):
            after_id = getattr(view, attr, None)
            if after_id:
                try:
                    view.frame.after_cancel(after_id)
                finally:
                    setattr(view, attr, None)
    app._shutdown_root()


@contextmanager
def _forbid_view_only_access(app, view):
    hits = []
    originals = []

    def _forbidden(label):
        def _raise(*_args, **_kwargs):
            hits.append(label)
            raise AssertionError(f"hover route accessed {label}")

        return _raise

    targets = (
        (app, "ws_a_val"),
        (app, "ws_b_val"),
        (app, "ws_base_val"),
        (app, "ws_a_edit"),
        (app, "ws_b_edit"),
        (app, "ws_base_edit"),
        (app, "_request_edit_preload"),
        (app, "_ensure_edit_loaded"),
        (app, "_load_edit_workbooks_owned"),
        (app, "_atomic_save"),
        (app, "_atomic_save_with_retry"),
        (app, "_atomic_replace_file_with_retry"),
        (app, "_try_alt_save"),
        (sm, "_atomic_save_wb"),
        (app, "build_manual_b_output_file"),
        (app, "save_a_inplace"),
        (app, "save_b_inplace"),
        (app, "save_merged_and_exit"),
        (view, "_run_copy_action_by_mode"),
        (view, "_apply_global_sheet_overwrite"),
        (view, "_apply_selected_column_block"),
    )
    try:
        for owner, name in targets:
            if hasattr(owner, name):
                originals.append((owner, name, getattr(owner, name)))
                setattr(owner, name, _forbidden(f"{type(owner).__name__}.{name}"))
        yield hits
    finally:
        for owner, name, original in reversed(originals):
            setattr(owner, name, original)


def _event_for_logical_cell(widget, view, line: int, logical_col: int):
    spans = view._spans_for_line()
    assert logical_col in spans, (logical_col, spans)
    start, end = spans[logical_col]
    char = start + 1 if end - start > 1 else start
    index = f"{line}.{max(0, int(char))}"
    widget.see(index)
    widget.update_idletasks()
    widget.update()
    bbox = widget.bbox(index)
    assert bbox is not None, (index, widget.index("@0,0"))
    x, y, width, height = bbox
    event = SimpleNamespace(
        x=int(x + max(1, width // 2)),
        y=int(y + max(1, height // 2)),
        x_root=int(widget.winfo_rootx() + x + max(1, width // 2)),
        y_root=int(widget.winfo_rooty() + y + max(1, height // 2)),
    )
    target_line, target_col = map(int, str(widget.index(f"@{event.x},{event.y}")).split("."))
    assert target_line == int(line), (target_line, line)
    assert start <= target_col < max(end, start + 1), (target_col, start, end)
    return event


def _panel_text(view) -> str:
    return str(view.hover_cmp_text.get("1.0", "end-1c"))


@contextmanager
def _observe_cached_only_diff_publisher(app, view):
    """Record real cache-only publications while delegating to production."""
    calls = []
    original = view._publish_prepared_cache_surface

    def _state():
        entry = dict(app._sheet_exact_entry(view.sheet) or {})
        return {
            "sheet": str(view.sheet),
            "selected_sheet": str(getattr(app, "selected_sheet", "") or ""),
            "compute_generation": int(
                (getattr(app, "_sheet_compute_generation", {}) or {}).get(view.sheet, -1)
            ),
            "exact_generation": int(entry.get("generation", -1)),
            "exact_state": str(entry.get("state") or ""),
            "mode_switch_seq": int(getattr(view, "_mode_switch_seq", -1)),
            "mode_switch_requested": getattr(view, "_mode_switch_requested_value", None),
            "mode_switch_pending": bool(getattr(view, "_mode_switch_pending", False)),
            "only_diff_build_seq": int(getattr(view, "_only_diff_async_build_seq", -1)),
            "only_diff_building": bool(getattr(view, "_only_diff_async_building", False)),
        }

    def _wrapped(*args, **kwargs):
        rows = tuple(int(pair_idx) for pair_idx in (kwargs.get("prepared_rows") or ()))
        record = {
            "before": _state(),
            "prepared_rows": rows,
            "prepared_rows_digest": _fingerprint(rows),
        }
        calls.append(record)
        try:
            result = original(*args, **kwargs)
        except Exception as exc:
            record["exception"] = f"{type(exc).__name__}: {exc}"
            record["after"] = _state()
            raise
        record["result"] = result
        record["after"] = _state()
        return result

    view._publish_prepared_cache_surface = _wrapped
    try:
        yield calls
    finally:
        view._publish_prepared_cache_surface = original


def _assert_single_cached_publication(
    calls,
    *,
    app,
    view,
    action: str,
    expected_rows,
    requested_value: int,
    expected_mode_switch_seq: int,
    expected_compute_generation: int,
    expected_exact_generation: int,
    expected_async_build_seq: int | None = None,
) -> None:
    assert len(calls) == 1, (action, calls)
    call = calls[0]
    before = dict(call.get("before") or {})
    rows = tuple(int(pair_idx) for pair_idx in (expected_rows or ()))
    assert call.get("result") is True, (action, call)
    assert tuple(call.get("prepared_rows") or ()) == rows, (action, call, rows)
    assert call.get("prepared_rows_digest") == _fingerprint(rows), (action, call, rows)
    assert before["sheet"] == str(view.sheet), (action, before)
    assert before["selected_sheet"] == str(view.sheet), (action, before)
    assert before["compute_generation"] == int(expected_compute_generation), (action, before)
    assert before["exact_generation"] == int(expected_exact_generation), (action, before)
    assert before["mode_switch_seq"] == int(expected_mode_switch_seq), (action, before)
    assert before["mode_switch_requested"] == int(requested_value), (action, before)
    assert before["mode_switch_pending"] is True, (action, before)
    if expected_async_build_seq is not None:
        assert int(before["only_diff_build_seq"]) == int(expected_async_build_seq), (
            action, before, expected_async_build_seq
        )
        assert not bool(before["only_diff_building"]), (action, before)


def _run_case(case_name: str) -> None:
    original_settings_path = sm._SETTINGS_PATH
    user_settings = Path(original_settings_path)
    user_settings_before = _path_snapshot(user_settings)
    environment_before = {
        key: value for key, value in os.environ.items() if key.startswith("SOW_")
    }
    app = None
    root_path = None
    try:
        with tempfile.TemporaryDirectory(prefix=f"sow_c_hover_{case_name}_") as root:
            root_path = Path(root)
            mine = str(root_path / "mine.xlsx")
            theirs = str(root_path / "theirs.xlsx")
            base = str(root_path / "base.xlsx")
            _make_book(mine, side="MINE")
            _make_book(theirs, side="THEIRS")
            input_paths = {"mine": mine, "theirs": theirs}
            if case_name == "three-way":
                _make_book(base, side="BASE")
                input_paths["base"] = base
            input_hashes = {name: _sha256(path) for name, path in input_paths.items()}
            settings_path = root_path / "settings.json"
            settings_path.write_text(json.dumps({"only_diff": 0}), encoding="utf-8")
            sm._SETTINGS_PATH = str(settings_path)
            deadline = time.monotonic() + 90.0
            print(f"C_HOVER_STAGE open-current-exact mode={case_name}", flush=True)
            if case_name == "three-way":
                app = sm.SowMergeApp(
                    mine,
                    theirs,
                    merge_mode=True,
                    base_path=base,
                    initial_sheet="S",
                )
            else:
                app = sm.SowMergeApp(mine, theirs, initial_sheet="S")
            view = app.sheet_views["S"]
            _wait(
                app,
                lambda: (
                    app._is_sheet_exact_current("S")
                    and bool(app._sheet_exact_entry("S").get("full_detail_terminal"))
                    and bool(view._prepared_complete)
                    and bool(view._data_ready)
                    and not bool(view._pending_exact_render)
                ),
                deadline=deadline,
                stage="selected exact prepared surface",
            )
            assert not bool(view.only_diff_var.get())
            assert view._is_exact_immutable_view_ready()
            pair_idx = view.row_a_to_pair_idx.get(3)
            next_pair_idx = view.row_a_to_pair_idx.get(4)
            assert pair_idx is not None and next_pair_idx is not None
            line = view.row_to_line.get(pair_idx)
            next_line = view.row_to_line.get(next_pair_idx)
            assert line is not None and next_line is not None
            widget = view.base if case_name == "three-way" else view.left
            side = "BASE" if case_name == "three-way" else "A"
            event = _event_for_logical_cell(widget, view, line, 2)
            next_event = _event_for_logical_cell(widget, view, next_line, 2)

            payload_calls = 0
            c_area_renders = 0
            original_payload = view._cmp_tooltip_payload_by_pair_col
            original_update_cursor = view._update_cursor_lines

            def _count_payload(*args, **kwargs):
                nonlocal payload_calls
                payload_calls += 1
                return original_payload(*args, **kwargs)

            def _count_c_area_render(*args, **kwargs):
                nonlocal c_area_renders
                c_area_renders += 1
                return original_update_cursor(*args, **kwargs)

            view._cmp_tooltip_payload_by_pair_col = _count_payload
            view._update_cursor_lines = _count_c_area_render
            hard_before_hover = _hard_mutation_snapshot(app, view, "S", input_paths)
            render_before_hover = _render_cache_version(view)
            try:
                print(f"C_HOVER_STAGE bound-main-hover mode={case_name}", flush=True)
                with _forbid_view_only_access(app, view) as hits, _observe_cached_only_diff_publisher(app, view) as hover_publications:
                    # The widget binding is the production lambda route; do not
                    # call the handler directly in this regression.
                    widget.event_generate("<Motion>", x=event.x, y=event.y)
                    _wait(
                        app,
                        lambda: bool(_panel_text(view)) and view.hover_pair_idx == pair_idx,
                        deadline=deadline,
                        stage="first bound hover panel",
                    )
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_hover,
                        _hard_mutation_snapshot(app, view, "S", input_paths),
                        action="bound hover event",
                    )
                    assert _render_cache_version(view) == render_before_hover
                    first_panel = _panel_text(view)
                    first_c_key = getattr(view, "_c_area_last_render_key", None)
                    widget.event_generate("<Motion>", x=event.x, y=event.y)
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_hover,
                        _hard_mutation_snapshot(app, view, "S", input_paths),
                        action="bound hover event",
                    )
                    assert _render_cache_version(view) == render_before_hover
                    assert payload_calls == 1 and c_area_renders == 1
                    assert getattr(view, "_c_area_last_render_key", None) == first_c_key
                    widget.event_generate("<Leave>")
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_hover,
                        _hard_mutation_snapshot(app, view, "S", input_paths),
                        action="bound hover event",
                    )
                    assert _render_cache_version(view) == render_before_hover
                    assert _panel_text(view) == first_panel
                    widget.event_generate("<Motion>", x=next_event.x, y=next_event.y)
                    _wait(
                        app,
                        lambda: (
                            payload_calls == 2
                            and c_area_renders == 2
                            and view.hover_pair_idx == next_pair_idx
                        ),
                        deadline=deadline,
                        stage="new-row bound hover panel",
                    )
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_hover,
                        _hard_mutation_snapshot(app, view, "S", input_paths),
                        action="bound hover event",
                    )
                    assert _render_cache_version(view) == render_before_hover
                assert not hits, hits
                assert hover_publications == [], hover_publications
                _assert_hard_mutation_unchanged(
                    hard_before_hover,
                    _hard_mutation_snapshot(app, view, "S", input_paths),
                    action="bound hover sequence",
                )
                assert _render_cache_version(view) == render_before_hover
            finally:
                view._cmp_tooltip_payload_by_pair_col = original_payload
                view._update_cursor_lines = original_update_cursor

            assert payload_calls == 2, payload_calls
            assert c_area_renders == 2, c_area_renders
            assert "MINE_FULL_" in first_panel and "THEIRS_FULL_" in first_panel, first_panel
            if case_name == "three-way":
                assert "BASE_FULL_" in first_panel, first_panel
                assert view._is_three_way_enabled()
                assert bool(view._base_diff_full_exact)
            else:
                assert "BASE_FULL_" not in first_panel, first_panel
            assert view.hover_pair_idx == next_pair_idx and int(view.hover_col_idx) == 2
            assert {name: _sha256(path) for name, path in input_paths.items()} == input_hashes
            assert not app._edit_workbooks_ready()
            environment_after = {
                key: value for key, value in os.environ.items() if key.startswith("SOW_")
            }
            assert environment_after == environment_before
            assert time.monotonic() <= deadline
            print(
                "GUI_SELF_TEST_C_HOVER_TOOLTIP_OK "
                + json.dumps(
                    {
                        "case": case_name,
                        "deadline_seconds": 90,
                        "payload_calls": payload_calls,
                        "c_area_render_calls": c_area_renders,
                        "input_hashes": input_hashes,
                    },
                    sort_keys=True,
                ),
                flush=True,
            )
    finally:
        _shutdown(app)
        sm._SETTINGS_PATH = original_settings_path
        assert _path_snapshot(user_settings) == user_settings_before
        if root_path is not None:
            assert not root_path.exists(), root_path


def main(argv=None) -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--list-cases", action="store_true")
    parser.add_argument("--case", choices=_CASES)
    args = parser.parse_args(argv)
    if args.list_cases:
        if args.case:
            parser.error("--list-cases cannot be combined with --case")
        for case_name in _CASES:
            print(case_name, flush=True)
        return
    selected = (args.case,) if args.case else _CASES
    for case_name in selected:
        _run_case(case_name)
    print(f"GUI_SELF_TEST_C_HOVER_TOOLTIP_SUITE_OK ({len(selected)} cases)", flush=True)


if __name__ == "__main__":
    main()
