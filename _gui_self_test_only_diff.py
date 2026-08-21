"""B3 GUI regression: public only-diff selection preservation stays snapshot-only."""

import argparse
from contextlib import contextmanager
import hashlib
import json
import os
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


_CASES = ("only-diff-selection-restore",)


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


def _make_book(path: str, *, changed: bool) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["id@id", "value"])
    ws.append(["int32", "string"])
    for row_id in range(1, 2201):
        value = f"value-{row_id}"
        if changed and row_id == 8:
            value += "-changed"
        ws.append([row_id, value])
    wb.save(path)
    wb.close()


def _pump(root, deadline: float) -> None:
    while time.monotonic() < deadline:
        root.update_idletasks()
        root.update()
        time.sleep(0.005)


def _wait(app, predicate, *, deadline: float, stage: str) -> None:
    while time.monotonic() < deadline:
        _pump(app.root, min(deadline, time.monotonic() + 0.05))
        if predicate():
            return
        view = app.sheet_views.get("S")
        if view is not None and getattr(view, "_derive_lifecycle_state", lambda: "")() in {
            "FAILED",
            "UNRESOLVED",
            "CANCELED",
            "CLOSING",
        }:
            raise AssertionError(
                f"{stage} terminal failure: "
                f"{view._derive_lifecycle_state()} {getattr(view, '_lifecycle_error', None)!r}"
            )
    view = app.sheet_views.get("S")
    raise AssertionError(
        f"timeout {stage}: "
        + repr(
            {
                "entry": app._sheet_exact_entry("S"),
                "selected": app.selected_sheet,
                "only_diff": None if view is None else int(view.only_diff_var.get()),
                "building": None if view is None else bool(view._only_diff_async_building),
                "ready": None if view is None else bool(view._data_ready),
                "cache": None if view is None else bool(
                    view._has_valid_only_diff_snapshot_cache()
                ),
            }
        )
    )


def _shutdown(app) -> None:
    if app is None:
        return
    for view in tuple(getattr(app, "sheet_views", {}).values()):
        if view is None:
            continue
        after_id = getattr(view, "_settings_save_id", None)
        if after_id:
            try:
                view.frame.after_cancel(after_id)
            finally:
                view._settings_save_id = None
    app._shutdown_root()


@contextmanager
def _forbid_view_only_access(app, view):
    hits = []
    originals = []

    def _forbidden(label):
        def _raise(*_args, **_kwargs):
            hits.append(label)
            raise AssertionError(f"only-diff view route accessed {label}")

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


def _assert_selected_diff(view, pair_idx: int, logical_col: int) -> None:
    assert view.has_explicit_cell_selection()
    assert int(view.selected_pair_idx) == int(pair_idx)
    assert int(view._main_sel_col) == int(logical_col)
    assert int(view._cursor_cmp_sel_col) == int(logical_col)
    assert int(view._main_sel_line) == int(view.row_to_line[pair_idx])


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


def _run_selection_restore() -> None:
    original_settings_path = sm._SETTINGS_PATH
    user_settings = Path(original_settings_path)
    user_settings_before = _path_snapshot(user_settings)
    app = None
    root_path = None
    try:
        with tempfile.TemporaryDirectory(prefix="sow_only_diff_selection_") as root:
            root_path = Path(root)
            mine = str(root_path / "mine.xlsx")
            theirs = str(root_path / "theirs.xlsx")
            _make_book(mine, changed=False)
            _make_book(theirs, changed=True)
            input_paths = {"mine": mine, "theirs": theirs}
            input_hashes = {name: _sha256(path) for name, path in input_paths.items()}
            settings_path = root_path / "settings.json"
            settings_path.write_text(json.dumps({"only_diff": 0}), encoding="utf-8")
            sm._SETTINGS_PATH = str(settings_path)
            deadline = time.monotonic() + 90.0
            print("ONLY_DIFF_STAGE open-current-exact", flush=True)
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
                    and str(view.only_diff_cb.cget("state")) == "normal"
                    and not bool(view.only_diff_var.get())
                ),
                deadline=deadline,
                stage="selected exact full surface",
            )
            assert view._is_exact_immutable_view_ready()
            assert not app._edit_workbooks_ready()

            pair_idx = view.row_a_to_pair_idx.get(10)
            assert pair_idx is not None and pair_idx in view.pair_diff_cols
            line = view.row_to_line.get(pair_idx)
            assert line is not None
            view._set_main_selected_cell(line, 2)
            view.selected_pair_idx = pair_idx
            view.selected_excel_row = 10
            view.selected_excel_row_a = 10
            view.selected_excel_row_b = 10
            view._cursor_cmp_sel_col = 2
            view._cursor_cmp_sel_line = 1
            view._update_cursor_lines()
            _pump(app.root, min(deadline, time.monotonic() + 0.05))
            _assert_selected_diff(view, pair_idx, 2)
            full_logical_rows_before = tuple(view._full_display_rows)
            assert full_logical_rows_before == tuple(range(len(view.row_pairs)))
            assert len(full_logical_rows_before) > sm._VIRTUAL_VIEWPORT_MAX_ROWS
            assert view._virtual_mode_active()

            print("ONLY_DIFF_STAGE public-enable-cache-only", flush=True)
            assert view._has_valid_only_diff_snapshot_cache(), (
                "selection selector requires the already-exact cache-only route"
            )
            hard_before_enable = _hard_mutation_snapshot(app, view, "S", input_paths)
            render_before_enable = _render_cache_version(view)
            enable_rows = tuple(view._only_diff_rows_with_touched(view._only_diff_rows_cache))
            enable_mode_switch_before = int(view._mode_switch_seq)
            enable_compute_generation = int(app._sheet_compute_generation["S"])
            enable_exact_generation = int(app._sheet_exact_entry("S")["generation"])
            with _forbid_view_only_access(app, view) as hits, _observe_cached_only_diff_publisher(app, view) as enable_publications:
                view.only_diff_cb.invoke()
                _wait(
                    app,
                    lambda: (
                        bool(view.only_diff_var.get())
                        and not bool(view._only_diff_async_building)
                        and not bool(view._mode_switch_pending)
                        and app._is_sheet_exact_current("S")
                        and bool(app._sheet_exact_entry("S").get("full_detail_terminal"))
                        and bool(view._data_ready)
                        and bool(view._prepared_complete)
                        and bool(view._only_diff_rows_exact)
                        and bool(view._has_valid_only_diff_snapshot_cache())
                        and str(view.only_diff_cb.cget("state")) == "normal"
                    ),
                    deadline=deadline,
                    stage="only-diff terminal cache publication",
                )
            _pump(app.root, min(deadline, time.monotonic() + 0.05))
            assert not hits, hits
            _assert_hard_mutation_unchanged(
                hard_before_enable,
                _hard_mutation_snapshot(app, view, "S", input_paths),
                action="only-diff cache-only enable",
            )
            assert _render_cache_version(view) > render_before_enable, (
                render_before_enable,
                _render_cache_version(view),
            )
            _assert_single_cached_publication(
                enable_publications,
                app=app,
                view=view,
                action="only-diff cache-only enable",
                expected_rows=enable_rows,
                requested_value=1,
                expected_mode_switch_seq=enable_mode_switch_before + 1,
                expected_compute_generation=enable_compute_generation,
                expected_exact_generation=enable_exact_generation,
            )
            assert all(view._pair_has_visual_diff(pair) for pair in view.display_rows)
            _assert_selected_diff(view, pair_idx, 2)

            print("ONLY_DIFF_STAGE public-disable-cache-only", flush=True)
            hard_before_disable = _hard_mutation_snapshot(app, view, "S", input_paths)
            render_before_disable = _render_cache_version(view)
            disable_rows = tuple(range(len(view.row_pairs)))
            disable_mode_switch_before = int(view._mode_switch_seq)
            disable_compute_generation = int(app._sheet_compute_generation["S"])
            disable_exact_generation = int(app._sheet_exact_entry("S")["generation"])
            with _forbid_view_only_access(app, view) as hits, _observe_cached_only_diff_publisher(app, view) as disable_publications:
                view.only_diff_cb.invoke()
                _wait(
                    app,
                    lambda: (
                        not bool(view.only_diff_var.get())
                        and not bool(view._mode_switch_pending)
                        and app._is_sheet_exact_current("S")
                        and bool(app._sheet_exact_entry("S").get("full_detail_terminal"))
                        and bool(view._data_ready)
                        and bool(view._prepared_complete)
                        and str(view.only_diff_cb.cget("state")) == "normal"
                    ),
                    deadline=deadline,
                    stage="full-mode terminal cache publication",
                )
            _pump(app.root, min(deadline, time.monotonic() + 0.05))
            assert not hits, hits
            _assert_hard_mutation_unchanged(
                hard_before_disable,
                _hard_mutation_snapshot(app, view, "S", input_paths),
                action="only-diff cache-only disable",
            )
            assert _render_cache_version(view) > render_before_disable, (
                render_before_disable,
                _render_cache_version(view),
            )
            _assert_single_cached_publication(
                disable_publications,
                app=app,
                view=view,
                action="only-diff cache-only disable",
                expected_rows=disable_rows,
                requested_value=0,
                expected_mode_switch_seq=disable_mode_switch_before + 1,
                expected_compute_generation=disable_compute_generation,
                expected_exact_generation=disable_exact_generation,
            )
            assert tuple(view._full_display_rows) == full_logical_rows_before == disable_rows
            viewport_cap = min(sm._VIRTUAL_VIEWPORT_MAX_ROWS, len(full_logical_rows_before))
            selected_offset = full_logical_rows_before.index(int(pair_idx))
            expected_window_start = max(
                0,
                min(selected_offset, max(0, len(full_logical_rows_before) - viewport_cap)),
            )
            assert int(view._virtual_window_start) == expected_window_start
            assert tuple(view.display_rows) == full_logical_rows_before[
                expected_window_start:expected_window_start + viewport_cap
            ]
            _assert_selected_diff(view, pair_idx, 2)
            assert {name: _sha256(path) for name, path in input_paths.items()} == input_hashes
            assert not app._edit_workbooks_ready()
            assert time.monotonic() <= deadline
            print(
                "GUI_SELF_TEST_ONLY_DIFF_OK "
                + json.dumps(
                    {
                        "case": _CASES[0],
                        "deadline_seconds": 90,
                        "input_hashes": input_hashes,
                        "only_diff_cache": bool(view._has_valid_only_diff_snapshot_cache()),
                        "selection_pair": int(pair_idx),
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
        print(_CASES[0], flush=True)
        return
    selected = (args.case,) if args.case else _CASES
    for case_name in selected:
        if case_name != _CASES[0]:
            raise AssertionError(case_name)
        _run_selection_restore()
    print(f"GUI_SELF_TEST_ONLY_DIFF_SUITE_OK ({len(selected)} cases)", flush=True)


if __name__ == "__main__":
    main()
