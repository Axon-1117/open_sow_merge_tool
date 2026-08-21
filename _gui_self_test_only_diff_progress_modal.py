"""B3 GUI regression: public only-diff progress cancel/retry is bounded and view-only."""

import argparse
from contextlib import contextmanager
import hashlib
import json
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


_CASES = ("only-diff-progress-cancel-retry",)


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
        "prepared_topology": _prepared_topology_snapshot(view),
        "immutable_surface": _immutable_surface_snapshot(view),
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


def _prepared_topology_snapshot(view) -> dict:
    """Immutable logical/physical mapping must not move during view-only work."""
    return {
        "row_pairs": tuple(tuple(pair) for pair in (view.row_pairs or ())),
        "row_a_to_pair_idx": {
            int(row): int(pair_idx)
            for row, pair_idx in (view.row_a_to_pair_idx or {}).items()
        },
        "row_b_to_pair_idx": {
            int(row): int(pair_idx)
            for row, pair_idx in (view.row_b_to_pair_idx or {}).items()
        },
        "mine_to_base_row": {
            int(row): int(base_row)
            for row, base_row in (view.mine_to_base_row or {}).items()
        },
        "theirs_to_base_row": {
            int(row): int(base_row)
            for row, base_row in (view.theirs_to_base_row or {}).items()
        },
        "pair_base_row_override": {
            int(pair_idx): (None if base_row is None else int(base_row))
            for pair_idx, base_row in (view.pair_base_row_override or {}).items()
        },
    }


def _column_confidence_semantics(confidence) -> dict:
    return {
        "score": float(getattr(confidence, "score", 0.0)),
        "reason": str(getattr(confidence, "reason", "") or ""),
        "evidence": tuple(
            str(item) for item in (getattr(confidence, "evidence", ()) or ())
        ),
        "cause_codes": tuple(
            str(item) for item in (getattr(confidence, "cause_codes", ()) or ())
        ),
        "ambiguous": bool(getattr(confidence, "ambiguous", False)),
    }


def _immutable_surface_snapshot(view) -> dict:
    """Canonical immutable column/cache semantics; no Tk or worksheet reads."""
    cache = getattr(view, "column_comparison_cache", None)
    projection = getattr(view, "column_projection", None)
    model = getattr(cache, "model", None)
    projection_model = getattr(projection, "model", None)
    assert cache is not None and model is not None and projection is not None
    assert projection_model is model
    slots = tuple(
        {
            "logical_idx": int(slot.logical_idx),
            "mine_col": None if slot.mine_col is None else int(slot.mine_col),
            "base_col": None if slot.base_col is None else int(slot.base_col),
            "theirs_col": None if slot.theirs_col is None else int(slot.theirs_col),
            "state": str(slot.state),
            "base_boundary": (
                None if slot.base_boundary is None else int(slot.base_boundary)
            ),
            "origin_side": slot.origin_side,
            "confidence": _column_confidence_semantics(slot.confidence),
        }
        for slot in tuple(model.slots)
    )
    cache_key = model.cache_key
    semantics = {
        "cache_key": {
            "sheet_name": str(cache_key.sheet_name),
            "row_model_version": int(cache_key.row_model_version),
            "column_model_version": int(cache_key.column_model_version),
            "mine_edit_version": int(cache_key.mine_edit_version),
            "base_edit_version": int(cache_key.base_edit_version),
            "theirs_edit_version": int(cache_key.theirs_edit_version),
        },
        "slots": slots,
        "blocks": tuple(
            {
                "ordinal": int(block.ordinal),
                "slot_indices": tuple(int(idx) for idx in block.slot_indices),
                "state": str(block.state),
                "confidence": _column_confidence_semantics(block.confidence),
            }
            for block in tuple(model.blocks)
        ),
        "physical_maps": {
            name: tuple(
                (int(key), int(value))
                for key, value in tuple(getattr(model, name).entries)
            )
            for name in (
                "mine_physical_to_logical",
                "base_physical_to_logical",
                "theirs_physical_to_logical",
                "mine_logical_to_physical",
                "base_logical_to_physical",
                "theirs_logical_to_physical",
            )
        },
        "model_confidence": _column_confidence_semantics(model.confidence),
        "structural_diff_cols": tuple(
            sorted(int(value) for value in cache.structural_diff_cols)
        ),
        "unresolved_cols": tuple(sorted(int(value) for value in cache.unresolved_cols)),
        "projection_block_ordinal_by_slot": tuple(
            int(value) for value in projection.block_ordinal_by_slot
        ),
        "two_way_alignment": cache.two_way_alignment is not None,
        "three_way_alignment": cache.three_way_alignment is not None,
    }
    return {
        "bounds": {
            "max_row": int(view.max_row),
            "max_col": int(view.max_col),
            "col_max_a": int(view.col_max_a),
            "col_max_b": int(view.col_max_b),
            "col_max_base": int(view.col_max_base),
        },
        "align_rows_enabled": bool(view._align_rows_enabled),
        "missing_base_row_map": {
            int(pair_idx): int(base_row)
            for pair_idx, base_row in (view._missing_base_row_map or {}).items()
        },
        "sheet_structural_diff": bool(view._sheet_structural_diff),
        "only_diff_source_version": int(view._only_diff_source_version),
        "cache_identity": id(cache),
        "projection_identity": id(projection),
        "model_identity": id(model),
        "semantics": semantics,
        "semantic_digest": _fingerprint(semantics),
    }


def _assert_two_way_fixture_surface(view) -> None:
    """Make Base absence and the two-column schema explicit in this fixture."""
    surface = _immutable_surface_snapshot(view)
    bounds = surface["bounds"]
    assert bounds == {
        "max_row": 2202,
        "max_col": 2,
        "col_max_a": 2,
        "col_max_b": 2,
        "col_max_base": 1,
    }, surface
    assert surface["align_rows_enabled"]
    assert surface["missing_base_row_map"] == {}
    assert not surface["sheet_structural_diff"]
    semantics = surface["semantics"]
    assert semantics["two_way_alignment"] and not semantics["three_way_alignment"]
    assert semantics["structural_diff_cols"] == ()
    assert semantics["unresolved_cols"] == ()
    assert all(slot["base_col"] is None for slot in semantics["slots"])
    assert semantics["physical_maps"]["base_physical_to_logical"] == ()
    assert semantics["physical_maps"]["base_logical_to_physical"] == ()
    assert getattr(view, "column_alignment_3way", None) is None

def _prepared_content_snapshot(view) -> dict:
    """Snapshot-only fragments/diffs, kept separate from immutable topology."""
    def _raw(parts):
        return {
            int(pair_idx): tuple(values)
            for pair_idx, values in (parts or {}).items()
        }

    def _diffs(columns):
        return {
            int(pair_idx): frozenset(int(column) for column in values)
            for pair_idx, values in (columns or {}).items()
        }

    return {
        "raw_a": _raw(getattr(view, "pair_raw_parts_a", None)),
        "raw_b": _raw(getattr(view, "pair_raw_parts_b", None)),
        "raw_base": _raw(getattr(view, "pair_raw_parts_base", None)),
        "pair_diff_cols": _diffs(getattr(view, "pair_diff_cols", None)),
        "pair_base_diff_cols": _diffs(
            getattr(view, "pair_base_diff_cols", None)
        ),
    }


def _assert_prepared_content_unchanged(before, after, *, action: str) -> None:
    diffs = _recursive_field_diffs(before, after, path="prepared_content")
    assert not diffs, (
        f"{action}: forbidden prepared cache drift: "
        f"{json.dumps(diffs, ensure_ascii=False, sort_keys=True)}"
    )


def _assert_retry_prepared_content(before, after, *, view, target_pair: int) -> None:
    """Allow only the fixture's one immutable async-only-diff publication."""
    target_pair = int(target_pair)
    base_diff_domain = set(range(len(view.row_pairs)))
    assert after["pair_diff_cols"] == {target_pair: frozenset({2})}, after
    assert set(before["pair_base_diff_cols"]) == base_diff_domain, before
    assert all(
        columns == frozenset()
        for columns in before["pair_base_diff_cols"].values()
    ), before
    assert after["pair_base_diff_cols"] == before["pair_base_diff_cols"], after
    assert before["raw_base"] == {}, before
    assert after["raw_base"] == {}, after
    expected_a = ("1900", "value-1900")
    expected_b = ("1900", "value-1900-changed")
    assert after["raw_a"].get(target_pair) == expected_a, after
    assert after["raw_b"].get(target_pair) == expected_b, after
    for side in ("raw_a", "raw_b"):
        before_raw = before[side]
        after_raw = after[side]
        assert set(after_raw) - set(before_raw) <= {target_pair}, (
            side, before_raw, after_raw,
        )
        assert set(before_raw) - set(after_raw) <= {target_pair}, (
            side, before_raw, after_raw,
        )
        for pair_idx in sorted((set(before_raw) | set(after_raw)) - {target_pair}):
            assert after_raw.get(pair_idx) == before_raw.get(pair_idx), (
                side,
                pair_idx,
                before_raw.get(pair_idx),
                after_raw.get(pair_idx),
            )


def _presentation_diagnostic(view) -> dict:
    """Useful failure context; intentionally excluded from hard mutation state."""
    return {
        "render_cache_version": _render_cache_version(view),
        "display_rows": tuple(getattr(view, "display_rows", ()) or ()),
        "only_diff_rows_cache": tuple(
            getattr(view, "_only_diff_rows_cache", ()) or ()
        ),
        "only_diff_rows_exact": bool(
            getattr(view, "_only_diff_rows_exact", False)
        ),
        "only_diff_cache_key": getattr(view, "_only_diff_rows_cache_key", None),
        "pending_exact_render": bool(getattr(view, "_pending_exact_render", False)),
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
    ws.title = "WorldMonster"
    ws.append(["id@id", "value"])
    ws.append(["int32", "string"])
    for row_id in range(1, 2201):
        value = f"value-{row_id}"
        if changed and row_id == 1900:
            value += "-changed"
        ws.append([row_id, value])
    wb.save(path)
    wb.close()


def _pump(app, deadline: float) -> None:
    while time.monotonic() < deadline:
        app.root.update_idletasks()
        app.root.update()
        time.sleep(0.005)


def _progress_dialog_facts(app) -> dict:
    win = getattr(app, "_only_diff_progress_win", None)
    facts = {
        "owner": getattr(app, "_only_diff_progress_owner", None),
        "visible_token": getattr(app, "_only_diff_progress_visible_token", None),
        "show_after_id": getattr(app, "_only_diff_progress_show_after_id", None),
        "show_token": getattr(app, "_only_diff_progress_show_token", None),
        "watchdog_after_id": getattr(
            app, "_only_diff_progress_watchdog_after_id", None
        ),
        "watchdog_token": getattr(
            app, "_only_diff_progress_watchdog_token", None
        ),
    }
    if win is None:
        facts["window"] = None
        return facts
    try:
        facts["window"] = {
            "exists": bool(win.winfo_exists()),
            "state": str(win.state()),
            "mapped": bool(win.winfo_ismapped()),
            "viewable": bool(win.winfo_viewable()),
            "grabbed": bool(win.grab_current() == win),
        }
    except Exception as exc:
        facts["window"] = f"{type(exc).__name__}: {exc}"
    return facts


def _wait(app, predicate, *, deadline: float, stage: str) -> None:
    while time.monotonic() < deadline:
        _pump(app, min(deadline, time.monotonic() + 0.05))
        if predicate():
            return
        view = app.sheet_views.get("WorldMonster")
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
    view = app.sheet_views.get("WorldMonster")
    raise AssertionError(
        f"timeout {stage}: "
        + repr(
            {
                "entry": app._sheet_exact_entry("WorldMonster"),
                "only_diff": None if view is None else int(view.only_diff_var.get()),
                "building": None if view is None else bool(view._only_diff_async_building),
                "cache": None if view is None else bool(
                    view._has_valid_only_diff_snapshot_cache()
                ),
                "progress": _progress_dialog_facts(app),
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
            raise AssertionError(f"only-diff progress route accessed {label}")

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


@contextmanager
def _observe_progress_show_schedule(app):
    """Observe the real after(0) show path without changing its scheduling."""
    records = []
    original = app._safe_root_after

    def _wrapped(delay_ms, callback):
        qualname = str(getattr(callback, "__qualname__", ""))
        kind = (
            "show"
            if qualname.endswith("._show_dialog")
            else "watchdog"
            if qualname.endswith("._show_watchdog")
            else None
        )
        if kind is None:
            return original(delay_ms, callback)
        record = {
            "kind": kind,
            "delay_ms": int(delay_ms),
            "callback_qualname": qualname,
            "after_id": None,
            "executed": 0,
        }
        records.append(record)

        def _observed_callback():
            record["executed"] += 1
            return callback()

        record["after_id"] = original(delay_ms, _observed_callback)
        return record["after_id"]

    app._safe_root_after = _wrapped
    try:
        yield records
    finally:
        app._safe_root_after = original


def _progress_schedule_record(records, kind: str) -> dict:
    matched = [record for record in records if record.get("kind") == kind]
    assert len(matched) == 1, (kind, records)
    return matched[0]


def _progress_dialog_visible(app, view, seq: int, records) -> bool:
    show = _progress_schedule_record(records, "show")
    win = getattr(app, "_only_diff_progress_win", None)
    if win is None:
        return False
    try:
        return bool(
            show.get("after_id") is not None
            and int(show.get("executed") or 0) == 1
            and app._only_diff_progress_owner == (view, int(seq))
            and app._only_diff_progress_visible_token == (view, int(seq))
            and str(win.state()) == "normal"
            and bool(win.winfo_ismapped())
            and bool(win.winfo_viewable())
            and win.grab_current() == win
            and str(app._only_diff_progress_cancel_btn.cget("state")) == "normal"
        )
    except Exception:
        return False


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


def _run_progress_cancel_retry() -> None:
    original_settings_path = sm._SETTINGS_PATH
    user_settings = Path(original_settings_path)
    user_settings_before = _path_snapshot(user_settings)
    app = None
    root_path = None
    try:
        with tempfile.TemporaryDirectory(prefix="sow_only_diff_progress_") as root:
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
            print("ONLY_DIFF_PROGRESS_STAGE open-current-exact", flush=True)
            app = sm.SowMergeApp(mine, theirs, initial_sheet="WorldMonster")
            view = app.sheet_views["WorldMonster"]
            _wait(
                app,
                lambda: (
                    app._is_sheet_exact_current(view.sheet)
                    and bool(app._sheet_exact_entry(view.sheet).get("full_detail_terminal"))
                    and bool(view._prepared_complete)
                    and bool(view._data_ready)
                    and not bool(view._pending_exact_render)
                    and not bool(view.only_diff_var.get())
                    and str(view.only_diff_cb.cget("state")) == "normal"
                ),
                deadline=deadline,
                stage="selected exact full surface",
            )
            assert view._is_exact_immutable_view_ready()
            _assert_two_way_fixture_surface(view)
            stable_full_rows = tuple(view.display_rows)
            prior_exact_entry = dict(app._sheet_exact_entry(view.sheet) or {})
            assert str(prior_exact_entry.get("state") or "") in sm._SHEET_EXACT_TERMINAL
            assert bool(prior_exact_entry.get("full_detail_terminal", False))
            original_start = view._start_async_large_only_diff_build
            fake_starts = []

            def _controlled_start(*, user_initiated=False):
                assert user_initiated is True
                view._only_diff_async_build_seq += 1
                seq = int(view._only_diff_async_build_seq)
                fake_starts.append(seq)
                # Exercise the same production state transition as a real
                # immutable worker start; the controlled path only omits disk IO.
                assert view._begin_only_diff_exact_transition(
                    seq, view._current_only_diff_cache_key()
                )
                app._begin_only_diff_progress(view, seq)
                return True

            print("ONLY_DIFF_PROGRESS_STAGE public-cancel", flush=True)
            view._invalidate_only_diff_snapshot_cache()
            view._start_async_large_only_diff_build = _controlled_start
            hard_before_start = _hard_mutation_snapshot(
                app, view, view.sheet, input_paths
            )
            render_before_start = _render_cache_version(view)
            try:
                with (
                    _forbid_view_only_access(app, view) as hits,
                    _observe_cached_only_diff_publisher(
                        app, view
                    ) as fake_publications,
                    _observe_progress_show_schedule(app) as progress_schedule,
                ):
                    view.only_diff_cb.invoke()
                    show_deadline = min(deadline, time.monotonic() + 0.5)
                    _wait(
                        app,
                        lambda: (
                            len(fake_starts) == 1
                            and _progress_dialog_visible(
                                app, view, fake_starts[0], progress_schedule
                            )
                            and str(view.only_diff_cb.cget("state")) == "disabled"
                            and bool(view.only_diff_var.get())
                        ),
                        deadline=show_deadline,
                        stage="progress dialog visible within 500ms",
                    )
                    show_record = _progress_schedule_record(
                        progress_schedule, "show"
                    )
                    watchdog_record = _progress_schedule_record(
                        progress_schedule, "watchdog"
                    )
                    assert show_record["after_id"] is not None, progress_schedule
                    assert show_record["executed"] == 1, progress_schedule
                    assert watchdog_record["after_id"] is not None, progress_schedule
                    assert app._only_diff_progress_visible_token == (
                        view,
                        fake_starts[0],
                    )
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_start,
                        _hard_mutation_snapshot(app, view, view.sheet, input_paths),
                        action="only-diff controlled start",
                    )
                    assert _render_cache_version(view) == render_before_start
                    seq = fake_starts[0]
                    cancel_button = app._only_diff_progress_cancel_btn
                    assert cancel_button is not None
                    hard_before_cancel = _hard_mutation_snapshot(
                        app, view, view.sheet, input_paths
                    )
                    prepared_before_cancel = _prepared_content_snapshot(view)
                    render_before_cancel = _render_cache_version(view)
                    cancel_button.invoke()
                    _wait(
                        app,
                        lambda: (
                            app._only_diff_progress_owner is None
                            and not bool(view._only_diff_async_building)
                            and not bool(view.only_diff_var.get())
                            and tuple(view.display_rows) == stable_full_rows
                            and dict(app._sheet_exact_entry(view.sheet) or {})
                            == prior_exact_entry
                            and str(view.only_diff_cb.cget("state")) == "normal"
                        ),
                        deadline=deadline,
                        stage="cancel terminal",
                    )
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_cancel,
                        _hard_mutation_snapshot(app, view, view.sheet, input_paths),
                        action="only-diff public cancel",
                    )
                    _assert_prepared_content_unchanged(
                        prepared_before_cancel,
                        _prepared_content_snapshot(view),
                        action="only-diff public cancel",
                    )
                    assert _render_cache_version(view) == render_before_cancel
                    # A scheduling failure is a real production dialog-failure
                    # disposition: the controlled start still uses the exact
                    # transition helper and must restore the same terminal entry.
                    print("ONLY_DIFF_PROGRESS_STAGE show-schedule-failure", flush=True)
                    view._invalidate_only_diff_snapshot_cache()
                    show_fail_prior = dict(app._sheet_exact_entry(view.sheet) or {})
                    safe_after_original = app._safe_root_after
                    target_show_calls = []

                    def _fail_show_schedule(delay_ms, callback):
                        callback_qualname = str(
                            getattr(callback, "__qualname__", "")
                        )
                        if (
                            int(delay_ms) == 0
                            and callback_qualname.endswith("._show_dialog")
                        ):
                            target_show_calls.append((int(delay_ms), callback))
                            return None
                        return safe_after_original(delay_ms, callback)

                    app._safe_root_after = _fail_show_schedule
                    hard_before_show_failure = _hard_mutation_snapshot(
                        app, view, view.sheet, input_paths
                    )
                    prepared_before_show_failure = _prepared_content_snapshot(view)
                    try:
                        view.only_diff_cb.invoke()
                        _wait(
                            app,
                            lambda: (
                                len(fake_starts) == 2
                                and app._only_diff_progress_owner is None
                                and not bool(view._only_diff_async_building)
                                and not bool(view.only_diff_var.get())
                                and dict(app._sheet_exact_entry(view.sheet) or {})
                                == show_fail_prior
                                and str(view.only_diff_cb.cget("state")) == "normal"
                            ),
                            deadline=min(deadline, time.monotonic() + 0.5),
                            stage="dialog show schedule fail restores exact terminal",
                        )
                        assert len(target_show_calls) == 1, target_show_calls
                        _pump(app, min(deadline, time.monotonic() + 0.05))
                        _assert_hard_mutation_unchanged(
                            hard_before_show_failure,
                            _hard_mutation_snapshot(
                                app, view, view.sheet, input_paths
                            ),
                            action="only-diff dialog show schedule failure",
                        )
                        _assert_prepared_content_unchanged(
                            prepared_before_show_failure,
                            _prepared_content_snapshot(view),
                            action="only-diff dialog show schedule failure",
                        )
                    finally:
                        app._safe_root_after = safe_after_original
                    hard_before_stale = _hard_mutation_snapshot(
                        app, view, view.sheet, input_paths
                    )
                    prepared_before_stale = _prepared_content_snapshot(view)
                    render_before_stale = _render_cache_version(view)
                    app._update_only_diff_progress(view, seq, "stale", 2200, 2200)
                    assert app._only_diff_progress_owner is None
                    _pump(app, min(deadline, time.monotonic() + 0.05))
                    _assert_hard_mutation_unchanged(
                        hard_before_stale,
                        _hard_mutation_snapshot(app, view, view.sheet, input_paths),
                        action="only-diff stale progress drop",
                    )
                    _assert_prepared_content_unchanged(
                        prepared_before_stale,
                        _prepared_content_snapshot(view),
                        action="only-diff stale progress drop",
                    )
                    assert _render_cache_version(view) == render_before_stale
                assert not hits, hits
                assert fake_publications == [], fake_publications
            finally:
                view._start_async_large_only_diff_build = original_start

            retry_starts = []
            original_retry_start = view._start_async_large_only_diff_build

            def _recording_retry(*, user_initiated=False):
                retry_starts.append(bool(user_initiated))
                return original_retry_start(user_initiated=user_initiated)

            print("ONLY_DIFF_PROGRESS_STAGE public-retry-terminal", flush=True)
            view._start_async_large_only_diff_build = _recording_retry
            assert not view._has_valid_only_diff_snapshot_cache(), (
                "retry must prove the real worker/apply path, not reuse an exact cache"
            )
            hard_before_retry = _hard_mutation_snapshot(
                app, view, view.sheet, input_paths
            )
            prepared_before_retry = _prepared_content_snapshot(view)
            target_physical_row = 1902
            target_pair = int(view.row_a_to_pair_idx[target_physical_row])
            assert tuple(view.row_pairs[target_pair]) == (
                target_physical_row,
                target_physical_row,
            )
            render_before_retry = _render_cache_version(view)
            presentation_before_retry = _presentation_diagnostic(view)
            retry_rows = tuple(
                pair_idx
                for pair_idx in range(len(view.row_pairs))
                if view._pair_has_visual_diff(pair_idx)
            )
            assert retry_rows == (target_pair,), (target_pair, retry_rows)
            retry_mode_switch_before = int(view._mode_switch_seq)
            retry_compute_generation = int(app._sheet_compute_generation[view.sheet])
            retry_exact_generation = int(app._sheet_exact_entry(view.sheet)["generation"])
            retry_build_seq_before = int(view._only_diff_async_build_seq)
            try:
                with _forbid_view_only_access(app, view) as hits, _observe_cached_only_diff_publisher(app, view) as retry_publications:
                    view.only_diff_cb.invoke()
                    _wait(
                        app,
                        lambda: (
                            retry_starts == [True]
                            and not bool(view._only_diff_async_building)
                            and app._only_diff_progress_owner is None
                            and bool(view.only_diff_var.get())
                            and bool(view._only_diff_rows_exact)
                            and bool(view._has_valid_only_diff_snapshot_cache())
                            and app._is_sheet_exact_current(view.sheet)
                            and bool(app._sheet_exact_entry(view.sheet).get("full_detail_terminal"))
                            and str(view.only_diff_cb.cget("state")) == "normal"
                        ),
                        deadline=deadline,
                        stage="retry current terminal",
                    )
                _pump(app, min(deadline, time.monotonic() + 0.05))
                assert not hits, hits
                _assert_hard_mutation_unchanged(
                    hard_before_retry,
                    _hard_mutation_snapshot(app, view, view.sheet, input_paths),
                    action="only-diff real retry apply and cache publication",
                )
                prepared_after_retry = _prepared_content_snapshot(view)
                _assert_retry_prepared_content(
                    prepared_before_retry,
                    prepared_after_retry,
                    view=view,
                    target_pair=target_pair,
                )
                assert tuple(view._only_diff_rows_cache or ()) == (target_pair,)
                assert view._only_diff_rows_cache_key == view._current_only_diff_cache_key()
                assert bool(view._only_diff_rows_exact)
                assert app._is_sheet_exact_current(view.sheet)
                _assert_two_way_fixture_surface(view)
                assert _render_cache_version(view) > render_before_retry, (
                    render_before_retry,
                    _render_cache_version(view),
                    presentation_before_retry,
                    _presentation_diagnostic(view),
                )
                _assert_single_cached_publication(
                    retry_publications,
                    app=app,
                    view=view,
                    action="only-diff real retry apply and cache publication",
                    expected_rows=retry_rows,
                    requested_value=1,
                    expected_mode_switch_seq=retry_mode_switch_before + 1,
                    expected_compute_generation=retry_compute_generation,
                    expected_exact_generation=retry_exact_generation,
                    expected_async_build_seq=retry_build_seq_before + 1,
                )
            finally:
                view._start_async_large_only_diff_build = original_retry_start

            popup = app._only_diff_progress_win
            assert popup is None or not bool(popup.winfo_viewable())
            assert tuple(view.display_rows) != stable_full_rows
            assert {name: _sha256(path) for name, path in input_paths.items()} == input_hashes
            assert not app._edit_workbooks_ready()
            assert time.monotonic() <= deadline
            print(
                "GUI_SELF_TEST_ONLY_DIFF_PROGRESS_MODAL_OK "
                + json.dumps(
                    {
                        "case": _CASES[0],
                        "deadline_seconds": 90,
                        "fake_cancel_seq": fake_starts[0],
                        "retry_starts": retry_starts,
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
        print(_CASES[0], flush=True)
        return
    selected = (args.case,) if args.case else _CASES
    for case_name in selected:
        if case_name != _CASES[0]:
            raise AssertionError(case_name)
        _run_progress_cancel_retry()
    print(f"GUI_SELF_TEST_ONLY_DIFF_PROGRESS_MODAL_SUITE_OK ({len(selected)} cases)", flush=True)


if __name__ == "__main__":
    main()
