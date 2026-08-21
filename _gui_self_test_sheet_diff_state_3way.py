"""Three-way exact Sheet-state regression using immutable typed snapshots."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


_CASE = "sheet-diff-state-3way"
_SHEETS = ("S_same", "S_base_diff", "S_ab_diff", "S_new_common")
_EXPECTED_EXACT = {
    "S_same": sm._SHEET_EXACT_SAME,
    "S_base_diff": sm._SHEET_EXACT_CHANGED,
    "S_ab_diff": sm._SHEET_EXACT_CHANGED,
    "S_new_common": sm._SHEET_EXACT_CHANGED,
}
_EXPECTED_BADGES = {
    "S_same": "精确相同",
    "S_base_diff": "精确变更",
    "S_ab_diff": "精确变更",
    "S_new_common": "精确变更",
}
_HIDDEN_SNAPSHOT_FAILURE = (
    "后台精确快照技术失败，已停止自动兼容读取以保持交互响应；"
    "请选择此 Sheet 后重试。"
)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _path_snapshot(path: Path) -> tuple[bool, bytes | None]:
    return (True, path.read_bytes()) if path.exists() else (False, None)


def _normalized_path(path: str | Path) -> str:
    return os.path.normcase(os.path.abspath(os.fspath(path)))


def _make_typed_book(path: Path, sheets: dict[str, tuple[tuple[str, str], ...]]) -> None:
    workbook = Workbook()
    try:
        workbook.remove(workbook.active)
        for sheet, records in sheets.items():
            worksheet = workbook.create_sheet(sheet)
            worksheet.append(["id@id", "value"])
            worksheet.append(["string", "string"])
            for record_id, value in records:
                worksheet.append([record_id, value])
        workbook.save(path)
    finally:
        workbook.close()


def _diagnostic(app) -> dict:
    return {
        "selected_sheet": getattr(app, "selected_sheet", None),
        "notebook_sheet": (
            app.nb.tab(app.nb.select(), "text") if app is not None and app.nb.select() else None
        ),
        "generations": dict(getattr(app, "_sheet_compute_generation", {}) or {}),
        "sheet_diff_state": dict(getattr(app, "sheet_diff_state", {}) or {}),
        "entries": {
            sheet: dict(app._sheet_exact_entry(sheet))
            for sheet in _SHEETS
        },
        "meta": {sheet: app.get_sheet_meta(sheet) for sheet in _SHEETS},
        "loaded": dict(getattr(app, "_sheet_loaded", {}) or {}),
    }


def _pump(app) -> None:
    app.root.update_idletasks()
    app.root.update()


def _wait_initial_summaries_with_hidden_failure(app, *, deadline: float) -> int:
    while time.monotonic() < deadline:
        _pump(app)
        same_entry = app._sheet_exact_entry("S_same")
        hidden_entries = {
            sheet: app._sheet_exact_entry(sheet)
            for sheet in ("S_base_diff", "S_ab_diff")
        }
        new_entry = app._sheet_exact_entry("S_new_common")
        new_generation = int(app._sheet_compute_generation["S_new_common"])
        if (
            str(getattr(app, "selected_sheet", "")) == "S_same"
            and app._is_sheet_exact_current("S_same")
            and same_entry.get("state") == sm._SHEET_EXACT_SAME
            and bool(same_entry.get("full_detail_terminal", False))
            and all(
                app._is_sheet_exact_current(sheet)
                and entry.get("state") == sm._SHEET_EXACT_CHANGED
                and int(entry.get("generation", -1))
                == int(app._sheet_compute_generation[sheet])
                and not bool(entry.get("full_detail_terminal", False))
                for sheet, entry in hidden_entries.items()
            )
            and int(new_entry.get("generation", -1)) == new_generation
            and new_entry.get("state") == sm._SHEET_EXACT_FAILED
            and not bool(new_entry.get("full_detail_terminal", False))
            and new_entry.get("reason") == _HIDDEN_SNAPSHOT_FAILURE
            and app.sheet_diff_state.get("S_new_common") == 0
            and not bool(getattr(app, "_sheet_loaded", {}).get("S_new_common", False))
            and app.get_sheet_meta("S_new_common")
            == {"has_a": True, "has_b": True, "has_base": False, "view_mode": "normal"}
        ):
            return new_generation
        time.sleep(0.01)
    raise AssertionError(
        "three-way initial summary/hidden-failure wait timed out: "
        + json.dumps(_diagnostic(app), default=str, sort_keys=True)
    )


def _select_full_detail(app, sheet: str, *, deadline: float):
    app.nb.select(app._sheet_containers[sheet])
    while time.monotonic() < deadline:
        _pump(app)
        entry = app._sheet_exact_entry(sheet)
        view = app.sheet_views.get(sheet)
        notebook_sheet = app.nb.tab(app.nb.select(), "text") if app.nb.select() else None
        if (
            notebook_sheet == sheet
            and str(getattr(app, "selected_sheet", "")) == sheet
            and app._is_sheet_exact_current(sheet)
            and entry.get("state") == _EXPECTED_EXACT[sheet]
            and int(entry.get("generation", -1))
            == int(app._sheet_compute_generation[sheet])
            and bool(entry.get("full_detail_terminal", False))
            and view is not None
            and bool(getattr(view, "_prepared_complete", False))
            and bool(getattr(view, "_data_ready", False))
            and bool(getattr(view, "_row_model_exact", False))
            and not bool(getattr(view, "_pending_exact_render", False))
        ):
            return view
        time.sleep(0.01)
    raise AssertionError(
        f"selected full-detail timed out: sheet={sheet} "
        + json.dumps(_diagnostic(app), default=str, sort_keys=True)
    )


def _install_exact_state_transition_spy(app, sheet: str, generation: int):
    original = app._set_sheet_exact_state
    transitions: list[dict] = []

    def _delegating(target_sheet, state, *args, **kwargs):
        result = original(target_sheet, state, *args, **kwargs)
        if str(target_sheet) == sheet:
            transitions.append(
                {
                    "sheet": str(target_sheet),
                    "state": str(state),
                    "generation": int(
                        kwargs.get("generation", app._sheet_compute_generation.get(sheet, -1))
                    ),
                    "accepted": bool(result),
                    "at": time.monotonic(),
                }
            )
        return result

    app._set_sheet_exact_state = _delegating
    return original, transitions


def _select_hidden_failure_retry_full_detail(
    app, sheet: str, *, generation: int, transitions: list[dict], deadline: float
):
    app.nb.select(app._sheet_containers[sheet])
    saw_selected = False
    saw_active = False
    while time.monotonic() < deadline:
        _pump(app)
        entry = app._sheet_exact_entry(sheet)
        view = app.sheet_views.get(sheet)
        notebook_sheet = app.nb.tab(app.nb.select(), "text") if app.nb.select() else None
        selected = notebook_sheet == sheet and str(getattr(app, "selected_sheet", "")) == sheet
        saw_selected = saw_selected or selected
        current_generation = int(app._sheet_compute_generation[sheet])
        saw_active = saw_active or (
            selected
            and current_generation == generation
            and int(entry.get("generation", -1)) == generation
            and entry.get("state") in (sm._SHEET_EXACT_PENDING, sm._SHEET_EXACT_CALCULATING)
        ) or any(
            transition.get("accepted")
            and transition.get("generation") == generation
            and transition.get("state")
            in (sm._SHEET_EXACT_PENDING, sm._SHEET_EXACT_CALCULATING)
            for transition in transitions
        )
        if (
            selected
            and saw_selected
            and saw_active
            and current_generation == generation
            and int(entry.get("generation", -1)) == generation
            and app._is_sheet_exact_current(sheet)
            and entry.get("state") == sm._SHEET_EXACT_CHANGED
            and bool(entry.get("full_detail_terminal", False))
            and view is not None
            and bool(getattr(view, "_prepared_complete", False))
            and bool(getattr(view, "_data_ready", False))
            and bool(getattr(view, "_row_model_exact", False))
            and not bool(getattr(view, "_pending_exact_render", False))
        ):
            return view
        time.sleep(0.01)
    raise AssertionError(
        "selected hidden-failure retry timed out: "
        + json.dumps(
            {
                "generation": generation,
                "saw_selected": saw_selected,
                "saw_active": saw_active,
                "transitions": transitions,
                "diagnostic": _diagnostic(app),
            },
            default=str,
            sort_keys=True,
        )
    )


def _assert_nav_and_current(app) -> None:
    expected_diff = {
        "S_same": 0,
        "S_base_diff": 2,
        "S_ab_diff": 2,
        "S_new_common": 2,
    }
    app.refresh_sheet_nav()
    for sheet in _SHEETS:
        entry = app._sheet_exact_entry(sheet)
        assert entry.get("state") == _EXPECTED_EXACT[sheet], entry
        assert int(entry.get("generation", -1)) == int(app._sheet_compute_generation[sheet]), entry
        assert app._is_sheet_exact_current(sheet), entry
        assert app.sheet_diff_state.get(sheet) == expected_diff[sheet], app.sheet_diff_state
        button = app._nav_buttons.get(sheet)
        assert button is not None and button.winfo_exists(), sheet
        text = str(button.cget("text"))
        assert text.startswith(f"{sheet} · ") and _EXPECTED_BADGES[sheet] in text, text


def _assert_full_three_way_projection(
    app, view, sheet: str, *, require_complete_base_fragments: bool = True
) -> None:
    entry = app._sheet_exact_entry(sheet)
    assert bool(getattr(app, "has_base", False))
    assert view._is_three_way_enabled()
    assert app._is_sheet_exact_current(sheet)
    assert bool(entry.get("full_detail_terminal", False))
    assert bool(getattr(view, "_prepared_complete", False))
    assert bool(getattr(view, "_data_ready", False))
    assert bool(getattr(view, "_row_model_exact", False))
    assert bool(getattr(view, "_pair_diff_full_exact", False))
    assert bool(getattr(view, "_base_diff_full_exact", False))
    if require_complete_base_fragments:
        assert view._has_complete_prepared_rows()
    assert view._column_mapping_is_current()


def _pair_for_record(view, record_id: str) -> int:
    for pair_idx, raw in dict(getattr(view, "pair_raw_parts_a", {}) or {}).items():
        if tuple(raw or ())[:1] == (record_id,):
            return int(pair_idx)
    raise AssertionError((record_id, dict(getattr(view, "pair_raw_parts_a", {}) or {})))


def _prepared_pair(view, pair_idx: int, side: str) -> tuple[str, str]:
    return tuple(
        view._prepared_value_for_logical_cell(pair_idx, side, logical_col)
        for logical_col in (1, 2)
    )


def _assert_base_relative_common_change(app, view) -> None:
    pair_idx = _pair_for_record(view, "base-diff-1")
    row_a, row_b = view.row_pairs[pair_idx]
    assert (row_a, row_b) == (3, 3)
    assert view.mine_to_base_row[row_a] == 3
    assert view.theirs_to_base_row[row_b] == 3
    assert view.pair_base_row_override[pair_idx] == 3
    raw_a = tuple(view.pair_raw_parts_a[pair_idx])
    raw_b = tuple(view.pair_raw_parts_b[pair_idx])
    raw_base = tuple(view.pair_raw_parts_base[pair_idx])
    assert raw_a == raw_b == ("base-diff-1", "new")
    assert raw_base == ("base-diff-1", "old")
    assert _prepared_pair(view, pair_idx, "A") == raw_a
    assert _prepared_pair(view, pair_idx, "B") == raw_b
    assert _prepared_pair(view, pair_idx, "BASE") == raw_base
    assert view.pair_diff_cols.get(pair_idx) == set()
    assert view.pair_base_diff_cols.get(pair_idx) == {2}
    comparison = sm.compare_logical_row_3way(
        view._active_column_comparison_cache(),
        raw_a,
        raw_base,
        raw_b,
        raw_a,
        raw_base,
        raw_b,
        mine_row=row_a,
        base_row=3,
        theirs_row=row_b,
    )
    assert comparison.mine_changed_cols == frozenset({2})
    assert comparison.theirs_changed_cols == frozenset({2})
    assert comparison.conflict_cols == frozenset()
    assert view._pair_has_visual_diff(pair_idx)
    assert not (getattr(app, "merge_conflict_cells_by_sheet", {}) or {}).get("S_base_diff")


def _assert_three_way_conflict_semantics(view) -> None:
    pair_idx = _pair_for_record(view, "ab-diff-1")
    row_a, row_b = view.row_pairs[pair_idx]
    assert (row_a, row_b) == (3, 3)
    assert view.pair_base_row_override[pair_idx] == 3
    raw_a = tuple(view.pair_raw_parts_a[pair_idx])
    raw_b = tuple(view.pair_raw_parts_b[pair_idx])
    raw_base = tuple(view.pair_raw_parts_base[pair_idx])
    assert raw_a == ("ab-diff-1", "mine")
    assert raw_base == ("ab-diff-1", "base")
    assert raw_b == ("ab-diff-1", "theirs")
    assert _prepared_pair(view, pair_idx, "A") == raw_a
    assert _prepared_pair(view, pair_idx, "B") == raw_b
    assert _prepared_pair(view, pair_idx, "BASE") == raw_base
    assert view.pair_diff_cols.get(pair_idx) == {2}
    assert view.pair_base_diff_cols.get(pair_idx) == {2}
    comparison = sm.compare_logical_row_3way(
        view._active_column_comparison_cache(),
        raw_a,
        raw_base,
        raw_b,
        raw_a,
        raw_base,
        raw_b,
        mine_row=row_a,
        base_row=3,
        theirs_row=row_b,
    )
    assert comparison.mine_changed_cols == frozenset({2})
    assert comparison.theirs_changed_cols == frozenset({2})
    assert comparison.conflict_cols == frozenset({2})
    assert view._pair_has_visual_diff(pair_idx)


def _assert_new_common_structural(app, view) -> None:
    meta = app.get_sheet_meta("S_new_common")
    assert meta == {"has_a": True, "has_b": True, "has_base": False, "view_mode": "normal"}, meta
    assert bool(getattr(view, "_sheet_structural_diff", False))
    pair_idx = _pair_for_record(view, "new-common-1")
    raw_a = tuple(view.pair_raw_parts_a[pair_idx])
    raw_b = tuple(view.pair_raw_parts_b[pair_idx])
    assert raw_a == raw_b == ("new-common-1", "new")
    assert _prepared_pair(view, pair_idx, "A") == raw_a
    assert _prepared_pair(view, pair_idx, "B") == raw_b
    assert all(not cols for cols in view.pair_diff_cols.values())
    assert all(not cols for cols in view.pair_base_diff_cols.values())
    assert view._pair_has_visual_diff(pair_idx) is False


def _assert_late_provisional_cannot_regress(app) -> None:
    before_entries = {sheet: dict(app._sheet_exact_entry(sheet)) for sheet in _SHEETS}
    before_states = {sheet: app.sheet_diff_state.get(sheet) for sheet in _SHEETS}
    for sheet, has in (("S_same", True), ("S_base_diff", False), ("S_ab_diff", False), ("S_new_common", False)):
        app.set_sheet_has_diff(sheet, has, confirmed=False)
    for sheet in _SHEETS:
        assert dict(app._sheet_exact_entry(sheet)) == before_entries[sheet], sheet
        assert app.sheet_diff_state.get(sheet) == before_states[sheet], sheet
        assert app._is_sheet_exact_current(sheet), sheet


def _assert_effective_startup_inputs(
    app, input_paths, input_hashes, startup_ledger: set[str]
) -> frozenset[str]:
    raw_paths = {
        "base": _normalized_path(input_paths["base"]),
        "mine": _normalized_path(input_paths["mine"]),
        "theirs": _normalized_path(input_paths["theirs"]),
    }
    effective_paths = {
        "base": _normalized_path(app.base_path),
        "mine": _normalized_path(app.file_a),
        "theirs": _normalized_path(app.file_b),
    }
    assert len(set(raw_paths.values())) == 3, raw_paths
    assert len(set(effective_paths.values())) == 3, effective_paths
    for side in ("base", "mine", "theirs"):
        raw_path = raw_paths[side]
        effective_path = effective_paths[side]
        assert raw_path != effective_path, (side, raw_path, effective_path)
        assert os.path.isfile(effective_path), (side, effective_path)
        assert _sha256(Path(effective_path)) == input_hashes[side], side
    owned_effective_paths = frozenset(effective_paths.values())
    assert app._owned_startup_temp_paths is startup_ledger
    assert set(app._owned_startup_temp_paths) == owned_effective_paths
    return owned_effective_paths


def _assert_owned_startup_cleanup_evidence(
    evidence_items, owned_effective_paths: frozenset[str]
) -> None:
    evidence_by_path = {
        _normalized_path(item["path"]): item
        for item in evidence_items
    }
    assert set(evidence_by_path) == set(owned_effective_paths), evidence_by_path
    for effective_path in owned_effective_paths:
        evidence = evidence_by_path[effective_path]
        assert evidence.get("removed") is True, evidence
        assert evidence.get("exists_after") is False, evidence
        assert not evidence.get("error"), evidence


def _assert_owned_startup_cleanup(app, owned_effective_paths: frozenset[str]) -> None:
    assert not app._owned_startup_temp_paths
    _assert_owned_startup_cleanup_evidence(
        app._owned_startup_temp_cleanup_evidence, owned_effective_paths
    )


def _shutdown_app(app) -> None:
    if app is None:
        return
    for view in tuple(getattr(app, "sheet_views", {}).values()):
        for attr in ("_settings_save_id", "_hover_debounce_id", "_diff_map_debounce_id"):
            after_id = getattr(view, attr, None)
            if not after_id:
                continue
            try:
                view.frame.after_cancel(after_id)
            except Exception:
                pass
            finally:
                setattr(view, attr, None)
    app._shutdown_root()


def _run_case() -> None:
    original_settings_path = sm._SETTINGS_PATH
    user_settings_path = Path(original_settings_path)
    user_settings_before = _path_snapshot(user_settings_path)
    temporary = tempfile.TemporaryDirectory(prefix="sow_sheet_diff_state_3way_")
    root = Path(temporary.name)
    temp_settings_path = root / "settings.json"
    merged = root / "merged.xlsx"
    input_paths: dict[str, Path] = {}
    input_hashes: dict[str, str] = {}
    startup_ledger: set[str] = set()
    owned_effective_paths: frozenset[str] = frozenset()
    app = None
    exact_state_original = None
    primary: BaseException | None = None
    cleanup_errors: list[str] = []
    try:
        deadline = time.monotonic() + 90.0
        temp_settings_path.write_text(json.dumps({"only_diff": 0}), encoding="utf-8")
        sm._SETTINGS_PATH = str(temp_settings_path)
        base = root / "base.xlsx"
        mine = root / "mine.xlsx"
        theirs = root / "theirs.xlsx"
        _make_typed_book(
            base,
            {
                "S_same": (("same-1", "same"),),
                "S_base_diff": (("base-diff-1", "old"),),
                "S_ab_diff": (("ab-diff-1", "base"),),
            },
        )
        _make_typed_book(
            mine,
            {
                "S_same": (("same-1", "same"),),
                "S_base_diff": (("base-diff-1", "new"),),
                "S_ab_diff": (("ab-diff-1", "mine"),),
                "S_new_common": (("new-common-1", "new"),),
            },
        )
        _make_typed_book(
            theirs,
            {
                "S_same": (("same-1", "same"),),
                "S_base_diff": (("base-diff-1", "new"),),
                "S_ab_diff": (("ab-diff-1", "theirs"),),
                "S_new_common": (("new-common-1", "new"),),
            },
        )
        input_paths = {"base": base, "mine": mine, "theirs": theirs}
        input_hashes = {name: _sha256(path) for name, path in input_paths.items()}
        assert not merged.exists()

        print("SHEET_DIFF_STATE_3WAY_STAGE app-created", flush=True)
        app = sm.SowMergeApp(
            str(mine),
            str(theirs),
            merge_mode=True,
            merged_path=str(merged),
            base_path=str(base),
            initial_sheet="S_same",
            startup_owned_paths=startup_ledger,
        )
        owned_effective_paths = _assert_effective_startup_inputs(
            app, input_paths, input_hashes, startup_ledger
        )
        print("SHEET_DIFF_STATE_3WAY_STAGE initial-summaries-hidden-failure", flush=True)
        new_generation = _wait_initial_summaries_with_hidden_failure(
            app, deadline=deadline
        )

        print("SHEET_DIFF_STATE_3WAY_STAGE same-full-detail", flush=True)
        same_view = _select_full_detail(app, "S_same", deadline=deadline)
        _assert_full_three_way_projection(app, same_view, "S_same")
        assert not any(bool(cols) for cols in same_view.pair_diff_cols.values())
        assert not any(bool(cols) for cols in same_view.pair_base_diff_cols.values())

        print("SHEET_DIFF_STATE_3WAY_STAGE base-common-full-detail", flush=True)
        base_view = _select_full_detail(app, "S_base_diff", deadline=deadline)
        _assert_full_three_way_projection(app, base_view, "S_base_diff")
        _assert_base_relative_common_change(app, base_view)

        print("SHEET_DIFF_STATE_3WAY_STAGE ab-conflict-full-detail", flush=True)
        ab_view = _select_full_detail(app, "S_ab_diff", deadline=deadline)
        _assert_full_three_way_projection(app, ab_view, "S_ab_diff")
        _assert_three_way_conflict_semantics(ab_view)

        new_failed_entry = app._sheet_exact_entry("S_new_common")
        assert str(getattr(app, "selected_sheet", "")) == "S_ab_diff"
        assert int(app._sheet_compute_generation["S_new_common"]) == new_generation
        assert int(new_failed_entry.get("generation", -1)) == new_generation
        assert new_failed_entry.get("state") == sm._SHEET_EXACT_FAILED
        assert not bool(new_failed_entry.get("full_detail_terminal", False))
        assert new_failed_entry.get("reason") == _HIDDEN_SNAPSHOT_FAILURE
        assert app.sheet_diff_state.get("S_new_common") == 0
        assert not bool(getattr(app, "_sheet_loaded", {}).get("S_new_common", False))
        print("SHEET_DIFF_STATE_3WAY_STAGE new-common-selected-retry", flush=True)
        exact_state_original, retry_transitions = _install_exact_state_transition_spy(
            app, "S_new_common", new_generation
        )
        new_view = _select_hidden_failure_retry_full_detail(
            app,
            "S_new_common",
            generation=new_generation,
            transitions=retry_transitions,
            deadline=deadline,
        )
        _assert_full_three_way_projection(
            app, new_view, "S_new_common", require_complete_base_fragments=False
        )
        _assert_new_common_structural(app, new_view)
        assert _assert_effective_startup_inputs(
            app, input_paths, input_hashes, startup_ledger
        ) == owned_effective_paths

        print("SHEET_DIFF_STATE_3WAY_STAGE late-provisional", flush=True)
        _assert_late_provisional_cannot_regress(app)
        _assert_nav_and_current(app)
        assert not merged.exists()
        assert time.monotonic() <= deadline, "three-way Sheet-state case exceeded 90 seconds"
        print(
            "SHEET_DIFF_STATE_3WAY_CASE_OK "
            + json.dumps(
                {
                    "case": _CASE,
                    "deadline_seconds": 90,
                    "states": dict(app.sheet_diff_state),
                    "generations": dict(app._sheet_compute_generation),
                    "structural_sheet": bool(new_view._sheet_structural_diff),
                    "merged_absent": not merged.exists(),
                },
                sort_keys=True,
            ),
            flush=True,
        )
    except BaseException as exc:
        primary = exc
    finally:
        try:
            if app is not None and exact_state_original is not None:
                app._set_sheet_exact_state = exact_state_original
        except Exception as exc:
            cleanup_errors.append(
                f"exact-state spy restore failed: {type(exc).__name__}: {exc}"
            )
        try:
            if app is not None:
                if not owned_effective_paths:
                    owned_effective_paths = frozenset(
                        _normalized_path(path) for path in startup_ledger
                    )
                _shutdown_app(app)
                _assert_owned_startup_cleanup(app, owned_effective_paths)
            elif startup_ledger:
                owned_effective_paths = frozenset(
                    _normalized_path(path) for path in startup_ledger
                )
                construction_failure_evidence: list[dict] = []
                sm._consume_owned_startup_temp_paths(
                    startup_ledger, construction_failure_evidence
                )
                assert not startup_ledger
                _assert_owned_startup_cleanup_evidence(
                    construction_failure_evidence, owned_effective_paths
                )
        except Exception as exc:
            cleanup_errors.append(f"app shutdown failed: {type(exc).__name__}: {exc}")
        try:
            for name, path in input_paths.items():
                if _sha256(path) != input_hashes[name]:
                    cleanup_errors.append(f"input hash changed: {name}")
        except Exception as exc:
            cleanup_errors.append(f"input hash audit failed: {type(exc).__name__}: {exc}")
        try:
            if merged.exists():
                cleanup_errors.append(f"merged output unexpectedly exists: {merged}")
        except Exception as exc:
            cleanup_errors.append(f"merged output audit failed: {type(exc).__name__}: {exc}")
        try:
            sm._SETTINGS_PATH = original_settings_path
            if _path_snapshot(user_settings_path) != user_settings_before:
                cleanup_errors.append("user settings changed")
        except Exception as exc:
            cleanup_errors.append(f"settings restore failed: {type(exc).__name__}: {exc}")
        try:
            temporary.cleanup()
            if os.path.lexists(root):
                cleanup_errors.append(f"owned temporary root retained: {root}")
        except Exception as exc:
            cleanup_errors.append(f"owned temporary cleanup failed: {type(exc).__name__}: {exc}")
    if primary is not None:
        for detail in cleanup_errors:
            try:
                primary.add_note(detail)
            except Exception:
                pass
        raise primary
    if cleanup_errors:
        raise AssertionError("; ".join(cleanup_errors))


def main(argv=None) -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--list-cases", action="store_true")
    parser.add_argument("--case", choices=(_CASE,))
    args = parser.parse_args(argv)
    if args.list_cases:
        if args.case:
            parser.error("--list-cases cannot be combined with --case")
        print(_CASE, flush=True)
        return
    selected = (args.case,) if args.case else (_CASE,)
    for _case in selected:
        _run_case()
    print(f"GUI_SELF_TEST_SHEET_DIFF_STATE_3WAY_OK ({len(selected)} cases)")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"GUI_SELF_TEST_SHEET_DIFF_STATE_3WAY_FAIL: {exc}", file=sys.stderr)
        raise
