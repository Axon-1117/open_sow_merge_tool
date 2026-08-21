"""GUI regression: exact Sheet states gate every operation in 2-way and 3-way."""

import argparse
import hashlib
import json
import os
import sys
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

from openpyxl import Workbook

import sow_merge_tool as sm


_CASES = ("two-way", "three-way")


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _path_snapshot(path: Path) -> tuple[bool, bytes | None]:
    if path.exists():
        return True, path.read_bytes()
    return False, None


def _shutdown_app(app) -> None:
    if app is None:
        return
    for view in tuple(getattr(app, "sheet_views", {}).values()):
        after_id = getattr(view, "_settings_save_id", None)
        if after_id:
            try:
                view.frame.after_cancel(after_id)
            finally:
                view._settings_save_id = None
    app._shutdown_root()


def _operation_snapshot(app, sheet: str) -> tuple:
    return (
        list(app.undo_stack),
        list(app.redo_stack),
        dict(app.manual_a_cell_ops),
        dict(app.manual_b_cell_ops),
        dict(app.sheet_operation_overlay(sheet).cells),
    )


def _make_book(path, suffix, *, changed=False):
    wb = Workbook()
    for sheet in ("S1", "S2"):
        ws = wb.active if sheet == "S1" else wb.create_sheet(sheet)
        ws.title = sheet
        ws.append(["id @id", "value"])
        for row in range(1, 81):
            value = f"{sheet}-{suffix}-{row}"
            if changed and sheet == "S1" and row == 40:
                value += "-changed"
            ws.append([row, value])
    wb.save(path)
    wb.close()


def _pump(app, deadline):
    while time.monotonic() < deadline:
        app.root.update_idletasks()
        app.root.update()
        time.sleep(0.01)


def _wait_all_sheet_summaries(app, sheets, *, deadline: float):
    while time.monotonic() < deadline:
        _pump(app, min(deadline, time.monotonic() + 0.08))
        if all(app._is_sheet_exact_current(sheet) for sheet in sheets):
            return
    states = {sheet: app._sheet_exact_entry(sheet) for sheet in sheets}
    raise AssertionError(f"exact summary scan did not finish: {states}")


def _wait_selected_full_detail(app, sheet: str, *, deadline: float):
    sheet = str(sheet)
    while time.monotonic() < deadline:
        _pump(app, min(deadline, time.monotonic() + 0.08))
        entry = app._sheet_exact_entry(sheet)
        view = app.sheet_views.get(sheet)
        selected_tab = app.nb.tab(app.nb.select(), "text")
        if (
            selected_tab == sheet
            and str(getattr(app, "selected_sheet", "")) == sheet
            and app._is_sheet_exact_current(sheet)
            and bool(entry.get("full_detail_terminal", False))
            and view is not None
            and bool(getattr(view, "_prepared_complete", False))
            and bool(getattr(view, "_data_ready", False))
            and not bool(getattr(view, "_pending_exact_render", False))
        ):
            return view
    raise AssertionError(
        "selected full detail did not finish: "
        + repr(
            {
                "sheet": sheet,
                "selected_sheet": getattr(app, "selected_sheet", None),
                "selected_tab": app.nb.tab(app.nb.select(), "text"),
                "entry": app._sheet_exact_entry(sheet),
                "view": {
                    "exists": app.sheet_views.get(sheet) is not None,
                    "prepared": bool(
                        getattr(app.sheet_views.get(sheet), "_prepared_complete", False)
                    ),
                    "data_ready": bool(
                        getattr(app.sheet_views.get(sheet), "_data_ready", False)
                    ),
                    "pending_exact_render": bool(
                        getattr(app.sheet_views.get(sheet), "_pending_exact_render", False)
                    ),
                },
            }
        )
    )


def _assert_full_current_view(app, view, sheet: str, *, expected_state: str | None = None):
    entry = app._sheet_exact_entry(sheet)
    assert int(entry["generation"]) == int(app._sheet_compute_generation[sheet]), entry
    assert app._is_sheet_exact_current(sheet), entry
    if expected_state is not None:
        assert entry["state"] == expected_state, entry
    assert bool(entry["full_detail_terminal"]), entry
    assert bool(getattr(view, "_prepared_complete", False))
    assert bool(getattr(view, "_data_ready", False))
    assert not bool(getattr(view, "_pending_exact_render", False))


def _assert_hidden_exact_summary(app, sheet: str, *, expected_state: str):
    entry = app._sheet_exact_entry(sheet)
    assert entry["state"] == expected_state, entry
    assert app._is_sheet_exact_current(sheet), entry
    assert not bool(entry.get("full_detail_terminal", False)), entry
    assert app.sheet_views.get(sheet) is None, app.sheet_views.get(sheet)
    assert _operation_snapshot(app, sheet) == ([], [], {}, {}, {})


@contextmanager
def _forbid_view_only_upgrade_paths(app, view):
    """Make the real tab promotion fail immediately on any edit/read/write path."""
    blocked = []
    originals = []

    def _blocked(label):
        def _raise(*_args, **_kwargs):
            blocked.append(label)
            raise AssertionError(f"view-only hidden-sheet upgrade called {label}")

        return _raise

    targets = (
        (app, "_request_edit_preload"),
        (view, "_request_edit_preload"),
        (app, "_ensure_edit_loaded"),
        (view, "_ensure_edit_loaded"),
        (app, "ws_a_val"),
        (app, "ws_b_val"),
        (app, "ws_base_val"),
        (app, "ws_a_edit"),
        (app, "ws_b_edit"),
        (app, "ws_base_edit"),
        (app, "_atomic_save"),
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
            if not hasattr(owner, name):
                continue
            originals.append((owner, name, getattr(owner, name)))
            setattr(owner, name, _blocked(f"{type(owner).__name__}.{name}"))
        yield blocked
    finally:
        for owner, name, original in reversed(originals):
            setattr(owner, name, original)


def _select_hidden_sheet_for_full_detail(
    app, sheet: str, *, deadline: float, input_paths, hashes_before
):
    assert app.sheet_views.get(sheet) is None
    assert not bool(app._sheet_loaded.get(sheet, False))
    before_ops = _operation_snapshot(app, sheet)
    current_view = app.sheet_views[str(getattr(app, "selected_sheet", ""))]
    print(f"EXACT_READINESS_STAGE select-hidden-full-detail sheet={sheet}", flush=True)
    with _forbid_view_only_upgrade_paths(app, current_view) as blocked:
        app.nb.select(app._sheet_containers[sheet])
        view = _wait_selected_full_detail(app, sheet, deadline=deadline)
    assert not blocked, blocked
    assert _operation_snapshot(app, sheet) == before_ops
    hashes_after = {name: _sha256(path) for name, path in input_paths.items()}
    assert hashes_after == hashes_before, (hashes_before, hashes_after)
    return view


def _assert_rejected_without_mutation(app, view, sheet):
    notices = []
    original = sm.messagebox.showwarning
    sm.messagebox.showwarning = lambda *args, **kwargs: notices.append((args, kwargs))
    try:
        before = (
            list(app.undo_stack), list(app.redo_stack),
            dict(app.manual_a_cell_ops), dict(app.manual_b_cell_ops),
            dict(app.sheet_operation_overlay(sheet).cells),
        )
        assert not view._guard_mutation_ready("测试覆盖")
        # B3 verifies the production non-ready modal and zero mutation without
        # introducing action, undo, or save workflows into this view-only run.
        after = (
            list(app.undo_stack), list(app.redo_stack),
            dict(app.manual_a_cell_ops), dict(app.manual_b_cell_ops),
            dict(app.sheet_operation_overlay(sheet).cells),
        )
        assert after == before, (before, after)
        assert notices, "rejected action did not open readiness modal"
        text = "\n".join(str(part) for call in notices for part in call[0])
        assert sheet in text and "状态" in text and "进度" in text and "重试" in text, text
    finally:
        sm.messagebox.showwarning = original


def _run_two_way(tmp, *, deadline: float):
    mine = os.path.join(tmp, "two-mine.xlsx")
    theirs = os.path.join(tmp, "two-theirs.xlsx")
    _make_book(mine, "same")
    _make_book(theirs, "same", changed=True)
    input_paths = {"mine": mine, "theirs": theirs}
    before_hashes = {name: _sha256(path) for name, path in input_paths.items()}
    app = None
    try:
        app = sm.SowMergeApp(mine, theirs, initial_sheet="S1")
        view = app.sheet_views["S1"]
        assert _operation_snapshot(app, "S1") == ([], [], {}, {}, {})

        print("EXACT_READINESS_STAGE initial-selected-full-detail sheet=S1", flush=True)
        view = _wait_selected_full_detail(app, "S1", deadline=deadline)
        _assert_full_current_view(
            app, view, "S1", expected_state=sm._SHEET_EXACT_CHANGED
        )
        assert "正在加载并计算差异" not in view.left.get("1.0", "end")

        print("EXACT_READINESS_STAGE all-sheet-summary", flush=True)
        _wait_all_sheet_summaries(app, ("S1", "S2"), deadline=deadline)
        _assert_full_current_view(
            app, view, "S1", expected_state=sm._SHEET_EXACT_CHANGED
        )
        _assert_hidden_exact_summary(
            app, "S2", expected_state=sm._SHEET_EXACT_SAME
        )

        # An old worker result cannot republish an earlier generation as ready.
        old_generation = app._sheet_exact_entry("S1")["generation"]
        with app._compute_lock:
            app._sheet_compute_generation["S1"] += 1
        assert not app._set_sheet_exact_state(
            "S1", sm._SHEET_EXACT_CHANGED, generation=old_generation
        )
        assert not app._is_sheet_exact_current("S1")
        _assert_rejected_without_mutation(app, view, "S1")
        current_generation = app._sheet_compute_generation["S1"]
        assert app._set_sheet_exact_state(
            "S1",
            sm._SHEET_EXACT_FAILED,
            generation=current_generation,
            stage="测试失败态",
            reason="测试故障",
        )
        status_text = app.exact_status_var.get()
        assert "未解决/失败" in status_text and "正在计算 S1" not in status_text, status_text
        assert app._set_sheet_exact_state(
            "S1",
            sm._SHEET_EXACT_UNRESOLVED,
            generation=current_generation,
            stage="测试歧义映射",
            reason="逻辑列映射仍有歧义",
        )
        view._show_exact_unavailable("未解决：逻辑列映射仍有歧义")
        assert not app._is_sheet_exact_current("S1")
        assert "无法发布可操作的精确结果" in view.left.get("1.0", "end")
        assert view._derive_lifecycle_state() == "UNRESOLVED"

        s2_view = _select_hidden_sheet_for_full_detail(
            app,
            "S2",
            deadline=deadline,
            input_paths=input_paths,
            hashes_before=before_hashes,
        )
        _assert_full_current_view(
            app, s2_view, "S2", expected_state=sm._SHEET_EXACT_SAME
        )
    finally:
        _shutdown_app(app)
        after_hashes = {name: _sha256(path) for name, path in input_paths.items()}
        assert after_hashes == before_hashes, (before_hashes, after_hashes)


def _assert_three_way_full_projection(app, view, sheet: str):
    assert bool(getattr(app, "has_base", False))
    assert view._is_three_way_enabled()
    assert bool(getattr(view, "_pair_diff_full_exact", False))
    assert bool(getattr(view, "_base_diff_full_exact", False))
    assert view._has_complete_prepared_rows()
    assert view._column_mapping_is_current()
    mine_to_base = dict(getattr(view, "mine_to_base_row", {}) or {})
    theirs_to_base = dict(getattr(view, "theirs_to_base_row", {}) or {})
    overrides = dict(getattr(view, "pair_base_row_override", {}) or {})
    assert mine_to_base and theirs_to_base
    assert all(isinstance(row, int) and row > 0 for row in mine_to_base.values())
    assert all(isinstance(row, int) and row > 0 for row in theirs_to_base.values())
    assert all(isinstance(pair, int) and pair >= 0 for pair in overrides)
    assert all(row is None or (isinstance(row, int) and row > 0) for row in overrides.values())


def _run_three_way(tmp, *, deadline: float):
    mine = os.path.join(tmp, "three-mine.xlsx")
    base = os.path.join(tmp, "three-base.xlsx")
    theirs = os.path.join(tmp, "three-theirs.xlsx")
    _make_book(base, "same")
    _make_book(mine, "same", changed=True)
    _make_book(theirs, "same", changed=True)
    input_paths = {"mine": mine, "base": base, "theirs": theirs}
    before_hashes = {name: _sha256(path) for name, path in input_paths.items()}
    app = None
    try:
        app = sm.SowMergeApp(
            mine,
            theirs,
            merge_mode=True,
            merged_path=os.path.join(tmp, "merged.xlsx"),
            base_path=base,
            initial_sheet="S1",
        )
        view = app.sheet_views["S1"]
        assert _operation_snapshot(app, "S1") == ([], [], {}, {}, {})

        print("EXACT_READINESS_STAGE initial-selected-full-detail sheet=S1", flush=True)
        view = _wait_selected_full_detail(app, "S1", deadline=deadline)
        _assert_full_current_view(app, view, "S1")
        _assert_three_way_full_projection(app, view, "S1")

        print("EXACT_READINESS_STAGE all-sheet-summary", flush=True)
        _wait_all_sheet_summaries(app, ("S1", "S2"), deadline=deadline)
        _assert_full_current_view(app, view, "S1")
        _assert_three_way_full_projection(app, view, "S1")
        _assert_hidden_exact_summary(
            app, "S2", expected_state=sm._SHEET_EXACT_SAME
        )

        s2_view = _select_hidden_sheet_for_full_detail(
            app,
            "S2",
            deadline=deadline,
            input_paths=input_paths,
            hashes_before=before_hashes,
        )
        _assert_full_current_view(
            app, s2_view, "S2", expected_state=sm._SHEET_EXACT_SAME
        )
        _assert_three_way_full_projection(app, s2_view, "S2")
    finally:
        _shutdown_app(app)
        after_hashes = {name: _sha256(path) for name, path in input_paths.items()}
        assert after_hashes == before_hashes, (before_hashes, after_hashes)


def _run_case(case_name: str) -> None:
    original_settings_path = sm._SETTINGS_PATH
    user_settings_path = Path(original_settings_path)
    user_settings_before = _path_snapshot(user_settings_path)
    environment_before = {
        key: value for key, value in os.environ.items() if key.startswith("SOW_")
    }
    root_path = None
    try:
        with tempfile.TemporaryDirectory(prefix=f"sow_exact_readiness_{case_name}_") as tmp:
            root_path = Path(tmp)
            temp_settings_path = root_path / "settings.json"
            temp_settings_path.write_text(json.dumps({"only_diff": 0}), encoding="utf-8")
            sm._SETTINGS_PATH = str(temp_settings_path)
            deadline = time.monotonic() + 40.0
            print(f"EXACT_READINESS_CASE_START {case_name}", flush=True)
            if case_name == "two-way":
                _run_two_way(tmp, deadline=deadline)
            else:
                _run_three_way(tmp, deadline=deadline)
            assert time.monotonic() <= deadline, f"{case_name} exceeded 40s"
            assert _path_snapshot(user_settings_path) == user_settings_before
            environment_after = {
                key: value for key, value in os.environ.items() if key.startswith("SOW_")
            }
            assert environment_after == environment_before, (environment_before, environment_after)
            print(
                "EXACT_READINESS_CASE_OK "
                + json.dumps(
                    {
                        "case": case_name,
                        "deadline_seconds": 40,
                        "settings_path_is_temp": str(sm._SETTINGS_PATH)
                        == str(temp_settings_path),
                        "user_settings_unchanged": _path_snapshot(user_settings_path)
                        == user_settings_before,
                        "sow_environment_unchanged": environment_after == environment_before,
                        "monkeypatches_restored": True,
                    },
                    sort_keys=True,
                ),
                flush=True,
            )
    finally:
        if root_path is not None:
            assert str(sm._SETTINGS_PATH) == str(root_path / "settings.json")
        sm._SETTINGS_PATH = original_settings_path
        assert _path_snapshot(user_settings_path) == user_settings_before
        if root_path is not None:
            assert not root_path.exists(), root_path


def main(argv=None):
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
    print(f"GUI_SELF_TEST_EXACT_SHEET_READINESS_OK ({len(selected)} cases)")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"GUI_SELF_TEST_EXACT_SHEET_READINESS_FAIL: {exc}", file=sys.stderr)
        raise
