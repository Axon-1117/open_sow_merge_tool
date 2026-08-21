"""GUI regression for current-generation Sheet diff and exact state badges."""
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


_CASE = "sheet-diff-state"
_SHEETS = ("S_ok", "S_diff", "S_diff2")
_EXPECTED_EXACT = {
    "S_ok": sm._SHEET_EXACT_SAME,
    "S_diff": sm._SHEET_EXACT_CHANGED,
    "S_diff2": sm._SHEET_EXACT_CHANGED,
}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _path_snapshot(path: Path) -> tuple[bool, bytes | None]:
    return (True, path.read_bytes()) if path.exists() else (False, None)


def _make_typed_book(path: Path, *, changed_rows: dict[str, int]) -> None:
    workbook = Workbook()
    try:
        workbook.remove(workbook.active)
        for sheet in _SHEETS:
            worksheet = workbook.create_sheet(sheet)
            worksheet.append(["id@id", "value"])
            worksheet.append(["string", "string"])
            for row in range(1, 81):
                value = f"{sheet}-value-{row}"
                if changed_rows.get(sheet) == row:
                    value += "-changed"
                worksheet.append([f"{sheet}-{row}", value])
        workbook.save(path)
    finally:
        workbook.close()


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


def _terminal_diagnostic(app) -> dict:
    return {
        "selected_sheet": getattr(app, "selected_sheet", None),
        "generations": dict(getattr(app, "_sheet_compute_generation", {}) or {}),
        "sheet_diff_state": dict(getattr(app, "sheet_diff_state", {}) or {}),
        "entries": {sheet: dict(app._sheet_exact_entry(sheet)) for sheet in _SHEETS},
        "loaded": dict(getattr(app, "_sheet_loaded", {}) or {}),
    }


def _wait_current_terminals(app, *, deadline: float):
    while time.monotonic() < deadline:
        app.root.update_idletasks()
        app.root.update()
        selected_entry = app._sheet_exact_entry("S_ok")
        selected_view = app.sheet_views.get("S_ok")
        if (
            all(
                app._is_sheet_exact_current(sheet)
                and app._sheet_exact_entry(sheet).get("state") == _EXPECTED_EXACT[sheet]
                and int(app._sheet_exact_entry(sheet).get("generation", -1))
                == int(app._sheet_compute_generation[sheet])
                for sheet in _SHEETS
            )
            and bool(selected_entry.get("full_detail_terminal", False))
            and selected_view is not None
            and bool(getattr(selected_view, "_prepared_complete", False))
            and bool(getattr(selected_view, "_data_ready", False))
        ):
            return
        time.sleep(0.01)
    raise AssertionError("exact terminal wait timed out: " + json.dumps(_terminal_diagnostic(app), default=str, sort_keys=True))


def _assert_current_badges(app) -> None:
    expected_diff = {"S_ok": 0, "S_diff": 2, "S_diff2": 2}
    expected_badge = {
        "S_ok": "精确相同",
        "S_diff": "精确变更",
        "S_diff2": "精确变更",
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
        assert text.startswith(f"{sheet} · ") and expected_badge[sheet] in text, text


def _assert_late_provisional_cannot_regress(app) -> None:
    before_entries = {sheet: dict(app._sheet_exact_entry(sheet)) for sheet in _SHEETS}
    before_states = {sheet: app.sheet_diff_state.get(sheet) for sheet in _SHEETS}
    for sheet, has in (("S_ok", True), ("S_diff", False), ("S_diff2", False)):
        app.set_sheet_has_diff(sheet, has, confirmed=False)
    for sheet in _SHEETS:
        assert dict(app._sheet_exact_entry(sheet)) == before_entries[sheet], sheet
        assert app.sheet_diff_state.get(sheet) == before_states[sheet], sheet
        assert app._is_sheet_exact_current(sheet), sheet


def _run_case() -> None:
    original_settings_path = sm._SETTINGS_PATH
    user_settings_path = Path(original_settings_path)
    user_settings_before = _path_snapshot(user_settings_path)
    temporary = tempfile.TemporaryDirectory(prefix="sow_sheet_diff_state_")
    root = Path(temporary.name)
    temp_settings_path = root / "settings.json"
    input_paths: dict[str, Path] = {}
    input_hashes: dict[str, str] = {}
    app = None
    primary: BaseException | None = None
    cleanup_errors: list[str] = []
    try:
        temp_settings_path.write_text(json.dumps({"only_diff": 0}), encoding="utf-8")
        sm._SETTINGS_PATH = str(temp_settings_path)
        mine = root / "mine.xlsx"
        theirs = root / "theirs.xlsx"
        _make_typed_book(mine, changed_rows={})
        _make_typed_book(theirs, changed_rows={"S_diff": 20, "S_diff2": 70})
        input_paths = {"mine": mine, "theirs": theirs}
        input_hashes = {name: _sha256(path) for name, path in input_paths.items()}

        fingerprints_mine = sm._xlsx_sheet_part_fingerprints(str(mine))
        fingerprints_theirs = sm._xlsx_sheet_part_fingerprints(str(theirs))
        assert fingerprints_mine["S_ok"] == fingerprints_theirs["S_ok"]
        assert fingerprints_mine["S_diff"] != fingerprints_theirs["S_diff"]
        assert fingerprints_mine["S_diff2"] != fingerprints_theirs["S_diff2"]

        deadline = time.monotonic() + 90.0
        print("SHEET_DIFF_STATE_STAGE app-created", flush=True)
        app = sm.SowMergeApp(str(mine), str(theirs), initial_sheet="S_ok")
        print("SHEET_DIFF_STATE_STAGE exact-terminals", flush=True)
        _wait_current_terminals(app, deadline=deadline)
        _assert_current_badges(app)
        print("SHEET_DIFF_STATE_STAGE late-provisional", flush=True)
        _assert_late_provisional_cannot_regress(app)
        assert time.monotonic() <= deadline, "sheet diff state exceeded 90 seconds"
    except BaseException as exc:
        primary = exc
    finally:
        try:
            _shutdown_app(app)
        except Exception as exc:
            cleanup_errors.append(f"app shutdown failed: {type(exc).__name__}: {exc}")
        try:
            for name, path in input_paths.items():
                if _sha256(path) != input_hashes[name]:
                    cleanup_errors.append(f"input hash changed: {name}")
        except Exception as exc:
            cleanup_errors.append(f"input hash audit failed: {type(exc).__name__}: {exc}")
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
    print(f"GUI_SELF_TEST_SHEET_DIFF_STATE_OK ({len(selected)} cases)")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"GUI_SELF_TEST_SHEET_DIFF_STATE_FAIL: {exc}", file=sys.stderr)
        raise
