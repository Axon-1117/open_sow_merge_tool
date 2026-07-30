"""OpenSpec 4.1 regressions for SVN launch roles and production dispatch.

These tests deliberately exercise both the pure launch-classification API and
``main()``.  The latter is important: keeping the raw paths in a data model is
not sufficient if the production launch path still replaces source Base with
the target working copy's pristine file before scanning or opening the UI.
"""

from __future__ import annotations

import inspect
import os
import sys
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


def _make_book(path: str, marker: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "marker"
    worksheet["A2"] = marker
    workbook.save(path)
    workbook.close()


def _scenario_name(value) -> str:
    if hasattr(value, "name"):
        value = value.name
    elif hasattr(value, "value"):
        value = value.value
    return str(value).strip().replace("-", "_").replace(" ", "_").upper()


def _path_of(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, (str, os.PathLike)):
        return os.path.abspath(os.fspath(value))
    if isinstance(value, dict):
        for key in ("path", "source_path", "raw_path", "file_path"):
            if value.get(key):
                return os.path.abspath(os.fspath(value[key]))
        return None
    for name in ("path", "source_path", "raw_path", "file_path"):
        candidate = getattr(value, name, None)
        if candidate:
            return os.path.abspath(os.fspath(candidate))
    return None


def _context_value(context, *names):
    if isinstance(context, dict):
        for name in names:
            if name in context:
                return context[name]
    for name in names:
        if hasattr(context, name):
            return getattr(context, name)
    return None


def _context_identity_path(context, role: str) -> str | None:
    aliases = {
        "base": (
            "source_base",
            "source_base_identity",
            "base",
            "base_identity",
            "source_base_path",
            "base_path",
        ),
        "mine": ("mine", "mine_identity", "mine_working", "mine_path"),
        "theirs": (
            "theirs",
            "theirs_identity",
            "theirs_incoming",
            "theirs_path",
        ),
        "target_pristine": (
            "target_pristine",
            "target_pristine_identity",
            "target_wc_pristine",
            "target_pristine_path",
        ),
    }
    value = _context_value(context, *aliases[role])
    direct = _path_of(value)
    if direct:
        return direct
    identities = _context_value(context, "identities", "versions")
    if isinstance(identities, dict):
        return _path_of(identities.get(role))
    return None


def _build_context(base, mine, theirs, merged, target_pristine):
    builder = getattr(smt, "build_merge_launch_context", None)
    assert callable(builder), "missing build_merge_launch_context API"
    kwargs = {
        "base_path": base,
        "source_base_path": base,
        "mine_path": mine,
        "theirs_path": theirs,
        "merged_path": merged,
        "target_pristine_path": target_pristine,
    }
    signature = inspect.signature(builder)
    supported = {
        name: value
        for name, value in kwargs.items()
        if name in signature.parameters
    }
    try:
        return builder(**supported)
    except TypeError:
        # The agreed API has the first four positional parameters; the
        # optional pristine identity may be positional or keyword-only.
        try:
            return builder(
                base,
                mine,
                theirs,
                merged,
                target_pristine_path=target_pristine,
            )
        except TypeError:
            return builder(base, mine, theirs, merged, target_pristine)


def test_four_launch_scenarios_and_identities() -> None:
    classifier = getattr(smt, "classify_merge_launch", None)
    assert callable(classifier), "missing classify_merge_launch API"

    root = make_temp_dir("sow_role_semantics_")
    mine = os.path.join(root, "Design.xlsx")
    merged = mine
    pristine = os.path.join(root, "Design.target-pristine.xlsx")
    update_base = os.path.join(root, "Design.xlsx.r101")
    update_theirs = os.path.join(root, "Design.xlsx.r109")
    legacy_update_base = os.path.join(root, "Legacy.xlsx.rOLD")
    legacy_update_theirs = os.path.join(root, "Legacy.xlsx.rNEW")
    branch_base = os.path.join(root, "Design.xlsx.merge-left.r210")
    branch_theirs = os.path.join(root, "Design.xlsx.merge-right.r240")
    unknown_base = os.path.join(root, "old-copy.xlsx")
    unknown_theirs = os.path.join(root, "incoming-copy.xlsx")

    for path, marker in (
        (mine, "mine"),
        (pristine, "target-pristine"),
        (update_base, "update-base"),
        (update_theirs, "update-theirs"),
        (legacy_update_base, "legacy-update-base"),
        (legacy_update_theirs, "legacy-update-theirs"),
        (branch_base, "branch-left"),
        (branch_theirs, "branch-right"),
        (unknown_base, "unknown-base"),
        (unknown_theirs, "unknown-theirs"),
    ):
        _make_book(path, marker)

    cases = (
        ((mine, None, None, None), "TWO_WAY"),
        ((update_base, mine, update_theirs, merged), "UPDATE_CONFLICT"),
        ((legacy_update_base, mine, legacy_update_theirs, merged), "UPDATE_CONFLICT"),
        ((branch_base, mine, branch_theirs, merged), "CROSS_BRANCH_MERGE"),
        ((unknown_base, mine, unknown_theirs, merged), "UNKNOWN_THREE_WAY"),
    )
    for arguments, expected in cases:
        classifier_args = arguments[: len(inspect.signature(classifier).parameters)]
        scenario = classifier(*classifier_args)
        assert _scenario_name(scenario) == expected, (
            arguments,
            scenario,
            expected,
        )

    for base, theirs, expected in (
        (update_base, update_theirs, "UPDATE_CONFLICT"),
        (branch_base, branch_theirs, "CROSS_BRANCH_MERGE"),
    ):
        context = _build_context(base, mine, theirs, merged, pristine)
        assert _scenario_name(_context_value(context, "scenario", "merge_scenario")) == expected
        assert _context_identity_path(context, "base") == os.path.abspath(base)
        assert _context_identity_path(context, "mine") == os.path.abspath(mine)
        assert _context_identity_path(context, "theirs") == os.path.abspath(theirs)
        assert _context_identity_path(context, "target_pristine") == os.path.abspath(pristine)
        assert _context_identity_path(context, "base") != _context_identity_path(
            context, "target_pristine"
        )


def _run_production_launch(case_name: str):
    root = make_temp_dir(f"sow_production_roles_{case_name}_")
    target = os.path.join(root, "Design.xlsx")
    pristine = os.path.join(root, "Design.target-pristine.xlsx")
    merged = target
    if case_name == "update":
        base = os.path.join(root, "Design.xlsx.r301")
        theirs = os.path.join(root, "Design.xlsx.r305")
        expected_scenario = "UPDATE_CONFLICT"
    else:
        base = os.path.join(root, "Design.xlsx.merge-left.r401")
        theirs = os.path.join(root, "Design.xlsx.merge-right.r409")
        expected_scenario = "CROSS_BRANCH_MERGE"
    for path, marker in (
        (base, f"{case_name}-raw-base"),
        (target, f"{case_name}-mine"),
        (theirs, f"{case_name}-theirs"),
        (pristine, f"{case_name}-target-pristine"),
    ):
        _make_book(path, marker)

    captured = SimpleNamespace(app_args=None, app_kwargs=None, scan_args=None, logs=[])

    class _FakeApp:
        def __init__(self, *args, **kwargs):
            captured.app_args = args
            captured.app_kwargs = kwargs

        def show_nonblocking_notice(self, *_args, **_kwargs):
            return None

        def run(self):
            return None

    def _scan(scan_base, scan_mine, scan_theirs):
        captured.scan_args = (
            os.path.abspath(scan_base),
            os.path.abspath(scan_mine),
            os.path.abspath(scan_theirs),
        )
        return [], {}

    def _startup(_title, _message, worker):
        return worker(lambda *_args, **_kwargs: None)

    argv = [
        "sow_merge_tool.py",
        "/base",
        base,
        "/mine",
        target,
        "/theirs",
        theirs,
        "/merged",
        merged,
    ]
    with ExitStack() as stack:
        stack.enter_context(patch.object(sys, "argv", argv))
        stack.enter_context(patch.object(smt, "SowMergeApp", _FakeApp))
        stack.enter_context(patch.object(smt, "_run_startup_progress_task", _startup))
        stack.enter_context(patch.object(smt, "_ensure_xlsx_copy", lambda path: path))
        stack.enter_context(
            patch.object(
                smt,
                "_try_export_svn_revision_from_merge_temp",
                lambda path: path,
            )
        )
        stack.enter_context(
            patch.object(
                smt,
                "_try_export_svn_base_from_working_copy",
                lambda _path: pristine,
            )
        )
        if hasattr(smt, "_copy_wc_pristine_local"):
            stack.enter_context(
                patch.object(smt, "_copy_wc_pristine_local", lambda _path: pristine)
            )
        stack.enter_context(patch.object(smt, "_scan_three_way_conflicts", _scan))
        stack.enter_context(
            patch.object(smt, "_dlog", lambda message: captured.logs.append(str(message)))
        )
        stack.enter_context(
            patch.object(
                smt,
                "_trace_launch",
                lambda message: captured.logs.append(str(message)),
            )
        )
        try:
            smt.main()
        except SystemExit as exc:
            assert exc.code in (None, 0), exc.code

    assert captured.app_kwargs is not None, "production launch did not construct SowMergeApp"
    if captured.scan_args is not None:
        assert captured.scan_args == (
            os.path.abspath(base),
            os.path.abspath(target),
            os.path.abspath(theirs),
        ), (
            "production scanner must receive original raw Base/Mine/Theirs",
            captured.scan_args,
        )

    app_kwargs = captured.app_kwargs
    assert os.path.abspath(app_kwargs.get("base_path")) == os.path.abspath(base), (
        "production launch replaced raw source Base",
        app_kwargs,
    )
    assert os.path.abspath(app_kwargs.get("raw_base")) == os.path.abspath(base), (
        "raw Base diagnostic identity was rewritten",
        app_kwargs,
    )
    context = app_kwargs.get("launch_context")
    assert context is not None, "production UI did not receive MergeLaunchContext"
    assert _scenario_name(_context_value(context, "scenario", "merge_scenario")) == expected_scenario
    assert _context_identity_path(context, "base") == os.path.abspath(base)
    assert _context_identity_path(context, "target_pristine") == os.path.abspath(pristine)
    assert _context_identity_path(context, "base") != os.path.abspath(pristine)
    return captured


def test_update_production_path_preserves_raw_old_base() -> None:
    _run_production_launch("update")


def test_branch_production_path_preserves_merge_left() -> None:
    _run_production_launch("branch")


def test_two_way_sidecar_is_normalized_without_rewriting_raw_identity() -> None:
    root = make_temp_dir("sow_two_way_sidecar_")
    base = os.path.join(root, "Design.xlsx.r301")
    mine = os.path.join(root, "Design.xlsx")
    _make_book(base, "base")
    _make_book(mine, "mine")
    captured = SimpleNamespace(app_args=None, app_kwargs=None)

    class _FakeApp:
        def __init__(self, *args, **kwargs):
            captured.app_args = args
            captured.app_kwargs = kwargs

        def run(self):
            return None

    argv = [
        "sow_merge_tool.py",
        "/base",
        base,
        "/mine",
        mine,
    ]
    with patch.object(sys, "argv", argv), patch.object(smt, "SowMergeApp", _FakeApp):
        smt.main()

    assert captured.app_args is not None
    ui_base, ui_mine = captured.app_args[:2]
    assert os.path.isfile(ui_base) and ui_base.lower().endswith(".xlsx"), ui_base
    assert os.path.abspath(ui_base) != os.path.abspath(base), (
        "raw .rN sidecar was not normalized to an openpyxl-compatible copy",
        ui_base,
    )
    assert os.path.abspath(ui_mine) == os.path.abspath(mine)
    assert os.path.abspath(captured.app_kwargs["raw_base"]) == os.path.abspath(base)
    assert os.path.abspath(captured.app_kwargs["raw_mine"]) == os.path.abspath(mine)


def test_auto_detected_conflict_preserves_raw_sidecar_identities() -> None:
    root = make_temp_dir("sow_auto_detected_raw_roles_")
    base = os.path.join(root, "Design.xlsx.merge-left.r401")
    mine = os.path.join(root, "Design.xlsx")
    theirs = os.path.join(root, "Design.xlsx.merge-right.r409")
    for path, marker in ((base, "base"), (mine, "mine"), (theirs, "theirs")):
        _make_book(path, marker)
    captured = SimpleNamespace(app_kwargs=None)

    class _FakeApp:
        def __init__(self, *_args, **kwargs):
            captured.app_kwargs = kwargs

        def show_startup_outcome_dialog_deferred(self):
            return None

        def run(self):
            return None

    def _analysis(context):
        return smt.StartupMergeAnalysis(
            context,
            smt.EquivalenceMatrix(),
            smt.StartupMergeOutcome(),
            [],
            {},
        )

    selection = ("merge", base, mine, theirs, mine, True)
    with (
        patch.object(sys, "argv", ["sow_merge_tool.py"]),
        patch.object(smt, "pick_files_or_conflict", lambda: selection),
        patch.object(smt, "run_startup_merge_analysis", _analysis),
        patch.object(
            smt,
            "_run_startup_progress_task",
            lambda _title, _message, worker: worker(lambda *_args, **_kwargs: None),
        ),
        patch.object(smt, "SowMergeApp", _FakeApp),
    ):
        try:
            smt.main()
        except SystemExit as exc:
            assert exc.code in (None, 0), exc.code

    context = captured.app_kwargs["launch_context"]
    assert _scenario_name(context.scenario) == "CROSS_BRANCH_MERGE"
    assert _context_identity_path(context, "base") == os.path.abspath(base)
    assert _context_identity_path(context, "mine") == os.path.abspath(mine)
    assert _context_identity_path(context, "theirs") == os.path.abspath(theirs)
    assert os.path.abspath(captured.app_kwargs["raw_base"]) == os.path.abspath(base)


def main() -> None:
    tests = (
        test_four_launch_scenarios_and_identities,
        test_update_production_path_preserves_raw_old_base,
        test_branch_production_path_preserves_merge_left,
        test_two_way_sidecar_is_normalized_without_rewriting_raw_identity,
        test_auto_detected_conflict_preserves_raw_sidecar_identities,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"PASS: SVN merge role semantics ({len(tests)} tests)")


if __name__ == "__main__":
    main()
