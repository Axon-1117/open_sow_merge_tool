"""OpenSpec 4.1 regressions for SVN launch roles and production dispatch.

These tests deliberately exercise both the pure launch-classification API and
``main()``.  The latter is important: keeping the raw paths in a data model is
not sufficient if the production launch path still replaces source Base with
the target working copy's pristine file before scanning or opening the UI.
"""

from __future__ import annotations

import hashlib
import inspect
import os
import sys
import tempfile
import time
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


_CASE_DEADLINE: float | None = None
_ACTIVE_OWNED_CASE = None


class _OwnedCase:
    def __init__(self, name: str):
        self._temporary = tempfile.TemporaryDirectory(prefix=f"sow_svn_roles_{name}_")
        self.root = self._temporary.name
        self._next_dir = 0
        self.input_hashes: dict[str, str] = {}

    def make_temp_dir(self, prefix: str) -> str:
        self._next_dir += 1
        path = os.path.join(self.root, f"{prefix}{self._next_dir}")
        os.makedirs(path, exist_ok=False)
        return path

    def record_input(self, path: str) -> None:
        absolute = os.path.abspath(path)
        self.input_hashes[absolute] = _sha256(absolute)

    def cleanup(self) -> None:
        self._temporary.cleanup()


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _checkpoint(label: str) -> None:
    deadline = _CASE_DEADLINE
    if deadline is not None and time.monotonic() > deadline:
        raise TimeoutError(f"SVN role case exceeded 90 seconds at {label}")


def _local_only_author_metadata(context):
    for identity in getattr(context, "identities", {}).values():
        if identity is None:
            continue
        identity.author = "未知"
        identity.author_status = "unavailable"
        identity.author_source = "test-local-no-network"
        identity.availability_reason = "test local metadata only"
    return getattr(context, "identities", {})


def _consume_fake_startup_ledger(kwargs: dict) -> tuple[dict, ...]:
    ledger = kwargs.get("startup_owned_paths")
    assert isinstance(ledger, set), kwargs
    evidence_sink: list[dict] = []
    evidence = smt._consume_owned_startup_temp_paths(ledger, evidence_sink)
    assert tuple(evidence_sink) == evidence, (evidence_sink, evidence)
    assert ledger == set(), ledger
    assert all(
        fact["removed"] and not fact["exists_after"] and not fact["error"]
        for fact in evidence
    ), evidence
    return evidence


def _make_book(path: str, marker: str, *, document_title: str | None = None) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "marker"
    worksheet["A2"] = marker
    if document_title is not None:
        # The branch Source After fixture deliberately carries one unsupported
        # document-property delta.  Cross-branch startup must fail closed
        # rather than silently dropping it while preparing Target Working.
        workbook.properties.title = str(document_title)
    workbook.save(path)
    workbook.close()
    active = _ACTIVE_OWNED_CASE
    if active is not None:
        active.record_input(path)


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
    startup_artifacts = os.path.join(root, "startup-artifacts")
    os.makedirs(startup_artifacts, exist_ok=False)
    target = os.path.join(root, "Design.xlsx")
    pristine = os.path.join(root, "Design.target-pristine.xlsx")
    merged = target
    if case_name == "update":
        base = os.path.join(root, "Design.xlsx.r301")
        theirs = os.path.join(root, "Design.xlsx.r305")
        expected_scenario = "UPDATE_CONFLICT"
        expects_scanner = True
    else:
        base = os.path.join(root, "Design.xlsx.merge-left.r401")
        theirs = os.path.join(root, "Design.xlsx.merge-right.r409")
        expected_scenario = "CROSS_BRANCH_MERGE"
        expects_scanner = False
    for path, marker in (
        (base, f"{case_name}-raw-base"),
        (target, f"{case_name}-mine"),
        (theirs, f"{case_name}-theirs"),
        (pristine, f"{case_name}-target-pristine"),
    ):
        _make_book(
            path,
            marker,
            document_title=(
                "branch-source-after-document-title"
                if case_name == "branch" and path == theirs
                else None
            ),
        )

    captured = SimpleNamespace(
        app_args=None,
        app_kwargs=None,
        cleanup_evidence=(),
        app_ledger=None,
        ensure_calls=[],
        scan_args=None,
        logs=[],
    )

    class _FakeApp:
        def __init__(self, *args, **kwargs):
            captured.app_args = args
            captured.app_kwargs = kwargs
            captured.app_ledger = kwargs.get("startup_owned_paths")

        def show_nonblocking_notice(self, *_args, **_kwargs):
            return None

        def run(self):
            captured.cleanup_evidence = _consume_fake_startup_ledger(
                captured.app_kwargs
            )
            return None

    def _scan(scan_base, scan_mine, scan_theirs, **_kwargs):
        captured.scan_args = (
            os.path.abspath(scan_base),
            os.path.abspath(scan_mine),
            os.path.abspath(scan_theirs),
        )
        return [], {}

    def _startup(_title, _message, worker):
        return worker(lambda *_args, **_kwargs: None)

    production_ensure_xlsx_copy = smt._ensure_xlsx_copy

    def _delegating_ensure_xlsx_copy(path, *, owned_paths=None, **kwargs):
        raw_path = os.path.normcase(os.path.abspath(os.fspath(path)))
        effective_path = production_ensure_xlsx_copy(
            path,
            owned_paths=owned_paths,
            **kwargs,
        )
        captured.ensure_calls.append(
            {
                "raw": raw_path,
                "effective": os.path.normcase(os.path.abspath(os.fspath(effective_path))),
                "ledger": owned_paths,
            }
        )
        return effective_path

    raw_base = os.path.normcase(os.path.abspath(base))
    raw_mine = os.path.normcase(os.path.abspath(target))
    raw_theirs = os.path.normcase(os.path.abspath(theirs))
    raw_pristine = os.path.normcase(os.path.abspath(pristine))

    def _is_startup_artifact(path) -> bool:
        try:
            return os.path.commonpath(
                (os.path.abspath(startup_artifacts), os.path.abspath(os.fspath(path)))
            ) == os.path.abspath(startup_artifacts)
        except (TypeError, ValueError):
            return False

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
        stack.enter_context(
            patch.object(smt, "resolve_cross_branch_source_metadata", lambda context: context)
        )
        stack.enter_context(
            patch.object(smt, "resolve_svn_author_metadata", _local_only_author_metadata)
        )
        stack.enter_context(
            patch.object(smt, "_ensure_xlsx_copy", _delegating_ensure_xlsx_copy)
        )
        stack.enter_context(
            patch.object(smt.tempfile, "gettempdir", lambda: startup_artifacts)
        )
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
                lambda _path, **_kwargs: pristine,
            )
        )
        if hasattr(smt, "_copy_wc_pristine_local"):
            stack.enter_context(
                patch.object(
                    smt,
                    "_copy_wc_pristine_local",
                    lambda _path, **_kwargs: pristine,
                )
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
    assert captured.app_ledger is captured.app_kwargs["startup_owned_paths"]
    assert captured.app_kwargs["startup_owned_paths"] == set()
    assert captured.ensure_calls, "production launch did not prepare effective inputs"
    assert all(call["ledger"] is captured.app_ledger for call in captured.ensure_calls), (
        "startup input preparation escaped the app-owned ledger",
        captured.ensure_calls,
    )
    assert all(
        call["raw"] == call["effective"] or _is_startup_artifact(call["effective"])
        for call in captured.ensure_calls
    ), captured.ensure_calls
    effective_by_raw = {
        raw: next(call["effective"] for call in captured.ensure_calls if call["raw"] == raw)
        for raw in (raw_base, raw_mine, raw_theirs)
    }
    assert effective_by_raw[raw_base] != raw_base
    assert effective_by_raw[raw_theirs] != raw_theirs
    assert effective_by_raw[raw_base].endswith(".xlsx")
    assert effective_by_raw[raw_theirs].endswith(".xlsx")
    assert _is_startup_artifact(effective_by_raw[raw_base])
    assert _is_startup_artifact(effective_by_raw[raw_theirs])
    if expects_scanner:
        assert captured.scan_args is not None, "update startup did not call the conflict scanner"
        normalized_scan_args = tuple(
            os.path.normcase(os.path.abspath(path))
            for path in captured.scan_args
        )
        assert normalized_scan_args == (
            effective_by_raw[raw_base],
            effective_by_raw[raw_mine],
            effective_by_raw[raw_theirs],
        ), (
            "production scanner must receive prepared Base/Mine/Theirs",
            normalized_scan_args,
        )
    else:
        assert captured.scan_args is None, captured.scan_args
        declined = [
            message
            for message in captured.logs
            if message.startswith("CROSS_BRANCH_SOURCE_DELTA_DECLINED ")
        ]
        assert len(declined) == 1, captured.logs
        assert (
            "reason=Unsupported Source OOXML representation change: docProps/core.xml"
            in declined[0]
        ), declined[0]
        final_candidate = os.path.normcase(os.path.abspath(captured.app_args[0]))
        evidence_by_path = {
            os.path.normcase(os.path.abspath(candidate["path"])): candidate
            for candidate in captured.cleanup_evidence
        }
        assert final_candidate in evidence_by_path, (
            "final UI Mine candidate was not transferred through the owned ledger",
            final_candidate,
            captured.cleanup_evidence,
        )
        assert any(
            os.path.basename(candidate["path"]).startswith(
                f"{smt.APP_NAME}_startup_candidate_mine_"
            )
            for candidate in captured.cleanup_evidence
        ), captured.cleanup_evidence
        assert not any(
            os.path.basename(candidate["path"]).startswith(
                f"{smt.APP_NAME}_cross_branch_candidate_"
            )
            for candidate in captured.cleanup_evidence
        ), captured.cleanup_evidence

    app_kwargs = captured.app_kwargs
    assert os.path.normcase(os.path.abspath(app_kwargs.get("base_path"))) == effective_by_raw[raw_base], (
        "production launch did not use the prepared Source Base",
        app_kwargs,
    )
    assert os.path.normcase(os.path.abspath(captured.app_args[1])) == effective_by_raw[raw_theirs]
    assert os.path.normcase(os.path.abspath(app_kwargs.get("raw_base"))) == raw_base, (
        "raw Base diagnostic identity was rewritten",
        app_kwargs,
    )
    assert os.path.normcase(os.path.abspath(app_kwargs.get("raw_mine"))) == raw_mine
    assert os.path.normcase(os.path.abspath(app_kwargs.get("raw_theirs"))) == raw_theirs
    context = app_kwargs.get("launch_context")
    assert context is not None, "production UI did not receive MergeLaunchContext"
    assert _scenario_name(_context_value(context, "scenario", "merge_scenario")) == expected_scenario
    assert os.path.normcase(_context_identity_path(context, "base")) == raw_base
    assert os.path.normcase(_context_identity_path(context, "mine")) == raw_mine
    assert os.path.normcase(_context_identity_path(context, "theirs")) == raw_theirs
    assert os.path.normcase(_context_identity_path(context, "target_pristine")) == raw_pristine
    assert os.path.normcase(_context_identity_path(context, "base")) != raw_pristine
    assert os.path.normcase(os.path.abspath(context.identity_for("base").effective_path)) == effective_by_raw[raw_base]
    assert os.path.normcase(os.path.abspath(context.identity_for("mine").effective_path)) == effective_by_raw[raw_mine]
    assert os.path.normcase(os.path.abspath(context.identity_for("theirs").effective_path)) == effective_by_raw[raw_theirs]
    assert captured.cleanup_evidence, captured.cleanup_evidence
    assert all(_is_startup_artifact(candidate["path"]) for candidate in captured.cleanup_evidence), (
        "source-created startup artifact escaped this owned case",
        captured.cleanup_evidence,
    )
    assert all(
        candidate["removed"]
        and not candidate["exists_after"]
        and not candidate["error"]
        for candidate in captured.cleanup_evidence
    ), captured.cleanup_evidence
    assert not any(
        os.path.normcase(os.path.abspath(candidate["path"]))
        in {raw_base, raw_mine, raw_theirs, raw_pristine}
        for candidate in captured.cleanup_evidence
    ), captured.cleanup_evidence
    assert not any(
        os.path.basename(candidate["path"]).startswith(
            f"{smt.APP_NAME}_cross_branch_candidate_"
        )
        for candidate in captured.cleanup_evidence
    ), captured.cleanup_evidence
    evidence_names = tuple(os.path.basename(candidate["path"]) for candidate in captured.cleanup_evidence)
    assert any(name.startswith(f"{smt.APP_NAME}_svn_") for name in evidence_names), evidence_names
    assert any(name.startswith(f"{smt.APP_NAME}_stable_") for name in evidence_names), evidence_names
    if case_name == "branch":
        assert any(
            name.startswith(f"{smt.APP_NAME}_startup_candidate_mine_")
            for name in evidence_names
        ), evidence_names
    assert not os.listdir(startup_artifacts), os.listdir(startup_artifacts)
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
    captured = SimpleNamespace(app_args=None, app_kwargs=None, cleanup_evidence=())

    class _FakeApp:
        def __init__(self, *args, **kwargs):
            captured.app_args = args
            captured.app_kwargs = kwargs

        def run(self):
            captured.cleanup_evidence = _consume_fake_startup_ledger(
                captured.app_kwargs
            )
            return None

    argv = [
        "sow_merge_tool.py",
        "/base",
        base,
        "/mine",
        mine,
    ]
    with (
        patch.object(sys, "argv", argv),
        patch.object(smt, "SowMergeApp", _FakeApp),
        patch.object(smt, "resolve_svn_author_metadata", lambda context: context.identities),
    ):
        smt.main()

    assert captured.app_args is not None
    ui_base, ui_mine = captured.app_args[:2]
    assert os.path.abspath(ui_base) == os.path.abspath(base), (
        "main normalized the raw sidecar before the app owned its ledger",
        ui_base,
    )
    assert os.path.abspath(ui_mine) == os.path.abspath(mine)
    assert os.path.abspath(captured.app_kwargs["raw_base"]) == os.path.abspath(base)
    assert os.path.abspath(captured.app_kwargs["raw_mine"]) == os.path.abspath(mine)


def test_two_way_production_path_preserves_author_metadata() -> None:
    root = make_temp_dir("sow_two_way_author_context_")
    base = os.path.join(root, "Design.xlsx.r301")
    mine = os.path.join(root, "Design.xlsx")
    _make_book(base, "base")
    _make_book(mine, "mine")
    captured = SimpleNamespace(app_kwargs=None, cleanup_evidence=())

    class _FakeApp:
        def __init__(self, *_args, **kwargs):
            captured.app_kwargs = kwargs

        def run(self):
            captured.cleanup_evidence = _consume_fake_startup_ledger(
                captured.app_kwargs
            )
            return None

    def _resolve(context):
        context.identity_for("base").author = "alice"
        context.identity_for("base").author_status = "resolved"
        context.identity_for("mine").author = "bob"
        context.identity_for("mine").author_status = "resolved"
        return context.identities

    argv = ["sow_merge_tool.py", "/base", base, "/mine", mine]
    with (
        patch.object(sys, "argv", argv),
        patch.object(smt, "SowMergeApp", _FakeApp),
        patch.object(smt, "resolve_svn_author_metadata", _resolve),
    ):
        smt.main()

    context = captured.app_kwargs["launch_context"]
    assert _scenario_name(context.scenario) == "TWO_WAY"
    assert context.identity_for("base").author == "alice"
    assert context.identity_for("mine").author == "bob"
    assert "Author = alice" in smt.format_compact_version_identity(context.identity_for("base"))
    assert "Author = bob" in smt.format_compact_version_identity(context.identity_for("mine"))
    assert os.path.abspath(captured.app_kwargs["raw_base"]) == os.path.abspath(base)
    assert os.path.abspath(captured.app_kwargs["raw_mine"]) == os.path.abspath(mine)


def test_two_way_local_author_fallback_is_visible() -> None:
    root = make_temp_dir("sow_two_way_author_fallback_")
    base = os.path.join(root, "old.xlsx")
    mine = os.path.join(root, "new.xlsx")
    _make_book(base, "base")
    _make_book(mine, "mine")
    context = smt.build_merge_launch_context(base, mine, None)
    with patch.object(smt, "_find_svn_wc_root_for_path", lambda _path: None):
        identities = smt.resolve_svn_author_metadata(context)
    assert identities["base"].author == "未知"
    assert identities["mine"].author == "未知"
    assert identities["base"].availability_reason
    assert identities["mine"].availability_reason
    assert "Author = 未知" in smt.format_compact_version_identity(identities["base"])
    assert "Author = 未知" in smt.format_compact_version_identity(identities["mine"])


def test_auto_detected_conflict_preserves_raw_sidecar_identities() -> None:
    root = make_temp_dir("sow_auto_detected_raw_roles_")
    base = os.path.join(root, "Design.xlsx.merge-left.r401")
    mine = os.path.join(root, "Design.xlsx")
    theirs = os.path.join(root, "Design.xlsx.merge-right.r409")
    for path, marker in ((base, "base"), (mine, "mine"), (theirs, "theirs")):
        _make_book(path, marker)
    captured = SimpleNamespace(app_kwargs=None, cleanup_evidence=())

    class _FakeApp:
        def __init__(self, *_args, **kwargs):
            captured.app_kwargs = kwargs

        def show_startup_outcome_dialog_deferred(self):
            return None

        def run(self):
            captured.cleanup_evidence = _consume_fake_startup_ledger(
                captured.app_kwargs
            )
            return None

    def _analysis(context, **_kwargs):
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


def _run_owned_case(test) -> None:
    global _ACTIVE_OWNED_CASE, _CASE_DEADLINE, make_temp_dir
    previous_case = _ACTIVE_OWNED_CASE
    previous_deadline = _CASE_DEADLINE
    previous_factory = make_temp_dir
    owned = _OwnedCase(test.__name__)
    primary = None
    cleanup_errors: list[str] = []
    _ACTIVE_OWNED_CASE = owned
    _CASE_DEADLINE = time.monotonic() + 90.0
    make_temp_dir = owned.make_temp_dir
    try:
        with ExitStack() as stack:
            stack.enter_context(
                patch.object(smt, "_SETTINGS_PATH", os.path.join(owned.root, "settings.json"))
            )
            stack.enter_context(
                patch.object(smt, "resolve_cross_branch_source_metadata", lambda context: context)
            )
            if test is not test_two_way_local_author_fallback_is_visible:
                stack.enter_context(
                    patch.object(smt, "resolve_svn_author_metadata", _local_only_author_metadata)
                )
            _checkpoint("before-test")
            test()
            _checkpoint("after-test")
    except BaseException as exc:
        primary = exc
    finally:
        make_temp_dir = previous_factory
        _ACTIVE_OWNED_CASE = previous_case
        _CASE_DEADLINE = previous_deadline
        try:
            for path, expected_hash in owned.input_hashes.items():
                actual_hash = _sha256(path)
                if actual_hash != expected_hash:
                    cleanup_errors.append(
                        f"input hash changed: {path} {expected_hash} -> {actual_hash}"
                    )
        except Exception as exc:
            cleanup_errors.append(f"input hash verification failed: {type(exc).__name__}: {exc}")
        try:
            owned.cleanup()
            if os.path.lexists(owned.root):
                cleanup_errors.append(f"owned temporary root retained: {owned.root}")
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


def main() -> None:
    tests = (
        test_four_launch_scenarios_and_identities,
        test_update_production_path_preserves_raw_old_base,
        test_branch_production_path_preserves_merge_left,
        test_two_way_sidecar_is_normalized_without_rewriting_raw_identity,
        test_two_way_production_path_preserves_author_metadata,
        test_two_way_local_author_fallback_is_visible,
        test_auto_detected_conflict_preserves_raw_sidecar_identities,
    )
    for test in tests:
        _run_owned_case(test)
        print(f"PASS: {test.__name__}")
    print(f"PASS: SVN merge role semantics ({len(tests)} tests)")


if __name__ == "__main__":
    main()
