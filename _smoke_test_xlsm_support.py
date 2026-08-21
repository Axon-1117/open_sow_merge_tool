"""Actual XLSM load/save gate with macro-package preservation."""

from __future__ import annotations

import json
import os
import re
import stat
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

import sow_merge_tool as mod
import _smoke_test_column_native_save_replay as column


_VBA_PART = "xl/vbaProject.bin"
_CONTENT_TYPES = "[Content_Types].xml"
_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_MACRO_WORKBOOK_CONTENT_TYPE = (
    b"application/vnd.ms-excel.sheet.macroEnabled.main+xml"
)
_VBA_CONTENT_TYPE = b"application/vnd.ms-office.vbaProject"
_VBA_REL_TYPE = b"http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def _checkpoint(deadline: float, stage: str) -> None:
    assert time.monotonic() < deadline, ("XLSM support deadline", stage)


def _make_macro_enabled_package(path: str) -> bytes:
    """Complete the reusable fixture's VBA sentinel into an XLSM package."""
    with zipfile.ZipFile(path, "r") as archive:
        payloads = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    assert _VBA_PART in payloads, path
    content_types = payloads[_CONTENT_TYPES]
    content_types, count = re.subn(
        rb'(<Override PartName="/xl/workbook\.xml" ContentType=")[^"]+("/>)',
        lambda match: match.group(1) + _MACRO_WORKBOOK_CONTENT_TYPE + match.group(2),
        content_types,
        count=1,
    )
    assert count == 1, path
    if b'/xl/vbaProject.bin' not in content_types:
        content_types = content_types.replace(
            b"</Types>",
            b'<Override PartName="/xl/vbaProject.bin" ContentType="'
            + _VBA_CONTENT_TYPE
            + b'"/></Types>',
            1,
        )
    relations = payloads[_WORKBOOK_RELS]
    if _VBA_REL_TYPE not in relations:
        relations = relations.replace(
            b"</Relationships>",
            b'<Relationship Id="rIdCodexVba" Type="'
            + _VBA_REL_TYPE
            + b'" Target="vbaProject.bin"/></Relationships>',
            1,
        )
    payloads[_CONTENT_TYPES] = content_types
    payloads[_WORKBOOK_RELS] = relations
    rewritten = path + ".macro-rewrite"
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(rewritten, "w") as target:
        for info in source.infolist():
            target.writestr(info, payloads[info.filename])
    os.replace(rewritten, path)
    return payloads[_VBA_PART]


def _assert_macro_package(
    path: str,
    expected_vba: bytes,
    *,
    expected_vba_form: str,
) -> None:
    """Validate one exact, form-specific VBA package declaration."""
    assert expected_vba_form in {"override", "default"}, expected_vba_form
    valid, reason = mod._validate_xlsx_package(path)
    assert valid, reason
    with zipfile.ZipFile(path, "r") as archive:
        assert archive.read(_VBA_PART) == expected_vba
        content_types = ET.fromstring(archive.read(_CONTENT_TYPES))
        relationships = ET.fromstring(archive.read(_WORKBOOK_RELS))

    def _local_name(node) -> str:
        return str(node.tag).rsplit("}", 1)[-1]

    macro_workbook_content_type = _MACRO_WORKBOOK_CONTENT_TYPE.decode("ascii")
    vba_content_type = _VBA_CONTENT_TYPE.decode("ascii")
    vba_relation_type = _VBA_REL_TYPE.decode("ascii")
    overrides = [node for node in content_types if _local_name(node) == "Override"]
    defaults = [node for node in content_types if _local_name(node) == "Default"]
    workbook_overrides = [
        node for node in overrides
        if node.get("PartName") == "/xl/workbook.xml"
    ]
    assert len(workbook_overrides) == 1, (path, workbook_overrides)
    assert workbook_overrides[0].get("ContentType") == macro_workbook_content_type
    macro_workbook_declarations = [
        node for node in (*overrides, *defaults)
        if node.get("ContentType") == macro_workbook_content_type
    ]
    assert len(macro_workbook_declarations) == 1
    assert macro_workbook_declarations[0] is workbook_overrides[0]

    vba_overrides = [
        node for node in overrides
        if node.get("PartName") == "/xl/vbaProject.bin"
    ]
    vba_defaults = [
        node for node in defaults
        if node.get("Extension", "").lower() == "bin"
    ]
    vba_declarations = [*vba_overrides, *vba_defaults]
    assert len(vba_declarations) == 1, (path, vba_declarations)
    vba_declaration = vba_declarations[0]
    assert vba_declaration.get("ContentType") == vba_content_type
    typed_vba_declarations = [
        node for node in (*overrides, *defaults)
        if node.get("ContentType") == vba_content_type
    ]
    assert len(typed_vba_declarations) == 1, (path, typed_vba_declarations)
    assert typed_vba_declarations[0] is vba_declaration
    if expected_vba_form == "override":
        assert vba_declaration in vba_overrides
    else:
        assert vba_declaration in vba_defaults

    relationship_nodes = [
        node for node in relationships
        if _local_name(node) == "Relationship"
    ]
    vba_type_relations = [
        node for node in relationship_nodes
        if node.get("Type") == vba_relation_type
    ]
    vba_target_relations = [
        node for node in relationship_nodes
        if node.get("Target") == "vbaProject.bin"
    ]
    assert len(vba_type_relations) == 1, (path, vba_type_relations)
    assert len(vba_target_relations) == 1, (path, vba_target_relations)
    assert vba_type_relations[0] is vba_target_relations[0]
    assert vba_type_relations[0].get("TargetMode", "Internal") == "Internal"


def main() -> None:
    case_deadline = time.monotonic() + 90.0
    fixture = None
    app = None
    primary = None
    owned_effective_paths = frozenset()
    original_loader = mod._openpyxl_load_workbook
    original_prompt_scheduler = mod.SowMergeApp._schedule_formula_cache_prompt
    original_settings_path = mod._SETTINGS_PATH
    original_settings_exists = os.path.lexists(original_settings_path)
    if original_settings_exists:
        with open(original_settings_path, "rb") as stream:
            original_settings_bytes = stream.read()
    else:
        original_settings_bytes = None
    calls = []

    def _spy_loader(filename, *args, **kwargs):
        calls.append((os.path.abspath(str(filename)), dict(kwargs)))
        return original_loader(filename, *args, **kwargs)

    try:
        _checkpoint(case_deadline, "before-fixture")
        fixture = column._fixture_set(".xlsm")
        assert fixture.expected is not None
        expected_original_sha = fixture.input_hashes[fixture.expected]
        assert column._sha256(fixture.expected) == expected_original_sha
        vba_by_path = {
            path: _make_macro_enabled_package(path)
            for path in (fixture.mine, fixture.theirs)
        }
        expected_vba = vba_by_path[fixture.mine]
        assert all(vba == expected_vba for vba in vba_by_path.values())
        assert column._sha256(fixture.expected) == expected_original_sha
        fixture.input_hashes.update({
            path: column._sha256(path)
            for path in (fixture.mine, fixture.theirs)
        })
        assert fixture.input_hashes[fixture.expected] == expected_original_sha
        _assert_macro_package(
            fixture.mine,
            expected_vba,
            expected_vba_form="override",
        )
        _assert_macro_package(
            fixture.theirs,
            expected_vba,
            expected_vba_form="override",
        )
        _checkpoint(case_deadline, "macro-fixture-ready")

        settings_path = os.path.join(fixture.root, "settings.json")
        with open(settings_path, "w", encoding="utf-8") as stream:
            json.dump({"only_diff": 0}, stream)
        mod._SETTINGS_PATH = settings_path
        mod._openpyxl_load_workbook = _spy_loader
        mod.SowMergeApp._schedule_formula_cache_prompt = lambda _self: None
        app = mod.SowMergeApp(fixture.mine, fixture.theirs)
        _checkpoint(case_deadline, "app-created")
        raw_a = os.path.abspath(fixture.mine)
        raw_b = os.path.abspath(fixture.theirs)
        effective_a = os.path.abspath(app.file_a)
        effective_b = os.path.abspath(app.file_b)
        assert raw_a != raw_b
        assert effective_a != effective_b
        assert effective_a != raw_a
        assert effective_b != raw_b
        owned_effective_paths = frozenset(
            os.path.normcase(path) for path in (effective_a, effective_b)
        )
        assert set(app._owned_startup_temp_paths) == owned_effective_paths
        stable_dir = os.path.normcase(os.path.abspath(tempfile.gettempdir()))
        stable_prefix = f"{mod.APP_NAME}_stable_{os.getpid()}_"
        for raw_path, effective_path in ((raw_a, effective_a), (raw_b, effective_b)):
            info = os.lstat(effective_path)
            assert stat.S_ISREG(info.st_mode)
            assert not os.path.islink(effective_path)
            assert not (int(getattr(info, "st_file_attributes", 0)) & 0x400)
            assert os.path.normcase(os.path.dirname(effective_path)) == stable_dir
            assert os.path.basename(effective_path).startswith(stable_prefix)
            assert effective_path.lower().endswith(".xlsm")
            assert column._sha256(effective_path) == fixture.input_hashes[raw_path]
            _assert_macro_package(
                effective_path,
                expected_vba,
                expected_vba_form="override",
            )
        while time.monotonic() < case_deadline:
            if callable(getattr(app, "_edit_preload_target", None)):
                break
            time.sleep(0.01)
        assert callable(getattr(app, "_edit_preload_target", None)), (
            "XLSM edit preload target unavailable",
            {
                "target": repr(getattr(app, "_edit_preload_target", None)),
                "ready": app._edit_workbooks_ready(),
                "loading_started": bool(getattr(app, "_edit_loading_started", False)),
                "owner": repr(getattr(app, "_edit_preload_thread", None)),
                "owner_alive": bool(
                    getattr(app, "_edit_preload_thread", None)
                    and app._edit_preload_thread.is_alive()
                ),
                "active": app._edit_preload_active_event.is_set(),
                "requests": tuple(getattr(app, "_edit_load_requests", ())),
                "loader_calls": tuple(calls),
                "edit_handles": tuple(
                    (name, handle is not None, getattr(handle, "read_only", None))
                    for name, handle in (
                        ("a", app._wb_a_edit),
                        ("b", app._wb_b_edit),
                    )
                ),
            },
        )
        assert not app._edit_workbooks_ready()
        requests_before = tuple(app._edit_load_requests)
        calls_before = len(calls)
        app._request_edit_preload(
            reason="test:xlsm-preservation",
            caller="_smoke_test_xlsm_support",
        )
        requests_after_request = tuple(app._edit_load_requests)
        assert len(requests_after_request) == len(requests_before) + 1
        request = requests_after_request[-1]
        assert request["reason"] == "test:xlsm-preservation"
        assert request["caller"] == "_smoke_test_xlsm_support"
        assert request["ready"] is False
        assert app._edit_loading_started is True
        edit_owner = app._edit_preload_thread
        assert edit_owner is not None

        edit_ready = False
        while time.monotonic() < case_deadline:
            assert app._edit_preload_thread is edit_owner
            if app._edit_workbooks_ready() and app._edit_loaded_event.is_set():
                edit_ready = True
                break
            time.sleep(0.01)
        assert edit_ready, (
            "XLSM edit preload timeout",
            {
                "ready": app._edit_workbooks_ready(),
                "event": app._edit_loaded_event.is_set(),
                "loading_started": bool(getattr(app, "_edit_loading_started", False)),
                "target": repr(getattr(app, "_edit_preload_target", None)),
                "owner": repr(getattr(app, "_edit_preload_thread", None)),
                "owner_alive": bool(edit_owner.is_alive()),
                "active": app._edit_preload_active_event.is_set(),
                "requests": tuple(getattr(app, "_edit_load_requests", ())),
                "loader_calls": tuple(calls[calls_before:]),
                "edit_handles": tuple(
                    (name, handle is not None, getattr(handle, "read_only", None))
                    for name, handle in (
                        ("a", app._wb_a_edit),
                        ("b", app._wb_b_edit),
                    )
                ),
            },
        )
        preload_calls = tuple(calls[calls_before:])
        assert app._edit_preload_thread is edit_owner
        assert preload_calls
        assert all(
            path.lower().endswith(".xlsm") and kwargs.get("keep_vba") is True
            for path, kwargs in preload_calls
        ), preload_calls
        editable_counts = Counter(
            os.path.normcase(path)
            for path, kwargs in preload_calls
            if kwargs.get("data_only") is False
        )
        assert editable_counts == Counter({
            os.path.normcase(effective_a): 1,
            os.path.normcase(effective_b): 1,
        })
        assert os.path.normcase(raw_a) not in editable_counts
        assert os.path.normcase(raw_b) not in editable_counts
        request_count_after_ready = len(app._edit_load_requests)
        app._ensure_edit_loaded()
        assert len(app._edit_load_requests) == request_count_after_ready
        assert app._edit_preload_thread is edit_owner
        output = os.path.join(fixture.root, "saved.xlsm")
        app._atomic_save(app._wb_a_edit, output)
        assert os.path.isfile(output), output
        _assert_macro_package(output, expected_vba, expected_vba_form="default")
        reopened = mod.load_workbook(output, data_only=False, keep_vba=True)
        try:
            assert getattr(reopened, "vba_archive", None) is not None
            assert reopened["S1"]["A1"].value == "A"
        finally:
            column._close_workbook(reopened)
        macro_calls = [
            (path, kwargs)
            for path, kwargs in calls
            if path.lower().endswith(".xlsm")
        ]
        assert macro_calls, calls
        assert all(kwargs.get("keep_vba") is True for _path, kwargs in macro_calls), macro_calls
        assert any(path == effective_a for path, _kwargs in macro_calls)
        assert any(path == effective_b for path, _kwargs in macro_calls)
        assert any(path == os.path.abspath(output) for path, _kwargs in macro_calls)
        _checkpoint(case_deadline, "save-reopen-verified")
    except BaseException as exc:
        primary = exc
        raise
    finally:
        cleanup_errors = []
        if app is not None:
            try:
                for candidate in (app, *getattr(app, "sheet_views", {}).values()):
                    if candidate is None:
                        continue
                    after_id = getattr(candidate, "_settings_save_id", None)
                    if after_id is not None:
                        app.root.after_cancel(after_id)
                        candidate._settings_save_id = None
                app._shutdown_root()
                assert not app._owned_startup_temp_paths
                evidence_by_path = {
                    os.path.normcase(os.path.abspath(item["path"])): item
                    for item in app._owned_startup_temp_cleanup_evidence
                }
                assert set(evidence_by_path) == owned_effective_paths
                for effective_path in owned_effective_paths:
                    evidence = evidence_by_path[effective_path]
                    assert evidence.get("removed") is True
                    assert evidence.get("exists_after") is False
                    assert not evidence.get("error")
            except BaseException as exc:
                cleanup_errors.append(f"shutdown: {exc!r}")

        def restore():
            mod._openpyxl_load_workbook = original_loader
            mod.SowMergeApp._schedule_formula_cache_prompt = original_prompt_scheduler
            mod._SETTINGS_PATH = original_settings_path
            if original_settings_exists:
                with open(original_settings_path, "rb") as stream:
                    assert stream.read() == original_settings_bytes
            else:
                assert not os.path.lexists(original_settings_path)

        if fixture is not None:
            try:
                column._cleanup_owned_fixture(fixture, primary, restore=restore)
            except BaseException as exc:
                cleanup_errors.append(f"owned fixture cleanup: {exc!r}")
        else:
            try:
                restore()
            except BaseException as exc:
                cleanup_errors.append(f"loader/settings restore: {exc!r}")
        if cleanup_errors:
            message = "XLSM support cleanup failed: " + "; ".join(cleanup_errors)
            if primary is not None:
                primary.add_note(message)
            else:
                raise AssertionError(message)
    print("SMOKE_XLSM_SUPPORT_OK")


if __name__ == "__main__":
    main()
